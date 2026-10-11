"""Cleaner screen: controls and rendering only; services own analysis and export."""

from collections.abc import Callable
from functools import partial
from pathlib import Path

from PySide6.QtCore import QSignalBlocker, Qt, QTimer, Slot
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTableView,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from excel_toolkit.application.cleaning import CleaningJob, CleaningService
from excel_toolkit.application.models import CsvOptions, Preview
from excel_toolkit.application.preview import PreviewService
from excel_toolkit.modules.cleaner import CleaningPlan
from excel_toolkit.ui.table_model import PreviewTableModel
from excel_toolkit.ui.worker import PreviewWorker, ServiceWorker, analyze_for_display, export_file

OPERATIONS = (
    ("trim_whitespace", "Trim Whitespace"),
    ("remove_empty_rows", "Remove Empty Rows"),
    ("remove_duplicates", "Remove Duplicates"),
    ("detect_missing_values", "Detect Missing Values"),
)
SUMMARY_FIELDS = (
    ("original_rows", "Original Rows"),
    ("final_rows", "Cleaned Rows"),
    ("removed_empty_rows", "Removed Empty Rows"),
    ("removed_duplicates", "Removed Duplicates"),
    ("changed_text_cells", "Trimmed Cells"),
    ("missing_values", "Missing Values"),
)


class MainWindow(QMainWindow):
    def __init__(
        self, service: PreviewService, cleaning_service: CleaningService | None = None
    ) -> None:
        super().__init__()
        self._service = service
        self._cleaning_service = cleaning_service
        self._worker: ServiceWorker | None = None
        self._operation: str | None = None
        self._path: Path | None = None
        self._job: CleaningJob | None = None
        self._revision = 0
        self._worker_revision = 0
        self._closing = False
        self.setWindowTitle("Excel Freelancer Toolkit")
        self.resize(1100, 900)
        self.setMinimumSize(760, 540)
        self._build_ui()
        self._set_busy(False)

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(14)
        self.navigation = QListWidget()
        self.navigation.addItems(["Cleaner", "Merger (скоро)", "Compare (скоро)"])
        for index in (1, 2):
            item = self.navigation.item(index)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEnabled)
        self.navigation.setCurrentRow(0)
        self.navigation.setFixedWidth(150)
        layout.addWidget(self.navigation)
        scroll = QScrollArea()
        self._scroll = scroll
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        page = QWidget()
        scroll.setWidget(page)
        layout.addWidget(scroll, 1)
        content = QVBoxLayout(page)
        content.setContentsMargins(0, 0, 4, 0)
        content.setSpacing(10)
        content.addWidget(QLabel("Cleaner — очистка таблиц"))
        hint = QLabel(
            "Предпросмотр: первые 100 строк, до 200 столбцов. Анализ обрабатывает весь файл."
        )
        hint.setWordWrap(True)
        content.addWidget(hint)
        content.addWidget(self._file_controls())
        content.addWidget(self._cleaning_controls())
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.hide()
        content.addWidget(self.progress)
        self.model = PreviewTableModel()
        self.cleaned_model = PreviewTableModel()
        self.table = self._table_view(self.model)
        self.cleaned_table = self._table_view(self.cleaned_model)
        self.preview_tabs = QTabWidget()
        self.preview_tabs.addTab(self.table, "Original Data")
        self.preview_tabs.addTab(self.cleaned_table, "Cleaned Data")
        self.preview_tabs.setTabEnabled(1, False)
        self.preview_tabs.currentChanged.connect(self._preview_changed)
        self.preview_tabs.setMinimumHeight(190)
        content.addWidget(self.preview_tabs, 1)
        self.preview_label = QLabel("Предпросмотр не загружен.")
        self.preview_label.setWordWrap(True)
        content.addWidget(self.preview_label)
        content.addWidget(self._summary_panel())
        self.status_label = QLabel("Выберите .xlsx или .csv для просмотра и очистки.")
        self.status_label.setTextFormat(Qt.TextFormat.PlainText)
        self.status_label.setWordWrap(True)
        content.addWidget(self.status_label)
        self.saved_label = QLabel()
        self.saved_label.setTextFormat(Qt.TextFormat.PlainText)
        self.saved_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.saved_label.setWordWrap(True)
        self.saved_label.hide()
        content.addWidget(self.saved_label)
        self._error_title = QLabel("Ошибки:")
        content.addWidget(self._error_title)
        self.errors = QPlainTextEdit()
        self.errors.setReadOnly(True)
        self.errors.setPlaceholderText("Ошибок нет.")
        self.errors.setMinimumHeight(60)
        self.errors.setMaximumHeight(100)
        content.addWidget(self.errors)
        self._clear_errors()

    def _file_controls(self) -> QGroupBox:
        group = QGroupBox("Файл и параметры загрузки")
        content = QVBoxLayout(group)
        actions = QHBoxLayout()
        self.open_button = QPushButton("Выбрать файл…")
        self.open_button.clicked.connect(self._choose_file)
        actions.addWidget(self.open_button)
        self.reload_button = QPushButton("Повторить загрузку")
        self.reload_button.clicked.connect(self._reload)
        actions.addWidget(self.reload_button)
        actions.addStretch()
        content.addLayout(actions)
        self.file_label = QLabel("Файл не выбран")
        self.file_label.setTextFormat(Qt.TextFormat.PlainText)
        self.file_label.setWordWrap(True)
        content.addWidget(self.file_label)
        options = QGridLayout()
        self.sheet_combo = QComboBox()
        self.sheet_combo.textActivated.connect(self._select_sheet)
        self.sheet_combo.currentTextChanged.connect(self._source_options_changed)
        self.encoding_combo = QComboBox()
        for label, encoding in (
            ("UTF-8 / UTF-8 BOM", "utf-8-sig"),
            ("Windows-1251", "cp1251"),
            ("Windows-1252", "cp1252"),
            ("UTF-16 с BOM", "utf-16"),
        ):
            self.encoding_combo.addItem(label, encoding)
        self.delimiter_combo = QComboBox()
        for label, delimiter in (
            ("Авто", None),
            ("Запятая", ","),
            ("Точка с запятой", ";"),
            ("Табуляция", "\t"),
            ("Вертикальная черта", "|"),
        ):
            self.delimiter_combo.addItem(label, delimiter)
        for column, (label, combo) in enumerate(
            (
                ("Лист XLSX:", self.sheet_combo),
                ("CSV — кодировка:", self.encoding_combo),
                ("CSV — разделитель:", self.delimiter_combo),
            )
        ):
            options.addWidget(QLabel(label), 0, column)
            combo.setMinimumWidth(120)
            combo.setMaximumWidth(250)
            combo.setSizeAdjustPolicy(
                QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
            )
            options.addWidget(combo, 1, column)
        self.encoding_combo.currentIndexChanged.connect(self._source_options_changed)
        self.delimiter_combo.currentIndexChanged.connect(self._source_options_changed)
        content.addLayout(options)
        hint = QLabel(
            "После изменения параметров загрузите предпросмотр заново или запустите анализ."
        )
        hint.setWordWrap(True)
        content.addWidget(hint)
        self.csv_hint = QLabel(
            "CSV не хранит типы: при открытии результата в Calc задайте для ID тип «Текст», "
            "чтобы сохранить ведущие нули."
        )
        self.csv_hint.setWordWrap(True)
        self.csv_hint.hide()
        content.addWidget(self.csv_hint)
        return group

    def _cleaning_controls(self) -> QGroupBox:
        group = QGroupBox("Операции очистки")
        content = QVBoxLayout(group)
        grid = QGridLayout()
        self.operation_checks: dict[str, QCheckBox] = {}
        for index, (name, title) in enumerate(OPERATIONS):
            checkbox = QCheckBox(title)
            checkbox.setChecked(True)
            checkbox.toggled.connect(self._plan_changed)
            self.operation_checks[name] = checkbox
            grid.addWidget(checkbox, index // 2, index % 2)
        content.addLayout(grid)
        hint = QLabel(
            "Порядок: Trim → Empty Rows → Duplicates → Missing Values. Первая строка — заголовок."
        )
        hint.setWordWrap(True)
        content.addWidget(hint)
        actions = QHBoxLayout()
        self.analyze_button = QPushButton("Analyze && Clean")
        self.analyze_button.clicked.connect(self._analyze)
        self.save_button = QPushButton("Save Cleaned File…")
        self.save_button.clicked.connect(self._choose_destination)
        actions.addWidget(self.analyze_button)
        actions.addWidget(self.save_button)
        actions.addStretch()
        content.addLayout(actions)
        return group

    def _summary_panel(self) -> QGroupBox:
        group = QGroupBox("Cleaning Summary")
        grid = QGridLayout(group)
        self.summary_values: dict[str, QLabel] = {}
        for index, (name, title) in enumerate(SUMMARY_FIELDS):
            row, column = index // 2, (index % 2) * 2
            grid.addWidget(QLabel(title), row, column)
            value = QLabel("—")
            value.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            self.summary_values[name] = value
            grid.addWidget(value, row, column + 1)
        self.report_details = QPlainTextEdit()
        self.report_details.setReadOnly(True)
        self.report_details.setPlaceholderText(
            "Выполните анализ, чтобы увидеть операции и предупреждения."
        )
        self.report_details.setMinimumHeight(60)
        self.report_details.setMaximumHeight(85)
        grid.addWidget(self.report_details, 3, 0, 1, 4)
        return group

    @staticmethod
    def _table_view(model: PreviewTableModel) -> QTableView:
        table = QTableView()
        table.setModel(model)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setWordWrap(False)
        table.horizontalHeader().setDefaultSectionSize(150)
        return table

    def _csv_options(self) -> CsvOptions:
        return CsvOptions(self.encoding_combo.currentData(), self.delimiter_combo.currentData())

    def cleaning_plan(self) -> CleaningPlan:
        return CleaningPlan(
            **{name: check.isChecked() for name, check in self.operation_checks.items()}
        )

    def _invalidate_result(self) -> None:
        self._revision += 1
        self._job = None
        self.cleaned_model.set_preview(None)
        self.preview_tabs.setCurrentIndex(0)
        self.preview_tabs.setTabEnabled(1, False)
        for label in self.summary_values.values():
            label.setText("—")
        self.report_details.clear()
        self.saved_label.clear()
        self.saved_label.hide()
        self._clear_errors()
        self._set_busy(self._worker is not None)

    @Slot()
    def _plan_changed(self) -> None:
        self._invalidate_result()
        if self._path is not None:
            self.status_label.setText("Операции изменены. Выполните анализ перед сохранением.")

    @Slot()
    def _source_options_changed(self) -> None:
        self._invalidate_result()
        self.model.set_preview(None)
        self._preview_changed()
        if self._path is not None:
            self.status_label.setText(
                "Параметры изменены. Обновите предпросмотр или выполните анализ."
            )

    @Slot()
    def _choose_file(self) -> None:
        if self._worker is not None or self._closing:
            return
        filename, _ = QFileDialog.getOpenFileName(
            self, "Выбрать таблицу", "", "Таблицы (*.xlsx *.csv);;Excel (*.xlsx);;CSV (*.csv)"
        )
        if filename:
            self.load_file(Path(filename))

    def load_file(self, path: Path, sheet: str | None = None) -> None:
        if self._worker is not None or self._closing:
            return
        self._path = path
        self._invalidate_result()
        self.model.set_preview(None)
        self._preview_changed()
        self.file_label.setText(f"Файл: {path.name}")
        self.file_label.setToolTip(str(path))
        self.csv_hint.setVisible(path.suffix.lower() == ".csv")
        with QSignalBlocker(self.sheet_combo):
            self.sheet_combo.clear()
        self.status_label.setText("Чтение файла…")
        worker = PreviewWorker(self._service, path, sheet, self._csv_options())
        self._start_worker(worker, "preview", self._loaded)

    def _start_worker(self, worker: ServiceWorker, operation: str, loaded: Callable) -> None:
        self._worker = worker
        self._operation = operation
        self._worker_revision = self._revision
        worker.loaded.connect(loaded, Qt.ConnectionType.QueuedConnection)
        worker.failed.connect(self._failed, Qt.ConnectionType.QueuedConnection)
        worker.finished.connect(self._finished, Qt.ConnectionType.QueuedConnection)
        self._set_busy(True)
        worker.start()

    @Slot()
    def _reload(self) -> None:
        if self._path is not None:
            self.load_file(self._path, self.sheet_combo.currentText() or None)

    @Slot(str)
    def _select_sheet(self, sheet: str) -> None:
        if self._path is not None:
            self.load_file(self._path, sheet)

    @Slot()
    def _analyze(self) -> None:
        if self._worker is not None or self._closing or self._path is None:
            return
        if self._cleaning_service is None:
            return
        self._invalidate_result()
        self.status_label.setText("Анализ и очистка полного файла…")
        task = partial(
            analyze_for_display,
            self._cleaning_service,
            self._path,
            self.cleaning_plan(),
            self.sheet_combo.currentText() or None,
            self._csv_options(),
        )
        self._start_worker(ServiceWorker(task, "анализ и очистка"), "analyze", self._analyzed)

    @Slot()
    def _choose_destination(self) -> None:
        if self._worker is not None or self._closing or self._job is None:
            return
        source = self._job.source.path
        suggested = source.with_name(f"{source.stem}_cleaned{source.suffix}")
        file_filter = "Excel (*.xlsx)" if source.suffix.lower() == ".xlsx" else "CSV (*.csv)"
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить очищенную таблицу в новый файл",
            str(suggested),
            file_filter,
            options=QFileDialog.Option.DontConfirmOverwrite,
        )
        if filename:
            destination = Path(filename)
            if not destination.suffix:
                destination = destination.with_suffix(source.suffix)
            self.save_cleaned_file(destination)

    def save_cleaned_file(self, destination: Path) -> None:
        if self._worker is not None or self._closing or self._job is None:
            return
        if self._cleaning_service is None:
            return
        self._clear_errors()
        self.saved_label.clear()
        self.saved_label.hide()
        self.status_label.setText("Сохранение очищенного файла…")
        destination = destination.absolute()
        task = partial(export_file, self._cleaning_service, self._job, destination)
        self._start_worker(ServiceWorker(task, "экспорт"), "export", self._exported)

    def _accept_result(self) -> bool:
        return not self._closing and self._worker_revision == self._revision

    def _show_original(self, preview: Preview) -> None:
        self.model.set_preview(preview)
        with QSignalBlocker(self.sheet_combo):
            self.sheet_combo.clear()
            self.sheet_combo.addItems(preview.sheet_names)
            if preview.selected_sheet is not None:
                self.sheet_combo.setCurrentText(preview.selected_sheet)

    @Slot(object)
    def _loaded(self, preview: Preview) -> None:
        if not self._accept_result():
            return
        self._show_original(preview)
        self._preview_changed()
        self.status_label.setText("Предпросмотр загружен. Выберите операции и запустите анализ.")

    @Slot(object)
    def _analyzed(self, payload: tuple[CleaningJob, Preview, Preview]) -> None:
        if not self._accept_result():
            return
        job, original, cleaned = payload
        self._job = job
        self._show_original(original)
        self.cleaned_model.set_preview(cleaned)
        report = job.result.report
        for name, label in self.summary_values.items():
            label.setText(str(getattr(report, name)))
        titles = dict(OPERATIONS)
        operations = ", ".join(titles.get(name, name) for name in report.performed_operations)
        details = f"Выполненные операции: {operations or 'не выбраны'}."
        if report.warnings:
            details += "\nПредупреждения:\n" + "\n".join(report.warnings)
        else:
            details += "\nПредупреждений нет."
        self.report_details.setPlainText(details)
        self.preview_tabs.setTabEnabled(1, True)
        self.preview_tabs.setCurrentIndex(1)
        self._preview_changed()
        self.status_label.setText("Анализ завершён. Результат можно сохранить в новый файл.")

    @Slot(object)
    def _exported(self, destination: Path) -> None:
        if not self._closing:
            self.saved_label.setText(f"Сохранено: {destination}")
            self.saved_label.show()
            self.status_label.setText("Очищенный файл сохранён.")
            QTimer.singleShot(0, self, self._reveal_feedback)

    def _clear_errors(self) -> None:
        self.errors.clear()
        self.errors.hide()
        self._error_title.hide()

    @Slot()
    def _reveal_feedback(self) -> None:
        if not self._closing:
            widget = self.errors if self.errors.isVisible() else self.saved_label
            self._scroll.ensureWidgetVisible(widget)

    @Slot()
    def _preview_changed(self) -> None:
        model = self.model if self.preview_tabs.currentIndex() == 0 else self.cleaned_model
        preview = model.preview
        if preview is None:
            self.preview_label.setText("Предпросмотр не загружен.")
            return
        message = f"Показано строк: {len(preview.rows)}, столбцов: {preview.column_count}."
        if not preview.rows:
            message += " Таблица пуста."
        if preview.more_rows:
            message += " Есть строки за пределами предпросмотра."
        if preview.more_columns:
            message += " Столбцы после 200-го скрыты."
        if preview.csv_options:
            message += (
                f" CSV: {preview.csv_options.encoding}, "
                f"разделитель {preview.csv_options.delimiter!r}."
            )
        self.preview_label.setText(message)

    @Slot(str)
    def _failed(self, message: str) -> None:
        if not self._accept_result():
            return
        self.errors.setPlainText(message)
        self._error_title.show()
        self.errors.show()
        QTimer.singleShot(0, self, self._reveal_feedback)
        descriptions = {
            "preview": "Предпросмотр не загружен. Проверьте файл и настройки.",
            "analyze": "Анализ не выполнен. Экспорт недоступен; причина указана ниже.",
            "export": "Файл не сохранён. Проверьте путь и выберите новое имя.",
        }
        self.status_label.setText(descriptions[self._operation])

    @Slot()
    def _finished(self) -> None:
        worker = self._worker
        self._worker = None
        self._operation = None
        if worker is not None:
            worker.deleteLater()
        self._set_busy(False)
        if self._closing:
            self.close()

    def _set_busy(self, busy: bool) -> None:
        ready = not busy and not self._closing
        self.progress.setVisible(busy)
        self.open_button.setEnabled(ready)
        self.reload_button.setEnabled(ready and self._path is not None)
        self.sheet_combo.setEnabled(ready and self.sheet_combo.count() > 1)
        self.encoding_combo.setEnabled(ready)
        self.delimiter_combo.setEnabled(ready)
        for check in self.operation_checks.values():
            check.setEnabled(ready)
        self.analyze_button.setEnabled(
            ready and self._path is not None and self._cleaning_service is not None
        )
        self.save_button.setEnabled(ready and self._job is not None)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._worker is not None:
            self._closing = True
            if self._operation == "preview":
                self._worker.requestInterruption()
            self.status_label.setText(
                "Ожидание безопасного завершения текущей операции перед закрытием…"
            )
            self._set_busy(True)
            event.ignore()
        else:
            event.accept()
