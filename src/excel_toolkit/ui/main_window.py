"""Cleaner preview screen. All filesystem work is delegated to the service."""

from pathlib import Path

from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from excel_toolkit.application.models import CsvOptions, Preview
from excel_toolkit.application.preview import PreviewService
from excel_toolkit.ui.table_model import PreviewTableModel
from excel_toolkit.ui.worker import PreviewWorker


class MainWindow(QMainWindow):
    def __init__(self, service: PreviewService) -> None:
        super().__init__()
        self._service = service
        self._worker: PreviewWorker | None = None
        self._path: Path | None = None
        self._closing = False
        self.setWindowTitle("Excel Freelancer Toolkit")
        self.resize(1100, 720)
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)
        self.navigation = QListWidget()
        self.navigation.addItems(["Cleaner", "Merger (скоро)", "Compare (скоро)"])
        for index in (1, 2):
            item = self.navigation.item(index)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEnabled)
        self.navigation.setCurrentRow(0)
        self.navigation.setFixedWidth(180)
        layout.addWidget(self.navigation)
        content = QVBoxLayout()
        layout.addLayout(content, 1)
        content.addWidget(QLabel("Cleaner — предпросмотр исходных данных"))
        content.addWidget(QLabel("Первые 100 строк, до 200 столбцов. Очистка пока недоступна."))
        actions = QHBoxLayout()
        self.open_button = QPushButton("Выбрать файл…")
        self.open_button.clicked.connect(self._choose_file)
        actions.addWidget(self.open_button)
        self.reload_button = QPushButton("Повторить загрузку")
        self.reload_button.setEnabled(False)
        self.reload_button.clicked.connect(self._reload)
        actions.addWidget(self.reload_button)
        actions.addStretch()
        content.addLayout(actions)
        self.file_label = QLabel("Файл не выбран")
        self.file_label.setTextFormat(Qt.TextFormat.PlainText)
        self.file_label.setWordWrap(True)
        content.addWidget(self.file_label)
        options = QHBoxLayout()
        options.addWidget(QLabel("Лист:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setEnabled(False)
        self.sheet_combo.textActivated.connect(self._select_sheet)
        options.addWidget(self.sheet_combo, 1)
        options.addWidget(QLabel("CSV — кодировка:"))
        self.encoding_combo = QComboBox()
        for label, encoding in (
            ("UTF-8 / UTF-8 BOM", "utf-8-sig"),
            ("Windows-1251", "cp1251"),
            ("Windows-1252", "cp1252"),
            ("UTF-16 с BOM", "utf-16"),
        ):
            self.encoding_combo.addItem(label, encoding)
        options.addWidget(self.encoding_combo)
        options.addWidget(QLabel("Разделитель:"))
        self.delimiter_combo = QComboBox()
        for label, delimiter in (
            ("Авто", None),
            ("Запятая", ","),
            ("Точка с запятой", ";"),
            ("Табуляция", "\t"),
            ("Вертикальная черта", "|"),
        ):
            self.delimiter_combo.addItem(label, delimiter)
        options.addWidget(self.delimiter_combo)
        content.addLayout(options)
        content.addWidget(QLabel("После изменения настроек CSV нажмите «Повторить загрузку»."))
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.hide()
        content.addWidget(self.progress)
        self.model = PreviewTableModel()
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.horizontalHeader().setDefaultSectionSize(150)
        content.addWidget(self.table, 1)
        self.status_label = QLabel("Выберите .xlsx или .csv для просмотра.")
        self.status_label.setWordWrap(True)
        content.addWidget(self.status_label)
        content.addWidget(QLabel("Ошибки:"))
        self.errors = QPlainTextEdit()
        self.errors.setReadOnly(True)
        self.errors.setMaximumHeight(90)
        content.addWidget(self.errors)

    @Slot()
    def _choose_file(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self, "Выбрать таблицу", "", "Таблицы (*.xlsx *.csv);;Excel (*.xlsx);;CSV (*.csv)"
        )
        if filename:
            self.load_file(Path(filename))

    def load_file(self, path: Path, sheet: str | None = None) -> None:
        if self._worker is not None or self._closing:
            return
        self._path = path
        self.errors.clear()
        self.model.set_preview(None)
        self.file_label.setText(f"Файл: {path.name}")
        self.file_label.setToolTip(str(path))
        self.sheet_combo.clear()
        self.status_label.setText("Чтение файла…")
        options = CsvOptions(self.encoding_combo.currentData(), self.delimiter_combo.currentData())
        worker = PreviewWorker(self._service, path, sheet, options)
        self._worker = worker
        worker.loaded.connect(self._loaded, Qt.ConnectionType.QueuedConnection)
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

    @Slot(object)
    def _loaded(self, preview: Preview) -> None:
        if self._closing:
            return
        self.model.set_preview(preview)
        self.sheet_combo.addItems(preview.sheet_names)
        if preview.selected_sheet is not None:
            self.sheet_combo.setCurrentText(preview.selected_sheet)
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
        self.status_label.setText(message)

    @Slot(str)
    def _failed(self, message: str) -> None:
        self.errors.setPlainText(message)
        self.status_label.setText("Предпросмотр не загружен. Проверьте файл и настройки.")

    @Slot()
    def _finished(self) -> None:
        worker = self._worker
        self._worker = None
        if worker is not None:
            worker.deleteLater()
        self._set_busy(False)
        if self._closing:
            self.close()

    def _set_busy(self, busy: bool) -> None:
        self.progress.setVisible(busy)
        self.open_button.setEnabled(not busy)
        self.reload_button.setEnabled(not busy and self._path is not None)
        self.sheet_combo.setEnabled(not busy and self.sheet_combo.count() > 1)
        self.encoding_combo.setEnabled(not busy)
        self.delimiter_combo.setEnabled(not busy)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._worker is not None:
            self._closing = True
            self._worker.requestInterruption()
            self.status_label.setText("Завершение чтения перед закрытием…")
            event.ignore()
        else:
            event.accept()
