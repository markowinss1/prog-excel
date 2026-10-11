import os
import time
from collections.abc import Callable, Iterator
from pathlib import Path
from threading import Event

import pytest
from PySide6.QtCore import Qt, QThread, QTimer
from PySide6.QtWidgets import QApplication, QFileDialog

from excel_toolkit.application.models import Cell, CsvOptions, Preview
from excel_toolkit.application.preview import CancelCheck, PreviewService, check_cancelled
from excel_toolkit.ui.main_window import MainWindow


@pytest.fixture(scope="module")
def app() -> QApplication:
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    os.environ["QT_QPA_PLATFORMTHEME"] = ""
    return QApplication.instance() or QApplication([])


def wait_until(app: QApplication, predicate: Callable[[], bool]) -> None:
    deadline = time.monotonic() + 5
    while not predicate():
        app.processEvents()
        if time.monotonic() > deadline:
            pytest.fail("Timed out waiting for the background preview")
        time.sleep(0.002)
    app.processEvents()


@pytest.fixture
def window(app: QApplication, service: PreviewService) -> Iterator[MainWindow]:
    widget = MainWindow(service)
    widget.show()
    yield widget
    widget.close()
    wait_until(app, lambda: widget._worker is None)
    widget.close()
    widget.deleteLater()
    app.processEvents()


def test_file_dialog_preview_sheet_switch_and_error(
    app: QApplication, window: MainWindow, xlsx_file: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (str(xlsx_file), ""))
    gui_threads = []
    window.model.modelReset.connect(lambda: gui_threads.append(QThread.currentThread()))
    window.open_button.click()
    wait_until(app, lambda: window._worker is None)
    assert window.isVisible()
    assert window.windowTitle() == "Excel Freelancer Toolkit"
    assert window.model.rowCount() == 2
    assert window.model.data(window.model.index(1, 0)) == "00123"
    assert window.sheet_combo.isEnabled()
    assert xlsx_file.name in window.file_label.text()
    assert not window.navigation.item(1).flags() & Qt.ItemFlag.ItemIsEnabled
    window.sheet_combo.textActivated.emit("Заказы")
    wait_until(app, lambda: window._worker is None)
    assert window.model.data(window.model.index(0, 0)) == "00007"
    assert window.sheet_combo.currentText() == "Заказы"
    assert all(thread == app.thread() for thread in gui_threads)
    window.load_file(xlsx_file.with_suffix(".xls"))
    wait_until(app, lambda: window._worker is None)
    assert ".xls" in window.errors.toPlainText()
    assert window.model.rowCount() == 0
    assert window.open_button.isEnabled()


def test_csv_retry_after_encoding_error(
    app: QApplication, window: MainWindow, tmp_path: Path
) -> None:
    path = tmp_path / "legacy.csv"
    path.write_text("Код;Имя\n001;Тест\n", encoding="cp1251")
    window.load_file(path)
    wait_until(app, lambda: window._worker is None)
    assert window.errors.toPlainText()
    window.encoding_combo.setCurrentIndex(window.encoding_combo.findData("cp1251"))
    window.reload_button.click()
    wait_until(app, lambda: window._worker is None)
    assert window.errors.toPlainText() == ""
    assert window.model.data(window.model.index(1, 0)) == "001"
    assert not window.sheet_combo.isEnabled()


def test_gui_remains_responsive_and_closes_safely(app: QApplication) -> None:
    started = Event()
    release = Event()

    class SlowReader:
        def read(
            self, path: Path, sheet: str | None, options: CsvOptions, cancelled: CancelCheck
        ) -> Preview:
            started.set()
            while not release.wait(0.01):
                check_cancelled(cancelled)
            return Preview(path, ((Cell("done"),),))

    window = MainWindow(PreviewService(SlowReader()))
    window.show()
    ticks = []
    timer = QTimer()
    timer.setInterval(5)
    timer.timeout.connect(lambda: ticks.append(True))
    timer.start()
    try:
        window.load_file(Path("slow.csv"))
        wait_until(app, lambda: started.is_set() and len(ticks) >= 3)
        assert window._worker is not None
        assert not window.open_button.isEnabled()
        window.close()
        wait_until(app, lambda: window._worker is None)
        assert not window.isVisible()
    finally:
        release.set()
        window.close()
        wait_until(app, lambda: window._worker is None)
        timer.stop()
        window.deleteLater()
