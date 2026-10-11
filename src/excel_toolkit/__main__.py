"""Composition root; run with python -m excel_toolkit."""

import sys

from PySide6.QtWidgets import QApplication

from excel_toolkit.application.preview import PreviewService
from excel_toolkit.infrastructure.files.reader import FilePreviewReader
from excel_toolkit.ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Freelancer Toolkit")
    window = MainWindow(PreviewService(FilePreviewReader()))
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
