"""One background load; signals are delivered to the GUI thread."""

import logging
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from excel_toolkit.application.models import CsvOptions, PreviewCancelled, PreviewError
from excel_toolkit.application.preview import PreviewService

logger = logging.getLogger(__name__)


class PreviewWorker(QThread):
    loaded = Signal(object)
    failed = Signal(str)

    def __init__(
        self, service: PreviewService, path: Path, sheet: str | None, options: CsvOptions
    ) -> None:
        super().__init__()
        self._service = service
        self._path = path
        self._sheet = sheet
        self._options = options

    def run(self) -> None:
        try:
            preview = self._service.load(
                self._path, self._sheet, self._options, self.isInterruptionRequested
            )
            if not self.isInterruptionRequested():
                self.loaded.emit(preview)
        except PreviewCancelled:
            pass
        except PreviewError as exc:
            self.failed.emit(str(exc))
        except Exception:
            logger.exception("Unexpected preview failure")
            self.failed.emit("Неожиданная ошибка чтения. Подробности записаны в терминал.")
