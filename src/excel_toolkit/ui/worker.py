"""One QThread lifecycle for preview, analysis and export; no widgets in workers."""

import logging
from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from excel_toolkit.application.cleaning import CleaningJob, CleaningService
from excel_toolkit.application.models import CsvOptions, Preview, PreviewCancelled, PreviewError
from excel_toolkit.application.preview import PreviewService
from excel_toolkit.modules.cleaner import CleaningError, CleaningPlan

logger = logging.getLogger(__name__)


class ServiceWorker(QThread):
    loaded = Signal(object)
    failed = Signal(str)

    def __init__(self, task: Callable[[], object], description: str) -> None:
        super().__init__()
        self._task = task
        self._description = description

    def run(self) -> None:
        try:
            result = self._task()
            if not self.isInterruptionRequested():
                self.loaded.emit(result)
        except PreviewCancelled:
            pass
        except (PreviewError, CleaningError) as exc:
            self.failed.emit(str(exc))
        except Exception:
            logger.exception("Unexpected %s failure", self._description)
            self.failed.emit(
                f"Неожиданная ошибка: {self._description}. Подробности записаны в терминал."
            )
        finally:
            # Release captured jobs/services, including the preview callback's reference to self.
            self._task = None


class PreviewWorker(ServiceWorker):
    def __init__(
        self, service: PreviewService, path: Path, sheet: str | None, options: CsvOptions
    ) -> None:
        super().__init__(
            lambda: service.load(path, sheet, options, self.isInterruptionRequested),
            "чтение предпросмотра",
        )


def analyze_for_display(
    service: CleaningService, path: Path, plan: CleaningPlan, sheet: str | None, options: CsvOptions
) -> tuple[CleaningJob, Preview, Preview]:
    """Prepare both views in the same worker that analyzes the full supported file."""
    job = service.analyze(path, plan, sheet=sheet, csv_options=options)
    original, cleaned = job.previews()
    return job, original, cleaned


def export_file(
    service: CleaningService,
    job: CleaningJob,
    destination: Path,
    *,
    confirm_conflicts: bool = False,
) -> Path:
    service.export(job, destination, confirm_conflicts=confirm_conflicts)
    return destination
