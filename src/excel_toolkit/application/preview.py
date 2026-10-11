"""Loading orchestration through a small, injectable reader contract."""

from collections.abc import Callable
from pathlib import Path
from typing import Protocol

from excel_toolkit.application.models import CsvOptions, Preview, PreviewCancelled, PreviewError

CancelCheck = Callable[[], bool]


def never_cancel() -> bool:
    return False


def check_cancelled(cancelled: CancelCheck) -> None:
    if cancelled():
        raise PreviewCancelled


class PreviewReader(Protocol):
    def read(
        self, path: Path, sheet: str | None, options: CsvOptions, cancelled: CancelCheck
    ) -> Preview: ...


class PreviewService:
    def __init__(self, reader: PreviewReader) -> None:
        self._reader = reader

    def load(
        self,
        path: str | Path,
        sheet: str | None = None,
        options: CsvOptions | None = None,
        cancelled: CancelCheck = never_cancel,
    ) -> Preview:
        source = Path(path)
        options = options if options is not None else CsvOptions()
        check_cancelled(cancelled)
        if source.suffix.lower() not in {".xlsx", ".csv"}:
            raise PreviewError("Поддерживаются только .xlsx и .csv. Формат .xls не поддерживается.")
        if options.encoding not in {"utf-8-sig", "cp1251", "cp1252", "utf-16"}:
            raise PreviewError("Неизвестная кодировка CSV.")
        if options.delimiter not in {None, ",", ";", "\t", "|"}:
            raise PreviewError("Недопустимый разделитель CSV.")
        return self._reader.read(source, sheet, options, cancelled)
