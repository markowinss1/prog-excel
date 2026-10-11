"""Immutable preview data. A preview is never a complete workbook for export."""

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path

ROW_LIMIT = 100
COLUMN_LIMIT = 200
CellValue = str | int | float | bool | datetime | date | time | timedelta | None


class PreviewError(Exception):
    """A loading failure suitable for display to the user."""


class PreviewCancelled(Exception):
    """Cooperative cancellation of an in-progress preview."""


@dataclass(frozen=True)
class CsvOptions:
    encoding: str = "utf-8-sig"
    delimiter: str | None = None


@dataclass(frozen=True)
class Cell:
    value: CellValue
    data_type: str = "s"
    number_format: str = "General"


@dataclass(frozen=True)
class Preview:
    path: Path
    rows: tuple[tuple[Cell, ...], ...]
    sheet_names: tuple[str, ...] = ()
    selected_sheet: str | None = None
    more_rows: bool = False
    more_columns: bool = False
    csv_options: CsvOptions | None = None

    @property
    def column_count(self) -> int:
        return max((len(row) for row in self.rows), default=0)
