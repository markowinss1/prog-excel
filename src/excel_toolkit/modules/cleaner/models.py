"""Complete, immutable tables and cleaning results (never a GUI Preview)."""

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta

from excel_toolkit.application.models import Cell


class CleaningError(Exception):
    """An invalid request or an operation that cannot preserve data safely."""


@dataclass(frozen=True)
class Table:
    rows: tuple[tuple[Cell, ...], ...]

    def __post_init__(self) -> None:
        object.__setattr__(self, "rows", tuple(tuple(row) for row in self.rows))
        if any(not isinstance(cell, Cell) for row in self.rows for cell in row):
            raise CleaningError("Table принимает строки из Cell, а не Preview или сырые значения.")
        supported = {str, int, float, bool, date, datetime, time, timedelta, type(None)}
        if any(type(cell.value) not in supported for row in self.rows for cell in row):
            raise CleaningError("Неподдерживаемый тип значения Cell в таблице.")

    @property
    def column_count(self) -> int:
        return max(map(len, self.rows), default=0)


@dataclass(frozen=True)
class CleaningPlan:
    trim_whitespace: bool = False
    remove_empty_rows: bool = False
    remove_duplicates: bool = False
    detect_missing_values: bool = False
    has_header: bool = True
    duplicate_columns: tuple[int, ...] | None = None  # Zero-based indexes.
    missing_detail_limit: int = 1000
    duplicate_detail_limit: int = 100
    duplicate_row_detail_limit: int = 20
    issue_detail_limit: int = 100

    def __post_init__(self) -> None:
        if any(
            type(flag) is not bool
            for flag in (
                self.trim_whitespace,
                self.remove_empty_rows,
                self.remove_duplicates,
                self.detect_missing_values,
                self.has_header,
            )
        ):
            raise CleaningError("Флаги операций и has_header должны быть bool.")
        for limit in (
            self.missing_detail_limit,
            self.duplicate_detail_limit,
            self.duplicate_row_detail_limit,
            self.issue_detail_limit,
        ):
            if type(limit) is not int or limit < 0:
                raise CleaningError("Лимит подробностей должен быть целым неотрицательным числом.")
        if self.duplicate_columns is not None:
            columns = tuple(self.duplicate_columns)
            if not columns or any(type(c) is not int or c < 0 for c in columns):
                raise CleaningError("Укажите непустой набор индексов столбцов от нуля.")
            if len(set(columns)) != len(columns):
                raise CleaningError("Индексы столбцов не должны повторяться.")
            object.__setattr__(self, "duplicate_columns", columns)


@dataclass(frozen=True)
class MissingValue:
    row: int  # One-based coordinates in the result.
    column: int
    source_row: int  # One-based coordinate before cleaning.


@dataclass(frozen=True)
class DataIssue:
    source_row: int  # One-based coordinate before cleaning.
    column: int  # One-based column coordinate.
    kind: str  # "Missing Key" or "Missing Value".


@dataclass(frozen=True)
class DuplicateGroup:
    key: tuple[Cell, ...]
    kept_source_row: int
    kept_values: tuple[Cell, ...]
    removed_source_rows: tuple[int, ...]
    removed_values: tuple[tuple[Cell, ...], ...]
    removed_count: int  # Full count, including removed rows beyond the sample limit.
    conflicting_columns: tuple[int, ...]  # Zero-based indexes, including unsampled rows.


@dataclass(frozen=True)
class CleaningReport:
    original_rows: int
    final_rows: int
    removed_empty_rows: int = 0
    removed_duplicates: int = 0
    changed_text_cells: int = 0
    missing_values: int = 0
    missing_details: tuple[MissingValue, ...] = ()
    missing_details_truncated: bool = False
    warnings: tuple[str, ...] = ()
    performed_operations: tuple[str, ...] = ()
    duplicate_columns: tuple[int, ...] = ()
    duplicate_groups: tuple[DuplicateGroup, ...] = ()
    duplicate_group_count: int = 0
    conflicting_duplicate_groups: int = 0
    missing_key_rows: int = 0
    issues: tuple[DataIssue, ...] = ()
    issue_count: int = 0
    issues_truncated: bool = False


@dataclass(frozen=True)
class CleaningResult:
    table: Table
    report: CleaningReport
    source_rows: tuple[int, ...]

    def preview(self, limit: int = 100) -> Table:
        if type(limit) is not int or limit < 0:
            raise CleaningError("Лимит предпросмотра должен быть целым неотрицательным числом.")
        return Table(self.table.rows[:limit])
