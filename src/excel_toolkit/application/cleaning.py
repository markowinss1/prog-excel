"""File cleaning workflow, independent of Qt and of a concrete storage adapter."""

from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Protocol

from excel_toolkit.application.models import COLUMN_LIMIT, ROW_LIMIT, CsvOptions, Preview
from excel_toolkit.modules.cleaner import CleaningPlan, CleaningResult, Table, clean


@dataclass(frozen=True)
class CleaningSource:
    path: Path
    table: Table
    sheet_names: tuple[str, ...] = ()
    selected_sheet: str | None = None
    csv_options: CsvOptions | None = None
    # Immutable source snapshot: export never reopens a potentially changed original.
    xlsx_snapshot: bytes | None = field(default=None, repr=False)
    sheet_part: str | None = None
    contains_formulas: bool = False
    # Small, validated view subtree; avoids reparsing worksheet data during validation.
    xlsx_sheet_views: bytes | None = field(default=None, repr=False)


@dataclass(frozen=True)
class CleaningJob:
    source: CleaningSource
    result: CleaningResult

    def previews(self) -> tuple[Preview, Preview]:
        """Bounded views of the exact analyzed snapshot, never inputs to export."""
        return tuple(
            Preview(
                path=self.source.path,
                rows=tuple(row[:COLUMN_LIMIT] for row in table.rows[:ROW_LIMIT]),
                sheet_names=self.source.sheet_names,
                selected_sheet=self.source.selected_sheet,
                more_rows=len(table.rows) > ROW_LIMIT,
                more_columns=any(len(row) > COLUMN_LIMIT for row in table.rows),
                csv_options=self.source.csv_options,
            )
            for table in (self.source.table, self.result.table)
        )


class CleaningStorage(Protocol):
    def read(self, path: Path, sheet: str | None, options: CsvOptions) -> CleaningSource: ...

    def validate(self, source: CleaningSource, result: CleaningResult) -> None: ...

    def write(self, job: CleaningJob, destination: Path) -> None: ...


class CleaningService:
    def __init__(self, storage: CleaningStorage) -> None:
        self._storage = storage

    def analyze(
        self,
        path: str | Path,
        plan: CleaningPlan,
        *,
        sheet: str | None = None,
        csv_options: CsvOptions | None = None,
    ) -> CleaningJob:
        source = self._storage.read(Path(path), sheet, csv_options or CsvOptions())
        result = clean(source.table, plan)
        self._storage.validate(source, result)
        if source.contains_formulas:
            result = replace(
                result,
                report=replace(
                    result.report,
                    warnings=result.report.warnings
                    + (
                        "Формулы и их кэш сохранены без вычисления. После изменения текста "
                        "зависимые результаты могут требовать пересчёта в Excel.",
                    ),
                ),
            )
        return CleaningJob(source, result)

    def export(self, job: CleaningJob, destination: str | Path) -> None:
        self._storage.write(job, Path(destination))
