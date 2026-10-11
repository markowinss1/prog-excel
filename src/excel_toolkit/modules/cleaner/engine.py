"""Pure cleaning in the fixed order: trim, empty rows, duplicates, missing values."""

from dataclasses import replace

from excel_toolkit.application.models import Cell
from excel_toolkit.modules.cleaner.models import (
    CleaningError,
    CleaningPlan,
    CleaningReport,
    CleaningResult,
    MissingValue,
    Table,
)


def is_missing(cell: Cell) -> bool:
    return cell.data_type != "f" and (
        cell.value is None or isinstance(cell.value, str) and not cell.value.strip()
    )


def _key(cell: Cell) -> tuple:
    # Python alone treats False == 0 == 0.0. Include types, without coercion.
    return cell.data_type, type(cell.value), cell.value


def clean(table: Table, plan: CleaningPlan | None = None) -> CleaningResult:
    """Analyze a full table without mutating it. All operations are opt-in.

    Formulas are explicit Cell(..., data_type="f"). Row deletion is rejected if
    formulas exist; file adapters additionally check the other workbook sheets.
    """
    if not isinstance(table, Table):
        raise CleaningError("Очистка требует полную Table; Preview не поддерживается.")
    plan = plan or CleaningPlan()
    width = table.column_count
    columns = plan.duplicate_columns
    if plan.remove_duplicates and columns and any(c >= width for c in columns):
        raise CleaningError("Индекс столбца выходит за пределы таблицы.")
    operations = tuple(
        name
        for name in (
            "trim_whitespace",
            "remove_empty_rows",
            "remove_duplicates",
            "detect_missing_values",
        )
        if getattr(plan, name)
    )
    output = []
    source_rows = []
    seen = set()
    trimmed = empty = duplicates = 0
    has_formulas = False
    for index, original in enumerate(table.rows, start=1):
        has_formulas |= any(cell.data_type == "f" for cell in original)
        row = original
        if plan.trim_whitespace:
            changed = []
            for cell in row:
                if isinstance(cell.value, str) and cell.data_type != "f":
                    value = cell.value.strip()
                    if value != cell.value:
                        cell = replace(cell, value=value)
                        trimmed += 1
                changed.append(cell)
            row = tuple(changed)
        header = plan.has_header and index == 1
        if not header and plan.remove_empty_rows and all(is_missing(cell) for cell in row):
            empty += 1
            continue
        if not header and plan.remove_duplicates:
            # Absent fields compare like explicit None; empty strings remain distinct.
            key = tuple(
                _key(row[c] if c < len(row) else Cell(None))
                for c in (columns if columns is not None else range(width))
            )
            if key in seen:
                duplicates += 1
                continue
            seen.add(key)
        output.append(row)
        source_rows.append(index)
    if has_formulas and (empty or duplicates):
        raise CleaningError("Удаление строк заблокировано: таблица содержит формулы.")
    missing = 0
    details = []
    if plan.detect_missing_values:
        # Use the original finite width, including columns whose blank rows were removed.
        for index, (row, source_row) in enumerate(zip(output, source_rows, strict=True), start=1):
            for column in range(width):
                if column >= len(row) or is_missing(row[column]):
                    missing += 1
                    if len(details) < plan.missing_detail_limit:
                        details.append(MissingValue(index, column + 1, source_row))
    truncated = missing > len(details)
    warnings = (
        ("Список пропусков усечён; общее количество указано полностью.",) if truncated else ()
    )
    return CleaningResult(
        Table(tuple(output)),
        CleaningReport(
            len(table.rows),
            len(output),
            empty,
            duplicates,
            trimmed,
            missing,
            tuple(details),
            truncated,
            warnings,
            operations,
        ),
        tuple(source_rows),
    )
