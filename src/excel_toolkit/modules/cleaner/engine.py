"""Pure cleaning in the fixed order: trim, empty rows, duplicates, missing values."""

from dataclasses import dataclass, field, replace

from excel_toolkit.application.models import Cell
from excel_toolkit.modules.cleaner.models import (
    CleaningError,
    CleaningPlan,
    CleaningReport,
    CleaningResult,
    DataIssue,
    DuplicateGroup,
    MissingValue,
    Table,
)

_EMPTY_CELL = Cell(None)


@dataclass
class _DuplicateDetail:
    removed_source_rows: list[int] = field(default_factory=list)
    removed_values: list[tuple[Cell, ...]] = field(default_factory=list)
    conflicting_columns: set[int] = field(default_factory=set)


@dataclass
class _SeenRow:
    source_row: int
    values: tuple[Cell, ...]
    removed_count: int = 0
    has_conflict: bool = False
    detail: _DuplicateDetail | None = None


def is_missing(cell: Cell) -> bool:
    return cell.data_type != "f" and (
        cell.value is None or isinstance(cell.value, str) and not cell.value.strip()
    )


def _key(cell: Cell) -> tuple:
    # Python alone treats False == 0 == 0.0. Include types, without coercion.
    return cell.data_type, type(cell.value), cell.value


def _cell(row: tuple[Cell, ...], column: int) -> Cell:
    # Absent fields compare like explicit None; empty strings remain distinct.
    return row[column] if column < len(row) else _EMPTY_CELL


def clean(table: Table, plan: CleaningPlan | None = None) -> CleaningResult:
    """Analyze a full table without mutating it. All operations are opt-in.

    Formulas are explicit Cell(..., data_type="f"). Row deletion is rejected if
    formulas exist; file adapters additionally check the other workbook sheets.
    Selected keys with missing fields are retained for review. Detailed duplicate
    groups and issues are bounded; aggregate counts always cover the full table.
    """
    if not isinstance(table, Table):
        raise CleaningError("Очистка требует полную Table; Preview не поддерживается.")
    plan = plan or CleaningPlan()
    width = table.column_count
    columns = plan.duplicate_columns
    if plan.remove_duplicates and columns and any(c >= width for c in columns):
        raise CleaningError("Индекс столбца выходит за пределы таблицы.")
    compared_columns = (
        (columns if columns is not None else tuple(range(width))) if plan.remove_duplicates else ()
    )
    selected_keys = plan.remove_duplicates and columns is not None
    key_columns = set(compared_columns)
    other_columns = tuple(c for c in range(width) if c not in key_columns)
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
    seen: dict[tuple, _SeenRow] = {}
    detailed_groups: list[_SeenRow] = []
    trimmed = empty = duplicates = group_count = conflicts = missing_keys = 0
    has_formulas = False
    for index, original in enumerate(table.rows, start=1):
        has_formulas |= any(cell.data_type == "f" for cell in original)
        row = original
        if plan.trim_whitespace:
            changed = None
            for column, cell in enumerate(row):
                if isinstance(cell.value, str) and cell.data_type != "f":
                    value = cell.value.strip()
                    if value != cell.value:
                        if changed is None:
                            changed = list(row)
                        changed[column] = replace(cell, value=value)
                        trimmed += 1
            if changed is not None:
                row = tuple(changed)
        header = plan.has_header and index == 1
        if not header and plan.remove_empty_rows and all(is_missing(cell) for cell in row):
            empty += 1
            continue
        if not header and plan.remove_duplicates:
            missing_key = selected_keys and any(
                is_missing(_cell(row, column)) for column in compared_columns
            )
            if missing_key:
                missing_keys += 1
            else:
                key = tuple(_key(_cell(row, column)) for column in compared_columns)
                first = seen.get(key)
                if first is None:
                    seen[key] = _SeenRow(index, row)
                else:
                    duplicates += 1
                    if first.removed_count == 0:
                        group_count += 1
                        if len(detailed_groups) < plan.duplicate_detail_limit:
                            first.detail = _DuplicateDetail()
                            detailed_groups.append(first)
                    first.removed_count += 1
                    detail = first.detail
                    if detail is not None:
                        if len(detail.removed_source_rows) < plan.duplicate_row_detail_limit:
                            detail.removed_source_rows.append(index)
                            detail.removed_values.append(row)
                    # Examine every removed row, including those beyond display limits.
                    if detail is not None or not first.has_conflict:
                        was_conflicting = first.has_conflict
                        for column in other_columns:
                            if _key(_cell(first.values, column)) != _key(_cell(row, column)):
                                first.has_conflict = True
                                if detail is not None:
                                    detail.conflicting_columns.add(column)
                                else:
                                    break
                        if first.has_conflict and not was_conflicting:
                            conflicts += 1
                    continue
        output.append(row)
        source_rows.append(index)
    if has_formulas and (empty or duplicates):
        raise CleaningError("Удаление строк заблокировано: таблица содержит формулы.")
    missing = issue_count = 0
    details = []
    issues = []
    if plan.detect_missing_values or selected_keys:
        # Keep the original finite width and the retained rows' source coordinates.
        for index, (row, source_row) in enumerate(zip(output, source_rows, strict=True), start=1):
            header = plan.has_header and source_row == 1
            for column in range(width):
                if not is_missing(_cell(row, column)):
                    continue
                if plan.detect_missing_values:
                    missing += 1
                    if len(details) < plan.missing_detail_limit:
                        details.append(MissingValue(index, column + 1, source_row))
                missing_key = selected_keys and not header and column in key_columns
                if missing_key or plan.detect_missing_values:
                    issue_count += 1
                    if len(issues) < plan.issue_detail_limit:
                        issues.append(
                            DataIssue(
                                source_row,
                                column + 1,
                                "Missing Key" if missing_key else "Missing Value",
                            )
                        )
    groups = []
    row_samples_truncated = False
    for first in detailed_groups:
        detail = first.detail
        assert detail is not None
        row_samples_truncated |= first.removed_count > len(detail.removed_source_rows)
        groups.append(
            DuplicateGroup(
                key=tuple(_cell(first.values, column) for column in compared_columns),
                kept_source_row=first.source_row,
                kept_values=first.values,
                removed_source_rows=tuple(detail.removed_source_rows),
                removed_values=tuple(detail.removed_values),
                removed_count=first.removed_count,
                conflicting_columns=tuple(sorted(detail.conflicting_columns)),
            )
        )
    truncated = missing > len(details)
    issues_truncated = issue_count > len(issues)
    warnings = []
    if conflicts:
        warnings.append(
            f"Конфликтующие группы дубликатов: {conflicts}. Значения вне ключевых столбцов "
            "различаются; сохранена первая строка. Экспорт требует подтверждения."
        )
    if missing_keys:
        warnings.append(f"Строки с отсутствующими ключами сохранены для проверки: {missing_keys}.")
    if group_count > len(groups):
        warnings.append(f"Показаны первые {len(groups)} из {group_count} групп дубликатов.")
    if row_samples_truncated:
        warnings.append("Примеры удаляемых строк усечены; общее количество указано полностью.")
    if truncated:
        warnings.append("Список пропусков усечён; общее количество указано полностью.")
    if issues_truncated:
        warnings.append(f"Showing first {len(issues)} of {issue_count} issues.")
    return CleaningResult(
        Table(tuple(output)),
        CleaningReport(
            original_rows=len(table.rows),
            final_rows=len(output),
            removed_empty_rows=empty,
            removed_duplicates=duplicates,
            changed_text_cells=trimmed,
            missing_values=missing,
            missing_details=tuple(details),
            missing_details_truncated=truncated,
            warnings=tuple(warnings),
            performed_operations=operations,
            duplicate_columns=compared_columns,
            duplicate_groups=tuple(groups),
            duplicate_group_count=group_count,
            conflicting_duplicate_groups=conflicts,
            missing_key_rows=missing_keys,
            issues=tuple(issues),
            issue_count=issue_count,
            issues_truncated=issues_truncated,
        ),
        tuple(source_rows),
    )
