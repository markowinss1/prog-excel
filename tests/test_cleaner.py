from datetime import date, datetime, time
from itertools import product

import pytest

from excel_toolkit.application.models import Cell, Preview
from excel_toolkit.modules.cleaner import CleaningError, CleaningPlan, Table, clean


def table(*rows):
    return Table(tuple(tuple(v if isinstance(v, Cell) else Cell(v) for v in row) for row in rows))


def values(result):
    return [[c.value for c in row] for row in result.table.rows]


def test_empty_rows_preserve_header_partial_rows_zero_and_false():
    source = table([None, " "], [None, "\t\n\u2003"], [], [0], [False], ["0"], [None, "x"])
    result = clean(source, CleaningPlan(remove_empty_rows=True))
    assert values(result) == [[None, " "], [0], [False], ["0"], [None, "x"]]
    assert result.report.removed_empty_rows == 2
    assert result.source_rows == (1, 4, 5, 6, 7)
    assert len(source.rows) == 7
    assert (
        clean(source, CleaningPlan(remove_empty_rows=True, has_header=False)).report.final_rows == 4
    )


def test_duplicates_are_stable_case_and_type_sensitive():
    source = table(["h"], ["ABC"], ["abc"], ["ABC"], [0], [False], [0.0], ["0"], ["001"])
    result = clean(source, CleaningPlan(remove_duplicates=True))
    assert values(result) == [["h"], ["ABC"], ["abc"], [0], [False], [0.0], ["0"], ["001"]]
    assert result.report.removed_duplicates == 1
    assert clean(table(["h"], ["h"]), CleaningPlan(remove_duplicates=True)).report.final_rows == 2


def test_duplicate_selected_columns_and_ragged_rows():
    source = table(["id", "name"], ["001", "one"], ["001", "two"], ["002"])
    result = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,)))
    assert values(result) == [["id", "name"], ["001", "one"], ["001", "two"], ["002"]]
    assert result.report.removed_duplicates == 0
    assert result.report.conflicting_duplicate_groups == 1
    assert clean(source, CleaningPlan(remove_duplicates=True)).report.removed_duplicates == 0
    with pytest.raises(CleaningError, match="пределы"):
        clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(2,)))


def test_trim_preserves_types_inner_spaces_formulas_unicode_and_ids():
    source = table(
        [" Заголовок "],
        [
            "\u00a0 001234 \t",
            " Алматы  Астана ",
            12,
            False,
            date(2026, 1, 2),
            datetime(2026, 1, 2),
            time(12),
            Cell('= A1 & " x " ', "f"),
        ],
    )
    result = clean(source, CleaningPlan(trim_whitespace=True))
    assert result.report.changed_text_cells == 3
    assert values(result)[1][:3] == ["001234", "Алматы  Астана", 12]
    assert result.table.rows[1][3:] == source.rows[1][3:]
    assert source.rows[1][0].value == "\u00a0 001234 \t"


def test_missing_in_finite_rectangle_with_source_coordinates_and_limit():
    source = table(["h", ""], [None, " "], ["x"], [False, 0])
    result = clean(
        source,
        CleaningPlan(remove_empty_rows=True, detect_missing_values=True, missing_detail_limit=1),
    )
    assert result.report.missing_values == 2
    assert result.report.missing_details[0].row == 1
    assert result.report.missing_details[0].column == 2
    assert result.report.missing_details_truncated
    assert result.report.warnings
    full = clean(source, CleaningPlan(remove_empty_rows=True, detect_missing_values=True))
    assert full.report.missing_details[1].source_row == 3
    assert full.report.missing_details[1].row == 2
    zero = clean(source, CleaningPlan(detect_missing_values=True, missing_detail_limit=0))
    assert zero.report.missing_values == 4
    assert zero.report.missing_details == ()
    assert zero.report.missing_details_truncated


def test_fixed_order_changes_duplicate_result_and_counts_trim_before_removal():
    source = table(["h"], [" a "], ["a"], [" "], [""])
    plan = CleaningPlan(True, True, True, True)
    result = clean(source, plan)
    assert values(result) == [["h"], ["a"]]
    assert result.report.changed_text_cells == 2
    assert result.report.removed_empty_rows == 2
    assert result.report.removed_duplicates == 1
    assert result.report.performed_operations == (
        "trim_whitespace",
        "remove_empty_rows",
        "remove_duplicates",
        "detect_missing_values",
    )
    reverse = clean(
        clean(source, CleaningPlan(remove_duplicates=True)).table,
        CleaningPlan(trim_whitespace=True),
    )
    assert reverse.report.final_rows == 5


@pytest.mark.parametrize("flags", list(product([False, True], repeat=4)))
def test_all_combinations_are_opt_in_and_idempotent(flags):
    source = table(["h", "b"], [" a ", ""], ["a", ""], [" ", ""], ["0", "z"])
    plan = CleaningPlan(*flags)
    first = clean(source, plan)
    second = clean(first.table, plan)
    assert second.table == first.table
    assert second.report.changed_text_cells == 0
    assert second.report.removed_empty_rows == 0
    assert second.report.removed_duplicates == 0
    report = first.report
    assert flags[0] or report.changed_text_cells == 0
    assert flags[1] or report.removed_empty_rows == 0
    assert flags[2] or report.removed_duplicates == 0
    assert flags[3] or report.missing_values == 0
    assert len(source.rows) == 5


@pytest.mark.parametrize("source", [Table(()), table(["h", "name"])])
def test_empty_and_header_only(source):
    result = clean(source, CleaningPlan(True, True, True, True))
    assert result.table == source
    assert result.report.original_rows == result.report.final_rows == len(source.rows)
    assert result.report.missing_values == 0


def test_formulas_prevent_actual_deletion_and_are_distinct_from_text():
    source = table(["h"], [Cell("=1", "f")], ["=1"], [])
    with pytest.raises(CleaningError, match="формулы"):
        clean(source, CleaningPlan(remove_empty_rows=True))
    assert clean(source, CleaningPlan(remove_duplicates=True)).report.removed_duplicates == 0
    assert clean(table(["h"], [Cell("=1", "f")]), CleaningPlan(True, True, True, True))


def test_preview_is_bounded_and_input_is_immutable():
    rows = [[Cell(" a ")], [Cell(" b ")]]
    source = Table(rows)
    rows[0].append(Cell("external mutation"))
    result = clean(source, CleaningPlan(trim_whitespace=True))
    assert result.preview(1).rows == ((Cell("a"),),)
    assert result.preview(0).rows == ()
    assert source.rows == ((Cell(" a "),), (Cell(" b "),))
    with pytest.raises(CleaningError):
        result.preview(-1)
    with pytest.raises(CleaningError):
        clean(Preview(None, ()))


@pytest.mark.parametrize(
    "kwargs",
    [
        {"duplicate_columns": ()},
        {"duplicate_columns": (-1,)},
        {"duplicate_columns": (0, 0)},
        {"duplicate_columns": (True,)},
        {"missing_detail_limit": -1},
        {"missing_detail_limit": 1.5},
    ],
)
def test_invalid_plans(kwargs):
    with pytest.raises(CleaningError):
        CleaningPlan(**kwargs)


def test_invalid_cell_values_and_boolean_options_are_rejected():
    with pytest.raises(CleaningError, match="тип"):
        Table(((Cell(["mutable"]),),))
    with pytest.raises(CleaningError, match="bool"):
        CleaningPlan(remove_duplicates="False")
