"""Business behavior for selected keys and bounded, complete cleaning analysis."""

from dataclasses import FrozenInstanceError
from datetime import date, datetime, time, timedelta

import pytest

from excel_toolkit.application.models import Cell
from excel_toolkit.modules.cleaner import (
    CleaningError,
    CleaningPlan,
    CleaningReport,
    DataIssue,
    Table,
    clean,
)


def table(*rows):
    return Table(tuple(tuple(v if isinstance(v, Cell) else Cell(v) for v in row) for row in rows))


def test_selected_column_keeps_first_record_and_original_order_with_conflict_details():
    source = table(
        ["ID", "Name", "City"],
        ["00123", "Alex", "Astana"],
        ["00007", "Dana", "Astana"],
        ["00123", "Alex", "Almaty"],
        ["00008", "Kai", "Astana"],
        ["00123", "Alex", "Astana"],
    )
    result = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,)))
    report = result.report
    assert result.source_rows == (1, 2, 3, 5)
    assert report.removed_duplicates == 2
    assert report.duplicate_columns == (0,)
    assert report.duplicate_group_count == report.conflicting_duplicate_groups == 1
    group = report.duplicate_groups[0]
    assert group.key == (Cell("00123"),)
    assert group.kept_source_row == 2
    assert group.kept_values is source.rows[1]
    assert group.removed_source_rows == (4, 6)
    assert group.removed_values[0] is source.rows[3]
    assert group.removed_count == 2
    assert group.conflicting_columns == (2,)
    assert result.table.rows[1] is source.rows[1]
    assert any("Конфликтующие группы дубликатов: 1" in text for text in report.warnings)
    assert all("00123" not in text and "Alex" not in text for text in report.warnings)


def test_multiple_keys_use_indices_and_preserve_requested_order():
    source = table(
        ["ID", "ID", ""],
        ["001", "Astana", "a"],
        ["001", "Almaty", "b"],
        ["001", "Astana", "c"],
        ["002", "Astana", "d"],
    )
    result = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(1, 0)))
    assert result.source_rows == (1, 2, 3, 5)
    assert result.report.duplicate_columns == (1, 0)
    assert result.report.duplicate_groups[0].key == (Cell("Astana"), Cell("001"))
    assert result.report.duplicate_groups[0].conflicting_columns == (2,)


def test_all_columns_preserves_legacy_matching_of_missing_and_ragged_records():
    source = table(["ID", "Name"], ["001", "a"], ["001", "b"], ["001", "a"], [], [None])
    result = clean(source, CleaningPlan(remove_duplicates=True))
    assert result.source_rows == (1, 2, 3, 5)
    assert result.report.duplicate_columns == (0, 1)
    assert result.report.removed_duplicates == result.report.duplicate_group_count == 2
    assert result.report.conflicting_duplicate_groups == 0
    assert result.report.missing_key_rows == result.report.issue_count == 0
    assert result.report.duplicate_groups[1].key == (Cell(None), Cell(None))
    assert result.report.duplicate_groups[1].kept_values == ()


def test_explicit_selection_of_every_column_still_protects_missing_keys():
    source = table(["ID", "Name"], ["001", ""], ["001", ""], ["002", "Kai"], ["002", "Kai"])
    result = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0, 1)))
    assert result.source_rows == (1, 2, 3, 4)
    assert result.report.removed_duplicates == 1
    assert result.report.missing_key_rows == 2
    assert result.report.issues == (DataIssue(2, 2, "Missing Key"), DataIssue(3, 2, "Missing Key"))


@pytest.mark.parametrize("missing_key", [None, "", " \t\n\u00a0"])
def test_any_missing_key_component_excludes_row_from_all_duplicate_groups(missing_key):
    source = table(
        ["ID", "Branch", "Name"],
        ["001", missing_key, "Alex"],
        ["001", missing_key, "Alex"],
        [missing_key, "Almaty", "Dana"],
        [missing_key, "Almaty", "Dana"],
    )
    result = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0, 1)))
    assert result.table == source
    assert result.report.missing_key_rows == result.report.issue_count == 4
    assert result.report.duplicate_group_count == result.report.removed_duplicates == 0
    assert result.report.duplicate_groups == ()
    assert result.report.issues == (
        DataIssue(2, 2, "Missing Key"),
        DataIssue(3, 2, "Missing Key"),
        DataIssue(4, 1, "Missing Key"),
        DataIssue(5, 1, "Missing Key"),
    )


def test_ragged_key_rows_remain_and_count_rows_separately_from_missing_cells():
    source = table(["ID", "Branch", "Name"], [], [], ["001"], ["001"])
    result = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0, 1)))
    assert result.table == source
    assert result.report.missing_key_rows == 4
    assert result.report.issue_count == 6
    assert result.report.issues[:2] == (
        DataIssue(2, 1, "Missing Key"),
        DataIssue(2, 2, "Missing Key"),
    )
    assert result.report.missing_values == 0
    assert result.report.missing_details == ()


def test_strict_types_case_and_text_identifiers_survive_selected_matching():
    keys = (
        1,
        True,
        1.0,
        "1",
        "001",
        "abc",
        "ABC",
        date(2026, 1, 2),
        datetime(2026, 1, 2),
        time(12),
        timedelta(days=1),
        Cell(1, "n"),
        0,
        False,
        0.0,
    )
    source = table(["ID", "Name"], *([key, "same"] for key in keys + keys))
    result = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,)))
    assert result.source_rows == tuple(range(1, len(keys) + 2))
    assert result.report.removed_duplicates == result.report.duplicate_group_count == len(keys)
    assert result.report.conflicting_duplicate_groups == result.report.missing_key_rows == 0
    for original, actual in zip(source.rows, result.table.rows, strict=False):
        assert actual is original


def test_nonkey_conflicts_also_use_strict_value_and_cell_data_types():
    source = table(
        ["ID", "Value"],
        ["a", 1],
        ["a", True],
        ["b", 1],
        ["b", 1.0],
        ["c", Cell(1, "s")],
        ["c", Cell(1, "n")],
        ["d", Cell(1, "n", "0")],
        ["d", Cell(1, "n", "000")],
    )
    report = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,))).report
    assert report.duplicate_group_count == 4
    assert report.conflicting_duplicate_groups == 3
    assert [group.conflicting_columns for group in report.duplicate_groups] == [
        (1,),
        (1,),
        (1,),
        (),
    ]


def test_trim_and_empty_rows_precede_matching_and_keep_original_source_coordinates():
    source = table(
        [" ID ", "Email"],
        [" ", " "],
        [" 001 ", " a@example.test "],
        ["001", "a@example.test"],
        [" ", "person@example.test"],
        ["", "person@example.test"],
        ["002", ""],
    )
    result = clean(source, CleaningPlan(True, True, True, True, True, (0,)))
    assert result.source_rows == (1, 3, 5, 6, 7)
    assert result.report.changed_text_cells == 6
    assert result.report.removed_empty_rows == result.report.removed_duplicates == 1
    assert result.report.duplicate_groups[0].kept_source_row == 3
    assert result.report.duplicate_groups[0].removed_source_rows == (4,)
    assert result.report.duplicate_groups[0].key == (Cell("001"),)
    assert result.report.conflicting_duplicate_groups == 0
    assert result.report.missing_key_rows == 2
    assert result.report.issues == (
        DataIssue(5, 1, "Missing Key"),
        DataIssue(6, 1, "Missing Key"),
        DataIssue(7, 2, "Missing Value"),
    )
    assert [(item.row, item.source_row, item.column) for item in result.report.missing_details] == [
        (3, 5, 1),
        (4, 6, 1),
        (5, 7, 2),
    ]
    assert source.rows[2][0].value == " 001 "


def test_missing_headers_are_values_and_no_header_data_can_have_missing_keys():
    source = table(["", ""], ["001", None])
    plan = CleaningPlan(remove_duplicates=True, duplicate_columns=(0,), detect_missing_values=True)
    result = clean(source, plan)
    assert result.report.missing_key_rows == 0
    assert result.report.issues == (
        DataIssue(1, 1, "Missing Value"),
        DataIssue(1, 2, "Missing Value"),
        DataIssue(2, 2, "Missing Value"),
    )
    result = clean(
        source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,), has_header=False)
    )
    assert result.report.missing_key_rows == 1
    assert result.report.issues == (DataIssue(1, 1, "Missing Key"),)


def test_inactive_duplicate_selection_does_not_emit_missing_key_issues():
    source = table(["ID"], [None], [None])
    report = clean(source, CleaningPlan(duplicate_columns=(0,), detect_missing_values=True)).report
    assert report.missing_key_rows == report.removed_duplicates == 0
    assert report.duplicate_columns == report.duplicate_groups == ()
    assert report.issues == (DataIssue(2, 1, "Missing Value"), DataIssue(3, 1, "Missing Value"))


def test_issue_limit_is_independent_of_legacy_missing_detail_limit():
    source = table(["ID", "Email"], [None, None], ["001", None], [None, None])
    report = clean(
        source,
        CleaningPlan(
            remove_duplicates=True,
            duplicate_columns=(0,),
            detect_missing_values=True,
            missing_detail_limit=1,
            issue_detail_limit=3,
        ),
    ).report
    assert report.missing_values == report.issue_count == 5
    assert report.missing_key_rows == 2
    assert len(report.missing_details) == 1
    assert report.missing_details_truncated and report.issues_truncated
    assert report.issues == (
        DataIssue(2, 1, "Missing Key"),
        DataIssue(2, 2, "Missing Value"),
        DataIssue(3, 2, "Missing Value"),
    )
    assert "Showing first 3 of 5 issues." in report.warnings


def test_conflicts_and_counts_cover_groups_and_rows_beyond_detail_limits():
    source = table(
        ["ID", "Name", "City"],
        ["a", "Alex", "Astana"],
        ["a", "Alex", "Astana"],
        ["b", "Dana", "Astana"],
        ["b", "Dana", "Astana"],
        ["a", "Alex", "Almaty"],
        ["b", "Dana", "Almaty"],
        ["a", "Alexander", "Astana"],
        ["b", "Daniel", "Astana"],
    )
    report = clean(
        source,
        CleaningPlan(
            remove_duplicates=True,
            duplicate_columns=(0,),
            duplicate_detail_limit=1,
            duplicate_row_detail_limit=1,
        ),
    ).report
    assert report.removed_duplicates == 6
    assert report.duplicate_group_count == report.conflicting_duplicate_groups == 2
    assert len(report.duplicate_groups) == 1
    group = report.duplicate_groups[0]
    assert group.removed_source_rows == (3,)
    assert group.removed_count == 3
    assert group.conflicting_columns == (1, 2)
    assert len(group.removed_values) == 1
    assert any("1 из 2 групп" in warning for warning in report.warnings)
    assert any("Примеры удаляемых строк усечены" in warning for warning in report.warnings)
    assert any("Конфликтующие группы дубликатов: 2" in warning for warning in report.warnings)


def test_zero_detail_limits_keep_complete_counts_and_warnings():
    source = table(["ID", "Name"], ["a", "x"], ["a", "y"], [None, None])
    report = clean(
        source,
        CleaningPlan(
            remove_duplicates=True,
            duplicate_columns=(0,),
            duplicate_detail_limit=0,
            duplicate_row_detail_limit=0,
            issue_detail_limit=0,
        ),
    ).report
    assert report.removed_duplicates == report.duplicate_group_count == 1
    assert report.conflicting_duplicate_groups == report.missing_key_rows == report.issue_count == 1
    assert report.duplicate_groups == report.issues == ()
    assert report.issues_truncated
    assert any("0 из 1 групп" in warning for warning in report.warnings)
    assert "Showing first 0 of 1 issues." in report.warnings


def test_zero_removed_row_sample_still_reports_group_conflicts():
    source = table(["ID", "Name"], ["a", "x"], ["a", "y"])
    report = clean(
        source,
        CleaningPlan(remove_duplicates=True, duplicate_columns=(0,), duplicate_row_detail_limit=0),
    ).report
    group = report.duplicate_groups[0]
    assert group.removed_source_rows == group.removed_values == ()
    assert group.removed_count == 1
    assert group.conflicting_columns == (1,)
    assert report.conflicting_duplicate_groups == 1


@pytest.mark.parametrize(
    "name", ["duplicate_detail_limit", "duplicate_row_detail_limit", "issue_detail_limit"]
)
@pytest.mark.parametrize("value", [-1, 1.5, True, None])
def test_new_limits_require_nonnegative_integers(name, value):
    with pytest.raises(CleaningError, match="Лимит"):
        CleaningPlan(**{name: value})


def test_existing_positional_report_and_plan_arguments_remain_compatible():
    report = CleaningReport(10, 8, 1, 1, 2, 3, (), False, ("warning",), ("remove_duplicates",))
    assert report.warnings == ("warning",)
    assert report.performed_operations == ("remove_duplicates",)
    assert report.duplicate_groups == report.issues == ()
    plan = CleaningPlan(True, True, True, True, False, (0,), 12)
    assert plan.missing_detail_limit == 12
    assert plan.duplicate_detail_limit == plan.issue_detail_limit == 100
    assert plan.duplicate_row_detail_limit == 20


def test_duplicate_groups_and_issues_are_immutable():
    source = table(["ID", "Name"], ["a", "x"], ["a", "y"], [None, "z"])
    report = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,))).report
    with pytest.raises(FrozenInstanceError):
        report.duplicate_groups[0].removed_count = 42
    with pytest.raises(FrozenInstanceError):
        report.issues[0].kind = "changed"


def test_ten_thousand_rows_have_bounded_samples_and_complete_totals():
    source = table(
        ["ID", "Value"],
        *([f"{index % 1000:05}", index // 1000] for index in range(10_000)),
    )
    result = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,)))
    report = result.report
    assert result.source_rows == tuple(range(1, 1002))
    assert report.removed_duplicates == 9000
    assert report.duplicate_group_count == report.conflicting_duplicate_groups == 1000
    assert len(report.duplicate_groups) == 100
    assert all(group.removed_count == 9 for group in report.duplicate_groups)
    assert all(len(group.removed_values) == 9 for group in report.duplicate_groups)
    assert all(group.conflicting_columns == (1,) for group in report.duplicate_groups)


def test_one_large_group_retains_only_default_number_of_removed_row_samples():
    source = table(["ID", "Value"], *(["001", "same"] for _ in range(100)), ["001", "late"])
    report = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,))).report
    group = report.duplicate_groups[0]
    assert report.removed_duplicates == group.removed_count == 100
    assert group.removed_source_rows == tuple(range(3, 23))
    assert group.conflicting_columns == (1,)
    assert report.conflicting_duplicate_groups == 1


def test_ragged_nonkey_none_matches_explicit_none_without_conflict():
    source = table(["ID", "Name"], ["a"], ["a", None])
    report = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,))).report
    assert report.removed_duplicates == 1
    assert report.conflicting_duplicate_groups == 0
    assert report.duplicate_groups[0].conflicting_columns == ()


def test_formula_keys_are_not_missing_and_formula_deletion_restriction_remains():
    source = table(["ID", "Value"], [Cell("", "f"), "x"], [None, "y"])
    report = clean(source, CleaningPlan(remove_duplicates=True, duplicate_columns=(0,))).report
    assert report.missing_key_rows == 1
    assert report.issues == (DataIssue(3, 1, "Missing Key"),)
    with pytest.raises(CleaningError, match="формулы"):
        clean(
            table(["ID", "Value"], [Cell("=1", "f"), "x"], [Cell("=1", "f"), "y"]),
            CleaningPlan(remove_duplicates=True, duplicate_columns=(0,)),
        )
