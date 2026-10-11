"""Keep the documented synthetic examples reproducible through real file exports."""

import hashlib
import runpy
from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import load_workbook

from excel_toolkit.application.cleaning import CleaningService
from excel_toolkit.application.models import CsvOptions
from excel_toolkit.infrastructure.files.cleaning import FileCleaningStorage
from excel_toolkit.modules.cleaner import CleaningPlan


@pytest.fixture
def sample_generator():
    script = Path(__file__).parents[1] / "examples" / "generate_advanced_samples.py"
    return runpy.run_path(str(script))["generate_samples"]


def fingerprint(path):
    return hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns


@pytest.mark.parametrize("extension", ["csv", "xlsx"])
@pytest.mark.parametrize(
    ("columns", "source_rows", "duplicates", "conflicts", "missing", "group_count"),
    [
        (None, (1, 2, 3, 4, 6, 8, 9, 10, 13, 14, 15), 3, 0, 4, 3),
        ((0,), (1, 2, 3, 4, 6, 7, 8, 9, 10, 13, 14, 15), 2, 2, 6, 4),
        ((0, 2), (1, 2, 3, 4, 6, 7, 8, 9, 10, 13, 14, 15), 2, 1, 6, 3),
    ],
)
def test_documented_sample_reports_and_exports(
    sample_generator,
    tmp_path,
    extension,
    columns,
    source_rows,
    duplicates,
    conflicts,
    missing,
    group_count,
):
    sample_generator(tmp_path)
    source = tmp_path / f"advanced_cleaning_demo.{extension}"
    original = fingerprint(source)
    service = CleaningService(FileCleaningStorage())
    plan = CleaningPlan(True, True, True, True, duplicate_columns=columns)
    options = CsvOptions(encoding="utf-8-sig", delimiter=";")
    job = service.analyze(source, plan, csv_options=options)
    report = job.result.report
    assert job.result.source_rows == source_rows
    assert report.original_rows == 15
    assert report.final_rows == len(source_rows)
    assert report.removed_empty_rows == 1
    assert report.removed_duplicates == duplicates
    assert report.changed_text_cells == 5
    assert report.missing_values == report.issue_count == missing
    assert report.duplicate_group_count == group_count
    assert report.conflicting_duplicate_groups == conflicts
    assert report.missing_key_rows == (2 if columns is not None else 0)
    if columns is not None:
        assert [(issue.source_row, issue.column, issue.kind) for issue in report.issues] == [
            (4, 4, "Missing Value"),
            (6, 1, "Missing Key"),
            (6, 4, "Missing Value"),
            (7, 1, "Missing Key"),
            (7, 4, "Missing Value"),
            (13, 4, "Missing Value"),
        ]
    if columns == (0,):
        assert [
            (group.kept_source_row, group.removed_source_rows, group.conflicting_columns)
            for group in report.duplicate_groups
        ] == [(2, (), (2,)), (4, (5,), ()), (10, (11,), ()), (13, (), (3,))]
        assert [group.retained_source_rows for group in report.duplicate_groups] == [
            (2, 3),
            (4,),
            (10,),
            (13, 14),
        ]

    destination = tmp_path / f"cleaned.{extension}"
    service.export(job, destination)
    assert fingerprint(source) == original
    reread = service.analyze(destination, plan, csv_options=options)
    assert reread.result.report.original_rows == len(source_rows)
    assert reread.result.report.removed_duplicates == 0
    assert reread.result.report.removed_empty_rows == 0
    assert reread.result.report.changed_text_cells == 0
    assert reread.result.report.conflicting_duplicate_groups == conflicts
    assert reread.result.table == job.result.table
    ids = [row[0].value for row in reread.result.table.rows[1:]]
    assert {"00123", "00007", "00042", "00000", "abc", "ABC"}.issubset(ids)
    assert all(value is None or isinstance(value, str) for value in ids)
    if extension == "xlsx":
        workbook = load_workbook(destination)
        try:
            assert workbook["Clients"].freeze_panes == "A2"
            assert workbook["Headers"].freeze_panes == "A2"
            assert workbook["Headers"]["B1"].value == workbook["Headers"]["D1"].value == "Name"
            assert workbook["Headers"]["C1"].value is None
        finally:
            workbook.close()
        with ZipFile(source) as before, ZipFile(destination) as after:
            assert after.testzip() is None
            assert before.namelist() == after.namelist()
            for part in before.namelist():
                if part != job.source.sheet_part:
                    assert before.read(part) == after.read(part)


def test_sample_headers_are_independent_keys(sample_generator, tmp_path):
    sample_generator(tmp_path)
    source = tmp_path / "advanced_cleaning_demo.xlsx"
    service = CleaningService(FileCleaningStorage())
    name_b = service.analyze(
        source, CleaningPlan(True, True, True, True, duplicate_columns=(1,)), sheet="Headers"
    )
    name_d = service.analyze(
        source, CleaningPlan(True, True, True, True, duplicate_columns=(3,)), sheet="Headers"
    )
    assert name_b.result.report.final_rows == 3
    assert name_b.result.report.removed_duplicates == 0
    assert name_b.result.report.conflicting_duplicate_groups == 1
    assert name_d.result.report.final_rows == 3
    assert name_d.result.report.removed_duplicates == 0
    for job in (name_b, name_d):
        assert [(i.source_row, i.column, i.kind) for i in job.result.report.issues] == [
            (1, 3, "Missing Value")
        ]


def test_sample_generator_never_overwrites_existing_files(sample_generator, tmp_path):
    created = sample_generator(tmp_path)
    assert len(created) == 2
    # Even unrelated contents under the reserved names must be preserved.
    created[0].write_bytes(b"local user file")
    before = {path: fingerprint(path) for path in created}
    assert sample_generator(tmp_path) == ()
    assert {path: fingerprint(path) for path in created} == before


def test_existing_stage4a_workbook_keeps_every_conflicting_variant(tmp_path):
    source = (
        Path(__file__).parents[1]
        / "examples"
        / "generated"
        / "stage4a"
        / "advanced_cleaning_demo.xlsx"
    )
    if not source.exists():
        pytest.skip("The local generated demonstration workbook is not checked into Git.")
    original = fingerprint(source)
    service = CleaningService(FileCleaningStorage())
    plan = CleaningPlan(True, True, True, True, duplicate_columns=(0,))
    job = service.analyze(source, plan, sheet="Clients")
    report = job.result.report
    assert (
        report.original_rows,
        report.final_rows,
        report.removed_empty_rows,
        report.removed_duplicates,
        report.conflicting_duplicate_groups,
        report.missing_key_rows,
        report.changed_text_cells,
        report.missing_values,
    ) == (15, 12, 1, 2, 2, 2, 5, 6)
    assert job.result.source_rows == (1, 2, 3, 4, 6, 7, 8, 9, 10, 13, 14, 15)
    assert report.duplicate_group_count == 4
    assert report.warnings
    rows = {
        source_row: tuple(cell.value for cell in row)
        for source_row, row in zip(job.result.source_rows, job.result.table.rows, strict=True)
    }
    assert rows[2] == ("00123", "Alex", "Astana", "alex@example.invalid")
    assert rows[3] == ("00123", "Alex", "Almaty", "alex@example.invalid")
    assert rows[13] == ("00999", "Lee", "Aktau", None)
    assert rows[14] == ("00999", "Lee", "Aktau", "lee@example.invalid")
    assert rows[4][0] == "00007" and rows[10][0] == "00042"
    assert rows[6][0] is None and rows[7][0] is None
    assert rows[8][0] == "abc" and rows[9][0] == "ABC"

    destination = tmp_path / "safe_cleaned.xlsx"
    service.export(job, destination)
    assert fingerprint(source) == original
    with ZipFile(source) as before, ZipFile(destination) as after:
        for part in before.namelist():
            if part != job.source.sheet_part:
                assert before.read(part) == after.read(part)
    repeated = service.analyze(destination, plan, sheet="Clients")
    assert repeated.result.table == job.result.table
    assert repeated.result.report.removed_duplicates == 0
    assert repeated.result.report.conflicting_duplicate_groups == 2
