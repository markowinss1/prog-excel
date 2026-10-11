"""Column metadata for selecting keys remains independent of bounded cell display."""

import csv

import pytest
from openpyxl import Workbook

from excel_toolkit.application.cleaning import CleaningService
from excel_toolkit.application.models import CsvOptions
from excel_toolkit.infrastructure.files.cleaning import FileCleaningStorage
from excel_toolkit.modules.cleaner import CleaningPlan


@pytest.mark.parametrize("suffix", [".csv", ".xlsx"])
def test_full_headers_available_outside_display_limit(service, tmp_path, suffix):
    path = tmp_path / f"wide{suffix}"
    headers = ["ID", "ID", ""] + [f"Field {i}" for i in range(3, 250)]
    if suffix == ".csv":
        with path.open("x", newline="") as stream:
            csv.writer(stream).writerow(headers)
    else:
        book = Workbook()
        book.active.append(headers)
        book.save(path)
        book.close()
    preview = service.load(path, options=CsvOptions(delimiter=","))
    assert preview.column_count == 200
    assert preview.available_column_count == 250
    assert len(preview.column_headers) == 250
    assert preview.column_headers[-1].value == "Field 249"
    assert [c.value for c in preview.column_headers[:2]] == ["ID", "ID"]
    assert preview.column_headers[2].value in ("", None)


def test_full_analysis_adds_headerless_columns_beyond_sample(service, tmp_path):
    path = tmp_path / "ragged.csv"
    with path.open("x", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(["ID"])
        writer.writerows([[str(i)] for i in range(100)])
        writer.writerow(["late", "new", "", "column"])
    preview = service.load(path, options=CsvOptions(delimiter=","))
    assert preview.available_column_count == 1
    assert preview.more_rows
    job = CleaningService(FileCleaningStorage()).analyze(
        path, CleaningPlan(), csv_options=CsvOptions(delimiter=",")
    )
    original, cleaned = job.previews()
    assert original.available_column_count == cleaned.available_column_count == 4
    assert len(original.rows) == len(cleaned.rows) == 100
    assert len(original.column_headers) == 1
