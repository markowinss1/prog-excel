from dataclasses import replace
from pathlib import Path

import pytest

from excel_toolkit.application.cleaning import CleaningService, CleaningSource
from excel_toolkit.application.models import Cell, CsvOptions
from excel_toolkit.modules.cleaner import CleaningError, CleaningPlan, Table


def test_application_passes_options_validates_and_exports_original_job():
    calls = []
    table = Table(((Cell("id"),), (Cell(" 001 "),), (Cell("001"),)))
    source = CleaningSource(Path("source.csv"), table, csv_options=CsvOptions("cp1251", ";"))

    class Storage:
        def read(self, path, sheet, options):
            calls.append(("read", path, sheet, options))
            return source

        def validate(self, original, result):
            calls.append(("validate", original, result))

        def write(self, job, destination):
            calls.append(("write", job, destination))

    service = CleaningService(Storage())
    options = CsvOptions("cp1251", ";")
    job = service.analyze("source.csv", CleaningPlan(True, True, True, True), csv_options=options)
    assert calls[0] == ("read", Path("source.csv"), None, options)
    assert calls[1] == ("validate", source, job.result)
    assert len(calls) == 2  # Analysis never writes.
    assert source.table.rows[1][0].value == " 001 "
    assert job.result.report.removed_duplicates == 1
    service.export(job, "new.csv")
    assert calls[-1] == ("write", job, Path("new.csv"))


def test_application_validation_failure_prevents_a_successful_result():
    source = CleaningSource(Path("source.xlsx"), Table(((Cell("h"),),)))

    class Storage:
        def read(self, *args):
            return source

        def validate(self, *args):
            raise CleaningError("Unsafe workbook")

        def write(self, *args):
            pytest.fail("Analysis must not export")

    with pytest.raises(CleaningError, match="Unsafe workbook"):
        CleaningService(Storage()).analyze("source.xlsx", CleaningPlan())


def test_previews_bound_both_dimensions_and_share_exact_analyzed_snapshot():
    # Wide data beyond the first 100 rows must still set more_columns.
    rows = [(Cell(" h "),)] * 105 + [tuple(Cell(" 001 ") for _ in range(205))]
    table = Table(tuple(rows))
    source = CleaningSource(Path("source.xlsx"), table, ("Data", "Other"), "Data")

    class Storage:
        def read(self, *args):
            return source

        def validate(self, *args):
            pass

    job = CleaningService(Storage()).analyze("source.xlsx", CleaningPlan(trim_whitespace=True))
    original, cleaned = job.previews()
    assert len(original.rows) == len(cleaned.rows) == 100
    assert original.more_rows and cleaned.more_rows
    assert original.more_columns and cleaned.more_columns
    assert original.sheet_names == cleaned.sheet_names == ("Data", "Other")
    assert original.selected_sheet == cleaned.selected_sheet == "Data"
    assert original.rows[0][0] is source.table.rows[0][0]
    assert original.rows[0][0].value == " h "
    assert cleaned.rows[0][0].value == "h"
    assert job.result.report.original_rows == 106
    assert len(job.result.table.rows[-1]) == 205
    wide_source = replace(source, table=Table((rows[-1],)))
    wide_job = replace(job, source=wide_source)
    assert wide_job.previews()[0].column_count == 200


def test_empty_previews_and_formula_warning():
    source = CleaningSource(Path("source.xlsx"), Table(()), contains_formulas=True)

    class Storage:
        def read(self, *args):
            return source

        def validate(self, *args):
            pass

    job = CleaningService(Storage()).analyze("source.xlsx", CleaningPlan())
    original, cleaned = job.previews()
    assert original.rows == cleaned.rows == ()
    assert not original.more_rows and not original.more_columns
    assert "кэш" in job.result.report.warnings[0]
