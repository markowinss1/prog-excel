import csv
import hashlib
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_toolkit.application.models import CsvOptions, PreviewCancelled, PreviewError
from excel_toolkit.application.preview import PreviewService


def test_xlsx_types_formulas_and_leading_zeroes(service: PreviewService, xlsx_file: Path) -> None:
    preview = service.load(xlsx_file)
    assert preview.sheet_names == ("Клиенты", "Заказы", "Пустой")
    assert preview.selected_sheet == "Клиенты"
    row = preview.rows[1]
    assert [cell.value for cell in row] == [
        "00123",
        "Алматы",
        12,
        "=C2*2",
        datetime(2026, 1, 2),
        True,
    ]
    assert row[0].data_type == "s"
    assert row[2].data_type == "n"
    assert row[2].number_format == "00000"
    assert row[3].data_type == "f"
    assert preview.more_rows is False


def test_sheet_selection(service: PreviewService, xlsx_file: Path) -> None:
    preview = service.load(xlsx_file, sheet="Заказы")
    assert preview.selected_sheet == "Заказы"
    assert preview.rows[0][0].value == "00007"
    assert service.load(xlsx_file, sheet="Пустой").rows == ()
    with pytest.raises(PreviewError, match="не найден"):
        service.load(xlsx_file, sheet="Несуществующий")


@pytest.mark.parametrize("delimiter", [",", ";", "\t", "|"])
@pytest.mark.parametrize("encoding", ["utf-8", "utf-8-sig"])
def test_csv_unicode_and_delimiters(
    service: PreviewService, tmp_path: Path, delimiter: str, encoding: str
) -> None:
    path = tmp_path / "данные.csv"
    with path.open("w", encoding=encoding, newline="") as stream:
        writer = csv.writer(stream, delimiter=delimiter)
        writer.writerows([["Код", "Город"], ["00012", "Алматы"], ["00100", "Астана"]])
    preview = service.load(path)
    assert preview.rows[0][0].value == "Код"
    assert preview.rows[1][0].value == "00012"
    assert preview.rows[1][1].value == "Алматы"
    assert preview.csv_options.delimiter == delimiter


@pytest.mark.parametrize(
    ("encoding", "value"), [("cp1251", "Привет"), ("cp1252", "café"), ("utf-16", "Қазақ")]
)
def test_explicit_alternative_encoding(
    service: PreviewService, tmp_path: Path, encoding: str, value: str
) -> None:
    path = tmp_path / "legacy.csv"
    path.write_text(f"id;name\n0001;{value}\n", encoding=encoding)
    with pytest.raises(PreviewError, match="кодировк"):
        service.load(path)
    preview = service.load(path, options=CsvOptions(encoding=encoding))
    assert preview.rows[1][1].value == value


def test_csv_quoted_multiline_and_ragged_rows(service: PreviewService, tmp_path: Path) -> None:
    path = tmp_path / "quoted.csv"
    path.write_text('001,"hello, world\nстрока"\n\n002,x,extra\n', encoding="utf-8")
    preview = service.load(path, options=CsvOptions(delimiter=","))
    assert preview.rows[0][1].value == "hello, world\nстрока"
    assert preview.rows[1] == ()
    assert preview.column_count == 3


@pytest.mark.parametrize("suffix", [".csv", ".xlsx"])
@pytest.mark.parametrize("count", [0, 100, 105])
def test_preview_row_limit(
    service: PreviewService, tmp_path: Path, suffix: str, count: int
) -> None:
    path = tmp_path / f"rows{suffix}"
    if suffix == ".csv":
        path.write_text("row\n" * count, encoding="utf-8")
    else:
        workbook = Workbook()
        for _ in range(count):
            workbook.active.append(["row"])
        workbook.save(path)
        workbook.close()
    preview = service.load(path)
    assert len(preview.rows) == min(count, 100)
    assert preview.more_rows == (count > 100)


@pytest.mark.parametrize("suffix", [".csv", ".xlsx"])
def test_preview_column_limit(service: PreviewService, tmp_path: Path, suffix: str) -> None:
    path = tmp_path / f"wide{suffix}"
    if suffix == ".csv":
        path.write_text(",".join(["x"] * 250), encoding="utf-8")
    else:
        workbook = Workbook()
        workbook.active.append(["x"] * 250)
        workbook.save(path)
        workbook.close()
    preview = service.load(path)
    assert preview.column_count == 200
    assert preview.more_columns is True


@pytest.mark.parametrize("suffix", [".xlsx", ".xls", ".txt"])
def test_damaged_or_unsupported_file(service: PreviewService, tmp_path: Path, suffix: str) -> None:
    path = tmp_path / f"bad{suffix}"
    path.write_bytes(b"not a workbook")
    with pytest.raises(PreviewError):
        service.load(path)


def test_missing_file_and_malformed_csv(service: PreviewService, tmp_path: Path) -> None:
    with pytest.raises(PreviewError, match="открыть"):
        service.load(tmp_path / "missing.xlsx")
    path = tmp_path / "bad.csv"
    path.write_text('a,"unfinished\n', encoding="utf-8")
    with pytest.raises(PreviewError, match="прочитать"):
        service.load(path, options=CsvOptions(delimiter=","))


def test_sources_unchanged(service: PreviewService, xlsx_file: Path, tmp_path: Path) -> None:
    csv_file = tmp_path / "source.csv"
    csv_file.write_text("Код;Город\n001;Алматы\n", encoding="utf-8-sig")
    for path in (xlsx_file, csv_file):
        before = hashlib.sha256(path.read_bytes()).digest(), path.stat().st_mtime_ns
        service.load(path)
        after = hashlib.sha256(path.read_bytes()).digest(), path.stat().st_mtime_ns
        assert before == after


def test_cooperative_cancellation(service: PreviewService, xlsx_file: Path) -> None:
    calls = 0

    def cancelled() -> bool:
        nonlocal calls
        calls += 1
        return calls >= 4  # Cancel inside worksheet iteration, after the workbook opens.

    with pytest.raises(PreviewCancelled):
        service.load(xlsx_file, cancelled=cancelled)
    assert service.load(xlsx_file).rows
