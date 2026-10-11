import csv
import hashlib
from dataclasses import replace
from datetime import datetime
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table as ExcelTable

from excel_toolkit.application.cleaning import CleaningService
from excel_toolkit.application.models import CsvOptions
from excel_toolkit.infrastructure.files.cleaning import FileCleaningStorage
from excel_toolkit.modules.cleaner import CleaningError, CleaningPlan


@pytest.fixture
def cleaner():
    return CleaningService(FileCleaningStorage())


def save_book(path, rows):
    book = Workbook()
    for row in rows:
        book.active.append(row)
    book.save(path)
    book.close()


def fingerprint(path):
    return hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns


def rewrite_zip(path, transform):
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    transform(parts)
    with ZipFile(path, "w") as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)


def test_xlsx_export_preserves_types_styles_other_sheets_and_original(cleaner, tmp_path):
    path = tmp_path / "original.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.append(["ID", "Город", "n", "b", "date"])
    sheet.append([" 001234 ", " Алматы  Астана ", 0, False, datetime(2026, 1, 2)])
    sheet.append([None, " "])
    sheet.append(["001234", "Алматы  Астана", 0, False, datetime(2026, 1, 2)])
    sheet["A2"].font = Font(bold=True, color="FF0000")
    sheet["A2"].fill = PatternFill("solid", fgColor="FFFF00")
    sheet["C2"].number_format = "00000"
    sheet.row_dimensions[2].height = 27
    sheet.column_dimensions["A"].width = 20
    sheet["J1000"].fill = PatternFill("solid", fgColor="00FF00")
    book.create_sheet("Другой").append([" untouched ", "00007"])
    book.save(path)
    book.close()
    before = fingerprint(path)
    job = cleaner.analyze(path, CleaningPlan(True, True, True, True))
    assert job.result.report.original_rows == 4  # Remote style is not part of the table.
    assert job.result.report.final_rows == 2
    assert job.result.report.missing_values == 0
    destination = tmp_path / "clean.xlsx"
    cleaner.export(job, destination)
    assert fingerprint(path) == before
    result = load_workbook(destination)
    try:
        assert result.active["A2"].value == "001234"
        assert result.active["A2"].data_type == "s"
        assert result.active["A2"].font.bold
        assert result.active["A2"].fill.fgColor.rgb == "00FFFF00"
        assert result.active["C2"].value == 0
        assert result.active["C2"].number_format == "00000"
        assert result.active["D2"].value is False
        assert result.active["E2"].value == datetime(2026, 1, 2)
        assert result.active.row_dimensions[2].height == 27
        assert result.active.column_dimensions["A"].width == 20
        assert result.active["J998"].fill.fgColor.rgb == "0000FF00"
        assert result["Другой"]["A1"].value == " untouched "
    finally:
        result.close()
    with ZipFile(path) as original, ZipFile(destination) as output:
        assert original.namelist() == output.namelist()
        for name in original.namelist():
            if name != job.source.sheet_part:
                assert original.read(name) == output.read(name)
    again = cleaner.analyze(destination, CleaningPlan(True, True, True, True))
    assert again.result.report.changed_text_cells == 0
    assert again.result.report.removed_duplicates == 0
    assert again.result.report.removed_empty_rows == 0


@pytest.mark.parametrize("delimiter", [",", ";", "\t", "|"])
@pytest.mark.parametrize("encoding", ["utf-8-sig", "cp1251", "utf-16"])
def test_csv_round_trip_unicode_quotes_leading_zeroes(cleaner, tmp_path, delimiter, encoding):
    path = tmp_path / "source.csv"
    rows = [
        ["id", "текст"],
        ["001234", ' Алматы; |,\t "слово"\nновая строка '],
        ["=1+1", " @text "],
        ["", " "],
    ]
    with path.open("w", encoding=encoding, newline="") as stream:
        csv.writer(stream, delimiter=delimiter).writerows(rows)
    before = fingerprint(path)
    job = cleaner.analyze(
        path,
        CleaningPlan(trim_whitespace=True, remove_empty_rows=True),
        csv_options=CsvOptions(encoding, delimiter),
    )
    destination = tmp_path / "output.csv"
    cleaner.export(job, destination)
    with destination.open(encoding=encoding, newline="") as stream:
        assert list(csv.reader(stream, delimiter=delimiter)) == [
            ["id", "текст"],
            ["001234", rows[1][1].strip()],
            ["=1+1", "@text"],
        ]
    assert fingerprint(path) == before


def test_formula_on_other_sheet_blocks_deletion_but_allows_trim(cleaner, tmp_path):
    path = tmp_path / "formula.xlsx"
    book = Workbook()
    book.active.append([" h "])
    book.active.append([" a "])
    book.active.append(["a"])
    book.create_sheet("Summary")["A1"] = "=COUNTA(Sheet!A:A)"
    book.save(path)
    book.close()
    before = fingerprint(path)
    with pytest.raises(CleaningError, match="формулы"):
        cleaner.analyze(path, CleaningPlan(trim_whitespace=True, remove_duplicates=True))
    job = cleaner.analyze(path, CleaningPlan(trim_whitespace=True, detect_missing_values=True))
    output = tmp_path / "trimmed.xlsx"
    cleaner.export(job, output)
    with ZipFile(path) as a, ZipFile(output) as b:
        assert a.read("xl/worksheets/sheet2.xml") == b.read("xl/worksheets/sheet2.xml")
    formula_job = cleaner.analyze(path, CleaningPlan(True, True, True, True), sheet="Summary")
    assert formula_job.result.table.rows[0][0].data_type == "f"
    assert fingerprint(path) == before


def test_formulas_and_formula_like_text_are_preserved(cleaner, tmp_path):
    path = tmp_path / "formula.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.append([" h ", "f"])
    sheet.append([" =1+1 ", '= A1 & " x " '])
    sheet["A3"] = "=literal"
    sheet["A3"].data_type = "s"
    book.save(path)
    book.close()
    job = cleaner.analyze(path, CleaningPlan(trim_whitespace=True))
    output = tmp_path / "out.xlsx"
    cleaner.export(job, output)
    result = load_workbook(output)
    try:
        assert result.active["A2"].value == "=1+1"
        assert result.active["A2"].data_type == "s"
        assert result.active["B2"].value == '= A1 & " x " '
        assert result.active["B2"].data_type == "f"
        assert result.active["A3"].data_type == "s"
    finally:
        result.close()


@pytest.mark.parametrize(
    "structure",
    [
        "merge",
        "table",
        "name",
        "filter",
        "freeze",
        "validation",
        "comment",
        "hyperlink",
        "protection",
        "conditional",
    ],
)
def test_unsafe_workbooks_rejected_even_on_unselected_sheet(cleaner, tmp_path, structure):
    from openpyxl.comments import Comment
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.worksheet.datavalidation import DataValidation

    path = tmp_path / "unsafe.xlsx"
    book = Workbook()
    book.active.append([" h "])
    sheet = book.create_sheet("Complex")
    sheet.append(["a", "b"])
    sheet.append(["x", "y"])
    if structure == "merge":
        sheet.merge_cells("A1:B1")
    elif structure == "table":
        sheet.add_table(ExcelTable(displayName="TestTable", ref="A1:B2"))
    elif structure == "name":
        book.defined_names.add(DefinedName("TestName", attr_text="Complex!$A$1"))
    elif structure == "filter":
        sheet.auto_filter.ref = "A1:B2"
    elif structure == "freeze":
        sheet.freeze_panes = "A2"
    elif structure == "validation":
        sheet.add_data_validation(DataValidation(type="whole", formula1=0, formula2=10))
    elif structure == "comment":
        sheet["A1"].comment = Comment("text", "author")
    elif structure == "hyperlink":
        sheet["A1"].hyperlink = "https://example.com"
    elif structure == "protection":
        sheet.protection.sheet = True
    elif structure == "conditional":
        sheet.conditional_formatting.add("A2", CellIsRule(operator="equal", formula=["1"]))
    book.save(path)
    book.close()
    before = fingerprint(path)
    with pytest.raises(CleaningError, match="Неподдерживаем"):
        cleaner.analyze(path, CleaningPlan())
    assert fingerprint(path) == before


@pytest.mark.parametrize("suffix", [".csv", ".xlsx"])
@pytest.mark.parametrize("rows", [[], [["h", "name"]]])
def test_empty_header_only_files(cleaner, tmp_path, suffix, rows):
    path = tmp_path / f"source{suffix}"
    if suffix == ".xlsx":
        save_book(path, rows)
    else:
        with path.open("w", newline="") as stream:
            csv.writer(stream).writerows(rows)
    job = cleaner.analyze(path, CleaningPlan(True, True, True, True))
    assert job.result.report.final_rows == len(rows)
    cleaner.export(job, tmp_path / f"output{suffix}")


def test_no_overwrite_aliases_or_format_conversion(cleaner, tmp_path):
    path = tmp_path / "original.csv"
    path.write_text("h\n a \n", encoding="utf-8")
    job = cleaner.analyze(path, CleaningPlan(trim_whitespace=True))
    hardlink = tmp_path / "hardlink.csv"
    hardlink.hardlink_to(path)
    symlink = tmp_path / "symlink.csv"
    symlink.symlink_to(path)
    dangling = tmp_path / "dangling.csv"
    dangling.symlink_to(tmp_path / "missing.csv")
    existing = tmp_path / "existing.csv"
    existing.write_bytes(b"keep me")
    before = fingerprint(path)
    for destination in (path, hardlink, symlink, dangling, existing, tmp_path / "out.xlsx"):
        with pytest.raises(CleaningError):
            cleaner.export(job, destination)
    assert fingerprint(path) == before
    assert existing.read_bytes() == b"keep me"
    assert not list(tmp_path.glob(".cleaner-*"))


def test_write_failure_and_publication_race_leave_no_partial_file(cleaner, tmp_path, monkeypatch):
    import excel_toolkit.infrastructure.files.cleaning as storage

    path = tmp_path / "source.csv"
    path.write_text("h\n a \n", encoding="utf-8")
    job = cleaner.analyze(path, CleaningPlan(trim_whitespace=True))
    destination = tmp_path / "out.csv"

    def broken_write(source, result, stream):
        stream.write(b"partial")
        raise OSError("simulated failure")

    with monkeypatch.context() as patch:
        patch.setattr(FileCleaningStorage, "_write_csv", staticmethod(broken_write))
        with pytest.raises(CleaningError, match="simulated failure"):
            cleaner.export(job, destination)
    assert not destination.exists()
    assert not list(tmp_path.glob(".cleaner-*"))
    real_link = storage.os.link

    def raced_link(source, target):
        target.write_bytes(b"another writer")
        real_link(source, target)

    monkeypatch.setattr(storage.os, "link", raced_link)
    with pytest.raises(CleaningError):
        cleaner.export(job, destination)
    assert destination.read_bytes() == b"another writer"
    assert not list(tmp_path.glob(".cleaner-*"))


def test_export_uses_snapshot_if_original_changes(cleaner, tmp_path):
    path = tmp_path / "source.xlsx"
    save_book(path, [[" h "], ["001"]])
    job = cleaner.analyze(path, CleaningPlan(trim_whitespace=True))
    path.write_bytes(b"changed after analysis")
    cleaner.export(job, tmp_path / "out.xlsx")
    result = load_workbook(tmp_path / "out.xlsx")
    try:
        assert result.active["A1"].value == "h"
        assert result.active["A2"].value == "001"
    finally:
        result.close()
    assert path.read_bytes() == b"changed after analysis"


def test_reject_forged_result(cleaner, tmp_path):
    path = tmp_path / "source.csv"
    path.write_text("h\nx\n", encoding="utf-8")
    job = cleaner.analyze(path, CleaningPlan())
    bad = replace(job, result=replace(job.result, source_rows=(2, 1)))
    with pytest.raises(CleaningError):
        cleaner.export(bad, tmp_path / "out.csv")


@pytest.mark.parametrize(
    "kind",
    [
        "unknown_part",
        "extension",
        "array_formula",
        "rich_text",
        "external_relationship",
        "invalid_dimension",
    ],
)
def test_xml_edge_cases(cleaner, tmp_path, kind):
    path = tmp_path / "source.xlsx"
    save_book(path, [[" h "], ["value"]])

    def transform(parts):
        key = "xl/worksheets/sheet1.xml"
        xml = parts[key]
        if kind == "unknown_part":
            parts["xl/unknown.xml"] = b"<unknown/>"
        elif kind == "extension":
            xml = xml.replace(b"</worksheet>", b"<extLst/></worksheet>")
        elif kind == "array_formula":
            xml = xml.replace(b"<is><t>value</t></is>", b'<f t="array" ref="A2">SUM(A1)</f>')
        elif kind == "rich_text":
            xml = xml.replace(b"<t>value</t>", b"<r><t>value</t></r>")
        elif kind == "external_relationship":
            rel = "xl/_rels/workbook.xml.rels"
            parts[rel] = parts[rel].replace(b'Type="', b'TargetMode="External" Type="', 1)
        elif kind == "invalid_dimension":
            xml = xml.replace(b'ref="A1:A2"', b'ref="A1:A1"')
        parts[key] = xml

    rewrite_zip(path, transform)
    if kind == "invalid_dimension":
        assert cleaner.analyze(path, CleaningPlan()).result.report.original_rows == 2
    else:
        with pytest.raises(CleaningError):
            cleaner.analyze(path, CleaningPlan())


def test_shared_strings_trim_does_not_change_other_cells(cleaner, tmp_path):
    path = tmp_path / "shared.xlsx"
    save_book(path, [[" h "], [" shared "]])

    def transform(parts):
        parts["xl/sharedStrings.xml"] = (
            b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            b'count="1" uniqueCount="1"><si><t xml:space="preserve"> shared </t></si></sst>'
        )
        rel = "xl/_rels/workbook.xml.rels"
        parts[rel] = parts[rel].replace(
            b"</Relationships>",
            (
                b'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                b'relationships/sharedStrings" Target="sharedStrings.xml" Id="rIdStrings"/>'
                b"</Relationships>"
            ),
        )
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
            b"</Types>",
            (
                b'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.'
                b'openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
            ),
        )
        key = "xl/worksheets/sheet1.xml"
        parts[key] = parts[key].replace(
            b'<c r="A2" t="inlineStr"><is><t xml:space="preserve"> shared </t></is></c>',
            b'<c r="A2" t="s"><v>0</v></c>',
        )

    rewrite_zip(path, transform)
    job = cleaner.analyze(path, CleaningPlan(trim_whitespace=True))
    assert job.result.table.rows[1][0].value == "shared"
    output = tmp_path / "out.xlsx"
    cleaner.export(job, output)
    with ZipFile(path) as a, ZipFile(output) as b:
        assert a.read("xl/sharedStrings.xml") == b.read("xl/sharedStrings.xml")
    assert cleaner.analyze(output, CleaningPlan()).result.table == job.result.table


def test_tens_of_thousands_of_rows(cleaner, tmp_path):
    path = tmp_path / "large.csv"
    path.write_text("id\n" + " 001 \n" * 30_000, encoding="utf-8")
    job = cleaner.analyze(path, CleaningPlan(trim_whitespace=True, remove_duplicates=True))
    assert job.result.report.original_rows == 30_001
    assert job.result.report.removed_duplicates == 29_999
    assert job.result.preview(1).rows[0][0].value == "id"


@pytest.mark.parametrize(
    "filename,content",
    [
        ("bad.xlsx", b"not ZIP"),
        ("empty.xlsx", b""),
        ("bad.csv", b'"unfinished'),
        ("bad.xls", b"old format"),
        ("encoding.csv", b"\xff"),
    ],
)
def test_bad_inputs(cleaner, tmp_path, filename, content):
    path = tmp_path / filename
    path.write_bytes(content)
    with pytest.raises(CleaningError):
        cleaner.analyze(path, CleaningPlan(), csv_options=CsvOptions(delimiter=","))


def test_internal_carriage_returns_survive_xml_rewrite(cleaner, tmp_path):
    path = tmp_path / "cr.xlsx"
    save_book(path, [["h"], [" a\rb\r\nc \n"], ["untouched\rvalue"], ['= "a\rb"']])

    # The stdlib openpyxl writer itself emits literal CR; construct valid preserved CR input.
    def preserve_cr(parts):
        key = "xl/worksheets/sheet1.xml"
        parts[key] = parts[key].replace(b"\r", b"&#13;")

    rewrite_zip(path, preserve_cr)
    job = cleaner.analyze(path, CleaningPlan(trim_whitespace=True))
    output = tmp_path / "out.xlsx"
    cleaner.export(job, output)
    result = load_workbook(output)
    try:
        assert result.active["A2"].value == "a\rb\r\nc"
        assert result.active["A3"].value == "untouched\rvalue"
        assert result.active["A4"].value == '= "a\rb"'
    finally:
        result.close()
    assert job.result.report.warnings


@pytest.mark.parametrize("kind", ["dtd", "utf16_dtd", "escape", "group", "book_protection"])
def test_additional_safety_rejections(cleaner, tmp_path, kind):
    path = tmp_path / "unsafe.xlsx"
    book = Workbook()
    book.active.append([" h "])
    if kind == "group":
        book.active.row_dimensions.group(1, 3)
    if kind == "book_protection":
        book.security.lockStructure = True
    book.save(path)
    book.close()

    def transform(parts):
        key = "xl/worksheets/sheet1.xml"
        if kind in {"dtd", "utf16_dtd"}:
            data = b'<!DOCTYPE worksheet [<!ENTITY value "test">]>' + parts[key]
            parts[key] = data if kind == "dtd" else data.decode().encode("utf-16")
        elif kind == "escape":
            parts[key] = parts[key].replace(b" h ", b" _x005F_x000A_ ")

    rewrite_zip(path, transform)
    with pytest.raises(CleaningError):
        cleaner.analyze(path, CleaningPlan())


def test_xlsx_bulk_row_deletion_and_styles_move_with_rows(cleaner, tmp_path):
    path = tmp_path / "bulk.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.append(["id"])
    for _ in range(20_000):
        sheet.append(["001"])
    sheet.append(["last"])
    sheet["A20002"].font = Font(bold=True)
    sheet.row_dimensions[20002].height = 40
    book.save(path)
    book.close()
    job = cleaner.analyze(path, CleaningPlan(remove_duplicates=True))
    assert job.result.report.removed_duplicates == 19_999
    cleaner.export(job, tmp_path / "out.xlsx")
    result = load_workbook(tmp_path / "out.xlsx")
    try:
        assert result.active.max_row == 3
        assert result.active["A3"].value == "last"
        assert result.active["A3"].font.bold
        assert result.active.row_dimensions[3].height == 40
    finally:
        result.close()


def test_xlsx_noop_export_leaves_all_member_payloads_identical(cleaner, tmp_path):
    path = tmp_path / "source.xlsx"
    save_book(path, [[" h "], ["001"]])
    job = cleaner.analyze(path, CleaningPlan())
    output = tmp_path / "out.xlsx"
    cleaner.export(job, output)
    with ZipFile(path) as a, ZipFile(output) as b:
        assert {n: a.read(n) for n in a.namelist()} == {n: b.read(n) for n in b.namelist()}


def test_reading_a_small_preview_still_uses_the_existing_bounded_reader(service, tmp_path):
    # A malformed record beyond the preview must not force full Cleaner loading.
    path = tmp_path / "preview.csv"
    path.write_text("h\n" * 101 + '"unfinished', encoding="utf-8")
    assert len(service.load(path, options=CsvOptions(delimiter=",")).rows) == 100


def test_size_limits_reject_instead_of_truncating(cleaner, tmp_path, monkeypatch):
    import excel_toolkit.infrastructure.files.cleaning as storage
    import excel_toolkit.infrastructure.files.cleaning_xlsx as xlsx

    csv_path = tmp_path / "source.csv"
    csv_path.write_text("a,b\nx,y\n", encoding="utf-8")
    monkeypatch.setattr(storage, "MAX_CELLS", 3)
    with pytest.raises(CleaningError, match="лимит"):
        cleaner.analyze(csv_path, CleaningPlan())
    path = tmp_path / "source.xlsx"
    save_book(path, [["a", "b"], ["x", "y"]])
    monkeypatch.setattr(xlsx, "MAX_CELLS", 3)
    with pytest.raises(CleaningError, match="лимит"):
        cleaner.analyze(path, CleaningPlan())
    monkeypatch.setattr(storage, "MAX_ARCHIVE_BYTES", 1)
    with pytest.raises(CleaningError, match="лимит"):
        cleaner.analyze(path, CleaningPlan())


@pytest.mark.parametrize("kind", ["excel_leap_day", "rounded_time", "rounded_number"])
def test_reader_normalization_cannot_create_false_duplicates(cleaner, tmp_path, kind):
    path = tmp_path / "ambiguous.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.append(["h"])
    if kind == "excel_leap_day":
        first, second, fmt = "59", "60", "yyyy-mm-dd"
    elif kind == "rounded_time":
        first, second, fmt = "0.000010000001", "0.000010000002", "hh:mm:ss.000"
    else:
        first, second, fmt = "1.00000000000000001", "1.00000000000000002", "General"
    sheet.append([123])
    sheet.append([456])
    sheet["A2"].number_format = fmt
    sheet["A3"].number_format = fmt
    book.save(path)
    book.close()

    def transform(parts):
        key = "xl/worksheets/sheet1.xml"
        parts[key] = (
            parts[key]
            .replace(b"<v>123</v>", f"<v>{first}</v>".encode())
            .replace(b"<v>456</v>", f"<v>{second}</v>".encode())
        )

    rewrite_zip(path, transform)
    with pytest.raises(CleaningError, match="неразличимы"):
        cleaner.analyze(path, CleaningPlan(remove_duplicates=True))
