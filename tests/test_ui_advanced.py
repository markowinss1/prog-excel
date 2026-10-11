"""Selected duplicate keys, bounded reports and safe conflict exports through real services."""

import csv
import os
import time
from threading import Event

import pytest
from openpyxl import Workbook
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QCheckBox, QFileDialog, QMessageBox

from excel_toolkit.application.cleaning import CleaningService
from excel_toolkit.application.models import Cell
from excel_toolkit.infrastructure.files.cleaning import FileCleaningStorage
from excel_toolkit.ui.main_window import MainWindow


@pytest.fixture(scope="module")
def app():
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    os.environ["QT_QPA_PLATFORMTHEME"] = ""
    return QApplication.instance() or QApplication([])


def wait_until(app, predicate):
    deadline = time.monotonic() + 8
    while not predicate():
        app.processEvents()
        if time.monotonic() > deadline:
            pytest.fail("Timed out waiting for advanced Cleaner GUI")
        time.sleep(0.002)
    app.processEvents()


@pytest.fixture
def window(app, service):
    widget = MainWindow(service, CleaningService(FileCleaningStorage()))
    widget.show()
    yield widget
    widget.close()
    wait_until(app, lambda: widget._worker is None)
    widget.close()
    widget.deleteLater()
    app.processEvents()


def load(app, window, path):
    window.load_file(path)
    wait_until(app, lambda: window._worker is None)
    assert not window.errors.toPlainText()


def analyze(app, window):
    window.analyze_button.click()
    wait_until(app, lambda: window._worker is None)


def select(window, *columns):
    window.duplicate_mode_combo.setCurrentIndex(1)
    for column in columns:
        window.column_model.setData(
            window.column_model.index(column, 0),
            Qt.CheckState.Checked,
            Qt.ItemDataRole.CheckStateRole,
        )


def test_duplicate_value_display_distinguishes_python_and_excel_types():
    numeric = MainWindow._cell_description(Cell(1, "n"))
    text_metadata = MainWindow._cell_description(Cell(1, "s"))
    assert numeric == "1 (int) [cell:n]"
    assert text_metadata == "1 (int) [cell:s]"


@pytest.fixture
def conflict_csv(tmp_path):
    path = tmp_path / "clients.csv"
    path.write_text(
        "ID,City,Email\n00123,Almaty,a@example.test\n00123,Astana,a@example.test\n"
        "00123,Almaty,a@example.test\n,Unknown,\n,Unknown,\n",
        encoding="utf-8",
    )
    return path


def test_column_labels_and_stable_indices_for_duplicate_and_empty_headers(app, window, tmp_path):
    path = tmp_path / "headers.csv"
    path.write_text("Name,Name,\na,x,1\na,y,2\n", encoding="utf-8")
    load(app, window, path)
    assert window.duplicate_mode_combo.currentText() == "All Columns"
    assert not window.column_list.isVisible()
    assert [window.column_model.data(window.column_model.index(i, 0)) for i in range(3)] == [
        "A — Name",
        "B — Name",
        "Column C",
    ]
    select(window, 1, 2)
    assert window.column_list.isVisible()
    assert window.cleaning_plan().duplicate_columns == (1, 2)
    analyze(app, window)
    assert window._job.result.report.duplicate_columns == (1, 2)
    assert window.column_model.selected_columns == (1, 2)


def test_empty_selection_blocks_analysis_but_disabled_duplicates_do_not(app, window, conflict_csv):
    load(app, window, conflict_csv)
    select(window)
    analyze(app, window)
    assert "выберите хотя бы один" in window.errors.toPlainText()
    assert window._worker is None
    assert window._job is None
    assert not window.save_button.isEnabled()
    window.operation_checks["remove_duplicates"].setChecked(False)
    analyze(app, window)
    assert window._job is not None
    assert not window.errors.toPlainText()
    assert not window.duplicate_mode_combo.isEnabled()


@pytest.mark.parametrize("change", ["mode", "keys"])
def test_key_changes_invalidate_every_detail_and_export(
    app, window, conflict_csv, tmp_path, change
):
    load(app, window, conflict_csv)
    select(window, 0)
    analyze(app, window)
    assert window._job.result.report.conflicting_duplicate_groups == 1
    if change == "mode":
        window.duplicate_mode_combo.setCurrentIndex(0)
    else:
        select(window, 1)
    assert window._job is None
    assert not window.save_button.isEnabled()
    assert window.issues_model.rowCount() == 0
    assert not window.duplicate_details.toPlainText()
    assert not window.preview_tabs.isTabEnabled(2)
    assert not window.preview_tabs.isTabEnabled(3)
    output = tmp_path / "stale.csv"
    window.save_cleaned_file(output)
    assert not output.exists()


def test_file_and_sheet_changes_refresh_keys(app, window, tmp_path):
    path = tmp_path / "sheets.xlsx"
    book = Workbook()
    book.active.append(["ID", "City"])
    book.active.append(["001", "Almaty"])
    second = book.create_sheet("Other")
    second.append([None, "Other ID", "Other ID"])
    second.append(["x", "002", "003"])
    book.save(path)
    book.close()
    load(app, window, path)
    select(window, 0)
    analyze(app, window)
    window.sheet_combo.setCurrentText("Other")
    assert window.column_model.rowCount() == 0
    assert not window.save_button.isEnabled()
    window.sheet_combo.textActivated.emit("Other")
    wait_until(app, lambda: window._worker is None)
    assert window.column_model.rowCount() == 3
    assert not window.column_model.selected_columns
    assert window.column_model.column_label(0) == "Column A"
    assert window.column_model.column_label(2) == "C — Other ID"
    another = tmp_path / "other.csv"
    another.write_text("Email\na@example.test\n", encoding="utf-8")
    load(app, window, another)
    assert window.column_model.rowCount() == 1
    assert window.column_model.column_label(0) == "A — Email"


def test_all_columns_vs_selected_details_and_source_issue_rows(app, window, conflict_csv):
    load(app, window, conflict_csv)
    analyze(app, window)
    report = window._job.result.report
    assert report.removed_duplicates == 2
    assert report.conflicting_duplicate_groups == 0
    select(window, 0)
    analyze(app, window)
    report = window._job.result.report
    assert report.removed_duplicates == 1
    assert report.missing_key_rows == 2
    assert report.conflicting_duplicate_groups == 1
    assert window._job.result.source_rows == (1, 2, 3, 5, 6)
    details = window.duplicate_details.toPlainText()
    for snippet in (
        "Comparison: Selected Columns — A — ID",
        "Duplicate Groups (repeated keys): 1",
        "Removed Duplicates (identical rows): 1",
        "Status: Conflict — Needs Review",
        "Key: A — ID = '00123' (str)",
        "Keep: Row 2",
        "Keep: Row 3",
        "Remove: Row 4",
        "Retained: 2 rows.",
        "Automatically Removed Identical Rows: 1.",
        "WARNING:",
        "B — City",
        "Almaty",
        "Astana",
        "Missing Key Rows: 2",
    ):
        assert snippet in details
    assert "Remove: Row 3" not in details
    assert "Конфликты не разрешены" in window.status_label.text()
    issues = {
        tuple(window.issues_model.data(window.issues_model.index(row, col)) for col in range(3))
        for row in range(window.issues_model.rowCount())
    }
    assert ("5", "A — ID", "Missing Key") in issues
    assert ("6", "A — ID", "Missing Key") in issues
    assert ("5", "C — Email", "Missing Value") in issues
    assert ("6", "C — Email", "Missing Value") in issues
    job = window._job
    window.preview_tabs.setCurrentIndex(2)
    window.preview_tabs.setCurrentIndex(3)
    assert window._worker is None
    assert window._job is job


def test_issues_show_original_rows_after_empty_and_duplicate_removal(app, window, tmp_path):
    path = tmp_path / "source-rows.csv"
    path.write_text("ID,Email\n,\n001,a\n001,a\n002,\n,missing-key\n", encoding="utf-8")
    load(app, window, path)
    select(window, 0)
    analyze(app, window)
    assert window._job.result.source_rows == (1, 3, 5, 6)
    actual = {
        tuple(window.issues_model.data(window.issues_model.index(row, col)) for col in range(3))
        for row in range(window.issues_model.rowCount())
    }
    assert ("5", "B — Email", "Missing Value") in actual
    assert ("6", "A — ID", "Missing Key") in actual
    assert all(row not in ("2", "4") for row, _, _ in actual)


@pytest.mark.parametrize("reanalyze", [False, True])
def test_conflicting_export_preserves_variants_without_consent_each_time(
    app, window, conflict_csv, tmp_path, monkeypatch, reanalyze
):
    load(app, window, conflict_csv)
    select(window, 0)
    analyze(app, window)
    before = conflict_csv.read_bytes()

    def unexpected(*args):
        pytest.fail("Safe conflict export must not offer a blanket Yes/No deletion confirmation")

    monkeypatch.setattr(QMessageBox, "warning", unexpected)
    for index in range(2):
        if reanalyze:
            analyze(app, window)
        output = tmp_path / f"out-{index}.csv"
        window.save_cleaned_file(output)
        wait_until(app, lambda: window._worker is None)
        assert output.exists()
        with output.open(encoding="utf-8-sig", newline="") as stream:
            assert list(csv.reader(stream)) == [
                ["ID", "City", "Email"],
                ["00123", "Almaty", "a@example.test"],
                ["00123", "Astana", "a@example.test"],
                ["", "Unknown", ""],
                ["", "Unknown", ""],
            ]
        assert "конфликты не разрешены. Конфликтующих групп: 1." in window.saved_label.text()
        assert "Конфликты не разрешены" in window.status_label.text()
        assert "Keep: Row 3" in window.duplicate_details.toPlainText()
        assert window._job.result.report.removed_duplicates == 1
    assert conflict_csv.read_bytes() == before
    assert window.save_button.isEnabled()


def test_safe_export_has_no_conflict_prompt(app, window, conflict_csv, tmp_path, monkeypatch):
    load(app, window, conflict_csv)
    analyze(app, window)

    def unexpected(*args):
        pytest.fail("All Columns export has no conflicting groups")

    monkeypatch.setattr(QMessageBox, "warning", unexpected)
    output = tmp_path / "safe.csv"
    window.save_cleaned_file(output)
    wait_until(app, lambda: window._worker is None)
    assert output.exists()


def test_export_warning_uses_analyzed_snapshot_if_settings_change_during_export(
    app, window, conflict_csv, tmp_path
):
    started, release = Event(), Event()

    class SlowExport(CleaningService):
        def export(self, *args, **kwargs):
            started.set()
            assert release.wait(6)
            return super().export(*args, **kwargs)

    window._cleaning_service = SlowExport(FileCleaningStorage())
    load(app, window, conflict_csv)
    select(window, 0)
    analyze(app, window)
    output = tmp_path / "snapshot.csv"
    window.save_cleaned_file(output)
    try:
        wait_until(app, started.is_set)
        # Disabled controls can still receive programmatic changes during background work.
        window.duplicate_mode_combo.setCurrentIndex(0)
        assert window._job is None
    finally:
        release.set()
        wait_until(app, lambda: window._worker is None)
    assert "конфликты не разрешены. Конфликтующих групп: 1." in window.saved_label.text()
    assert not window.save_button.isEnabled()
    with output.open(encoding="utf-8-sig", newline="") as stream:
        exported = list(csv.reader(stream))
    assert exported[1:3] == [
        ["00123", "Almaty", "a@example.test"],
        ["00123", "Astana", "a@example.test"],
    ]


@pytest.mark.parametrize("change", ["mode", "keys"])
def test_dialog_setting_change_cannot_export_stale_job(
    app, window, conflict_csv, tmp_path, monkeypatch, change
):
    load(app, window, conflict_csv)
    select(window, 0)
    analyze(app, window)
    output = tmp_path / "stale.csv"

    def changed_destination(*args, **kwargs):
        if change == "mode":
            window.duplicate_mode_combo.setCurrentIndex(0)
        else:
            select(window, 1)
        return str(output), ""

    monkeypatch.setattr(QFileDialog, "getSaveFileName", changed_destination)
    window.save_button.click()
    assert window._worker is None
    assert not output.exists()
    assert window._job is None
    assert not window.save_button.isEnabled()


def test_bounded_groups_samples_issues_and_full_conflict_warning(
    app, window, tmp_path, monkeypatch
):
    path = tmp_path / "many.csv"
    rows = ["ID,City"]
    for index in range(105):
        rows.extend([f"{index},Almaty", f"{index},Astana"])
    rows.extend(["0,Almaty"] * 24)
    rows.extend([",Missing ID"] * 110)
    path.write_text("\n".join(rows) + "\n", encoding="utf-8")
    load(app, window, path)
    select(window, 0)
    analyze(app, window)
    details = window.duplicate_details.toPlainText()
    assert "Showing first 100 of 105 duplicate groups." in details
    assert "Showing first 20 of 24 removed rows." in details
    assert "Conflicting Groups: 105" in details
    assert window.issues_model.rowCount() == 100
    assert "Showing first 100 of 110 issues." in window.issues_label.text()
    assert "Duplicate Group 101\n" not in details

    def unexpected(*args):
        pytest.fail("Unresolved conflicts are retained without a deletion confirmation")

    monkeypatch.setattr(QMessageBox, "warning", unexpected)
    output = tmp_path / "preserved.csv"
    window.save_cleaned_file(output)
    wait_until(app, lambda: window._worker is None)
    assert output.exists()
    assert "конфликты не разрешены. Конфликтующих групп: 105." in window.saved_label.text()
    with output.open(encoding="utf-8-sig", newline="") as stream:
        assert len(list(csv.reader(stream))) == 321


def test_retained_conflict_samples_are_bounded_without_dropping_export_rows(app, window, tmp_path):
    path = tmp_path / "variants.csv"
    rows = [["ID", "City"], *[["001", f"City {index}"] for index in range(26)]]
    with path.open("w", encoding="utf-8", newline="") as stream:
        csv.writer(stream).writerows(rows)
    load(app, window, path)
    select(window, 0)
    analyze(app, window)
    details = window.duplicate_details.toPlainText()
    assert "Status: Conflict — Needs Review" in details
    assert "Showing first 21 of 26 retained rows." in details
    assert "Automatically Removed Identical Rows: 0." in details
    assert "Keep: Row 22 |" in details
    assert "Keep: Row 23 |" not in details
    assert window._job.result.source_rows == tuple(range(1, 28))
    output = tmp_path / "all-variants.csv"
    window.save_cleaned_file(output)
    wait_until(app, lambda: window._worker is None)
    with output.open(encoding="utf-8-sig", newline="") as stream:
        assert list(csv.reader(stream)) == rows


def test_full_width_column_selection_uses_model_not_individual_widgets(app, window, tmp_path):
    path = tmp_path / "wide.csv"
    headers = [f"Header {index}" for index in range(1500)]
    row = ["x"] * 1500
    row[1200] = "00123"
    with path.open("w", encoding="utf-8", newline="") as stream:
        csv.writer(stream).writerows([headers, row, row])
    load(app, window, path)
    assert window.model.columnCount() == 200
    assert window.column_model.rowCount() == 1500
    assert len(window.findChildren(QCheckBox)) == 4
    select(window, 1200)
    analyze(app, window)
    assert window._job.result.report.duplicate_columns == (1200,)
    assert window._job.result.report.removed_duplicates == 1
    assert window.column_model.selected_columns == (1200,)
    assert "'00123' (str)" in window.duplicate_details.toPlainText()
