"""Repeatable warm-cache benchmark; run with the same interpreter on both revisions."""

import argparse
import json
import platform
import statistics
import tempfile
from datetime import datetime
from pathlib import Path
from time import perf_counter

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from excel_toolkit.application.cleaning import CleaningService
from excel_toolkit.infrastructure.files.cleaning import FileCleaningStorage
from excel_toolkit.modules.cleaner import CleaningPlan


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--freeze", default=None, help="Optional freeze coordinate, e.g. A2")
    parser.add_argument("--repeats", type=int, default=5)
    parser.add_argument("--dataset", choices=("duplicates", "clients"), default="duplicates")
    parser.add_argument("--tail", choices=("bare", "space", "formatted"), default="bare")
    parser.add_argument(
        "--save-sample", type=Path, help="Keep input under a new, exclusive filename"
    )
    args = parser.parse_args()
    if args.repeats < 1:
        parser.error("--repeats must be positive")
    measurements = []
    with tempfile.TemporaryDirectory(prefix="cleaner-benchmark-") as directory:
        path = Path(directory) / "input.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.append(["ID", "Name", "Number", "Date", "Flag"])
        if args.dataset == "clients":
            for index in range(10_000):
                sheet.append(
                    [f" {index:06d} ", f" Name {index} ", index, datetime(2026, 1, 2), True]
                )
            for _ in range(199):
                sheet.append([None])
            for index in range(400):
                sheet.append(
                    [f" {index:06d} ", f" Name {index} ", index, datetime(2026, 1, 2), True]
                )
            sheet.append([" " if args.tail == "space" else None])
            if args.tail == "formatted":
                sheet["A10601"].font = Font(bold=True)
            expected = (
                (10_601, 10_001, 200, 400) if args.tail == "space" else (10_600, 10_001, 199, 400)
            )
        else:
            for index in range(10_000):
                value = index // 2
                sheet.append(
                    [None, " "]
                    if index % 20 == 19
                    else [f" {value:06d} ", f" Name {value} ", value, datetime(2026, 1, 2), True]
                )
            expected = (10_001, 5_001, 500, 4_500)
        sheet.freeze_panes = args.freeze
        book.create_sheet("Other").append([" untouched ", "00001"])
        book.save(path)
        book.close()
        original = path.read_bytes()
        if args.save_sample:
            with args.save_sample.open("xb") as stream:
                stream.write(original)
        service = CleaningService(FileCleaningStorage())
        for repetition in range(args.repeats + 1):  # First run is warm-up.
            start = perf_counter()
            job = service.analyze(path, CleaningPlan(True, True, True, True))
            analyzed = perf_counter()
            output = Path(directory) / f"out-{repetition}.xlsx"
            service.export(job, output)
            exported = perf_counter()
            report = job.result.report
            assert (
                report.original_rows,
                report.final_rows,
                report.removed_empty_rows,
                report.removed_duplicates,
            ) == expected
            if repetition:
                measurements.append({"analyze": analyzed - start, "export": exported - analyzed})
        assert path.read_bytes() == original
        book = load_workbook(output)
        try:
            physical_last_cell = expected[1] + int(
                args.dataset == "clients" and args.tail == "formatted"
            )
            assert book.active.max_row == physical_last_cell
            assert book.active.cell(expected[1], 1).value == (
                "009999" if args.dataset == "clients" else "004999"
            )
            assert book.active.freeze_panes == args.freeze
            assert book["Other"]["B1"].value == "00001"
        finally:
            book.close()
        again = service.analyze(output, CleaningPlan(True, True, True, True))
        assert again.result.report.original_rows == expected[1]
        assert again.result.report.removed_empty_rows == 0
        assert again.result.report.removed_duplicates == 0
    print(
        json.dumps(
            {
                "python": platform.python_version(),
                "openpyxl": openpyxl.__version__,
                "platform": platform.platform(),
                "freeze": args.freeze,
                "records": 10_000,
                "dataset": args.dataset,
                "tail": args.tail if args.dataset == "clients" else None,
                "report": {
                    name: getattr(report, name)
                    for name in (
                        "original_rows",
                        "final_rows",
                        "removed_empty_rows",
                        "removed_duplicates",
                    )
                },
                "samples_seconds": measurements,
                "median_seconds": {
                    key: statistics.median(sample[key] for sample in measurements)
                    for key in ("analyze", "export")
                },
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
