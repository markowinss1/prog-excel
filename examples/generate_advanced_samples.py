"""Generate synthetic Stage 4A samples; never overwrite an existing file."""

import argparse
import csv
from pathlib import Path

from openpyxl import Workbook

CLIENT_ROWS = (
    ("ID", "Name", "City", "Email"),
    (" 00123 ", " Alex ", "Astana", "alex@example.invalid"),
    ("00123", "Alex", "Almaty", "alex@example.invalid"),
    ("00007", "Dana", "Astana", ""),
    ("00007", "Dana", "Astana", ""),
    ("", "No ID", "Astana", ""),
    ("", "No ID", "Astana", ""),
    ("abc", "Sam", "Astana", "sam@example.invalid"),
    ("ABC", "Sam", "Astana", "sam@example.invalid"),
    (" 00042 ", " Mira ", " Shymkent ", "mira@example.invalid"),
    ("00042", "Mira", "Shymkent", "mira@example.invalid"),
    ("", "", "", ""),
    ("00999", "Lee", "Aktau", ""),
    ("00999", "Lee", "Aktau", "lee@example.invalid"),
    ("00000", "Zero", "Astana", "zero@example.invalid"),
)

HEADER_ROWS = (
    ("ID", "Name", "", "Name"),
    ("00011", "Same", "Astana", "First"),
    ("00012", "Same", "Almaty", "Second"),
)


def _write_csv(path: Path) -> bool:
    try:
        stream = path.open("x", encoding="utf-8-sig", newline="")
    except FileExistsError:
        return False
    with stream:
        csv.writer(stream, delimiter=";").writerows(CLIENT_ROWS)
    return True


def _write_xlsx(path: Path) -> bool:
    try:
        stream = path.open("xb")
    except FileExistsError:
        return False
    workbook = Workbook()
    try:
        sheet = workbook.active
        sheet.title = "Clients"
        for row in CLIENT_ROWS:
            sheet.append(row)
        sheet.freeze_panes = "A2"
        for column, width in (("A", 16), ("B", 20), ("C", 18), ("D", 30)):
            sheet.column_dimensions[column].width = width
        headers = workbook.create_sheet("Headers")
        for row in HEADER_ROWS:
            headers.append(row)
        headers.freeze_panes = "A2"
        with stream:
            workbook.save(stream)
    finally:
        stream.close()
        workbook.close()
    return True


def generate_samples(directory: Path) -> tuple[Path, ...]:
    """Return newly created paths; existing contents and timestamps stay intact."""
    directory.mkdir(parents=True, exist_ok=True)
    created = []
    for filename, writer in (
        ("advanced_cleaning_demo.csv", _write_csv),
        ("advanced_cleaning_demo.xlsx", _write_xlsx),
    ):
        path = directory / filename
        if writer(path):
            created.append(path)
    return tuple(created)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).parent / "generated" / "stage4a",
        help="Destination directory (default: examples/generated/stage4a).",
    )
    arguments = parser.parse_args()
    created = generate_samples(arguments.output_dir)
    print(f"Created {len(created)} files; existing files were kept: {arguments.output_dir}")


if __name__ == "__main__":
    main()
