"""Bounded CSV preview without numeric coercion or encoding guesses."""

import csv
from pathlib import Path

from excel_toolkit.application.models import (
    COLUMN_LIMIT,
    ROW_LIMIT,
    Cell,
    CsvOptions,
    Preview,
    PreviewError,
)
from excel_toolkit.application.preview import CancelCheck, check_cancelled

SAMPLE_SIZE = 65536


def detect_delimiter(sample: str) -> str:
    if not sample:
        return ","
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        if not any(separator in sample for separator in ",;\t|"):
            return ","  # A one-column table needs no separator.
        raise PreviewError(
            "Не удалось определить разделитель CSV. Выберите его вручную и повторите загрузку."
        ) from None


def read_csv(path: Path, options: CsvOptions, cancelled: CancelCheck) -> Preview:
    rows: list[tuple[Cell, ...]] = []
    more_rows = more_columns = False
    with path.open("r", encoding=options.encoding, newline="") as stream:
        check_cancelled(cancelled)
        delimiter = options.delimiter
        if delimiter is None:
            sample = stream.read(SAMPLE_SIZE)
            # Avoid feeding a cut-off last record to the heuristic where possible.
            if len(sample) == SAMPLE_SIZE and "\n" in sample:
                sample = sample.rsplit("\n", 1)[0]
            delimiter = detect_delimiter(sample)
            stream.seek(0)
        reader = csv.reader(stream, delimiter=delimiter, strict=True)
        for index, row in enumerate(reader):
            check_cancelled(cancelled)
            if index == ROW_LIMIT:
                more_rows = True
                break
            more_columns |= len(row) > COLUMN_LIMIT
            rows.append(tuple(Cell(value) for value in row[:COLUMN_LIMIT]))
    return Preview(
        path=path,
        rows=tuple(rows),
        more_rows=more_rows,
        more_columns=more_columns,
        csv_options=CsvOptions(options.encoding, delimiter),
    )
