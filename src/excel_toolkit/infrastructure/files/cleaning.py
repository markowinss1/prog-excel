"""Full-table readers and exclusive, atomic export for Cleaner."""

import csv
import os
import tempfile
import zlib
from pathlib import Path
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from excel_toolkit.application.cleaning import CleaningJob, CleaningSource
from excel_toolkit.application.models import Cell, CsvOptions, PreviewError
from excel_toolkit.infrastructure.files.cleaning_xlsx import (
    MAX_ARCHIVE_BYTES,
    MAX_CELLS,
    inspect_archive,
    read_table,
    validate_result,
    write_xlsx,
)
from excel_toolkit.infrastructure.files.csv_reader import SAMPLE_SIZE, detect_delimiter
from excel_toolkit.modules.cleaner import CleaningError, CleaningResult, Table

FORMAT_ERRORS = (
    OSError,
    UnicodeError,
    csv.Error,
    BadZipFile,
    ParseError,
    InvalidFileException,
    KeyError,
    ValueError,
    TypeError,
    OverflowError,
    PreviewError,
    EOFError,
    IndexError,
    RuntimeError,
    zlib.error,
)


class FileCleaningStorage:
    def read(self, path: Path, sheet: str | None, options: CsvOptions) -> CleaningSource:
        path = path.resolve()
        try:
            if path.suffix.lower() == ".xlsx":
                with path.open("rb") as stream:
                    snapshot = stream.read(MAX_ARCHIVE_BYTES + 1)
                if len(snapshot) > MAX_ARCHIVE_BYTES:
                    raise CleaningError("XLSX превышает лимит 128 MiB.")
                sheets, formulas, views = inspect_archive(snapshot)
                if not sheets:
                    raise CleaningError("В книге нет листов.")
                selected = sheet if sheet is not None else next(iter(sheets))
                if selected not in sheets:
                    raise CleaningError(f"Лист «{selected}» не найден.")
                table, trailing_rows = read_table(snapshot, selected, sheets[selected])
                return CleaningSource(
                    path,
                    table,
                    tuple(sheets),
                    selected,
                    xlsx_snapshot=snapshot,
                    sheet_part=sheets[selected],
                    contains_formulas=formulas,
                    xlsx_sheet_views=views.get(sheets[selected]),
                    xlsx_trailing_rows=trailing_rows,
                )
            if path.suffix.lower() != ".csv":
                raise CleaningError("Cleaner поддерживает только .xlsx и .csv.")
            if sheet is not None:
                raise CleaningError("У CSV нет листов.")
            if options.encoding not in {"utf-8-sig", "cp1251", "cp1252", "utf-16"}:
                raise CleaningError("Неизвестная кодировка CSV.")
            if options.delimiter not in {None, ",", ";", "\t", "|"}:
                raise CleaningError("Недопустимый разделитель CSV.")
            rows = []
            width = 0
            with path.open("r", encoding=options.encoding, newline="") as stream:
                delimiter = options.delimiter
                if delimiter is None:
                    sample = stream.read(SAMPLE_SIZE)
                    if len(sample) == SAMPLE_SIZE and "\n" in sample:
                        sample = sample.rsplit("\n", 1)[0]
                    delimiter = detect_delimiter(sample)
                    stream.seek(0)
                for row in csv.reader(stream, delimiter=delimiter, strict=True):
                    width = max(width, len(row), 1)
                    if (len(rows) + 1) * width > MAX_CELLS:
                        raise CleaningError("Область данных CSV превышает лимит 2 000 000 ячеек.")
                    rows.append(tuple(Cell(value) for value in row))
            return CleaningSource(
                path, Table(tuple(rows)), csv_options=CsvOptions(options.encoding, delimiter)
            )
        except FORMAT_ERRORS as exc:
            raise CleaningError(f"Не удалось прочитать таблицу без потери данных: {exc}") from exc

    def validate(self, source: CleaningSource, result: CleaningResult) -> None:
        validate_result(source, result)

    def write(self, job: CleaningJob, destination: Path) -> None:
        source, result = job.source, job.result
        self.validate(source, result)
        try:
            # Do not resolve the final component: a dangling symlink must also count as existing.
            destination = destination.absolute()
            if destination.resolve() == source.path.resolve():
                raise CleaningError("Нельзя сохранять результат поверх оригинала.")
            if os.path.lexists(destination):
                raise CleaningError("Выходной файл уже существует; выберите новое имя.")
            if destination.suffix.lower() != source.path.suffix.lower():
                raise CleaningError(
                    "Экспорт поддерживает только исходный формат; конвертация отключена."
                )
            temporary = None
            try:
                # Same filesystem: publish the complete file with exclusive hard-link creation.
                with tempfile.NamedTemporaryFile(
                    dir=destination.parent, prefix=".cleaner-", suffix=".tmp", delete=False
                ) as stream:
                    temporary = Path(stream.name)
                    if source.xlsx_snapshot is not None:
                        write_xlsx(source, result, stream)
                    else:
                        self._write_csv(source, result, stream)
                    stream.flush()
                    os.fsync(stream.fileno())
                os.link(temporary, destination)  # Atomic, fails if target appeared in the meantime.
            finally:
                if temporary is not None:
                    temporary.unlink(missing_ok=True)
        except FORMAT_ERRORS as exc:
            raise CleaningError(f"Не удалось безопасно экспортировать таблицу: {exc}") from exc

    @staticmethod
    def _write_csv(source: CleaningSource, result: CleaningResult, stream) -> None:
        import io

        options = source.csv_options
        if options is None or options.delimiter is None:
            raise CleaningError("Не определены настройки CSV.")
        # CSV loaded by this adapter contains only strings; no implicit type conversion.
        if any(not isinstance(c.value, str) for row in result.table.rows for c in row):
            raise CleaningError("CSV-экспорт принимает только строковые значения.")
        wrapper = io.TextIOWrapper(stream, encoding=options.encoding, newline="")
        try:
            writer = csv.writer(wrapper, delimiter=options.delimiter)
            writer.writerows([cell.value for cell in row] for row in result.table.rows)
            wrapper.flush()
        finally:
            wrapper.detach()
