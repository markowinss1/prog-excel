"""Translate format/IO errors at the application boundary."""

import csv
from pathlib import Path
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from excel_toolkit.application.models import CsvOptions, Preview, PreviewError
from excel_toolkit.application.preview import CancelCheck
from excel_toolkit.infrastructure.files.csv_reader import read_csv
from excel_toolkit.infrastructure.files.xlsx_reader import read_xlsx


class FilePreviewReader:
    def read(
        self, path: Path, sheet: str | None, options: CsvOptions, cancelled: CancelCheck
    ) -> Preview:
        try:
            if path.suffix.lower() == ".xlsx":
                return read_xlsx(path, sheet, cancelled)
            if path.suffix.lower() == ".csv":
                return read_csv(path, options, cancelled)
            raise PreviewError("Поддерживаются только .xlsx и .csv.")
        except UnicodeError as exc:
            raise PreviewError(
                "Не удалось декодировать CSV. Выберите кодировку (например, Windows-1251) "
                "и откройте файл повторно. Автоматическая замена символов отключена."
            ) from exc
        except OSError as exc:
            raise PreviewError(f"Не удалось открыть файл: {exc}") from exc
        except (
            BadZipFile,
            InvalidFileException,
            ParseError,
            KeyError,
            ValueError,
            csv.Error,
        ) as exc:
            raise PreviewError(f"Не удалось прочитать таблицу: {exc}") from exc
