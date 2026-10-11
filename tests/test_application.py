from pathlib import Path

import pytest

from excel_toolkit.application.models import CsvOptions, Preview, PreviewError
from excel_toolkit.application.preview import CancelCheck, PreviewService


def test_service_uses_reader_contract_without_filesystem() -> None:
    calls = []
    expected = Preview(Path("virtual.xlsx"), ())

    class Reader:
        def read(
            self, path: Path, sheet: str | None, options: CsvOptions, cancelled: CancelCheck
        ) -> Preview:
            calls.append((path, sheet, options))
            return expected

    service = PreviewService(Reader())
    assert service.load("virtual.xlsx", sheet="Data") is expected
    assert calls == [(Path("virtual.xlsx"), "Data", CsvOptions())]
    with pytest.raises(PreviewError):
        service.load("unsupported.xls")
    assert len(calls) == 1
