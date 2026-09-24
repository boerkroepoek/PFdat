"""Tests voor apphulpfuncties."""

from datetime import datetime, timedelta
from io import BytesIO
import zipfile

from app import create_zip
from converter import ExcelResult, LoggerMetadata, Measurement


def test_create_zip() -> None:
    metadata = LoggerMetadata(
        datetime(2026, 1, 1), timedelta(hours=1), 10, "header",
        (datetime(2026, 1, 1),)
    )
    measurement = Measurement(datetime(2026, 1, 1), 1.0, 1, 2)
    result = ExcelResult("meting.xlsx", b"excel", (measurement,), metadata, None)
    with zipfile.ZipFile(BytesIO(create_zip([result]))) as archive:
        assert archive.read("meting.xlsx") == b"excel"
