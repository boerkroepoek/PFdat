"""Tests voor apphulpfuncties."""

from datetime import datetime, timedelta
from io import BytesIO
import zipfile

from app import create_zip
from converter import ExcelResult, LoggerMetadata


def test_create_zip() -> None:
    metadata = LoggerMetadata(datetime(2026, 1, 1), timedelta(hours=1), 10, "header")
    result = ExcelResult("meting.xlsx", b"excel-bytes", 1, metadata)
    archive_bytes = create_zip([result])
    with zipfile.ZipFile(BytesIO(archive_bytes)) as archive:
        assert archive.namelist() == ["meting.xlsx"]
        assert archive.read("meting.xlsx") == b"excel-bytes"
