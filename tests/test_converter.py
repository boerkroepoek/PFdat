"""Regressie- en eenheidstests voor de converter."""

from datetime import datetime, timedelta
from io import BytesIO
import struct

import pytest
from openpyxl import load_workbook

from converter import (
    Measurement,
    build_excel,
    compare_measurements,
    excel_serial_to_datetime,
    parse_reference_data,
    raw_pressure_to_hpa,
)


def sample_file(records: list[tuple[int, int]] | None = None) -> bytes:
    """Maak representatieve loggerbytes met twee datumvelden en 2-uursinterval."""
    header = b"Logger 00:00:12 17/03/26 00 02:00:00 0T end 12:00:00 17/03/26"
    payload = b"".join(struct.pack("<HH", *item) for item in (records or [(100, 200)]))
    return header + payload


def test_excel_serial_reference_date() -> None:
    assert excel_serial_to_datetime(46098.5) == datetime(2026, 3, 17, 12, 0)


def test_reference_parser_accepts_decimal_comma_and_tabs() -> None:
    values = parse_reference_data("46098,5\t1214,8\n46098,58333\t1213,692")
    assert values[0].timestamp == datetime(2026, 3, 17, 12, 0)
    assert values[0].pressure_hpa == pytest.approx(1214.8)
    assert values[1].timestamp == datetime(2026, 3, 17, 13, 59, 59, 712000)


def test_reference_parser_rejects_empty_input() -> None:
    with pytest.raises(ValueError, match="geen geldige"):
        parse_reference_data("")


def test_pressure_scale_and_offset() -> None:
    assert raw_pressure_to_hpa(100, 0.1, 2.0) == pytest.approx(11.80665)


def test_comparison_reports_time_and_pressure_differences() -> None:
    calculated = [Measurement(datetime(2026, 3, 17, 12), 1214.7, 1, 2)]
    reference = parse_reference_data("46098,5\t1214,8")
    summary = compare_measurements(calculated, reference, 0.2, 1.0)
    assert summary.pressure_within_tolerance
    assert summary.time_within_tolerance
    assert summary.mean_pressure_difference_hpa == pytest.approx(-0.1)


def test_build_excel_uses_second_header_datetime_when_selected() -> None:
    reference = parse_reference_data("46098,5\t9,80665")
    result = build_excel(
        sample_file(), "logger.dat", datetime_field_index=1, reference=reference
    )
    assert result.measurements[0].timestamp == datetime(2026, 3, 17, 12)
    assert result.metadata.sample_interval == timedelta(hours=2)
    workbook = load_workbook(BytesIO(result.content), read_only=True)
    try:
        assert workbook.sheetnames == ["Metingen", "Metadata", "Referentiecontrole"]
    finally:
        workbook.close()


def test_build_excel_rejects_invalid_header() -> None:
    with pytest.raises(ValueError, match="Geen datum/tijdvelden"):
        build_excel(b"ongeldig", "fout.dat")
