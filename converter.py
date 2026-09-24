"""Bedrijfslogica voor het converteren van peilfilterloggerbestanden."""

from __future__ import annotations

import re
import struct
from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Iterator

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

CMH2O_TO_HPA = 0.980665
DEFAULT_PRESSURE_SCALE = 0.1
EXCEL_MAX_ROWS = 1_048_576
RECORD_STRUCT = struct.Struct("<HH")
DATETIME_PATTERN = re.compile(
    rb"(?P<time>\d{2}:\d{2}:\d{2})\s+(?P<date>\d{2}/\d{2}/\d{2})"
)
INTERVAL_PATTERN = re.compile(
    rb"(?P<days>\d{2})\s+(?P<hours>\d{2}):(?P<minutes>\d{2}):"
    rb"(?P<seconds>\d{2})\s+\dT"
)


@dataclass(frozen=True)
class LoggerMetadata:
    """Metadata uit de header van een loggerbestand."""

    start_datetime: datetime
    sample_interval: timedelta
    data_offset: int
    header_text: str


@dataclass(frozen=True)
class Measurement:
    """Eén drukmeting."""

    timestamp: datetime
    pressure_hpa: float
    raw_pressure: int
    raw_temperature: int


@dataclass(frozen=True)
class ExcelResult:
    """Gegenereerde werkmap en samenvatting."""

    filename: str
    content: bytes
    measurement_count: int
    metadata: LoggerMetadata


def parse_logger_datetime(time_bytes: bytes, date_bytes: bytes) -> datetime:
    """Converteer DD/MM/YY en HH:MM:SS naar een datetime."""
    value = f"{date_bytes.decode('ascii')} {time_bytes.decode('ascii')}"
    return datetime.strptime(value, "%d/%m/%y %H:%M:%S")


def parse_interval(header: bytes) -> timedelta:
    """Lees en valideer de meetinterval uit de header."""
    match = INTERVAL_PATTERN.search(header)
    if match is None:
        raise ValueError("De meetinterval kon niet uit de header worden gelezen.")
    interval = timedelta(
        days=int(match.group("days")),
        hours=int(match.group("hours")),
        minutes=int(match.group("minutes")),
        seconds=int(match.group("seconds")),
    )
    if interval <= timedelta(0):
        raise ValueError(f"Ongeldige meetinterval gevonden: {interval}.")
    return interval


def detect_metadata(file_data: bytes, absolute_data_offset: int | None = None) -> LoggerMetadata:
    """Detecteer starttijd, interval en beginpositie van de meetdata."""
    if not file_data:
        raise ValueError("Het invoerbestand is leeg.")
    matches = list(DATETIME_PATTERN.finditer(file_data[:1024]))
    if not matches:
        raise ValueError("Geen datum/tijdvelden gevonden in de eerste 1024 bytes.")
    start = parse_logger_datetime(matches[0].group("time"), matches[0].group("date"))
    if absolute_data_offset is not None:
        offset = absolute_data_offset
    elif len(matches) >= 2:
        offset = matches[1].end()
    else:
        raise ValueError(
            "Het einde van de header kon niet automatisch worden bepaald. "
            "Vul een handmatige byte-offset in."
        )
    if offset < 0 or offset >= len(file_data):
        raise ValueError(
            f"Ongeldige data-offset {offset}; bestandsgrootte is {len(file_data)} bytes."
        )
    header = file_data[:offset]
    return LoggerMetadata(start, parse_interval(header), offset, header.decode("cp1252", errors="replace"))


def raw_pressure_to_hpa(raw_pressure: int, pressure_scale: float) -> float:
    """Converteer een ruwe drukwaarde via cmH2O naar hPa."""
    if pressure_scale <= 0:
        raise ValueError("De drukschaalfactor moet groter zijn dan nul.")
    return raw_pressure * pressure_scale * CMH2O_TO_HPA


def iter_measurements(
    file_data: bytes,
    metadata: LoggerMetadata,
    pressure_scale: float = DEFAULT_PRESSURE_SCALE,
) -> Iterator[Measurement]:
    """Itereer over alle volledige records; losse eindbytes worden genegeerd."""
    binary_data = file_data[metadata.data_offset:]
    for index in range(len(binary_data) // RECORD_STRUCT.size):
        raw_pressure, raw_temperature = RECORD_STRUCT.unpack_from(
            binary_data, index * RECORD_STRUCT.size
        )
        yield Measurement(
            timestamp=metadata.start_datetime + index * metadata.sample_interval,
            pressure_hpa=raw_pressure_to_hpa(raw_pressure, pressure_scale),
            raw_pressure=raw_pressure,
            raw_temperature=raw_temperature,
        )


def format_timedelta(value: timedelta) -> str:
    """Formatteer een timedelta als dagen en HH:MM:SS."""
    total = int(value.total_seconds())
    days, rest = divmod(total, 86_400)
    hours, rest = divmod(rest, 3_600)
    minutes, seconds = divmod(rest, 60)
    prefix = f"{days} dag(en), " if days else ""
    return f"{prefix}{hours:02d}:{minutes:02d}:{seconds:02d}"


def _style_header(worksheet) -> None:
    """Pas de gedeelde kopopmaak toe."""
    fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in worksheet[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def build_excel(
    file_data: bytes,
    input_filename: str,
    pressure_scale: float = DEFAULT_PRESSURE_SCALE,
    absolute_data_offset: int | None = None,
) -> ExcelResult:
    """Converteer bytes naar een gevalideerde Excel-werkmap in het geheugen."""
    metadata = detect_metadata(file_data, absolute_data_offset)
    expected_count = (len(file_data) - metadata.data_offset) // RECORD_STRUCT.size
    if expected_count == 0:
        raise ValueError("Geen volledige meetrecords gevonden.")
    if expected_count + 1 > EXCEL_MAX_ROWS:
        raise ValueError("Het aantal metingen overschrijdt de Excel-rijlimiet.")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Metingen"
    worksheet.sheet_view.showGridLines = False
    worksheet.freeze_panes = "A2"
    worksheet.append(["Datum/tijd", "Druk (hPa)"])
    _style_header(worksheet)

    count = 0
    for measurement in iter_measurements(file_data, metadata, pressure_scale):
        worksheet.append([measurement.timestamp, measurement.pressure_hpa])
        worksheet.cell(worksheet.max_row, 1).number_format = "dd-mm-yyyy hh:mm:ss"
        worksheet.cell(worksheet.max_row, 2).number_format = "0.000"
        count += 1
    worksheet.column_dimensions["A"].width = 22
    worksheet.column_dimensions["B"].width = 16
    table = Table(displayName="TabelMetingen", ref=f"A1:B{worksheet.max_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    worksheet.add_table(table)

    metadata_sheet = workbook.create_sheet("Metadata")
    metadata_sheet.sheet_view.showGridLines = False
    metadata_sheet.freeze_panes = "A2"
    rows = [
        ("Eigenschap", "Waarde"),
        ("Invoerbestand", Path(input_filename).name),
        ("Startdatum/tijd", metadata.start_datetime),
        ("Meetinterval", format_timedelta(metadata.sample_interval)),
        ("Aantal metingen", count),
        ("Begin meetdata, byte-offset", metadata.data_offset),
        ("Ruwe druk naar cmH2O", pressure_scale),
        ("cmH2O naar hPa", CMH2O_TO_HPA),
        ("Drukberekening", "ruwe druk × schaalfactor × 0,980665"),
    ]
    for row in rows:
        metadata_sheet.append(row)
    _style_header(metadata_sheet)
    metadata_sheet.column_dimensions["A"].width = 32
    metadata_sheet.column_dimensions["B"].width = 60
    metadata_sheet["B3"].number_format = "dd-mm-yyyy hh:mm:ss"
    for row in metadata_sheet.iter_rows(min_row=2, max_col=2):
        row[0].font = Font(bold=True)
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    workbook.properties.title = f"Peilfiltermetingen - {Path(input_filename).stem}"
    workbook.properties.creator = "Peilfilterlogger Streamlit-app"
    output = BytesIO()
    workbook.save(output)
    workbook.close()

    content = output.getvalue()
    validation = load_workbook(BytesIO(content), read_only=True, data_only=True)
    try:
        if set(validation.sheetnames) != {"Metingen", "Metadata"}:
            raise ValueError("De gegenereerde werkmap bevat niet de verwachte werkbladen.")
        if max(validation["Metingen"].max_row - 1, 0) != count:
            raise ValueError("Validatie van het aantal Excel-meetregels is mislukt.")
    finally:
        validation.close()

    return ExcelResult(
        filename=f"{Path(input_filename).stem}_hpa.xlsx",
        content=content,
        measurement_count=count,
        metadata=metadata,
    )
