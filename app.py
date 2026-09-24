"""Streamlit-interface voor conversie en referentiecontrole."""

from __future__ import annotations

import io
import zipfile
from datetime import datetime, time

import streamlit as st

from converter import (
    DEFAULT_PRESSURE_OFFSET_HPA,
    DEFAULT_PRESSURE_SCALE,
    ExcelResult,
    build_excel,
    parse_reference_data,
    suggested_pressure_offset,
)


def create_zip(results: list[ExcelResult]) -> bytes:
    """Bundel Excel-resultaten in één ZIP-bestand."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for result in results:
            archive.writestr(result.filename, result.content)
    return buffer.getvalue()


def initialize_state() -> None:
    """Initialiseer status die Streamlit-reruns moet overleven."""
    st.session_state.setdefault("results", [])
    st.session_state.setdefault("errors", [])


def main() -> None:
    """Render de Streamlit-app en verwerk gebruikersacties."""
    st.set_page_config(page_title="Peilfilterlogger naar Excel", page_icon="💧", layout="wide")
    initialize_state()
    st.title("Peilfilterlogger naar Excel")
    st.write(
        "Converteer binaire loggerbestanden en controleer de uitkomst optioneel "
        "tegen geëxporteerde Excel-datum- en hPa-kolommen."
    )

    with st.sidebar:
        st.header("Conversie")
        pressure_scale = st.number_input(
            "Ruwe druk naar cmH2O", min_value=0.000001,
            value=DEFAULT_PRESSURE_SCALE, format="%.6f"
        )
        pressure_offset_hpa = st.number_input(
            "Drukoffset (hPa)", value=DEFAULT_PRESSURE_OFFSET_HPA,
            format="%.6f", help="Wordt na de schaalconversie opgeteld."
        )
        datetime_field_number = st.number_input(
            "Te gebruiken datumveld uit header", min_value=1, value=1, step=1
        )
        use_manual_offset = st.checkbox("Handmatige data-offset")
        absolute_offset = st.number_input(
            "Byte-offset eerste record", min_value=0, value=0,
            step=1, disabled=not use_manual_offset
        )
        override_start = st.checkbox("Startdatum/tijd handmatig instellen")
        start_date = st.date_input("Startdatum", disabled=not override_start)
        start_time = st.time_input("Starttijd", value=time(0, 0), disabled=not override_start)

        st.header("Toleranties")
        pressure_tolerance = st.number_input(
            "Maximale drukafwijking (hPa)", min_value=0.0, value=0.1, format="%.3f"
        )
        time_tolerance = st.number_input(
            "Maximale tijdafwijking (seconden)", min_value=0.0, value=1.0, format="%.1f"
        )

    logger_files = st.file_uploader(
        "Loggerbestand(en)", type=["dat", "txt"], accept_multiple_files=True
    )
    reference_file = st.file_uploader(
        "Optioneel referentiebestand met Excel-datum en hPa",
        type=["txt", "csv", "tsv"],
        help="Ondersteunt tab, puntkomma of spaties en Nederlandse decimale komma's."
    )
    reference_text = st.text_area(
        "Of plak referentiedata",
        height=140,
        placeholder="46098,5\t1214,8\n46098,58333\t1213,692",
    )

    if st.button("Converteer en controleer", type="primary", disabled=not logger_files):
        results: list[ExcelResult] = []
        errors: list[str] = []
        try:
            uploaded_reference = (
                reference_file.getvalue().decode("utf-8-sig") if reference_file else ""
            )
            reference_source = reference_text.strip() or uploaded_reference.strip()
            reference = parse_reference_data(reference_source) if reference_source else None
        except (UnicodeDecodeError, ValueError) as exc:
            reference = None
            errors.append(f"Referentiedata: {exc}")

        if not errors:
            for uploaded_file in logger_files or []:
                try:
                    result = build_excel(
                        file_data=uploaded_file.getvalue(),
                        input_filename=uploaded_file.name,
                        pressure_scale=float(pressure_scale),
                        pressure_offset_hpa=float(pressure_offset_hpa),
                        absolute_data_offset=int(absolute_offset) if use_manual_offset else None,
                        start_datetime_override=(
                            datetime.combine(start_date, start_time) if override_start else None
                        ),
                        datetime_field_index=int(datetime_field_number) - 1,
                        reference=reference,
                        pressure_tolerance_hpa=float(pressure_tolerance),
                        time_tolerance_seconds=float(time_tolerance),
                    )
                    results.append(result)
                except (ValueError, OSError, struct.error) as exc:  # type: ignore[name-defined]
                    errors.append(f"{uploaded_file.name}: {exc}")
        st.session_state.results = results
        st.session_state.errors = errors

    for error in st.session_state.errors:
        st.error(error)

    results = st.session_state.results
    if results:
        st.success(f"{len(results)} bestand(en) geconverteerd.")
        for index, result in enumerate(results):
            st.subheader(result.filename)
            col1, col2, col3 = st.columns(3)
            col1.metric("Metingen", len(result.measurements))
            col2.metric("Meetinterval", str(result.metadata.sample_interval))
            col3.metric("Data-offset", f"{result.metadata.data_offset} bytes")
            st.caption(
                "Gevonden datumvelden: "
                + ", ".join(value.strftime("%d-%m-%Y %H:%M:%S") for value in result.metadata.detected_datetimes)
            )
            if result.validation:
                validation = result.validation
                status_ok = validation.pressure_within_tolerance and validation.time_within_tolerance
                if status_ok:
                    st.success("Referentiecontrole valt binnen beide toleranties.")
                else:
                    st.warning("Referentiecontrole valt buiten één of beide toleranties.")
                metrics = st.columns(4)
                metrics[0].metric("Vergeleken", validation.compared_count)
                metrics[1].metric(
                    "Gem. absolute drukafwijking",
                    f"{validation.mean_absolute_pressure_difference_hpa:.3f} hPa",
                )
                metrics[2].metric(
                    "Max. absolute drukafwijking",
                    f"{validation.maximum_absolute_pressure_difference_hpa:.3f} hPa",
                )
                metrics[3].metric(
                    "Max. tijdafwijking",
                    f"{validation.maximum_absolute_time_difference_seconds:.1f} s",
                )
                st.info(
                    "Voorgestelde aanvullende drukoffset op basis van het gemiddelde: "
                    f"{suggested_pressure_offset(validation):+.6f} hPa. "
                    "Dit corrigeert alleen een constante afwijking, niet een onjuiste schaalfactor."
                )
                preview = [
                    {
                        "Regel": row.index,
                        "Berekende tijd": row.calculated_timestamp,
                        "Referentietijd": row.reference_timestamp,
                        "Tijdverschil (s)": row.time_difference_seconds,
                        "Berekende druk (hPa)": row.calculated_pressure_hpa,
                        "Referentiedruk (hPa)": row.reference_pressure_hpa,
                        "Drukverschil (hPa)": row.pressure_difference_hpa,
                    }
                    for row in validation.rows[:10]
                ]
                st.dataframe(preview, use_container_width=True, hide_index=True)
                if validation.calculated_count != validation.reference_count:
                    st.warning(
                        "Het aantal berekende en referentieregels verschilt: "
                        f"{validation.calculated_count} tegenover {validation.reference_count}."
                    )
            st.download_button(
                "Download Excel", result.content, result.filename,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{index}_{result.filename}",
            )
        if len(results) > 1:
            st.download_button(
                "Download alles als ZIP", create_zip(results),
                "peilfilter_conversies.zip", "application/zip"
            )


if __name__ == "__main__":
    import struct

    main()
