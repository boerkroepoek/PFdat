"""Streamlit-interface voor de peilfilterloggerconverter."""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass

import streamlit as st

from converter import DEFAULT_PRESSURE_SCALE, ExcelResult, build_excel


@dataclass(frozen=True)
class ConversionError:
    """Veilige foutmelding voor één invoerbestand."""

    filename: str
    message: str


def create_zip(results: list[ExcelResult]) -> bytes:
    """Bundel meerdere Excel-bestanden in één ZIP-archief."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for result in results:
            archive.writestr(result.filename, result.content)
    return buffer.getvalue()


def initialize_state() -> None:
    """Initialiseer resultaten die tussen Streamlit-reruns behouden blijven."""
    st.session_state.setdefault("results", [])
    st.session_state.setdefault("errors", [])


def main() -> None:
    """Render en bestuur de Streamlit-app."""
    st.set_page_config(page_title="Peilfilterlogger naar Excel", page_icon="💧")
    initialize_state()

    st.title("Peilfilterlogger naar Excel")
    st.write(
        "Upload één of meerdere `.dat`- of `.txt`-bestanden. "
        "De app leest de header, converteert druk naar hPa en maakt per bestand een Excel-bestand."
    )

    with st.sidebar:
        st.header("Conversie-instellingen")
        pressure_scale = st.number_input(
            "Schaalfactor ruwe druk naar cmH2O",
            min_value=0.000001,
            value=DEFAULT_PRESSURE_SCALE,
            format="%.6f",
        )
        use_manual_offset = st.checkbox("Handmatige data-offset gebruiken")
        absolute_offset = st.number_input(
            "Byte-offset eerste record",
            min_value=0,
            value=0,
            step=1,
            disabled=not use_manual_offset,
        )
        st.caption("Conversie: ruwe druk × schaalfactor × 0,980665.")

    uploaded_files = st.file_uploader(
        "Loggerbestanden",
        type=["dat", "txt"],
        accept_multiple_files=True,
        help="Bestandstypecontrole op extensie is geen inhoudelijke veiligheidscontrole.",
    )

    if st.button("Converteer", type="primary", disabled=not uploaded_files):
        results: list[ExcelResult] = []
        errors: list[ConversionError] = []
        with st.spinner("Bestanden converteren..."):
            for uploaded_file in uploaded_files or []:
                try:
                    result = build_excel(
                        file_data=uploaded_file.getvalue(),
                        input_filename=uploaded_file.name,
                        pressure_scale=float(pressure_scale),
                        absolute_data_offset=int(absolute_offset) if use_manual_offset else None,
                    )
                    results.append(result)
                except (ValueError, UnicodeError, OSError) as exc:
                    errors.append(ConversionError(uploaded_file.name, str(exc)))
        st.session_state.results = results
        st.session_state.errors = errors

    for error in st.session_state.errors:
        st.error(f"{error.filename}: {error.message}")

    results = st.session_state.results
    if results:
        st.success(f"{len(results)} bestand(en) succesvol geconverteerd.")
        for index, result in enumerate(results):
            with st.expander(f"{result.filename} · {result.measurement_count} metingen", expanded=True):
                st.write(f"Start: {result.metadata.start_datetime:%d-%m-%Y %H:%M:%S}")
                st.write(f"Data-offset: {result.metadata.data_offset} bytes")
                st.download_button(
                    "Download Excel",
                    data=result.content,
                    file_name=result.filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{index}_{result.filename}",
                )
        if len(results) > 1:
            st.download_button(
                "Download alle bestanden als ZIP",
                data=create_zip(results),
                file_name="peilfilter_conversies.zip",
                mime="application/zip",
                key="download_zip",
            )


if __name__ == "__main__":
    main()
