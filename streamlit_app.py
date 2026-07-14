"""
streamlit_app.py

Streamlit web front-end for the Price Update Automation Tool.

Deploy by pushing this repository to GitHub and connecting it on
Streamlit Community Cloud (share.streamlit.io) — see README.md for
the full deployment walkthrough. Google service-account credentials
are read from Streamlit secrets (st.secrets["gcp_service_account"]);
no local credentials file is required in the cloud.
"""

from __future__ import annotations

import logging
from datetime import datetime
from io import StringIO
from typing import Optional

import pandas as pd
import streamlit as st

from config import APP_TITLE, APP_VERSION, PriceType, UpdateType
from excel_reader import ExcelReadError, ExcelReader
from excel_writer import ExcelWriteError, ExcelWriter
from google_sheet import GoogleSheetClient, GoogleSheetError
from validator import PriceValidator, ValidationError, ValidationResult

st.set_page_config(page_title=APP_TITLE, page_icon="🛒", layout="wide")

# --------------------------------------------------------------------------
# In-memory log capture (mirrors the desktop app's log window)
# --------------------------------------------------------------------------
_LOG_STREAM_KEY = "_log_stream"
_LOGGER_NAME = "PriceUpdateAutomation.streamlit"


def _get_logger() -> logging.Logger:
    """Return a logger that writes into a per-session StringIO buffer."""
    if _LOG_STREAM_KEY not in st.session_state:
        st.session_state[_LOG_STREAM_KEY] = StringIO()

    logger = logging.getLogger(_LOGGER_NAME)
    logger.setLevel(logging.INFO)

    # Always point the handler at the current session's buffer, and make
    # sure exactly one handler is attached (Streamlit re-runs this module
    # on every interaction).
    stream = st.session_state[_LOG_STREAM_KEY]
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
    handler = logging.StreamHandler(stream)
    handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S"))
    logger.addHandler(handler)
    logger.propagate = False
    return logger


def _log_text() -> str:
    stream: StringIO = st.session_state.get(_LOG_STREAM_KEY, StringIO())
    return stream.getvalue()


# --------------------------------------------------------------------------
# Credential resolution: Streamlit secrets first, local file as fallback
# --------------------------------------------------------------------------
def _resolve_credentials():
    """
    Return Google service-account credentials, preferring Streamlit
    secrets (cloud deployment) and falling back to a local credentials
    file (desktop/dev usage).
    """
    from config import CREDENTIALS_FILE

    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
    except FileNotFoundError:
        # No secrets.toml present at all (e.g. local dev without secrets
        # configured) - fall back to the local credentials file.
        pass
    except Exception:  # noqa: BLE001 - any other secrets-access failure
        pass

    return CREDENTIALS_FILE


# --------------------------------------------------------------------------
# Session state helpers
# --------------------------------------------------------------------------
def _init_session_state() -> None:
    defaults = {
        "validation_result": None,
        "master_row_count": None,
        "marketplace_row_count": None,
        "last_run_selections": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def _status_style(status: str) -> str:
    """Return a CSS background color for a given Validation Status."""
    colors = {
        "Ready for Upload": "background-color: #C6EFCE",
        "RRP Mismatch": "background-color: #FFEB9C",
        "SRP Mismatch": "background-color: #FFEB9C",
        "Both Mismatch": "background-color: #FFC7CE",
        "SKU Not Found": "background-color: #FFC7CE",
        "Duplicate SKU": "background-color: #D9D2E9",
        "Missing Data": "background-color: #F4CCCC",
        "Invalid Price": "background-color: #F4CCCC",
    }
    return colors.get(status, "")


def _highlight_status_rows(row: pd.Series) -> list[str]:
    style = _status_style(row.get("Validation Status", ""))
    return [style] * len(row)


# --------------------------------------------------------------------------
# Main app
# --------------------------------------------------------------------------
def main() -> None:
    _init_session_state()
    logger = _get_logger()

    st.title(f"🛒 {APP_TITLE}")
    st.caption(f"v{APP_VERSION} · Validate marketplace prices before uploading a Price Update file.")

    with st.sidebar:
        st.header("1. Data Sources")
        sheet_url = st.text_input(
            "Google Sheets URL",
            help="Full URL of the Master sheet (Seller SKU, RRP, SRP, Price Type).",
        )
        uploaded_file = st.file_uploader(
            "Marketplace Report (.xlsx)",
            type=["xlsx", "xlsm"],
            help="Must contain Seller SKU, Price, Special Price columns.",
        )

        st.header("2. Validation Options")
        price_types = st.multiselect(
            "Price Type(s)",
            options=PriceType.values(),
            default=[],
        )
        update_type = st.selectbox("Update Type", options=UpdateType.values(), index=0)

        st.header("3. Actions")
        validate_clicked = st.button("Validate", type="primary", use_container_width=True)
        reset_clicked = st.button("Reset", use_container_width=True)

    if reset_clicked:
        st.session_state["validation_result"] = None
        st.session_state[_LOG_STREAM_KEY] = StringIO()
        st.rerun()

    if validate_clicked:
        _run_validation(logger, sheet_url, uploaded_file, price_types, update_type)

    _render_results()
    _render_log_panel()


def _run_validation(
    logger: logging.Logger,
    sheet_url: str,
    uploaded_file,
    price_types_raw: list[str],
    update_type_raw: str,
) -> None:
    if not sheet_url or not sheet_url.strip():
        st.error("Please enter a Google Sheets URL.")
        return
    if not price_types_raw:
        st.error("Please select at least one Price Type.")
        return
    if uploaded_file is None:
        st.error("Please upload a Marketplace Report (.xlsx) file.")
        return

    price_types = [PriceType(v) for v in price_types_raw]
    update_type = UpdateType(update_type_raw)
    start_time = datetime.now()

    with st.spinner("Validating prices..."):
        try:
            logger.info("Connecting to Google Sheets...")
            credentials_source = _resolve_credentials()
            sheet_client = GoogleSheetClient(credentials_source)
            master_df = sheet_client.load_master_data(sheet_url)
            logger.info("Loaded %d Master row(s) from Google Sheets.", len(master_df))

            logger.info("Reading Marketplace report: %s", uploaded_file.name)
            marketplace_df = _read_uploaded_marketplace_report(uploaded_file)
            logger.info("Loaded %d Marketplace report row(s).", len(marketplace_df))

            logger.info(
                "Validating with Update Type='%s', Price Type(s)=%s",
                update_type.value,
                [pt.value for pt in price_types],
            )
            validator = PriceValidator(update_type=update_type, price_types=price_types)
            result = validator.validate(master_df, marketplace_df)

            duration = (datetime.now() - start_time).total_seconds()
            logger.info(
                "Validation complete in %.2fs. Total=%d Passed=%d Failed=%d",
                duration,
                result.total_rows,
                result.passed_rows,
                result.failed_rows,
            )
            for status, count in result.status_counts.items():
                logger.info("  %s: %d", status, count)

            st.session_state["validation_result"] = result
            st.session_state["master_row_count"] = len(master_df)
            st.session_state["marketplace_row_count"] = len(marketplace_df)
            st.session_state["last_run_selections"] = {
                "sheet_url": sheet_url,
                "price_types": [pt.value for pt in price_types],
                "update_type": update_type.value,
                "marketplace_file": uploaded_file.name,
            }
            st.success(
                f"Validation complete: {result.passed_rows} passed, "
                f"{result.failed_rows} failed out of {result.total_rows} total rows."
            )
        except (GoogleSheetError, ExcelReadError, ValidationError) as exc:
            logger.error("Validation failed: %s", exc)
            st.error(f"Validation failed: {exc}")
        except Exception as exc:  # noqa: BLE001 - surface unexpected errors safely
            logger.exception("Unexpected error during validation.")
            st.error(f"Unexpected error: {exc}")


def _read_uploaded_marketplace_report(uploaded_file) -> pd.DataFrame:
    """
    Adapt ExcelReader (which expects a filesystem Path) to Streamlit's
    in-memory UploadedFile by writing to a temporary file first.
    """
    import tempfile
    from pathlib import Path

    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = Path(tmp.name)

    try:
        return ExcelReader.read_marketplace_report(tmp_path)
    finally:
        tmp_path.unlink(missing_ok=True)


def _render_results() -> None:
    result: Optional[ValidationResult] = st.session_state.get("validation_result")
    if result is None:
        st.info("Enter your data sources and click **Validate** to get started.")
        return

    st.subheader("Validation Summary")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Rows", result.total_rows)
    col2.metric("Ready for Upload", result.passed_rows)
    col3.metric("Failed", result.failed_rows)

    if result.status_counts:
        status_df = pd.DataFrame(
            sorted(result.status_counts.items(), key=lambda kv: kv[0]),
            columns=["Validation Status", "Count"],
        )
        st.dataframe(status_df, use_container_width=True, hide_index=True)

    tab1, tab2, tab3 = st.tabs(
        ["Validation Report", "Upload File (Ready for Upload)", "Error Report"]
    )
    with tab1:
        st.dataframe(
            result.validation_report_df.style.apply(_highlight_status_rows, axis=1),
            use_container_width=True,
            hide_index=True,
        )
    with tab2:
        st.dataframe(result.upload_df, use_container_width=True, hide_index=True)
    with tab3:
        if result.error_report_df.empty:
            st.success("No errors found.")
        else:
            st.dataframe(result.error_report_df, use_container_width=True, hide_index=True)

    st.subheader("Generate Output Files")
    _render_download_buttons(result)


def _render_download_buttons(result: ValidationResult) -> None:
    from config import (
        ERROR_REPORT_FILE_NAME,
        UPLOAD_FILE_NAME,
        VALIDATION_REPORT_FILE_NAME,
        OUTPUT_DIR,
    )

    try:
        writer = ExcelWriter(OUTPUT_DIR)
        upload_bytes = writer.upload_file_bytes(result.upload_df)
        validation_bytes = writer.validation_report_bytes(result.validation_report_df)
        error_bytes = writer.error_report_bytes(result.error_report_df)
    except ExcelWriteError as exc:
        st.error(f"Failed to generate output files: {exc}")
        return

    col1, col2, col3 = st.columns(3)
    col1.download_button(
        "⬇ Download Upload_File.xlsx",
        data=upload_bytes,
        file_name=UPLOAD_FILE_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    col2.download_button(
        "⬇ Download Validation_Report.xlsx",
        data=validation_bytes,
        file_name=VALIDATION_REPORT_FILE_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    col3.download_button(
        "⬇ Download Error_Report.xlsx",
        data=error_bytes,
        file_name=ERROR_REPORT_FILE_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def _render_log_panel() -> None:
    with st.expander("Log", expanded=False):
        st.text_area(
            "automation.log",
            value=_log_text(),
            height=240,
            disabled=True,
            label_visibility="collapsed",
        )


if __name__ == "__main__":
    main()
