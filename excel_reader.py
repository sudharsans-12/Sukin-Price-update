"""
excel_reader.py

Reads the Marketplace Report (.xlsx) uploaded by the user and returns
a validated pandas DataFrame ready for use by the validator.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd

from config import MARKETPLACE_REQUIRED_COLUMNS


class ExcelReadError(Exception):
    """Raised for any error encountered while reading the marketplace report."""


class ExcelReader:
    """Reads and structurally validates marketplace report Excel files."""

    @staticmethod
    def read_marketplace_report(
        file_path: Path, sheet_name: Optional[str] = 0
    ) -> pd.DataFrame:
        """
        Read the marketplace report Excel file into a DataFrame.

        Args:
            file_path: Path to the uploaded .xlsx marketplace report.
            sheet_name: Sheet name or index to read (defaults to the
                first sheet).

        Returns:
            DataFrame containing the marketplace report data with
            columns stripped of surrounding whitespace, containing at
            least the columns defined in MARKETPLACE_REQUIRED_COLUMNS.

        Raises:
            ExcelReadError: If the file does not exist, is not a
                supported Excel format, cannot be parsed, is empty, or
                is missing required columns.
        """
        if not isinstance(file_path, Path):
            file_path = Path(file_path)

        if not file_path.exists():
            raise ExcelReadError(f"Marketplace report not found: '{file_path}'")

        if file_path.suffix.lower() not in (".xlsx", ".xlsm"):
            raise ExcelReadError(
                f"Unsupported file format '{file_path.suffix}'. "
                "Please upload a .xlsx file."
            )

        try:
            report_df = pd.read_excel(
                file_path, sheet_name=sheet_name, engine="openpyxl", dtype=object
            )
        except ValueError as exc:
            raise ExcelReadError(
                f"Failed to read sheet from '{file_path.name}': {exc}"
            ) from exc
        except Exception as exc:  # noqa: BLE001 - surface any parser failure
            raise ExcelReadError(
                f"Failed to read Excel file '{file_path.name}': {exc}"
            ) from exc

        if isinstance(report_df, dict):
            # sheet_name=None would return a dict of DataFrames; guard
            # against misuse even though the default here is 0.
            raise ExcelReadError(
                "Multiple sheets returned; please specify a single sheet name."
            )

        if report_df.empty:
            raise ExcelReadError(
                f"Marketplace report '{file_path.name}' contains no data rows."
            )

        report_df.columns = [str(col).strip() for col in report_df.columns]

        missing_columns = [
            col for col in MARKETPLACE_REQUIRED_COLUMNS if col not in report_df.columns
        ]
        if missing_columns:
            raise ExcelReadError(
                f"Marketplace report '{file_path.name}' is missing required "
                f"column(s): {', '.join(missing_columns)}"
            )

        return report_df
