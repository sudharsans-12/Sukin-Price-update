"""
config.py

Centralized configuration for the Price Update Automation Tool.
Holds filesystem paths, enumerations, column definitions, and
application-wide constants. No business logic lives here.
"""

from __future__ import annotations

from enum import Enum
from pathlib import Path


# --------------------------------------------------------------------------
# Filesystem paths
# --------------------------------------------------------------------------
BASE_DIR: Path = Path(__file__).resolve().parent
CREDENTIALS_DIR: Path = BASE_DIR / "credentials"
INPUT_DIR: Path = BASE_DIR / "input"
OUTPUT_DIR: Path = BASE_DIR / "output"
LOGS_DIR: Path = BASE_DIR / "logs"

CREDENTIALS_FILE: Path = CREDENTIALS_DIR / "credentials.json"
LOG_FILE: Path = LOGS_DIR / "automation.log"

REQUIRED_DIRS = (CREDENTIALS_DIR, INPUT_DIR, OUTPUT_DIR, LOGS_DIR)


def ensure_directories() -> None:
    """Create all required application directories if they do not exist."""
    for directory in REQUIRED_DIRS:
        directory.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------------------------------
# Application metadata
# --------------------------------------------------------------------------
APP_TITLE: str = "Price Update Automation Tool"
APP_VERSION: str = "1.0.0"


# --------------------------------------------------------------------------
# Enumerations
# --------------------------------------------------------------------------
class PriceType(str, Enum):
    """Price types that can be selected by the user (multi-select)."""

    BAU = "BAU"
    A_PLUS = "A+"
    MEGA = "Mega"

    @classmethod
    def values(cls) -> list[str]:
        return [member.value for member in cls]


class UpdateType(str, Enum):
    """Update type selected by the user (single select)."""

    NORMAL = "Normal"
    CLEARANCE = "Clearance"
    EXCLUSION = "Exclusion"

    @classmethod
    def values(cls) -> list[str]:
        return [member.value for member in cls]


class ValidationStatus(str, Enum):
    """Possible outcomes of the per-row validation logic."""

    READY_FOR_UPLOAD = "Ready for Upload"
    RRP_MISMATCH = "RRP Mismatch"
    SRP_MISMATCH = "SRP Mismatch"
    BOTH_MISMATCH = "Both Mismatch"
    SKU_NOT_FOUND = "SKU Not Found"
    DUPLICATE_SKU = "Duplicate SKU"
    MISSING_DATA = "Missing Data"
    INVALID_PRICE = "Invalid Price"


# --------------------------------------------------------------------------
# Column name constants
# --------------------------------------------------------------------------
# Google Sheet (Master) columns
COL_SELLER_SKU: str = "Seller SKU"
COL_RRP: str = "RRP"
COL_SRP: str = "SRP"
COL_PRICE_TYPE: str = "Price Type"

MASTER_REQUIRED_COLUMNS: list[str] = [
    COL_SELLER_SKU,
    COL_RRP,
    COL_SRP,
    COL_PRICE_TYPE,
]

# Marketplace report columns
COL_PRICE: str = "Price"
COL_SPECIAL_PRICE: str = "Special Price"

MARKETPLACE_REQUIRED_COLUMNS: list[str] = [
    COL_SELLER_SKU,
    COL_PRICE,
    COL_SPECIAL_PRICE,
]

# Validation report columns
COL_MASTER_RRP: str = "Master RRP"
COL_MASTER_SRP: str = "Master SRP"
COL_MARKETPLACE_PRICE: str = "Marketplace Price"
COL_MARKETPLACE_SPECIAL_PRICE: str = "Marketplace Special Price"
COL_VALIDATION_STATUS: str = "Validation Status"
COL_REMARKS: str = "Remarks"

VALIDATION_REPORT_COLUMNS: list[str] = [
    COL_SELLER_SKU,
    COL_MASTER_RRP,
    COL_MASTER_SRP,
    COL_MARKETPLACE_PRICE,
    COL_MARKETPLACE_SPECIAL_PRICE,
    COL_VALIDATION_STATUS,
    COL_REMARKS,
]

# Upload file columns
UPLOAD_FILE_COLUMNS: list[str] = [
    COL_SELLER_SKU,
    COL_RRP,
    COL_SRP,
    COL_PRICE_TYPE,
]

# Error report columns
COL_ERROR_REASON: str = "Error Reason"

ERROR_REPORT_COLUMNS: list[str] = [
    COL_SELLER_SKU,
    COL_ERROR_REASON,
]


# --------------------------------------------------------------------------
# Output file names
# --------------------------------------------------------------------------
UPLOAD_FILE_NAME: str = "Upload_File.xlsx"
VALIDATION_REPORT_FILE_NAME: str = "Validation_Report.xlsx"
ERROR_REPORT_FILE_NAME: str = "Error_Report.xlsx"


# --------------------------------------------------------------------------
# Validation constants
# --------------------------------------------------------------------------
PRICE_TOLERANCE: float = 0.01  # Absolute tolerance for float price comparisons

# --------------------------------------------------------------------------
# Google Sheets constants
# --------------------------------------------------------------------------
GOOGLE_SHEETS_SCOPES: list[str] = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]
