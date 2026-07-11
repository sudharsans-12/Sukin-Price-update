from pathlib import Path

# Project Root
BASE_DIR = Path(__file__).resolve().parent

# Folder Structure
INPUT_FOLDER = BASE_DIR / "input"
OUTPUT_FOLDER = BASE_DIR / "output"
LOG_FOLDER = BASE_DIR / "logs"
CREDENTIAL_FOLDER = BASE_DIR / "credentials"

# Google Credential JSON
GOOGLE_CREDENTIAL = CREDENTIAL_FOLDER / "credentials.json"

# Output Files
VALIDATION_REPORT = OUTPUT_FOLDER / "Validation_Report.xlsx"
UPLOAD_FILE = OUTPUT_FOLDER / "Upload_File.xlsx"
ERROR_REPORT = OUTPUT_FOLDER / "Error_Report.xlsx"

# Required Columns
MASTER_COLUMNS = [
    "Seller SKU",
    "RRP",
    "SRP",
    "Price Type"
]

MARKETPLACE_COLUMNS = [
    "Seller SKU",
    "Price",
    "Special Price"
]

# Validation Status
READY = "Ready for Upload"
