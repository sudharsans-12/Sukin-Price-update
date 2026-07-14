Price Update Automation Tool
Validates e-commerce marketplace prices (Lazada, Shopee, TikTok, Zalora, etc.)
against a Google Sheets Master file before uploading a Price Update file.
Flags mismatches, missing data, duplicates, and invalid prices, then produces
three Excel deliverables: an upload-ready file, a full validation report, and
an error report.
The project ships in two interchangeable front ends that share the same
core logic (`config.py`, `google_sheet.py`, `excel_reader.py`, `validator.py`,
`excel_writer.py`, `logger.py`):
Front end	File	Best for
Desktop app (Tkinter)	`main.py` / `gui.py`	Running locally, no browser needed
Web app (Streamlit)	`streamlit_app.py`	Sharing with a team via a browser link, deployed from GitHub
---
1. Project Structure
```text
PriceUpdateAutomation/
│── main.py              # Desktop app entry point
│── gui.py                # Tkinter GUI
│── streamlit_app.py      # Streamlit web app entry point
│── google_sheet.py       # Google Sheets (Master) loader
│── excel_reader.py       # Marketplace report (.xlsx) reader
│── validator.py          # Vectorized price validation engine
│── excel_writer.py       # Styled Excel output generator
│── config.py              # Paths, enums, column names, constants
│── logger.py              # Logging + run-statistics summary
│── requirements.txt
│── README.md
│── .gitignore
│
├── .streamlit/
│   ├── config.toml               # Theme + server settings
│   └── secrets.toml.example      # Template for Google credentials (Cloud)
│
├── credentials/          # Local service-account JSON (desktop app only)
├── input/                # Sample/uploaded marketplace reports
├── output/               # Generated Upload_File / Validation_Report / Error_Report
└── logs/                 # automation.log
```
---
2. How Validation Works
Master columns (Google Sheet): `Seller SKU`, `RRP`, `SRP`, `Price Type`
Marketplace report columns (.xlsx): `Seller SKU`, `Price`, `Special Price`
Update Type	Compares
Normal	Master RRP = Marketplace Price and Master SRP = Marketplace Special Price
Clearance	Master RRP = Marketplace Price only (SRP ignored)
Exclusion	Master RRP = Marketplace Price only (SRP ignored)
Each Master row (filtered to the selected Price Type(s)) is classified as one
of: `Ready for Upload`, `RRP Mismatch`, `SRP Mismatch`, `Both Mismatch`,
`SKU Not Found`, `Duplicate SKU`, `Missing Data`, `Invalid Price`.
Outputs:
`Upload_File.xlsx` — Ready for Upload rows only (`Seller SKU`, `RRP`, `SRP`, `Price Type`)
`Validation_Report.xlsx` — every row with its status and remarks
`Error_Report.xlsx` — every non-Ready row with a concise error reason
Validation is fully vectorized (pandas/numpy, no per-row Python loops) and
comfortably handles marketplace reports of 100,000+ rows.
---
3. Google Service Account Setup (required for both front ends)
In Google Cloud Console, create a
project (or reuse one) and enable the Google Sheets API and
Google Drive API.
Create a Service Account (IAM & Admin > Service Accounts), then
generate a JSON key for it.
Open your Master Google Sheet and share it with the service account's
`client_email` (found in the JSON key) as Viewer.
Where you put that JSON key depends on which front end you use — see
below.
---
4. Desktop App (Tkinter)
Install
```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```
Configure credentials
Save the service-account JSON key as:
```
credentials/credentials.json
```
Run
```bash
python main.py
```
Using the app
Paste the Google Sheets URL.
Click Browse Excel... and select the Marketplace Report (.xlsx).
Check one or more Price Type boxes (BAU / A+ / Mega).
Choose an Update Type (Normal / Clearance / Exclusion).
Click Validate — results and row counts appear in the log window.
Click Generate Output to write the three Excel files to `output/`.
Reset clears all fields; Exit closes the app.
All actions run on background threads, so the window stays responsive even
on large files. A full run summary (selections, row counts, output files,
errors) is appended to `logs/automation.log` on every generation.
---
5. Web App (Streamlit) — Deploying from GitHub
This is the standard "Streamlit + GitHub" workflow: push the repo to GitHub,
then point Streamlit Community Cloud at it. Streamlit Cloud has no
persistent local disk, so the web app reads credentials from Streamlit
secrets instead of `credentials/credentials.json`, and delivers output
files via in-browser download buttons instead of writing to `output/`.
5.1 Push this project to GitHub
```bash
cd PriceUpdateAutomation
git init
git add .
git commit -m "Initial commit: Price Update Automation Tool"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```
`.gitignore` already excludes `credentials/*.json`, `.streamlit/secrets.toml`,
and everything in `output/`/`logs/`/`input/` — only source code is pushed.
5.2 Deploy on Streamlit Community Cloud
Go to share.streamlit.io and sign in with
GitHub.
Click New app, select your repository, branch `main`, and set the
main file path to `streamlit_app.py`.
Before (or after) deploying, open Advanced settings > Secrets and
paste in your credentials using the format from
`.streamlit/secrets.toml.example`:
```toml
   [gcp_service_account]
   type = "service_account"
   project_id = "..."
   private_key_id = "..."
   private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
   client_email = "...@....iam.gserviceaccount.com"
   client_id = "..."
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
   universe_domain = "googleapis.com"
   ```
Click Deploy. Streamlit Cloud installs `requirements.txt` and starts
`streamlit_app.py` automatically. Any future `git push` to `main`
redeploys the app.
5.3 Run the Streamlit app locally (optional, before deploying)
```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# then edit .streamlit/secrets.toml with your real credentials
streamlit run streamlit_app.py
```
If no `.streamlit/secrets.toml` is present, the app automatically falls back
to `credentials/credentials.json`, so local desktop-style credentials work
too.
5.4 Using the web app
In the sidebar, paste the Google Sheets URL and upload the Marketplace
Report.
Select Price Type(s) and an Update Type.
Click Validate — a summary, status breakdown, and three tabs
(Validation Report / Upload File / Error Report) appear with
color-coded rows.
Click each Download button to save `Upload_File.xlsx`,
`Validation_Report.xlsx`, and `Error_Report.xlsx` to your machine.
Expand the Log panel at the bottom for a run log equivalent to the
desktop app's log window. Reset clears the session.
---
6. Validation Status Reference
Status	Meaning
Ready for Upload	RRP/SRP (per Update Type) match the marketplace report
RRP Mismatch	RRP differs from the marketplace Price
SRP Mismatch	SRP differs from the marketplace Special Price (Normal only)
Both Mismatch	Both RRP and SRP differ (Normal only)
SKU Not Found	Seller SKU exists in Master but not in the marketplace report
Duplicate SKU	Seller SKU appears more than once in Master and/or the report
Missing Data	RRP, SRP, Price, or Special Price is blank
Invalid Price	A price value is negative
---
7. Error Handling
Both front ends validate and clearly report:
Missing required columns in the Master sheet or Marketplace report
Blank or duplicate Seller SKUs
Blank or negative RRP/SRP/Price/Special Price
Google Sheets connection/permission errors (sheet not found, no access,
bad URL, API errors)
Excel file errors (wrong format, unreadable, empty, missing columns)
---
8. Logging
Every run appends a structured summary to `logs/automation.log` (desktop
app) or the in-session Log panel (web app), including: start/end time, user
selections, total/passed/failed row counts, output files created, and any
errors encountered.
---
9. Tech Stack
Python 3.12+, Tkinter, Streamlit, pandas, openpyxl, gspread, google-auth,
pathlib, logging, threading.
