@echo off
REM ─────────────────────────────────────────────────────────────────────────
REM  Hitech Citadel Phase 2 – Apartment Management System
REM  One-click startup script for Windows (Command Prompt)
REM ─────────────────────────────────────────────────────────────────────────

REM ══════════════════════════════════════════════════════════════════════════
REM  GOOGLE SHEETS CONFIGURATION  (Step 5 of GOOGLE_SETUP.md)
REM
REM  To use Google Sheets instead of local Excel files:
REM    1. Remove the "REM " from the start of each of the four set lines below.
REM    2. Replace the PASTE_... placeholder with your actual value.
REM    3. Save this file and double-click it to start the app.
REM
REM  Where to find each value:
REM
REM  GOOGLE_SHEETS_CREDENTIALS_FILE
REM    The full path to the JSON key file you downloaded in Step 2c.
REM    Example:  C:\Users\YourName\Downloads\apartment-mgmt-abc123.json
REM
REM  GOOGLE_SHEET_CONTRIBUTIONS_ID
REM    Open "Monthly Contribution_2025_26" in Google Sheets.
REM    Look at the browser address bar – the URL is:
REM      https://docs.google.com/spreadsheets/d/YOUR_ID_IS_HERE/edit
REM    Copy the long string between /d/ and /edit.
REM
REM  GOOGLE_SHEET_BALANCE_ID
REM    Same idea, but for the "Current Balance_Sheet_2024_25_26" spreadsheet.
REM
REM  GOOGLE_SHEET_CRDR_ID
REM    Same idea, but for the "balance sheet Cr Dr" spreadsheet.
REM
REM  GOOGLE_SHEET_USERS_ID
REM    A dedicated Google Sheet for user credentials (hashed passwords).
REM    Create a new blank Google Sheet, share it with the service account,
REM    and paste its Spreadsheet ID here.  Passwords are always stored as
REM    SHA-256 hashes — plain-text passwords are never written to the sheet.
REM ══════════════════════════════════════════════════════════════════════════
REM set GOOGLE_SHEETS_CREDENTIALS_FILE=C:\Users\YourName\Downloads\apartment-mgmt-abc123.json
REM set GOOGLE_SHEET_CONTRIBUTIONS_ID=PASTE_SPREADSHEET_ID_FOR_MONTHLY_CONTRIBUTION
REM set GOOGLE_SHEET_BALANCE_ID=PASTE_SPREADSHEET_ID_FOR_CURRENT_BALANCE_SHEET
REM set GOOGLE_SHEET_CRDR_ID=PASTE_SPREADSHEET_ID_FOR_BALANCE_SHEET_CR_DR
REM set GOOGLE_SHEET_USERS_ID=PASTE_SPREADSHEET_ID_FOR_USER_CREDENTIALS
REM ══════════════════════════════════════════════════════════════════════════

cd /d "%~dp0"

echo ================================================
echo   Hitech Citadel Phase 2
echo   Apartment Management System
echo ================================================
echo.

REM ── 1. Check Python ────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo.
    echo   Download Python 3 from: https://www.python.org/downloads/
    echo   During installation, check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PY_VER=%%i
echo Found %PY_VER%

REM ── 2. Install dependencies ────────────────────────────────────────────────
echo.
echo Installing dependencies...
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt
echo Dependencies installed.

REM ── 3. Set port (default 5000, override by setting PORT before running) ────
if not defined PORT set PORT=5000

REM ── 4. Start the app ───────────────────────────────────────────────────────
echo.
echo Starting the app on http://localhost:%PORT% ...
echo.
echo   Open your browser and go to:  http://localhost:%PORT%
echo   Default login:  admin / Welcome  (or flat number e.g. RG1 / Welcome)
echo.
echo   Press Ctrl+C to stop the server.
echo.

REM Open browser after a short delay
timeout /t 2 /nobreak >nul
start "" "http://localhost:%PORT%"

python app.py
pause
