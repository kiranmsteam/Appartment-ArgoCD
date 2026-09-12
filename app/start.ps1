# ─────────────────────────────────────────────────────────────────────────────
#  Hitech Citadel Phase 2 – Apartment Management System
#  One-click startup script for Windows PowerShell
# ─────────────────────────────────────────────────────────────────────────────

# ══════════════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS CONFIGURATION  (Step 5 of GOOGLE_SETUP.md)
#
#  To use Google Sheets instead of local Excel files:
#    1. Remove the leading '#' from each of the four $env: lines below.
#    2. Replace the PASTE_... placeholder with your actual value.
#    3. Save this file and run it (right-click → Run with PowerShell).
#
#  Where to find each value:
#
#  GOOGLE_SHEETS_CREDENTIALS_FILE
#    Full path to the JSON key file downloaded in Step 2c.
#    Example:  C:\Users\YourName\Downloads\apartment-mgmt-abc123.json
#
#  GOOGLE_SHEET_CONTRIBUTIONS_ID
#    Open "Monthly Contribution_2025_26" in Google Sheets.
#    The browser URL looks like:
#      https://docs.google.com/spreadsheets/d/YOUR_ID_IS_HERE/edit
#    Copy the long string between /d/ and /edit.
#
#  GOOGLE_SHEET_BALANCE_ID
#    Same idea, but for "Current Balance_Sheet_2024_25_26".
#
#  GOOGLE_SHEET_CRDR_ID
#    Same idea, but for "balance sheet Cr Dr".
#
#  GOOGLE_SHEET_USERS_ID
#    A dedicated Google Sheet for user credentials (hashed passwords).
#    Create a new blank Google Sheet, share it with the service account,
#    and paste its Spreadsheet ID here.  Passwords are always stored as
#    SHA-256 hashes — plain-text passwords are never written to the sheet.
# ══════════════════════════════════════════════════════════════════════════════
# $env:GOOGLE_SHEETS_CREDENTIALS_FILE = "C:\Users\YourName\Downloads\apartment-mgmt-abc123.json"
# $env:GOOGLE_SHEET_CONTRIBUTIONS_ID  = "PASTE_SPREADSHEET_ID_FOR_MONTHLY_CONTRIBUTION"
# $env:GOOGLE_SHEET_BALANCE_ID        = "PASTE_SPREADSHEET_ID_FOR_CURRENT_BALANCE_SHEET"
# $env:GOOGLE_SHEET_CRDR_ID           = "PASTE_SPREADSHEET_ID_FOR_BALANCE_SHEET_CR_DR"
# $env:GOOGLE_SHEET_USERS_ID          = "PASTE_SPREADSHEET_ID_FOR_USER_CREDENTIALS"
# ══════════════════════════════════════════════════════════════════════════════

Set-Location -Path $PSScriptRoot

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Hitech Citadel Phase 2" -ForegroundColor Cyan
Write-Host "  Apartment Management System" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# ── 1. Check Python ────────────────────────────────────────────────────────────
$python = $null
foreach ($cmd in @("python", "python3", "py")) {
    try {
        $ver = & $cmd -c "import sys; print(sys.version_info.major)" 2>$null
        if ($ver -ge 3) { $python = $cmd; break }
    } catch {}
}

if (-not $python) {
    Write-Host "ERROR: Python 3 is not installed or not in PATH." -ForegroundColor Red
    Write-Host ""
    Write-Host "  Download Python 3 from: https://www.python.org/downloads/"
    Write-Host "  During installation, check 'Add Python to PATH'"
    Read-Host "Press Enter to exit"
    exit 1
}

$pyVer = & $python --version 2>&1
Write-Host "✓ Found $pyVer" -ForegroundColor Green

# ── 2. Install dependencies ────────────────────────────────────────────────────
Write-Host ""
Write-Host "Installing dependencies..." -ForegroundColor Yellow
& $python -m pip install --quiet --upgrade pip
& $python -m pip install --quiet -r requirements.txt
Write-Host "✓ Dependencies installed" -ForegroundColor Green

# ── 3. Resolve port (default 5000, override by setting $env:PORT before running)
$port = if ($env:PORT) { $env:PORT } else { "5000" }
$env:PORT = $port

# ── 4. Launch browser then start the app ──────────────────────────────────────
Write-Host ""
Write-Host "Starting the app on http://localhost:$port ..." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Open your browser and go to:  " -NoNewline
Write-Host "http://localhost:$port" -ForegroundColor Cyan
Write-Host "  Default login:  admin / Welcome  (or flat number e.g. RG1 / Welcome)"
Write-Host ""
Write-Host "  Press Ctrl+C to stop the server."
Write-Host ""

# Open browser after 2-second delay (non-blocking)
Start-Job -ScriptBlock {
    param($p)
    Start-Sleep -Seconds 2
    Start-Process "http://localhost:$p"
} -ArgumentList $port | Out-Null

& $python app.py
