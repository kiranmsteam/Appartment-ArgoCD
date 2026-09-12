#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
#  Hitech Citadel Phase 2 – Apartment Management System
#  One-click startup script for Linux / macOS
# ─────────────────────────────────────────────────────────────────────────────

# ══════════════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS CONFIGURATION  (Step 5 of GOOGLE_SETUP.md)
#
#  To use Google Sheets instead of local Excel files:
#    1. Remove the leading '#' from each of the four export lines below.
#    2. Replace the PASTE_... placeholder with your actual value.
#    3. Save this file and run:  bash start.sh
#
#  Where to find each value:
#
#  GOOGLE_SHEETS_CREDENTIALS_FILE
#    Absolute path to the JSON key file downloaded in Step 2c.
#    Example:  /home/yourname/apartment-mgmt-abc123.json
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
# export GOOGLE_SHEETS_CREDENTIALS_FILE="/home/yourname/apartment-mgmt-abc123.json"
# export GOOGLE_SHEET_CONTRIBUTIONS_ID="PASTE_SPREADSHEET_ID_FOR_MONTHLY_CONTRIBUTION"
# export GOOGLE_SHEET_BALANCE_ID="PASTE_SPREADSHEET_ID_FOR_CURRENT_BALANCE_SHEET"
# export GOOGLE_SHEET_CRDR_ID="PASTE_SPREADSHEET_ID_FOR_BALANCE_SHEET_CR_DR"
# export GOOGLE_SHEET_USERS_ID="PASTE_SPREADSHEET_ID_FOR_USER_CREDENTIALS"
# ══════════════════════════════════════════════════════════════════════════════

set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "================================================"
echo "  Hitech Citadel Phase 2"
echo "  Apartment Management System"
echo "================================================"
echo ""

# ── 1. Check Python ──────────────────────────────────────────────────────────
PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" &>/dev/null; then
        ver=$("$cmd" -c "import sys; print(sys.version_info.major)")
        if [ "$ver" -ge 3 ]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo "ERROR: Python 3 is not installed."
    echo ""
    echo "  Linux (Ubuntu/Debian):  sudo apt install python3 python3-pip"
    echo "  macOS:                  brew install python"
    echo "  Or download from:       https://www.python.org/downloads/"
    exit 1
fi

echo "✓ Found Python: $($PYTHON --version)"

# ── 2. Install / upgrade dependencies ────────────────────────────────────────
echo ""
echo "Installing dependencies..."
$PYTHON -m pip install --quiet --upgrade pip
$PYTHON -m pip install --quiet -r requirements.txt
echo "✓ Dependencies installed"

# ── 3. Start the app ─────────────────────────────────────────────────────────
PORT=${PORT:-5000}
echo ""
echo "Starting the app on http://localhost:$PORT ..."
echo ""
echo "  Open your browser and go to:  http://localhost:$PORT"
echo "  Default login:  admin / Welcome  (or flat number e.g. RG1 / Welcome)"
echo ""
echo "  Press Ctrl+C to stop the server."
echo ""

# Try to open the browser (best-effort, don't fail if unavailable)
(sleep 2 && \
    if command -v xdg-open &>/dev/null; then xdg-open "http://localhost:$PORT"; \
    elif command -v open &>/dev/null; then open "http://localhost:$PORT"; \
    fi) &>/dev/null &

PORT=$PORT $PYTHON app.py
