"""google_sheets.py — Google Sheets backend for the Apartment Management app.

Drop-in replacements for every function in app.py that currently reads or
writes local Excel files.  The module is completely optional: if the required
environment variables are not set it returns empty / "not configured" results
and the app falls back to its local-file behaviour automatically.

Required environment variables
───────────────────────────────
GOOGLE_SHEETS_CREDENTIALS_FILE
    Absolute path to the Service Account JSON key file downloaded from
    Google Cloud Console.

GOOGLE_SHEET_CONTRIBUTIONS_ID
    Spreadsheet ID of the "Monthly Contribution_2025_26" Google Sheet.
    (The long alphanumeric string in the sheet's URL.)

GOOGLE_SHEET_BALANCE_ID
    Spreadsheet ID of the "Current Balance_Sheet_2024_25_26" Google Sheet.

GOOGLE_SHEET_CRDR_ID
    Spreadsheet ID of the "balance sheet Cr Dr" Google Sheet.

All three sheet IDs default to empty string, which disables their respective
features gracefully.
"""

import os
import re
import traceback
from datetime import datetime, date

try:
    import gspread
    from gspread.exceptions import APIError as _GSpreadAPIError
    from google.oauth2.service_account import Credentials
    _GSPREAD_AVAILABLE = True
except ImportError:
    _GSPREAD_AVAILABLE = False
    _GSpreadAPIError = None

# ─── OAuth scopes ─────────────────────────────────────────────────────────────

_SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# ─── Module-level diagnostics store ──────────────────────────────────────────
# Populated by the read functions so the admin status page can explain
# exactly why data is missing.

_errors = {}   # key → last error string
_info  = {}    # key → informational message (e.g. row count)


def _format_exc(exc):
    """Format an exception for display, avoiding a doubled class-name prefix.

    gspread's APIError.__str__() already returns 'APIError: [400]: ...',
    so naively prepending type(exc).__name__ would produce
    'APIError: APIError: [400]: ...'.  This helper only adds the prefix when
    the string representation does not already start with the class name.
    """
    s = str(exc)
    cls = type(exc).__name__
    return s if s.startswith(cls) else f'{cls}: {s}'


def _record_error(key, exc):
    _errors[key] = _format_exc(exc)


def _record_info(key, msg):
    _info[key] = msg


# ─── Month helpers (mirrors app.py) ──────────────────────────────────────────

_MONTH_NAME_TO_NUM = {
    'apr': 4, 'april': 4,
    'may': 5,
    'jun': 6, 'june': 6,
    'jul': 7, 'july': 7,
    'aug': 8, 'august': 8,
    'sep': 9, 'sept': 9, 'september': 9,
    'oct': 10, 'october': 10,
    'nov': 11, 'november': 11,
    'dec': 12, 'december': 12,
    'jan': 1, 'january': 1,
    'feb': 2, 'febuary': 2, 'february': 2,
    'mar': 3, 'march': 3,
}

_MONTH_NUM_TO_NAME = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April',
    5: 'May', 6: 'June', 7: 'July', 8: 'August',
    9: 'September', 10: 'October', 11: 'November', 12: 'December',
}

# Monthly Contributions sheet: column index (1-based, A=1) for each month.
# The sheet layout is April in col G (=7) through March in col R (=18).
_CONTRIBUTION_MONTH_COL = {
    4: 7,   # April   → G
    5: 8,   # May     → H
    6: 9,   # June    → I
    7: 10,  # July    → J
    8: 11,  # August  → K
    9: 12,  # Sep     → L
    10: 13, # Oct     → M
    11: 14, # Nov     → N
    12: 15, # Dec     → O
    1: 16,  # Jan     → P
    2: 17,  # Feb     → Q
    3: 18,  # Mar     → R
}


def _fy_from_month_year(month_num, year):
    """Return (fy_label, fy_start_year) for the given month and calendar year.

    Financial year runs April–March:
      month 4-12 in year Y  → FY "Y-(last 2 digits of Y+1)"  e.g. April 2025 → "2025-26"
      month 1-3  in year Y  → FY "(Y-1)-(last 2 digits of Y)" e.g. March 2026 → "2025-26"
    """
    if month_num >= 4:
        fy_start, fy_end = year, year + 1
    else:
        fy_start, fy_end = year - 1, year
    fy_label = f'{fy_start}-{str(fy_end)[2:]}'
    return fy_label, fy_start


def _contributions_tab_for_fy(fy_label, all_ws):
    """Return the worksheet for *fy_label* in the contributions spreadsheet,
    or None if not found.

    Convention:
      - The first/original term ("Sheet1" in the spreadsheet) is returned
        for the initial FY (_CONTRIBUTIONS_INITIAL_FY).
      - Subsequent terms are stored in tabs named exactly after the FY label
        (e.g. "2026-27").
    """
    # Prefer an exact match on the FY label first
    ws = _find_worksheet(None, fy_label, _worksheets=all_ws)
    if ws is not None:
        return ws
    # For the original term, fall back to "Sheet1" (the pre-existing tab name)
    if fy_label == _CONTRIBUTIONS_INITIAL_FY:
        return _find_worksheet(None, 'Sheet1', _worksheets=all_ws)
    return None


# The FY label of the initial contributions spreadsheet (its tab is named
# "Sheet1" rather than the FY label).  New terms get a tab named after their
# FY label (e.g. "2026-27").
_CONTRIBUTIONS_INITIAL_FY = '2025-26'

# ─── Client / auth ────────────────────────────────────────────────────────────

def _get_client():
    """Return an authenticated gspread client, or None if not configured.

    Credentials are resolved from the first available source:
      1. GOOGLE_SHEETS_CREDENTIALS_FILE – path to a local JSON key file
         (for local / self-hosted deployments).
      2. GOOGLE_SHEETS_CREDENTIALS_JSON – the *contents* of the JSON key file
         as a single environment variable string (recommended for cloud
         platforms like Render, Railway, Heroku where you cannot store files).
    """
    if not _GSPREAD_AVAILABLE:
        _record_info('client', 'gspread / google-auth not installed — run: pip install -r requirements.txt')
        return None

    creds_file = os.environ.get('GOOGLE_SHEETS_CREDENTIALS_FILE', '')
    creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS_JSON', '')

    if not creds_file and not creds_json:
        _record_info('client',
            'Neither GOOGLE_SHEETS_CREDENTIALS_FILE nor GOOGLE_SHEETS_CREDENTIALS_JSON is set')
        return None

    try:
        if creds_json:
            # Cloud deployment: JSON content stored directly in an env var
            import json as _json
            key_data = _json.loads(creds_json)
            creds = Credentials.from_service_account_info(key_data, scopes=_SCOPES)
        else:
            if not os.path.exists(creds_file):
                _record_info('client', f'Credentials file not found: {creds_file}')
                return None
            creds = Credentials.from_service_account_file(creds_file, scopes=_SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        _record_error('client', e)
        return None


def is_configured():
    """Return True when the Google Sheets backend is ready to use."""
    return _get_client() is not None


def _open_sheet(client, env_var):
    """Open a spreadsheet by the ID stored in *env_var*, or return None."""
    sheet_id = os.environ.get(env_var, '').strip()
    if not sheet_id:
        _record_info(env_var, f'{env_var} is not set — spreadsheet will not be read')
        return None
    try:
        return client.open_by_key(sheet_id)
    except Exception as e:
        _record_error(env_var, e)
        return None


def _is_not_supported_error(exc):
    """Return True when *exc* is the Google Sheets API error that means the
    spreadsheet is in Excel format and does not support write operations.

    Prefers checking the HTTP status code via gspread's APIError when
    available; falls back to string matching for robustness.
    """
    if _GSpreadAPIError is not None and isinstance(exc, _GSpreadAPIError):
        try:
            return exc.response.status_code == 400 and 'not supported' in str(exc).lower()
        except Exception:
            pass
    return '400' in str(exc) and 'not supported' in str(exc).lower()


def _is_excel_format(spreadsheet):
    """Return True if *spreadsheet* is an Excel file in Drive (not yet converted
    to native Google Sheets format).

    Native Google Sheets support an empty batchUpdate request (a no-op).  Excel
    files in Drive respond with HTTP 400 "This operation is not supported for
    this document".  We use this as a lightweight probe — it modifies nothing.
    """
    try:
        spreadsheet.batch_update({'requests': []})
        return False   # native Google Sheets — write operations are allowed
    except Exception as e:
        if _is_not_supported_error(e):
            return True   # Excel / non-native format
        # Some other error (network, auth) — don't claim it's Excel format
        return False


def _safe_float(val):
    """Convert a cell value to float, return None on failure."""
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        clean = val.replace(',', '').replace('₹', '').strip()
        try:
            return float(clean) if clean else None
        except ValueError:
            return None
    return None


def _find_worksheet(spreadsheet, name, _worksheets=None):
    """Find a worksheet by name.

    Tries:
    1. Exact match
    2. Case-insensitive + whitespace-stripped match
    3. Case-insensitive + collapsed-whitespace match (to handle multiple spaces)

    *_worksheets* is an optional pre-fetched list of Worksheet objects; when
    provided it avoids an extra API round-trip.

    Returns the Worksheet or None.
    """
    worksheets = _worksheets if _worksheets is not None else spreadsheet.worksheets()
    name_stripped = name.strip()
    for ws in worksheets:
        if ws.title == name_stripped:
            return ws
    name_lower = name_stripped.lower()
    for ws in worksheets:
        if ws.title.strip().lower() == name_lower:
            return ws
    # Fallback to collapsed whitespace match (e.g. "April  2026" vs "April 2026")
    name_norm = re.sub(r'\s+', ' ', name_lower)
    for ws in worksheets:
        ws_norm = re.sub(r'\s+', ' ', ws.title.strip().lower())
        if ws_norm == name_norm:
            return ws
    return None


def _parse_sheet_date(name):
    """Parse a sheet-tab name to a date, or return None.

    Handles multiple formats:
    - 'Jul 2024'    → date(2024, 7, 1)   (standard spaced)
    - 'July 2024'   → date(2024, 7, 1)   (full month name)
    - 'Jul24'       → date(2024, 7, 1)   (compact no-space, 2-digit year)
    - 'July24'      → date(2024, 7, 1)   (compact full-name, 2-digit year)
    - 'March 2026'  → date(2026, 3, 1)   (full month + 4-digit year)
    """
    name = name.strip()

    # ── spaced format: "Jul 2024" or "March 2026" ─────────────────────────
    parts = name.split()
    if len(parts) >= 2:
        month_num = _MONTH_NAME_TO_NUM.get(parts[0].lower())
        if month_num is not None:
            try:
                year_s = parts[-1]
                year = int(year_s) if len(year_s) == 4 else (2000 + int(year_s))
                return date(year, month_num, 1)
            except (ValueError, TypeError):
                pass

    # ── compact format: "Jul24" or "July24" or "March2026" ────────────────
    if len(parts) == 1:
        # Try longest-prefix match first so "september" beats "sep"
        for prefix in sorted(_MONTH_NAME_TO_NUM.keys(), key=len, reverse=True):
            if name.lower().startswith(prefix):
                rest = name[len(prefix):]
                try:
                    year = int(rest) if len(rest) == 4 else (2000 + int(rest))
                    return date(year, _MONTH_NAME_TO_NUM[prefix], 1)
                except (ValueError, TypeError):
                    pass

    return None

# ─── Monthly Contributions ────────────────────────────────────────────────────

def get_monthly_contributions(fy_label=None):
    """Mirrors app.get_monthly_contributions() using Google Sheets.

    Expected spreadsheet layout (matches the Excel file):
    - Column E  (index 4): Flat number  (e.g. "RS3")
    - Columns G–R (indices 6–17): monthly amounts for
      April, May, June, July, August, Sep, Oct, Nov, DEC, JAN, FEB, MAR

    Tab selection: uses the tab for *fy_label* when provided, otherwise
    uses the current financial year (determined from today's date),
    falling back to "Sheet1" for 2025-26.
    """
    _errors.pop('contributions', None)
    _info.pop('contributions', None)
    client = _get_client()
    if client is None:
        return {}
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_CONTRIBUTIONS_ID')
    if spreadsheet is None:
        return {}

    months = ['April', 'May', 'June', 'July', 'August', 'Sep', 'Oct', 'Nov',
              'DEC', 'JAN', 'FEB', 'MAR']
    result = {}
    try:
        all_ws = spreadsheet.worksheets()

        # Determine which FY tab to use
        if fy_label is None:
            today = date.today()
            fy_label, _ = _fy_from_month_year(today.month, today.year)
        current_fy_label = fy_label
        ws = _contributions_tab_for_fy(current_fy_label, all_ws)
        tab_used = ws.title if ws else None

        # Final fallback: Sheet1
        if ws is None:
            ws = _find_worksheet(None, 'Sheet1', _worksheets=all_ws)
            tab_used = 'Sheet1 (fallback)'

        if ws is None:
            tab_names = [w.title for w in all_ws]
            _record_info('contributions',
                f'No tab found for FY {current_fy_label} and "Sheet1" is '
                f'also missing. Available tabs: {tab_names}')
            return {}

        rows = ws.get_all_values()
        skipped = 0
        for row in rows:
            flat = row[4].strip() if len(row) > 4 else ''
            if not flat or flat.upper() != flat or len(flat) > 4:
                skipped += 1
                continue
            contrib = {}
            for i, month in enumerate(months):
                raw = row[6 + i] if 6 + i < len(row) else ''
                contrib[month] = _safe_float(raw)
            result[flat] = contrib
        _record_info('contributions',
            f'Read {len(result)} flats from "{tab_used}" (FY {current_fy_label}), '
            f'{skipped} rows skipped')
    except Exception as e:
        _record_error('contributions', e)
    return result


def get_monthly_contributions_last_updated():
    """Return a human-readable 'last fetched' timestamp."""
    if not is_configured():
        return None
    return 'Google Sheets (live)'


def update_monthly_contribution(month_num, year, flat, amount):
    """Write a flat's monthly contribution amount to the correct cell.

    The Monthly Contributions sheet is a grid where:
      - Column E holds the flat number (e.g. "RS3")
      - Columns G–R hold the monthly amounts April–March

    This function locates the row for *flat* (column E) and the column for
    *month_num* (columns G–R, see _CONTRIBUTION_MONTH_COL) and writes
    *amount* to that intersection — aligning the new entry the same way all
    previous entries are arranged in the sheet.

    Tab selection (only-create-when-needed):
      - Determines the financial year from *month_num*/*year*.
      - Uses the existing tab for that FY if one exists ("Sheet1" for
        the 2025-26 term, or a tab named after the FY label for later terms).
      - Creates a new tab only when no tab for this FY exists yet, by
        duplicating the first sheet ("Sheet1") to preserve its format and
        flat list, clearing only the monthly-amount columns (G–R) so the
        new term starts blank.

    Returns None on success or an error string on failure.
    """
    client = _get_client()
    if client is None:
        return None
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_CONTRIBUTIONS_ID')
    if spreadsheet is None:
        return None

    if _is_excel_format(spreadsheet):
        err_msg = (
            '[400] Cannot write to the Monthly Contributions spreadsheet — '
            'it is still in Excel (.xlsx) format in Google Drive. '
            'Fix: File \u2192 Save as Google Sheets, then update '
            'GOOGLE_SHEET_CONTRIBUTIONS_ID and restart the app.'
        )
        _record_error('contributions_write', Exception(err_msg))
        return err_msg

    col_index = _CONTRIBUTION_MONTH_COL.get(month_num)
    if col_index is None:
        return f'Invalid month number: {month_num}'

    fy_label, _ = _fy_from_month_year(month_num, year)

    try:
        all_ws = spreadsheet.worksheets()

        # ── Find or create the FY tab ─────────────────────────────────────────
        ws = _contributions_tab_for_fy(fy_label, all_ws)

        if ws is None:
            # New term: create a tab by duplicating Sheet1 (preserves the flat
            # list, column layout, and formatting) then clear the monthly-amount
            # cells (G2:R<last_row>) so the new term starts fresh.
            sheet1 = _find_worksheet(None, 'Sheet1', _worksheets=all_ws)
            if sheet1 is not None:
                spreadsheet.duplicate_sheet(
                    source_sheet_id=sheet1.id,
                    insert_sheet_index=len(all_ws),
                    new_sheet_name=fy_label,
                )
                ws = spreadsheet.worksheet(fy_label)
                # Surgically clear only the monthly-amount cells (cols G–R)
                # for rows that contain actual flat data (identified by col E
                # holding an all-uppercase alphanumeric code of ≤4 chars).
                # This preserves every header/title row — including the row
                # that labels columns G–R with month names (April … MAR) —
                # which would be wiped by a blanket G2:R{n} clear.
                all_rows = ws.get_all_values()
                flat_row_indices = [
                    i for i, row in enumerate(all_rows, start=1)
                    if len(row) > 4
                    and row[4].strip()
                    and row[4].strip() == row[4].strip().upper()
                    and row[4].strip().isalnum()
                    and len(row[4].strip()) <= 4
                ]
                if flat_row_indices:
                    ws.batch_clear([f'G{r}:R{r}' for r in flat_row_indices])
            else:
                # No Sheet1 to copy — create a minimal sheet
                ws = spreadsheet.add_worksheet(
                    title=fy_label, rows=60, cols=18
                )
                ws.update('E1', [['Flat']])

        # ── Locate the flat's row (search column E) ───────────────────────────
        flat_upper = flat.strip().upper()
        col_e_values = ws.col_values(5)   # column E, 1-indexed
        row_index = None
        for i, cell_val in enumerate(col_e_values, start=1):
            if cell_val.strip().upper() == flat_upper:
                row_index = i
                break

        if row_index is None:
            err_msg = (
                f'Flat "{flat}" not found in column E of the '
                f'"{ws.title}" tab in the Monthly Contributions sheet. '
                'Check that the flat number matches exactly.'
            )
            _record_error('contributions_write', Exception(err_msg))
            return err_msg

        # ── Write the amount to (row_index, col_index) ────────────────────────
        ws.update_cell(row_index, col_index, amount)
        return None  # success

    except Exception as e:
        if _is_not_supported_error(e):
            err_msg = (
                '[400] Cannot write to the Monthly Contributions spreadsheet — '
                'it is in Excel format. Convert to native Google Sheets '
                '(File \u2192 Save as Google Sheets) and update '
                'GOOGLE_SHEET_CONTRIBUTIONS_ID.'
            )
            _record_error('contributions_write', Exception(err_msg))
            return err_msg
        err_msg = _format_exc(e)
        _record_error('contributions_write', e)
        return err_msg




# These mirror the constants in app.py
_FY_CONFIG = [
    ('2026-27', '26-27 Report',    2026, 2027),
    ('2025-26', '25-26 Report',    2025, 2026),
    ('2024-25', '2024-25 Report',  2024, 2025),
]

_BALANCE_SHEET_MONTH_SHEETS = {
    ('2024-25',  4): 'April24',
    ('2024-25',  5): 'May24',
    ('2024-25',  6): 'June24',
    ('2024-25',  7): 'July24',
    ('2024-25',  8): 'Aug24',
    ('2024-25',  9): 'Sept24',
    ('2024-25', 10): 'Oct24',
    ('2024-25', 11): 'Nov24',
    ('2024-25', 12): 'Dec24',
    ('2024-25',  1): 'Jan 25',
    ('2024-25',  2): 'Feb 25',
    ('2024-25',  3): 'Mar 25',
    ('2025-26',  4): 'Apr 25',
    ('2025-26',  5): 'May 25',
    ('2025-26',  6): 'June 25',
    ('2025-26',  7): 'July 25',
    ('2025-26',  8): 'Aug 25',
    ('2025-26',  9): 'Sept 2025',
    ('2025-26', 10): 'Oct 2025',
    ('2025-26', 11): 'Nov 2025',
    ('2025-26', 12): 'Dec 2025',
    ('2025-26',  1): 'Jan 2026',
    ('2025-26',  2): 'Feb 2026',
    ('2025-26',  3): 'March 2026',
    # FY 2026-27
    ('2026-27',  4): 'April 2026',
    ('2026-27',  5): 'May 2026',
    ('2026-27',  6): 'June 2026',
    ('2026-27',  7): 'July 2026',
    ('2026-27',  8): 'August 2026',
    ('2026-27',  9): 'Sep 2026',
    ('2026-27', 10): 'Oct 2026',
    ('2026-27', 11): 'Nov 2026',
    ('2026-27', 12): 'Dec 2026',
    ('2026-27',  1): 'Jan 2027',
    ('2026-27',  2): 'Feb 2027',
    ('2026-27',  3): 'Mar 2027',
}


def _derive_income_expense_from_month_rows(rows):
    """Scan a per-month detail tab and return (income, expense) totals.

    The tab layout written by generate_*_balance_sheet.py:
      col C (index 2): description text
      col D (index 3): collected / income amounts
      col E (index 4): expense amounts

    Income  = value on the row whose col-C description is "sub total"  (col D).
    Expense = sum of rows whose col-C description is "total of b" or
              "total of c"  (col E).

    All comparisons are case-insensitive so minor formatting differences in the
    sheet do not break the lookup.
    """
    income = 0.0
    expense = 0.0
    for row in rows:
        if len(row) < 3:
            continue
        desc = row[2].strip().lower()
        if desc == 'sub total':
            val = _safe_float(row[3]) if len(row) > 3 else None
            if val is not None:
                income = val
        elif desc in ('total of b', 'total of c'):
            val = _safe_float(row[4]) if len(row) > 4 else None
            if val is not None:
                expense += val
    return income, expense


def get_balance_sheets(fy_filter=None):
    """Mirrors app.get_balance_sheets() using Google Sheets.

    Expected spreadsheet layout (matches the Excel file):
    - Report sheets like "25-26 Report", "2024-25 Report"
    - Each report sheet: row 1 is a header, then rows with
      [month_name, ?, income, expense]
    """
    _errors.pop('balance', None)
    _info.pop('balance', None)
    client = _get_client()
    if client is None:
        return [], []
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_BALANCE_ID')
    if spreadsheet is None:
        return [], []

    all_tab_names = []
    all_entries = []
    available_fys = []
    missing_reports = []

    try:
        all_ws = spreadsheet.worksheets()
        all_tab_names = [ws.title for ws in all_ws]

        for fy_label, report_sheet, fy_start, fy_end in _FY_CONFIG:
            ws = _find_worksheet(spreadsheet, report_sheet, _worksheets=all_ws)
            if ws is None:
                missing_reports.append(report_sheet)
                continue
            available_fys.append(fy_label)
            if fy_filter and fy_label != fy_filter:
                continue
            try:
                rows = ws.get_all_values()
                monthly = {}
                for row in rows[1:]:  # skip header
                    if not row or not row[0]:
                        continue
                    month_str = row[0].strip().lower()
                    month_num = _MONTH_NAME_TO_NUM.get(month_str)
                    if month_num is None:
                        continue
                    income  = _safe_float(row[2]) or 0 if len(row) > 2 else 0
                    expense = _safe_float(row[3]) or 0 if len(row) > 3 else 0
                    if month_num not in monthly:
                        cal_year = fy_start if month_num >= 4 else fy_end
                        monthly[month_num] = {
                            'fy':            fy_label,
                            'month_num':     month_num,
                            'year':          cal_year,
                            'display_month': _MONTH_NUM_TO_NAME[month_num] + ' ' + str(cal_year),
                            'income':        0,
                            'expense':       0,
                            'sheet_ref':     _BALANCE_SHEET_MONTH_SHEETS.get((fy_label, month_num), ''),
                        }
                    monthly[month_num]['income']  += income
                    monthly[month_num]['expense'] += expense
                all_entries.extend(monthly.values())
            except Exception as e:
                _record_error('balance', e)

        # ── Fallback: synthesise entries from per-month detail tabs ──────────
        # If a month has a per-month detail tab in the spreadsheet but no
        # corresponding row in a report tab (e.g. the "26-27 Report" summary
        # tab hasn't been created or updated yet), derive the income/expense
        # totals directly from the detail tab so the overview still shows it.
        covered = {(e['fy'], e['month_num']) for e in all_entries}
        fallback_added = []
        for (fy_label, month_num), tab_name in _BALANCE_SHEET_MONTH_SHEETS.items():
            if (fy_label, month_num) in covered:
                continue
            detail_ws = _find_worksheet(spreadsheet, tab_name, _worksheets=all_ws)
            if detail_ws is None:
                continue
            # This FY has at least one detail tab → include it in available_fys.
            if fy_label not in available_fys:
                available_fys.append(fy_label)
            # Respect fy_filter for the entries list.
            if fy_filter and fy_label != fy_filter:
                continue
            try:
                rows = detail_ws.get_all_values()
                income, expense = _derive_income_expense_from_month_rows(rows)
                if income == 0 and expense == 0:
                    continue  # tab exists but has no usable data yet
                fy_cfg = next((c for c in _FY_CONFIG if c[0] == fy_label), None)
                if fy_cfg is None:
                    continue
                _, _, fy_start, fy_end = fy_cfg
                cal_year = fy_start if month_num >= 4 else fy_end
                all_entries.append({
                    'fy':            fy_label,
                    'month_num':     month_num,
                    'year':          cal_year,
                    'display_month': _MONTH_NUM_TO_NAME[month_num] + ' ' + str(cal_year),
                    'income':        income,
                    'expense':       expense,
                    'sheet_ref':     tab_name,
                })
                covered.add((fy_label, month_num))
                fallback_added.append(tab_name)
            except Exception as e:
                _record_error('balance', e)

        # Keep available_fys in _FY_CONFIG order (newest first).
        _fy_order = {c[0]: i for i, c in enumerate(_FY_CONFIG)}
        available_fys.sort(key=lambda x: _fy_order.get(x, 999))

        msg = f'Tabs in spreadsheet: {all_tab_names}. Found {len(all_entries)} months.'
        if missing_reports:
            msg += f' Report tabs not found: {missing_reports} (checked case-insensitively).'
        if fallback_added:
            msg += f' Derived from detail tabs (no report row): {fallback_added}.'
        _record_info('balance', msg)
    except Exception as e:
        _record_error('balance', e)

    all_entries.sort(key=lambda x: (x['year'], x['month_num']), reverse=True)
    return all_entries, available_fys


def get_balance_sheet_detail(fy, month_num):
    """Mirrors app.get_balance_sheet_detail() using Google Sheets.

    Expected layout for per-month tabs (matches the Excel file):
    [skip_col, date, description, income, expense, ...]
    """
    _errors.pop('balance_detail', None)
    _info.pop('balance_detail', None)
    client = _get_client()
    if client is None:
        return []
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_BALANCE_ID')
    if spreadsheet is None:
        return []

    sheet_name = _BALANCE_SHEET_MONTH_SHEETS.get((fy, month_num))
    if not sheet_name:
        _record_info('balance_detail',
            f'No sheet mapping for fy={fy} month={month_num} in BALANCE_SHEET_MONTH_SHEETS')
        return []

    rows_out = []
    try:
        all_ws = spreadsheet.worksheets()
        ws = _find_worksheet(spreadsheet, sheet_name, _worksheets=all_ws)
        if ws is None:
            all_tab_names = [w.title for w in all_ws]
            _record_info('balance_detail',
                f'Tab "{sheet_name}" not found (fy={fy}, month={month_num}). '
                f'Available tabs: {all_tab_names}')
            return []
        rows = ws.get_all_values()
        for row in rows:
            if len(row) < 4:
                continue
            desc    = row[2]
            inc_raw = row[3] if len(row) > 3 else ''
            exp_raw = row[4] if len(row) > 4 else ''
            if not desc and not inc_raw and not exp_raw:
                continue
            date_raw = row[1]
            rows_out.append({
                'date':        date_raw,
                'description': str(desc).strip(),
                'income':      _safe_float(inc_raw),
                'expense':     _safe_float(exp_raw),
            })
        _record_info('balance_detail', f'Read {len(rows_out)} rows from "{sheet_name}"')
    except Exception as e:
        _record_error('balance_detail', e)
    return rows_out

# ─── Cr Dr Transactions ───────────────────────────────────────────────────────

def get_all_transactions(month_filter=None):
    """Mirrors app.get_all_transactions() using Google Sheets.

    Expected layout per monthly tab:
    Row 1: header  [Apartment Number, Date, Name, Dr, Cr]
    Subsequent rows: data

    available_months is returned newest-first (matches the Excel behaviour).
    """
    _errors.pop('crdr', None)
    _info.pop('crdr', None)
    client = _get_client()
    if client is None:
        return [], []
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_CRDR_ID')
    if spreadsheet is None:
        return [], []

    _skip = {'Template', 'Refernce Sheet'}  # 'Refernce' is intentionally misspelled — matches the original Excel tab name

    # Common strings that appear in column-A header rows (used to skip stray header rows)
    _header_values = {
        'apartment number', 'apartment', 'flat', 'flat no', 'flat no.',
        'apt', 'apt no', 'apt no.', 'sl no', 'sl.no.', 'no', 'no.',
    }

    result = []
    available_months = []

    try:
        all_ws = spreadsheet.worksheets()
        all_tab_names = [ws.title for ws in all_ws]
        dated_sheets = []
        unparsed = []
        for ws in all_ws:
            if ws.title in _skip:
                continue
            d = _parse_sheet_date(ws.title)
            if d:
                dated_sheets.append((d, ws.title, ws))
            else:
                unparsed.append(ws.title)
        dated_sheets.sort(key=lambda x: x[0], reverse=True)
        available_months = [sn for _, sn, _ in dated_sheets]

        for _, sheet_name, ws in dated_sheets:
            if month_filter and sheet_name != month_filter:
                continue
            rows = ws.get_all_values()
            for row in rows[1:]:  # skip header row
                if not row or all(c == '' for c in row):
                    continue
                if len(row) < 4:
                    continue
                apt = row[0].strip()
                if not apt or apt.lower() in _header_values or apt.startswith('=') or apt.startswith('#'):
                    continue
                extra_val = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ''
                result.append({
                    'month':     sheet_name,
                    'apartment': apt,
                    'name':      row[2] if len(row) > 2 else '',
                    'debit':     _safe_float(row[3]) or 0,
                    'credit':    _safe_float(row[4]) or 0 if len(row) > 4 else 0,
                    'date':      row[1] if len(row) > 1 else '',
                    'extra':     extra_val,
                })

        msg = (f'All tabs: {all_tab_names}. '
               f'Parsed as month tabs: {[sn for _, sn, _ in dated_sheets]}. '
               f'Skipped (not parseable as month): {unparsed}. '
               f'Total transaction rows: {len(result)}.')
        _record_info('crdr', msg)
    except Exception as e:
        _record_error('crdr', e)

    return result, available_months


def get_flat_payment_history(flat_no):
    """Mirrors app.get_flat_payment_history() using Google Sheets."""
    txs, _ = get_all_transactions()
    return [t for t in txs if t['apartment'].upper() == flat_no.upper()]


def append_transaction(month_num, year, apartment, date_str, description, debit, credit):
    """Mirrors app.append_to_crdr_xlsx() — writes a new row to the Cr Dr sheet.

    * If a tab for this month/year already exists (matched by date regardless
      of the exact tab name), it appends to that tab.
    * If no tab exists for this month/year, it duplicates the "Template" tab
      and renames it to "MonthName Year" (e.g. "March 2026"), then inserts it
      in chronological order before "Refernce Sheet".

    Returns None on success or an error string on failure so the caller can
    surface the problem to the user.
    """
    client = _get_client()
    if client is None:
        return None
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_CRDR_ID')
    if spreadsheet is None:
        return None

    # ── Proactive Excel-format check ──────────────────────────────────────────
    # Exact cause of APIError [400]: the spreadsheet stored in Google Drive is
    # still in Excel (.xlsx) format (i.e. it was uploaded but never converted
    # to a native Google Sheets document via File → Save as Google Sheets).
    # Google Drive allows *reading* such files through the Sheets API, but
    # *all write operations* (append_row, duplicate_sheet, add_worksheet,
    # reorder_worksheets, batch_update) are blocked and return HTTP 400
    # "This operation is not supported for this document".
    # We probe this before attempting any write so the error message is
    # immediate and actionable rather than a raw API exception.
    if _is_excel_format(spreadsheet):
        err_msg = (
            '[400] Cannot write to the Cr Dr spreadsheet — it is still in Excel '
            '(.xlsx) format in Google Drive. This is the exact cause of the '
            '"This operation is not supported for this document" error. '
            'Fix: open the file in Google Sheets, click '
            'File \u2192 Save as Google Sheets to create a native copy, then '
            'update GOOGLE_SHEET_CRDR_ID in your startup script to the new '
            "file's Spreadsheet ID and restart the app."
        )
        _record_error('append_transaction', Exception(err_msg))
        return err_msg

    _skip = {'Template', 'Refernce Sheet'}
    target_ws = None
    new_sheet_name = f'{_MONTH_NUM_TO_NAME[month_num]} {year}'

    try:
        all_ws = spreadsheet.worksheets()

        # Try to find an existing tab for this month/year
        for ws in all_ws:
            if ws.title in _skip:
                continue
            d = _parse_sheet_date(ws.title)
            if d and d.month == month_num and d.year == year:
                target_ws = ws
                break

        if target_ws is None:
            # Always start fresh — never copy the Template tab
            target_ws = spreadsheet.add_worksheet(
                title=new_sheet_name,
                rows=100,
                cols=5,
            )
            target_ws.append_row(
                ['Apartment Number', 'Date', 'Name', 'Dr', 'Cr'],
                value_input_option='RAW',
            )

            # Reorder: move new sheet to correct chronological position
            _reorder_crdr_sheets(spreadsheet, new_sheet_name, month_num, year)

        # Build the row values.
        # Date: parse from any recognised format and reformat as DD/MM/YYYY so
        # that Google Sheets stores it as a proper date value in every locale,
        # matching the way dates appear in manually-entered existing rows.
        formatted_date = date_str or ''
        if date_str:
            date_parsed = False
            for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d %b %Y', '%d-%m-%Y'):
                try:
                    formatted_date = datetime.strptime(date_str, fmt).strftime('%d/%m/%Y')
                    date_parsed = True
                    break
                except ValueError:
                    pass
            if not date_parsed:
                _record_error('append_transaction',
                    Exception(f'Unrecognised date format: {date_str!r} — storing as-is'))

        # Numeric values: pass as floats rather than strings so that Google
        # Sheets stores them as numbers (not left-aligned text), consistent
        # with manually-entered rows.
        dr_num = float(debit)  if isinstance(debit,  (int, float)) and debit  > 0 else ''
        cr_num = float(credit) if isinstance(credit, (int, float)) and credit > 0 else ''
        row = [apartment or '', formatted_date, description or '', dr_num, cr_num]
        target_ws.append_row(row, value_input_option='USER_ENTERED')
        return None  # success

    except Exception as e:
        if _is_not_supported_error(e):
            err_msg = (
                '[400] Cannot write to this spreadsheet — it is still in Excel '
                'format in Google Drive. Open it in Google Sheets, then click '
                'File → Save as Google Sheets to create a native copy, and '
                "update GOOGLE_SHEET_CRDR_ID to the new file's Spreadsheet ID."
            )
            _record_error('append_transaction', Exception(err_msg))
            return err_msg
        err_msg = _format_exc(e)
        _record_error('append_transaction', e)
        return err_msg


def delete_transaction(month_num, year, apartment, date_str, description, debit, credit):
    """Delete the first row in the Cr Dr tab that matches all provided fields.

    Matches on apartment (case-insensitive), description (case-insensitive),
    and Dr/Cr amounts (within ±0.01 tolerance).  Date is used as a secondary
    tiebreaker when multiple candidate rows exist.

    Returns None on success (including "no matching row found"), or an error
    string on failure so the caller can surface it to the user.
    """
    client = _get_client()
    if client is None:
        return None
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_CRDR_ID')
    if spreadsheet is None:
        return None

    try:
        all_ws = spreadsheet.worksheets()
        _skip = {'Template', 'Refernce Sheet'}
        target_ws = None
        for ws in all_ws:
            if ws.title in _skip:
                continue
            d = _parse_sheet_date(ws.title)
            if d and d.month == month_num and d.year == year:
                target_ws = ws
                break

        if target_ws is None:
            return None  # tab doesn't exist — nothing to delete

        rows = target_ws.get_all_values()
        flat_upper = (apartment or '').strip().upper()
        desc_lower = (description or '').strip().lower()
        dr_val = float(debit)  if isinstance(debit,  (int, float)) and debit  > 0 else 0.0
        cr_val = float(credit) if isinstance(credit, (int, float)) and credit > 0 else 0.0

        row_to_delete = None
        for i, row in enumerate(rows, start=1):
            if i == 1:   # skip header row
                continue
            if len(row) < 4:
                continue
            if not row[0].strip():
                continue
            row_apt  = row[0].strip().upper()
            row_desc = row[2].strip().lower() if len(row) > 2 else ''
            row_dr   = _safe_float(row[3]) or 0.0
            row_cr   = (_safe_float(row[4]) or 0.0) if len(row) > 4 else 0.0
            if (row_apt == flat_upper
                    and row_desc == desc_lower
                    and abs(row_dr - dr_val) < 0.01
                    and abs(row_cr - cr_val) < 0.01):
                row_to_delete = i
                break  # delete first match

        if row_to_delete is not None:
            target_ws.delete_rows(row_to_delete)
        return None  # success

    except Exception as e:
        err_msg = _format_exc(e)
        _record_error('delete_transaction', e)
        return err_msg


def clear_monthly_contribution(month_num, year, flat):
    """Clear a flat's monthly contribution cell (set to empty string).

    This is the inverse of update_monthly_contribution and is called when
    a payment entry is deleted so the Monthly Contributions sheet stays in
    sync with the app's payment records.

    Returns None on success or an error string on failure.
    """
    client = _get_client()
    if client is None:
        return None
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_CONTRIBUTIONS_ID')
    if spreadsheet is None:
        return None

    col_index = _CONTRIBUTION_MONTH_COL.get(month_num)
    if col_index is None:
        return f'Invalid month number: {month_num}'

    fy_label, _ = _fy_from_month_year(month_num, year)

    try:
        all_ws = spreadsheet.worksheets()
        ws = _contributions_tab_for_fy(fy_label, all_ws)
        if ws is None:
            return None  # tab doesn't exist — nothing to clear

        flat_upper = flat.strip().upper()
        col_e_values = ws.col_values(5)   # column E, 1-indexed
        row_index = None
        for i, cell_val in enumerate(col_e_values, start=1):
            if cell_val.strip().upper() == flat_upper:
                row_index = i
                break

        if row_index is None:
            return None  # flat not found — nothing to clear

        ws.update_cell(row_index, col_index, '')
        return None  # success

    except Exception as e:
        err_msg = _format_exc(e)
        _record_error('contributions_clear', e)
        return err_msg


def _reorder_crdr_sheets(spreadsheet, sheet_name, month_num, year):
    """Move *sheet_name* to its correct chronological position."""
    _skip = {'Template', 'Refernce Sheet', sheet_name}
    all_ws = spreadsheet.worksheets()
    new_date = date(year, month_num, 1)

    dated = []
    ref_index = None
    for ws in all_ws:
        if ws.title == 'Refernce Sheet':
            ref_index = ws.index
        if ws.title in _skip:
            continue
        d = _parse_sheet_date(ws.title)
        if d:
            dated.append((d, ws.index, ws.title))

    dated.sort(key=lambda x: x[0])

    # Find first sheet that is after the new date
    insert_before = None
    for d, idx, _ in dated:
        if d > new_date:
            insert_before = idx
            break

    target_index = insert_before if insert_before is not None else (
        ref_index if ref_index is not None else len(all_ws) - 1
    )

    try:
        target_ws = spreadsheet.worksheet(sheet_name)
        spreadsheet.reorder_worksheets(
            [ws for ws in spreadsheet.worksheets()
             if ws.title != sheet_name] + [target_ws]
        )
        # Use the Sheets API to move to exact position
        body = {
            'requests': [{
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': target_ws.id,
                        'index': target_index,
                    },
                    'fields': 'index',
                }
            }]
        }
        spreadsheet.batch_update(body)
    except Exception as e:
        if _is_not_supported_error(e):
            _record_error('reorder_sheets', Exception(
                '[400] Cannot reorder sheets — spreadsheet is in Excel format. '
                'Convert it to native Google Sheets (File \u2192 Save as Google Sheets) '
                'and update GOOGLE_SHEET_CRDR_ID.'
            ))
        else:
            _record_error('reorder_sheets', e)


# ─── Diagnostics ─────────────────────────────────────────────────────────────

# ─── Users sheet ─────────────────────────────────────────────────────────────

_USERS_HEADER = ['username', 'password_hash', 'role', 'name', 'flat', 'block',
                 'must_change_password', 'email']


def get_users():
    """Read all users from the Users Google Sheet.

    Returns a dict in the same format as data/users.json, or {} if the sheet
    is not configured or cannot be reached.

    Sheet layout (Row 1 = header row, Row 2+ = data):
        username | password_hash | role | name | flat | block | must_change_password | email
    The ``email`` column is optional for backwards compatibility.
    """
    client = _get_client()
    if client is None:
        return {}
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_USERS_ID')
    if spreadsheet is None:
        return {}
    try:
        ws = spreadsheet.sheet1
        rows = ws.get_all_values()
        if len(rows) < 2:
            return {}
        header = [h.strip().lower() for h in rows[0]]
        try:
            i_user = header.index('username')
            i_pwd  = header.index('password_hash')
            i_role = header.index('role')
            i_name = header.index('name')
            i_flat = header.index('flat')
            i_blk  = header.index('block')
            i_must = header.index('must_change_password')
        except ValueError as exc:
            _record_error('users', exc)
            return {}
        i_email = header.index('email') if 'email' in header else None
        users = {}
        for row in rows[1:]:
            def _c(i, _row=row):
                return _row[i].strip() if i < len(_row) else ''
            username = _c(i_user)
            if not username:
                continue
            must_raw = _c(i_must).upper()
            users[username] = {
                'role':                 _c(i_role),
                'name':                 _c(i_name),
                'flat':                 _c(i_flat) or None,
                'block':                _c(i_blk) or None,
                'password':             _c(i_pwd),
                'must_change_password': must_raw in ('TRUE', '1', 'YES'),
                'email':                _c(i_email) if i_email is not None else '',
            }
        return users
    except Exception as exc:
        _record_error('users', exc)
        return {}


def save_users(users_dict):
    """Write the complete users dict to the Users Google Sheet.

    Clears the sheet and rewrites header + one row per user.
    Passwords must already be hashed before calling this function —
    plain-text passwords are never written to the sheet.
    Returns True on success, False otherwise.
    """
    client = _get_client()
    if client is None:
        return False
    spreadsheet = _open_sheet(client, 'GOOGLE_SHEET_USERS_ID')
    if spreadsheet is None:
        return False
    try:
        ws = spreadsheet.sheet1
        rows = [_USERS_HEADER]
        for username, user in users_dict.items():
            rows.append([
                username,
                user.get('password', ''),
                user.get('role', ''),
                user.get('name', ''),
                user.get('flat') or '',
                user.get('block') or '',
                str(user.get('must_change_password', True)),
                user.get('email') or '',
            ])
        ws.clear()
        ws.update('A1', rows)
        return True
    except Exception as exc:
        _record_error('users', exc)
        return False


def _probe_users(spreadsheet):
    """Return a one-line summary of what was found in the users sheet."""
    try:
        ws = spreadsheet.sheet1
        rows = ws.get_all_values()
        if not rows:
            return 'Sheet is empty — no header row found'
        header = [h.strip().lower() for h in rows[0]]
        required = set(_USERS_HEADER)
        missing = required - set(header)
        if missing:
            return f'Header row is missing columns: {sorted(missing)}'
        user_count = sum(1 for r in rows[1:] if r and any(c.strip() for c in r))
        return f'Header row ✓ — {user_count} user row(s) found'
    except Exception as exc:
        return f'Error reading sheet: {_format_exc(exc)}'


# ─── Diagnostics ──────────────────────────────────────────────────────────────

def get_diagnostics():
    """Return a structured dict that describes the current connection state.

    Called by the admin status page.  Does NOT raise exceptions — every
    error is captured inside the returned dict so the page can display it.
    """
    result = {
        'gspread_installed': _GSPREAD_AVAILABLE,
        'client_ok': False,
        'client_error': '',
        'sheets': {},
    }

    client = _get_client()
    if client is None:
        result['client_error'] = _errors.get('client', '') or _info.get('client', 'Unknown error')
        return result

    result['client_ok'] = True

    sheet_specs = [
        ('users', 'GOOGLE_SHEET_USERS_ID',
         'User Credentials (hashed passwords)',
         lambda sp: _probe_users(sp)),
        ('contributions', 'GOOGLE_SHEET_CONTRIBUTIONS_ID',
         'Monthly Contributions',
         lambda sp: _probe_contributions(sp)),
        ('balance', 'GOOGLE_SHEET_BALANCE_ID',
         'Balance Sheet (25-26 Report / 2024-25 Report)',
         lambda sp: _probe_balance(sp)),
        ('crdr', 'GOOGLE_SHEET_CRDR_ID',
         'Cr Dr Transactions',
         lambda sp: _probe_crdr(sp)),
    ]

    for key, env_var, label, probe_fn in sheet_specs:
        sheet_id = os.environ.get(env_var, '').strip()
        entry = {
            'label': label,
            'env_var': env_var,
            'sheet_id': sheet_id,
            'ok': False,
            'excel_format': False,
            'tab_names': [],
            'summary': '',
            'error': '',
        }
        if not sheet_id:
            entry['error'] = f'{env_var} is not set'
        else:
            try:
                spreadsheet = client.open_by_key(sheet_id)
                entry['tab_names'] = [ws.title for ws in spreadsheet.worksheets()]
                entry['excel_format'] = _is_excel_format(spreadsheet)
                entry['ok'] = True
                entry['summary'] = probe_fn(spreadsheet)
            except Exception as e:
                entry['error'] = _format_exc(e)
        result['sheets'][key] = entry

    return result


def _probe_contributions(spreadsheet):
    """Return a one-line summary of what was found in the contributions sheet."""
    all_ws = spreadsheet.worksheets()
    today = date.today()
    current_fy_label, _ = _fy_from_month_year(today.month, today.year)
    ws = _contributions_tab_for_fy(current_fy_label, all_ws)
    if ws is None:
        tab_names = [w.title for w in all_ws]
        return (f'No tab found for current FY {current_fy_label} '
                f'(checked "{current_fy_label}" and "Sheet1"). '
                f'Available tabs: {tab_names}')
    rows = ws.get_all_values()
    flat_count = sum(
        1 for row in rows
        if len(row) > 4 and row[4].strip() and
           row[4].strip().upper() == row[4].strip() and
           len(row[4].strip()) <= 5
    )
    return (f'Tab "{ws.title}" found ✓ (FY {current_fy_label}) — '
            f'{flat_count} flat rows detected')


def _probe_balance(spreadsheet):
    all_ws = spreadsheet.worksheets()
    lines = []
    for fy_label, report_sheet, _, _ in _FY_CONFIG:
        ws = _find_worksheet(spreadsheet, report_sheet, _worksheets=all_ws)
        if ws is None:
            lines.append(f'Report tab "{report_sheet}" NOT FOUND')
        else:
            rows = ws.get_all_values()
            month_rows = sum(
                1 for r in rows[1:]
                if r and r[0].strip().lower() in _MONTH_NAME_TO_NUM
            )
            lines.append(f'"{report_sheet}" ✓ — {month_rows} month rows')
    return '; '.join(lines) if lines else 'No report tabs found'


def _probe_crdr(spreadsheet):
    _skip = {'Template', 'Refernce Sheet'}
    all_ws = spreadsheet.worksheets()
    dated = [ws.title for ws in all_ws
             if ws.title not in _skip and _parse_sheet_date(ws.title) is not None]
    unparsed = [ws.title for ws in all_ws
                if ws.title not in _skip and _parse_sheet_date(ws.title) is None]
    lines = []
    if dated:
        lines.append(f'{len(dated)} month tab(s) recognised: {dated[:5]}{"…" if len(dated) > 5 else ""}')
    if unparsed:
        lines.append(f'{len(unparsed)} tab(s) NOT recognised as months: {unparsed}')
    return '; '.join(lines) if lines else 'No tabs found'
