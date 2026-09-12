"""
generate_monthly_balance_sheet.py

Creates a monthly tab inside the "Current Balance_Sheet_2024_25_26" Google Sheet
for Hitech Citadel Phase 2 Flat Owners Association.

All data is read from and written to Google Drive — no local Excel files
are read or produced.

Usage:
  python generate_monthly_balance_sheet.py --month May --year 2026
"""

import json
import os
import re
import sys
import argparse
from datetime import datetime

try:
    import gspread
    from google.oauth2.service_account import Credentials
    _GSPREAD_AVAILABLE = True
except ImportError:
    _GSPREAD_AVAILABLE = False

# ---------------------------------------------------------------------------
# Configuration & Constants
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_FALLBACK_CREDS = os.path.join(SCRIPT_DIR, "hitech-citadel-965501da1a49.json")

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_MONTH_NAME_TO_NUM = {
    'apr': 4, 'april': 4, 'may': 5, 'jun': 6, 'june': 6, 'jul': 7, 'july': 7,
    'aug': 8, 'august': 8, 'sep': 9, 'sept': 9, 'september': 9, 'oct': 10, 'october': 10,
    'nov': 11, 'november': 11, 'dec': 12, 'december': 12, 'jan': 1, 'january': 1,
    'feb': 2, 'febuary': 2, 'february': 2, 'mar': 3, 'march': 3,
}

_MONTH_NUM_TO_NAME = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April',
    5: 'May', 6: 'June', 7: 'July', 8: 'August',
    9: 'September', 10: 'October', 11: 'November', 12: 'December',
}

_FY_CONFIG = [
    ('2026-27', '26-27 Report',    2026, 2027),
    ('2025-26', '25-26 Report',    2025, 2026),
    ('2024-25', '2024-25 Report',  2024, 2025),
]

_BALANCE_SHEET_MONTH_SHEETS = {
    # FY 2024-25
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
    # FY 2025-26
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

_VALID_FLATS = {
    'RG1', 'RG2', 'RG3', 'RF1', 'RF2', 'RF3', 'RS1', 'RS2', 'RS3',
    'VG1', 'VG2', 'VG3', 'VF1', 'VF2', 'VF3', 'VS1', 'VS2', 'VS3',
    'KG1', 'KG2', 'KG3', 'KG4', 'KF1', 'KF2', 'KF3', 'KF4', 'KS1', 'KS2', 'KS3', 'KS4'
}

# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------
def _get_client():
    if not _GSPREAD_AVAILABLE:
        raise RuntimeError("gspread / google-auth not installed.")

    creds_file = os.environ.get("GOOGLE_SHEETS_CREDENTIALS_FILE", "").strip()
    creds_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS_JSON", "").strip()

    if creds_json:
        key_data = json.loads(creds_json)
        creds = Credentials.from_service_account_info(key_data, scopes=_SCOPES)
    elif creds_file and os.path.exists(creds_file):
        creds = Credentials.from_service_account_file(creds_file, scopes=_SCOPES)
    elif os.path.exists(_FALLBACK_CREDS):
        print(f"  Using bundled credentials: {os.path.basename(_FALLBACK_CREDS)}")
        creds = Credentials.from_service_account_file(_FALLBACK_CREDS, scopes=_SCOPES)
    else:
        raise RuntimeError(
            "No Google credentials found. Set GOOGLE_SHEETS_CREDENTIALS_FILE or GOOGLE_SHEETS_CREDENTIALS_JSON."
        )
    return gspread.authorize(creds)

def _open_sheet(gc, env_var):
    sheet_id = os.environ.get(env_var, "").strip()
    if not sheet_id:
        raise RuntimeError(f"Environment variable {env_var} is not set.")
    return gc.open_by_key(sheet_id)

def _find_ws(spreadsheet, name):
    """Return a worksheet by name (whitespace-collapsed case-insensitive), or None."""
    name_norm = re.sub(r'\s+', ' ', name.strip().lower())
    for ws in spreadsheet.worksheets():
        ws_norm = re.sub(r'\s+', ' ', ws.title.strip().lower())
        if ws_norm == name_norm:
            return ws
    return None

def _safe_float(val):
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        val_clean = val.strip().replace(",", "")
        if not val_clean:
            return None
        # strip currencies
        val_clean = val_clean.replace("₹", "").strip()
        try:
            return float(val_clean)
        except ValueError:
            pass
    return None

def _is_cash_apartment(apartment_number: str) -> bool:
    return "cash" in str(apartment_number).strip().lower()

def _is_cash_withdrawal(apartment_number: str, description: str) -> bool:
    val = (str(apartment_number) + " " + str(description)).upper()
    return "WITHRAWAL" in val or "WITHDRAWAL" in val

def get_prev_month_year(month_num: int, year: int) -> tuple[int, int]:
    prev_month = month_num - 1
    prev_year = year
    if prev_month == 0:
        prev_month = 12
        prev_year = year - 1
    return prev_month, prev_year

def get_fy_label(month_num: int, year: int) -> str:
    if month_num >= 4:
        return f"{year}-{str(year+1)[2:]}"
    else:
        return f"{year-1}-{str(year)[2:]}"

def get_sheet_name(month_num: int, year: int) -> str:
    fy = get_fy_label(month_num, year)
    name = _BALANCE_SHEET_MONTH_SHEETS.get((fy, month_num))
    if not name:
        name = f"{_MONTH_NUM_TO_NAME[month_num]} {year}"
    return name

# ---------------------------------------------------------------------------
# Step 1 – Read Opening Balances from Previous Month
# ---------------------------------------------------------------------------
def get_previous_month_values(gc, prev_month_num: int, prev_year: int) -> tuple[float, float]:
    prev_sheet_name = get_sheet_name(prev_month_num, prev_year)
    print(f"  Looking up previous month sheet: '{prev_sheet_name}'")
    spreadsheet = _open_sheet(gc, "GOOGLE_SHEET_BALANCE_ID")
    ws = _find_ws(spreadsheet, prev_sheet_name)
    if ws is None:
        print(f"  WARNING: Previous month sheet '{prev_sheet_name}' not found. Defaulting balances to 0.0.")
        return 0.0, 0.0

    rows = ws.get_all_values()
    opening_balance = None
    cash_in_hand = None

    for row in rows:
        if len(row) < 4:
            continue
        label = row[2].strip().lower()
        d_val = _safe_float(row[3])

        if "balance in bank with corpus" in label and d_val is not None:
            opening_balance = d_val
        if "cash in hand" in label and d_val is not None:
            cash_in_hand = d_val

    if opening_balance is None:
        print("  WARNING: 'Balance in Bank with Corpus Amount' row not found. Defaulting to 0.0.")
        opening_balance = 0.0
    if cash_in_hand is None:
        cash_in_hand = 0.0

    print(f"  Opening Balance (from {prev_sheet_name}) = {opening_balance:,.2f}")
    print(f"  Cash in Hand   (from {prev_sheet_name}) = {cash_in_hand:,.2f}")
    return opening_balance, cash_in_hand

# ---------------------------------------------------------------------------
# Step 2 – Read Monthly Contributions and Debits from CR/DR tab
# ---------------------------------------------------------------------------
def read_crdr_data(gc, month_num: int, year: int) -> tuple[float, float, list, list]:
    # Check for both "Month Year" format (e.g. "May 2026") and short format (e.g. "May 26")
    crdr_sheet = _open_sheet(gc, "GOOGLE_SHEET_CRDR_ID")
    month_name = _MONTH_NUM_TO_NAME[month_num]
    ws_names_to_try = [
        f"{month_name} {year}",
        f"{month_name[:3]} {year}",
        f"{month_name[:3]} {str(year)[2:]}",
        f"{month_name} {str(year)[2:]}",
    ]
    
    ws = None
    for ws_name in ws_names_to_try:
        ws = _find_ws(crdr_sheet, ws_name)
        if ws:
            break

    if ws is None:
        print(f"  WARNING: CR/DR tab for {month_name} {year} not found. Returning empty values.")
        return 0.0, 0.0, [], []

    print(f"  Reading transaction data from Cr Dr tab: '{ws.title}'")
    total_credits = 0.0
    bank_interest = 0.0
    cheque_items = []
    cash_items = []

    rows = ws.get_all_values()
    if not rows:
        return 0.0, 0.0, [], []

    # Detect header columns dynamically
    header = [h.strip().lower() for h in rows[0]]
    apt_idx = header.index("apartment number") if "apartment number" in header else (header.index("appartment") if "appartment" in header else 0)
    date_idx = header.index("date") if "date" in header else 1
    name_idx = header.index("name") if "name" in header else 2
    dr_idx = header.index("dr") if "dr" in header else 3
    cr_idx = header.index("cr") if "cr" in header else 4

    for row in rows[1:]:
        if not row:
            continue
        
        apt_no = row[apt_idx].strip() if len(row) > apt_idx else ""
        apt_clean = apt_no.upper()
        if not apt_clean:
            continue

        # 1. Sum up credits (Contributions / Bank Interest)
        credit = _safe_float(row[cr_idx]) if len(row) > cr_idx else None
        if credit:
            if apt_clean in _VALID_FLATS:
                total_credits += credit
            elif "INTEREST" in apt_clean:
                bank_interest += credit

        # 2. Process debits (Expenses)
        debit = _safe_float(row[dr_idx]) if len(row) > dr_idx else None
        if not debit:
            continue

        raw_date = row[date_idx].strip() if len(row) > date_idx else ""
        desc = row[name_idx].strip() if len(row) > name_idx else ""

        # Format date cleanly (e.g., "5-May-2026")
        formatted_date = ""
        if raw_date:
            for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d %b %Y', '%d-%m-%Y', '%d %B %Y', '%d/%m/%y'):
                try:
                    dt = datetime.strptime(raw_date, fmt)
                    formatted_date = f"{dt.day}-{dt.strftime('%b')}-{dt.year}"
                    break
                except ValueError:
                    pass
            if not formatted_date:
                formatted_date = raw_date  # fallback as-is

        # Is it cash or cheque?
        is_cash = _is_cash_apartment(apt_no)
        
        if _is_cash_withdrawal(apt_no, desc):
            label = "Cash Withdrawal"
        else:
            label = desc or "Payment"
        
        if is_cash:
            cash_items.append((formatted_date, label, debit))
        else:
            cheque_items.append((formatted_date, label, debit))

    # Consolidate BESCOM cheque entries
    bescom_total = 0.0
    bescom_date = ""
    bescom_insert_idx = None
    consolidated_cheques = []
    
    for dt, desc, amt in cheque_items:
        if "BESCOM" in desc.upper():
            bescom_total += amt
            if not bescom_date:
                bescom_date = dt
            if bescom_insert_idx is None:
                bescom_insert_idx = len(consolidated_cheques)
            continue
        consolidated_cheques.append((dt, desc, amt))

    if bescom_total:
        if bescom_insert_idx is None:
            bescom_insert_idx = len(consolidated_cheques)
        consolidated_cheques.insert(
            bescom_insert_idx, (bescom_date, "Payment to BESCOM", bescom_total)
        )
    cheque_items = consolidated_cheques

    print(f"  Monthly Expense Contribution (Credits) = {total_credits:,.2f}")
    if bank_interest:
        print(f"  Bank Interest = {bank_interest:,.2f}")
    print(f"  Debits: {len(cheque_items)} cheque, {len(cash_items)} cash items")
    return total_credits, bank_interest, cheque_items, cash_items

# ---------------------------------------------------------------------------
# Step 3 – Build and Format the Sheet
# ---------------------------------------------------------------------------
def build_balance_sheet(
    gc,
    month_num: int,
    year: int,
    opening_balance: float,
    monthly_contribution: float,
    bank_interest: float,
    cash_in_hand: float,
    cheque_items,
    cash_items,
):
    target_sheet_name = get_sheet_name(month_num, year)
    spreadsheet = _open_sheet(gc, "GOOGLE_SHEET_BALANCE_ID")

    # Find cash withdrawal amount dynamically from cheque_items
    cash_wd_amount = 0.0
    for dt, desc, amt in cheque_items:
        if "CASH WITHDRAWAL" in desc.upper():
            cash_wd_amount = amt
            break

    # Delete existing sheets if they exist (handles potential duplicate sheets with different spacing)
    while True:
        existing = _find_ws(spreadsheet, target_sheet_name)
        if existing is None:
            break
        print(f"  Deleting existing '{existing.title}' tab …")
        spreadsheet.del_worksheet(existing)

    N = ""
    rows_data = []

    def _add(*cells):
        row = list(cells) + [N] * (5 - len(cells))
        rows_data.append(row[:5])
        return len(rows_data)

    month_name = _MONTH_NUM_TO_NAME[month_num]
    month_abbr = month_name[:3]
    last_day = 31
    if month_num in (4, 6, 9, 11):
        last_day = 30
    elif month_num == 2:
        last_day = 29 if year % 4 == 0 else 28

    # ── Fixed opening block (rows 1-14) ──
    _add()                                                                                          # 1
    _add(N, "Hitech Citadel Phase 2 Flat Owners Association")                                      # 2
    _add(N, f"Monthly Collection & Expenditure Details - {month_name} {year}")                      # 3
    _add(N, "Date", "Details", "Amount  Collected/Balance", "Expenses Amount")                      # 4
    r_opening = _add(N, f"01-{month_abbr}-{year}", "Opening Balance in Bank", opening_balance)      # 5
    r_corpus = _add(N, N, "Corpus Amount", 0)                                                      # 6
    _add(N, N, "Expense Available in Bank", f"=D{r_opening}-D{r_corpus}")                           # 7
    _add()                                                                                          # 8
    _add(N, N, "Collection")                                                                        # 9
    r_mc = _add(N, N, "Monthly Expense Contribution", monthly_contribution)                         # 10
    r_bank_int = _add(N, N, "Bank Interest", bank_interest)                                                    # 11
    r_subtotal = _add(N, N, "Sub Total", f"=SUM(D{r_mc}:D{r_bank_int})")                           # 12
    r_total_a = _add(N, N, "Total of A", f"=D{r_opening}+D{r_subtotal}")                          # 13
    _add()                                                                                          # 14

    # ── Cheque section ──
    _add(N, N, "Expenditure Paid by Cheque")                                                        # 15
    cheque_start = len(rows_data) + 1
    if cheque_items:
        for dt, desc, amt in cheque_items:
            _add(N, dt, desc, N, amt)
    else:
        _add(N, N, "Payment to Housekeeping (Mariswamy)")
        _add(N, N, "BESCOM")
        _add()
    cheque_end = len(rows_data)
    _add()
    r_total_b = _add(N, N, "Total of B", N, f"=SUM(E{cheque_start}:E{cheque_end})")
    r_closing_b = _add(N, N, "Closing Balance (A-B)", f"=D{r_total_a}-E{r_total_b}")
    _add()

    # ── Cash section ──
    _add(N, N, "Cash Expense")
    # Date label for cash wd and carry fwd is from previous month
    prev_month_num, prev_year = get_prev_month_year(month_num, year)
    prev_month_abbr = _MONTH_NUM_TO_NAME[prev_month_num][:3]
    
    r_cash_wd = _add(N, N, f"Cash Withdrawal in {month_abbr} {str(year)[2:]}", N, cash_wd_amount)
    r_carry_fwd = _add(N, N, f"Cash Carry Forward from {prev_month_abbr} {str(prev_year)[2:]}", N, cash_in_hand)
    _add()
    
    cash_start = len(rows_data) + 1
    if cash_items:
        for dt, desc, amt in cash_items:
            _add(N, dt, desc, N, amt)
    else:
        _add(N, N, "Towards Garbage Collection")
        _add()
    cash_end = len(rows_data)
    _add()
    _add()
    r_total_c = _add(N, N, "Total of C", N, f"=SUM(E{cash_start}:E{cash_end})")
    _add()
    r_cash_ih = _add(N, N, "Cash in Hand", f"=E{r_cash_wd}+E{r_carry_fwd}-E{r_total_c}")
    _add()

    # ── Closing rows ──
    closing_date = f"{last_day}-{month_abbr}-{year}"
    r_closing_f = _add(N, closing_date, "Closing Balance", f"=D{r_closing_b}")
    r_bal_corp = _add(N, closing_date, "Balance in Bank with Corpus Amount", f"=D{r_closing_f}+D{r_corpus}")
    
    for _ in range(5):
        _add()
    r_treasurer = _add(N, N, N, N, "Treasurer")

    total_rows = len(rows_data)

    print(f"  Creating '{target_sheet_name}' tab …")
    ws = spreadsheet.add_worksheet(title=target_sheet_name, rows=max(total_rows + 5, 60), cols=6)
    sheet_id = ws.id

    ws.update(f"A1:E{total_rows}", rows_data, value_input_option="USER_ENTERED")

    # Formatting
    _HEADER_BG = {"red": 0.851, "green": 0.882, "blue": 0.949}
    _TOTAL_BG = {"red": 0.949, "green": 0.949, "blue": 0.949}
    _WARN_RED = {"red": 1.0, "green": 0.0, "blue": 0.0}
    _ACCT = "#,##0.00"
    _DATE_FMT = "DD-MMM-YYYY"

    def _rng(r0, r1, c0, c1):
        return {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1, "startColumnIndex": c0, "endColumnIndex": c1}

    def _cell_fmt(r, c, fmt, fields):
        return {"repeatCell": {"range": _rng(r, r + 1, c, c + 1), "cell": {"userEnteredFormat": fmt}, "fields": fields}}

    def _row_fmt(r, c0, c1, fmt, fields):
        return {"repeatCell": {"range": _rng(r, r + 1, c0, c1), "cell": {"userEnteredFormat": fmt}, "fields": fields}}

    def _merge(r0, r1, c0, c1):
        return {"mergeCells": {"range": _rng(r0, r1, c0, c1), "mergeType": "MERGE_ALL"}}

    def _col_width(c, px):
        return {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": c, "endIndex": c + 1}, "properties": {"pixelSize": px}, "fields": "pixelSize"}}

    def _row_height(r, px):
        return {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": r, "endIndex": r + 1}, "properties": {"pixelSize": px}, "fields": "pixelSize"}}

    def _border_cell(r, c):
        return {"updateBorders": {"range": _rng(r, r + 1, c, c + 1), "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}}}

    def _num_fmt(r, c, pattern):
        return _cell_fmt(r, c, {"numberFormat": {"type": "NUMBER", "pattern": pattern}}, "userEnteredFormat.numberFormat")

    requests = []

    # Widths
    for col, px in [(0, 24), (1, 110), (2, 320), (3, 165), (4, 140)]:
        requests.append(_col_width(col, px))

    requests.append(_row_height(1, 26))
    requests.append(_row_height(2, 22))
    requests.append(_merge(1, 2, 1, 5))
    requests.append(_merge(2, 3, 1, 5))

    # Society Title
    requests.append(_row_fmt(1, 1, 5, {"textFormat": {"bold": True, "fontSize": 14, "fontFamily": "Calibri"}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}, "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)"))
    # Month Title
    requests.append(_row_fmt(2, 1, 5, {"textFormat": {"bold": True, "fontSize": 12, "fontFamily": "Calibri"}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}, "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)"))

    # Column Headers
    for c in range(1, 5):
        requests.append(_cell_fmt(3, c, {"textFormat": {"bold": True, "fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "CENTER", "backgroundColor": _HEADER_BG}, "userEnteredFormat(textFormat,horizontalAlignment,backgroundColor)"))
        requests.append(_border_cell(3, c))

    def _label(r):
        return _cell_fmt(r, 2, {"textFormat": {"fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "LEFT"}, "userEnteredFormat(textFormat,horizontalAlignment)")

    def _bold_label(r):
        return _cell_fmt(r, 2, {"textFormat": {"bold": True, "fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "LEFT"}, "userEnteredFormat(textFormat,horizontalAlignment)")

    def _section_hdr(r):
        return _cell_fmt(r, 2, {"textFormat": {"bold": True, "underline": True, "fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "LEFT"}, "userEnteredFormat(textFormat,horizontalAlignment)")

    def _total_d(r):
        return [
            _cell_fmt(r, 3, {"textFormat": {"bold": True, "fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "RIGHT", "backgroundColor": _TOTAL_BG, "numberFormat": {"type": "NUMBER", "pattern": _ACCT}}, "userEnteredFormat(textFormat,horizontalAlignment,backgroundColor,numberFormat)"),
            _border_cell(r, 3),
        ]

    def _total_e(r):
        return [
            _cell_fmt(r, 4, {"textFormat": {"bold": True, "fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "RIGHT", "backgroundColor": _TOTAL_BG, "numberFormat": {"type": "NUMBER", "pattern": _ACCT}}, "userEnteredFormat(textFormat,horizontalAlignment,backgroundColor,numberFormat)"),
            _border_cell(r, 4),
        ]

    # Format Rows
    r0 = r_opening - 1
    requests.append(_cell_fmt(r0, 1, {"numberFormat": {"type": "DATE", "pattern": _DATE_FMT}, "horizontalAlignment": "CENTER"}, "userEnteredFormat(numberFormat,horizontalAlignment)"))
    requests.append(_label(r0))
    requests.append(_num_fmt(r0, 3, _ACCT))

    requests.append(_label(r_corpus - 1))
    requests.append(_num_fmt(r_corpus - 1, 3, _ACCT))
    requests.append(_label(r_corpus))

    requests.append(_section_hdr(8))  # Collection header

    requests.append(_label(r_mc - 1))
    txt_fmt = {"fontSize": 11, "fontFamily": "Calibri"}
    if monthly_contribution == 0:
        txt_fmt["foregroundColor"] = _WARN_RED
    requests.append(_cell_fmt(r_mc - 1, 3, {"textFormat": txt_fmt, "horizontalAlignment": "RIGHT", "numberFormat": {"type": "NUMBER", "pattern": _ACCT}}, "userEnteredFormat(textFormat,horizontalAlignment,numberFormat)"))

    requests.append(_label(r_bank_int - 1))
    requests.append(_num_fmt(r_bank_int - 1, 3, _ACCT))

    requests.append(_bold_label(r_subtotal - 1))
    requests.append(_cell_fmt(r_subtotal - 1, 3, {"textFormat": {"bold": True, "fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "RIGHT", "backgroundColor": _TOTAL_BG, "numberFormat": {"type": "NUMBER", "pattern": _ACCT}}, "userEnteredFormat(textFormat,horizontalAlignment,backgroundColor,numberFormat)"))

    requests.append(_bold_label(r_total_a - 1))
    requests += _total_d(r_total_a - 1)

    # Cheque expenditure
    requests.append(_section_hdr(14))
    for r in range(cheque_start - 1, cheque_end):
        requests.append(_cell_fmt(r, 1, {"numberFormat": {"type": "DATE", "pattern": _DATE_FMT}, "horizontalAlignment": "CENTER"}, "userEnteredFormat(numberFormat,horizontalAlignment)"))
        requests.append(_label(r))
        requests.append(_num_fmt(r, 4, _ACCT))

    requests.append(_bold_label(r_total_b - 1))
    requests += _total_e(r_total_b - 1)

    requests.append(_bold_label(r_closing_b - 1))
    requests += _total_d(r_closing_b - 1)

    # Cash expenditure
    requests.append(_section_hdr(r_closing_b))
    requests.append(_label(r_cash_wd - 1))
    requests.append(_num_fmt(r_cash_wd - 1, 4, _ACCT))
    requests.append(_label(r_carry_fwd - 1))
    requests.append(_num_fmt(r_carry_fwd - 1, 4, _ACCT))

    for r in range(cash_start - 1, cash_end):
        requests.append(_cell_fmt(r, 1, {"numberFormat": {"type": "DATE", "pattern": _DATE_FMT}, "horizontalAlignment": "CENTER"}, "userEnteredFormat(numberFormat,horizontalAlignment)"))
        requests.append(_label(r))
        requests.append(_num_fmt(r, 4, _ACCT))

    requests.append(_bold_label(r_total_c - 1))
    requests += _total_e(r_total_c - 1)

    requests.append(_bold_label(r_cash_ih - 1))
    requests.append(_cell_fmt(r_cash_ih - 1, 3, {"textFormat": {"bold": True, "fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "RIGHT", "backgroundColor": _TOTAL_BG, "numberFormat": {"type": "NUMBER", "pattern": _ACCT}}, "userEnteredFormat(textFormat,horizontalAlignment,backgroundColor,numberFormat)"))

    # Closing Balances
    requests.append(_cell_fmt(r_closing_f - 1, 1, {"numberFormat": {"type": "DATE", "pattern": _DATE_FMT}, "horizontalAlignment": "CENTER"}, "userEnteredFormat(numberFormat,horizontalAlignment)"))
    requests.append(_bold_label(r_closing_f - 1))
    requests += _total_d(r_closing_f - 1)

    requests.append(_cell_fmt(r_bal_corp - 1, 1, {"numberFormat": {"type": "DATE", "pattern": _DATE_FMT}, "horizontalAlignment": "CENTER"}, "userEnteredFormat(numberFormat,horizontalAlignment)"))
    requests.append(_bold_label(r_bal_corp - 1))
    requests += _total_d(r_bal_corp - 1)

    requests.append(_cell_fmt(r_treasurer - 1, 4, {"textFormat": {"bold": True, "fontSize": 11, "fontFamily": "Calibri"}, "horizontalAlignment": "CENTER"}, "userEnteredFormat(textFormat,horizontalAlignment)"))

    print("  Applying cell formatting properties …")
    spreadsheet.batch_update({"requests": requests})
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}/edit#gid={sheet_id}"
    print(f"  Successfully built tab '{target_sheet_name}' in Google Sheets!")
    print(f"  URL: {sheet_url}")

# ---------------------------------------------------------------------------
# Main Execution Entry Point
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Generate Monthly Balance Sheet in Google Sheets")
    parser.add_argument("--month", required=True, help="Month name (e.g. May) or number (1-12)")
    parser.add_argument("--year", type=int, required=True, help="Calendar year (e.g. 2026)")
    args = parser.parse_args()

    # Determine month number
    month_input = args.month.strip().lower()
    if month_input.isdigit():
        month_num = int(month_input)
    else:
        month_num = _MONTH_NAME_TO_NUM.get(month_input)

    if not month_num or month_num < 1 or month_num > 12:
        print(f"Error: Invalid month input '{args.month}'")
        sys.exit(1)

    year = args.year
    month_name = _MONTH_NUM_TO_NAME[month_num]

    print("=" * 60)
    print(f"  Generating {month_name} {year} Balance Sheet")
    print("=" * 60)

    print("\n[1] Authenticating with Google Drive API …")
    gc = _get_client()
    print("  Authenticated.")

    prev_month_num, prev_year = get_prev_month_year(month_num, year)
    print(f"\n[2] Reading opening balances from previous month ({_MONTH_NUM_TO_NAME[prev_month_num]} {prev_year}) …")
    opening_balance, cash_in_hand = get_previous_month_values(gc, prev_month_num, prev_year)

    print(f"\n[3] Reading monthly credits and debits from Cr Dr tab …")
    monthly_contribution, bank_interest, cheque_items, cash_items = read_crdr_data(gc, month_num, year)

    print(f"\n[4] Rebuilding and formatting sheet …")
    build_balance_sheet(
        gc,
        month_num,
        year,
        opening_balance,
        monthly_contribution,
        bank_interest,
        cash_in_hand,
        cheque_items,
        cash_items
    )
    print("\nGeneration process complete.")

if __name__ == "__main__":
    main()
