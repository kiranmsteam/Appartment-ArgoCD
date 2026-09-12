import os
import json
import hashlib
import secrets
import base64
import smtplib
import urllib.request
import urllib.error
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, date
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_from_directory
import openpyxl
from openpyxl import Workbook
from werkzeug.utils import secure_filename

# Optional Google Sheets backend — activated when env-vars are set.
import google_sheets as _gs

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'apartment-mgmt-secret-2025')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
APARTMENT_DIR = os.path.join(BASE_DIR, 'Apartment')
UPLOAD_FOLDER = os.path.join(APARTMENT_DIR, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

USERS_FILE = os.path.join(DATA_DIR, 'users.json')
ANNOUNCEMENTS_FILE = os.path.join(DATA_DIR, 'announcements.json')
PAYMENTS_FILE = os.path.join(DATA_DIR, 'payments.json')

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
ANNOUNCEMENT_UPLOAD_FOLDER = os.path.join(APARTMENT_DIR, 'announcement_uploads')
os.makedirs(ANNOUNCEMENT_UPLOAD_FOLDER, exist_ok=True)

# ─── Month / FY constants ─────────────────────────────────────────────────────

MONTH_NAME_TO_NUM = {
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

MONTH_NUM_TO_NAME = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April',
    5: 'May', 6: 'June', 7: 'July', 8: 'August',
    9: 'September', 10: 'October', 11: 'November', 12: 'December',
}

# Full month names used for new Cr Dr sheet titles (e.g. "March 2026")
MONTH_FULL_NAMES = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April',
    5: 'May', 6: 'June', 7: 'July', 8: 'August',
    9: 'September', 10: 'October', 11: 'November', 12: 'December',
}

# Available financial years: (fy_label, report_sheet, fy_start_year, fy_end_year)
# Newest first
FY_CONFIG = [
    ('2026-27', '26-27 Report',    2026, 2027),
    ('2025-26', '25-26 Report',    2025, 2026),
    ('2024-25', '2024-25 Report',  2024, 2025),
]

# Maps (fy_label, calendar_month_num) → per-month sheet name in Current Balance Sheet
BALANCE_SHEET_MONTH_SHEETS = {
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

# ─── Helpers ──────────────────────────────────────────────────────────────────

def hash_password(pwd):
    return hashlib.sha256(pwd.encode()).hexdigest()

def normalize_username(username):
    """Normalise a username from a login/forgot-password form.

    Converts to uppercase except for the special 'admin' and 'admin1' accounts.
    """
    upper = username.strip().upper()
    if upper == 'ADMIN':
        return 'admin'
    if upper == 'ADMIN1':
        return 'admin1'
    return upper

def generate_random_password(length=16):
    """Return a cryptographically random password of the given length.

    Uses only unambiguous letters and digits (no characters that look
    similar, e.g. O/0 or l/1/I) so the password is easy to read and
    type from an email.
    """
    alphabet = (
        'abcdefghjkmnpqrstuvwxyz'   # lowercase, no ambiguous chars
        'ABCDEFGHJKLMNPQRSTUVWXYZ'  # uppercase, no ambiguous chars
        '23456789'                   # digits, no 0/1
    )
    return ''.join(secrets.choice(alphabet) for _ in range(length))

def send_email(to_address, subject, body):
    """Send a plain-text e-mail.

    Strategy (tried in order):
      1. Gmail API via the Google service account (HTTPS, no SMTP port needed).
         Only attempted when EMAIL_USE_GMAIL_API=true is explicitly set.
         Requires a Google Workspace admin account with domain-wide delegation.
         Not suitable for regular @gmail.com accounts.
      2. Gmail SMTP using EMAIL_SENDER + EMAIL_APP_PASSWORD.
         Tries port 465 (SSL) first, then port 587 (STARTTLS) as a fallback.
         Works with any regular Gmail account + App Password.
         Note: some cloud platforms block all outbound SMTP (e.g. Render free
         tier) — use Strategy 3 instead when this happens.
      3. SendGrid HTTP API using SENDGRID_API_KEY.
         Uses HTTPS — works on every platform, including those that block SMTP.
         Free tier: 100 emails/day at https://sendgrid.com
         EMAIL_SENDER must be a verified sender in your SendGrid account.

    Returns True on success, or an error-message string on failure.
    """
    sender = os.environ.get('EMAIL_SENDER', '').strip()
    if not sender:
        return (
            'EMAIL_SENDER is not configured. '
            'Please set it to the email address emails should be sent from.'
        )

    # ── Strategy 1: Gmail API via service account ──────────────────────────
    # Must be explicitly opted in to avoid using Sheets credentials for Gmail.
    gmail_api_result = None
    if os.environ.get('EMAIL_USE_GMAIL_API', '').strip().lower() == 'true':
        gmail_api_result = _send_via_gmail_api(to_address, subject, body, sender)
        if gmail_api_result is True:
            return True

    # ── Strategy 2: Gmail SMTP (port 465 SSL, then port 587 STARTTLS) ───────
    smtp_result = None
    app_password = os.environ.get('EMAIL_APP_PASSWORD', '').strip()
    if app_password:
        smtp_result = _send_via_smtp(to_address, subject, body, sender, app_password)
        if smtp_result is True:
            return True

    # ── Strategy 3: SendGrid HTTP API ─────────────────────────────────────
    # Works on platforms that block all outbound SMTP (e.g. Render free tier).
    # Sign up for a free API key at https://sendgrid.com (100 emails/day).
    # EMAIL_SENDER must be a verified sender identity in your SendGrid account.
    sendgrid_api_key = os.environ.get('SENDGRID_API_KEY', '').strip()
    if sendgrid_api_key:
        sg_result = _send_via_sendgrid(to_address, subject, body, sender, sendgrid_api_key)
        if sg_result is True:
            return True
        errors = []
        if gmail_api_result is not None:
            errors.append(f'Gmail API: {gmail_api_result}')
        if smtp_result is not None:
            errors.append(f'SMTP: {smtp_result}')
        errors.append(f'SendGrid: {sg_result}')
        return ' | '.join(errors)

    # No SendGrid configured — return the best error we have.
    if smtp_result is not None:
        if gmail_api_result is not None:
            return f'Gmail API: {gmail_api_result} | SMTP: {smtp_result}'
        return (
            f'SMTP error: {smtp_result}. '
            'Your hosting platform appears to block outbound SMTP. '
            'Set SENDGRID_API_KEY (free at https://sendgrid.com) to send '
            'email over HTTPS instead, or use EMAIL_USE_GMAIL_API=true with '
            'a Google Workspace account.'
        )

    if gmail_api_result is not None:
        return (
            f'Gmail API failed: {gmail_api_result}. '
            'Set EMAIL_APP_PASSWORD for SMTP or SENDGRID_API_KEY for SendGrid.'
        )
    return (
        'Email is not configured. '
        'Set EMAIL_APP_PASSWORD (Gmail App Password from '
        'https://myaccount.google.com/apppasswords) or '
        'SENDGRID_API_KEY (free at https://sendgrid.com).'
    )


def _build_mime_message(to_address, subject, body, sender):
    """Return a MIMEMultipart email ready to send."""
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to_address
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    return msg


def _send_via_gmail_api(to_address, subject, body, sender):
    """Send via Gmail API using the configured Google service account.

    The service account must have domain-wide delegation for the
    'https://www.googleapis.com/auth/gmail.send' scope, with *sender*
    as the delegated (impersonated) user.

    Returns True on success or an error string on failure.
    """
    try:
        from google.oauth2 import service_account as _sa
        import google.auth.transport.requests as _gatr
    except ImportError:
        return 'google-auth not installed'

    creds_json_str = os.environ.get('GOOGLE_SHEETS_CREDENTIALS_JSON', '').strip()
    creds_file = os.environ.get('GOOGLE_SHEETS_CREDENTIALS_FILE', '').strip()

    if creds_json_str:
        try:
            key_data = json.loads(creds_json_str)
        except json.JSONDecodeError as exc:
            return f'Malformed JSON in GOOGLE_SHEETS_CREDENTIALS_JSON: {exc}'
        except Exception as exc:
            return f'Cannot parse GOOGLE_SHEETS_CREDENTIALS_JSON: {exc}'
    elif creds_file and os.path.exists(creds_file):
        try:
            with open(creds_file) as f:
                key_data = json.load(f)
        except Exception as exc:
            return f'Cannot read credentials file: {exc}'
    else:
        return 'No Google service account credentials configured'

    try:
        creds = _sa.Credentials.from_service_account_info(
            key_data,
            scopes=['https://www.googleapis.com/auth/gmail.send'],
            subject=sender,
        )
        creds.refresh(_gatr.Request())
    except Exception as exc:
        return (
            f'Service account auth failed (ensure domain-wide delegation is '
            f'enabled for https://www.googleapis.com/auth/gmail.send and '
            f'{sender!r} is an authorised delegated user): {exc}'
        )

    msg = _build_mime_message(to_address, subject, body, sender)
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()

    try:
        data = json.dumps({'raw': raw}).encode('utf-8')
        req = urllib.request.Request(
            'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
            data=data,
            method='POST',
        )
        req.add_header('Authorization', f'Bearer {creds.token}')
        req.add_header('Content-Type', 'application/json')
        with urllib.request.urlopen(req, timeout=20):
            return True
    except urllib.error.HTTPError as exc:
        err = exc.read().decode('utf-8', errors='replace')
        return f'Gmail API HTTP {exc.code}: {err}'
    except Exception as exc:
        return str(exc)


def _send_via_smtp(to_address, subject, body, sender, app_password):
    """Send via Gmail SMTP, trying port 465 (SSL) then port 587 (STARTTLS).

    Some cloud platforms block port 465 but allow 587, so both are tried.
    Returns True on success or an error string describing all failures.
    """
    msg = _build_mime_message(to_address, subject, body, sender)
    raw = msg.as_string()

    # ── Try port 465 (SSL) ────────────────────────────────────────────────
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=10) as server:
            server.login(sender, app_password)
            server.sendmail(sender, to_address, raw)
        return True
    except Exception as exc_465:
        err_465 = str(exc_465)

    # ── Try port 587 (STARTTLS) ───────────────────────────────────────────
    try:
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender, app_password)
            server.sendmail(sender, to_address, raw)
        return True
    except Exception as exc_587:
        err_587 = str(exc_587)

    return f'port 465: {err_465} | port 587: {err_587}'


def _send_via_sendgrid(to_address, subject, body, sender, api_key):
    """Send via the SendGrid v3 Mail Send HTTP API.

    Uses only stdlib urllib — no extra packages required.
    Returns True on success or an error string on failure.

    Prerequisites (one-time setup at https://app.sendgrid.com):
      1. Create a free account and generate an API key with 'Mail Send' permission.
      2. Add EMAIL_SENDER as a verified Sender Identity (Settings → Sender Authentication).
      3. Set SENDGRID_API_KEY to the generated key in your environment / hosting dashboard.
    """
    payload = json.dumps({
        'personalizations': [{'to': [{'email': to_address}]}],
        'from': {'email': sender},
        'subject': subject,
        'content': [{'type': 'text/plain', 'value': body}],
    }).encode('utf-8')

    req = urllib.request.Request(
        'https://api.sendgrid.com/v3/mail/send',
        data=payload,
        method='POST',
    )
    req.add_header('Authorization', f'Bearer {api_key}')
    req.add_header('Content-Type', 'application/json')

    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status in (200, 202):
                return True
            return f'unexpected HTTP {resp.status}'
    except urllib.error.HTTPError as exc:
        err = exc.read().decode('utf-8', errors='replace')
        return f'HTTP {exc.code}: {err}'
    except Exception as exc:
        return str(exc)


def load_json(path, default):
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return default

def save_json(path, data):
    with open(path, 'w') as f:
        json.dump(data, f, indent=2)

def load_users():
    """Load users from Google Sheets (if configured) falling back to local JSON."""
    users = {}
    if os.environ.get('GOOGLE_SHEET_USERS_ID', '').strip() and _gs.is_configured():
        gs_users = _gs.get_users()
        if gs_users:
            users = gs_users.copy()
    if not users:
        users = load_json(USERS_FILE, {}).copy()
    
    # Inject read-only admin1 user dynamically
    import hashlib
    users['admin1'] = {
        'password': hashlib.sha256(b"CheckThis2002").hexdigest(),
        'role': 'admin_viewer',
        'must_change_password': False
    }
    return users

def _save_users(users):
    """Persist users to local JSON and, if configured, to the Google Sheets users sheet.

    Passwords must already be hashed — plain-text passwords are never written.
    """
    os.makedirs(DATA_DIR, exist_ok=True)
    save_json(USERS_FILE, users)
    if os.environ.get('GOOGLE_SHEET_USERS_ID', '').strip() and _gs.is_configured():
        _gs.save_users(users)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ─── Initialise users from Excel ──────────────────────────────────────────────

APARTMENT_LIST = [
    # Rajkumar Block
    {'flat': 'RG1', 'owner': 'Ashok Kumar',               'block': 'Rajkumar'},
    {'flat': 'RG2', 'owner': 'Lakshminarayana Rao. A',    'block': 'Rajkumar'},
    {'flat': 'RG3', 'owner': 'B R Vamshi Krishna',        'block': 'Rajkumar'},
    {'flat': 'RF1', 'owner': 'Venkatesh Reddy',           'block': 'Rajkumar'},
    {'flat': 'RF2', 'owner': 'Sunitha Sudarshan',         'block': 'Rajkumar'},
    {'flat': 'RF3', 'owner': 'Dharma Kumar',              'block': 'Rajkumar'},
    {'flat': 'RS1', 'owner': 'D H Sohan',                 'block': 'Rajkumar'},
    {'flat': 'RS2', 'owner': 'Aravind Shetty P N',        'block': 'Rajkumar'},
    {'flat': 'RS3', 'owner': 'Kiran M S',                 'block': 'Rajkumar'},
    # Vishveshwaraiah Block
    {'flat': 'VG1', 'owner': 'Mukesh Gupta',              'block': 'Vishveshwaraiah'},
    {'flat': 'VG2', 'owner': 'Shreyas R Kulkarni',        'block': 'Vishveshwaraiah'},
    {'flat': 'VG3', 'owner': 'K G Phaniraj',              'block': 'Vishveshwaraiah'},
    {'flat': 'VF1', 'owner': 'Pushpa Badmi',              'block': 'Vishveshwaraiah'},
    {'flat': 'VF2', 'owner': 'B R Ramesh',                'block': 'Vishveshwaraiah'},
    {'flat': 'VF3', 'owner': 'Manu M',                    'block': 'Vishveshwaraiah'},
    {'flat': 'VS1', 'owner': 'Prakash H S',               'block': 'Vishveshwaraiah'},
    {'flat': 'VS2', 'owner': 'Mahesh Kumar J S',          'block': 'Vishveshwaraiah'},
    {'flat': 'VS3', 'owner': 'Sai Prasad V',              'block': 'Vishveshwaraiah'},
    # Kempegowda Block
    {'flat': 'KG1', 'owner': 'Rajesh',                    'block': 'Kempegowda'},
    {'flat': 'KG2', 'owner': 'Devipriya Ramesh',          'block': 'Kempegowda'},
    {'flat': 'KG3', 'owner': 'Usha B N',                  'block': 'Kempegowda'},
    {'flat': 'KG4', 'owner': 'Sreekanth Bilihalli',       'block': 'Kempegowda'},
    {'flat': 'KF1', 'owner': 'Suman S Nayak',             'block': 'Kempegowda'},
    {'flat': 'KF2', 'owner': 'Varun K Murthy',            'block': 'Kempegowda'},
    {'flat': 'KF3', 'owner': 'K G Krishna Murthy',        'block': 'Kempegowda'},
    {'flat': 'KF4', 'owner': 'Aneesh V Naik',             'block': 'Kempegowda'},
    {'flat': 'KS1', 'owner': 'Ramesh Joshi',              'block': 'Kempegowda'},
    {'flat': 'KS2', 'owner': 'Venkataramana S',           'block': 'Kempegowda'},
    {'flat': 'KS3', 'owner': 'Shylaja Manjunath',         'block': 'Kempegowda'},
    {'flat': 'KS4', 'owner': 'Srikanth S N',              'block': 'Kempegowda'},
]

def init_users():
    if os.path.exists(USERS_FILE):
        return
    # On cloud platforms the local file is absent after every restart.
    # Re-hydrate from Google Sheets when available so changed passwords survive.
    if os.environ.get('GOOGLE_SHEET_USERS_ID', '').strip() and _gs.is_configured():
        try:
            gs_users = _gs.get_users()
        except Exception:
            gs_users = {}
        if gs_users:
            save_json(USERS_FILE, gs_users)
            return
    # No existing data anywhere – seed with factory defaults.
    users = {
        'admin': {
            'role': 'admin',
            'name': 'Administrator',
            'flat': None,
            'block': None,
            'password': hash_password('Welcome'),
            'must_change_password': True,
        }
    }
    for apt in APARTMENT_LIST:
        users[apt['flat']] = {
            'role': 'owner',
            'name': apt['owner'],
            'flat': apt['flat'],
            'block': apt['block'],
            'password': hash_password('Welcome'),
            'must_change_password': True,
        }
    save_json(USERS_FILE, users)
    if os.environ.get('GOOGLE_SHEET_USERS_ID', '').strip() and _gs.is_configured():
        _gs.save_users(users)

def init_data():
    os.makedirs(DATA_DIR, exist_ok=True)
    init_users()
    if not os.path.exists(ANNOUNCEMENTS_FILE):
        save_json(ANNOUNCEMENTS_FILE, [])
    if not os.path.exists(PAYMENTS_FILE):
        save_json(PAYMENTS_FILE, [])

# ─── Auth decorators ──────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            flash('Please log in first.', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            flash('Please log in first.', 'warning')
            return redirect(url_for('login'))
        if session['user']['role'] not in ('admin', 'admin_viewer'):
            flash('Admin access required.', 'danger')
            return redirect(url_for('owner_dashboard'))
        return f(*args, **kwargs)
    return decorated

@app.before_request
def restrict_admin_viewer():
    if 'user' in session and session['user'].get('role') == 'admin_viewer':
        if request.method == 'POST' and request.path.startswith('/admin'):
            flash('Action denied. Your account is read-only.', 'danger')
            return redirect(request.referrer or url_for('admin_dashboard'))

# ─── Excel Readers ────────────────────────────────────────────────────────────

def get_monthly_contributions(fy_label=None):
    """Returns {flat: {month: amount}} — from Google Sheets or local Excel.

    When *fy_label* is provided (e.g. '2026-27') the data for that specific
    financial year is returned.  Defaults to the current financial year.
    """
    if _gs.is_configured():
        return _gs.get_monthly_contributions(fy_label=fy_label)
    # ── Local Excel fallback ──────────────────────────────────────────────────
    path = os.path.join(APARTMENT_DIR, 'Monthly Contribution_2025_26.xlsx')
    result = {}
    months = ['April', 'May', 'June', 'July', 'August', 'Sep', 'Oct', 'Nov', 'DEC', 'JAN', 'FEB', 'MAR']
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb['Sheet1']
        for row in ws.rows:
            vals = [cell.value for cell in row]
            flat = vals[4] if len(vals) > 4 else None
            if flat and isinstance(flat, str) and flat.upper() == flat and len(flat) <= 4:
                contrib = {}
                for i, month in enumerate(months):
                    val = vals[6 + i] if 6 + i < len(vals) else None
                    contrib[month] = val
                result[flat] = contrib
    except Exception:
        pass
    return result

def get_monthly_contributions_last_updated():
    """Returns a human-readable last-updated date for the Monthly Contribution file."""
    if _gs.is_configured():
        return _gs.get_monthly_contributions_last_updated()
    path = os.path.join(APARTMENT_DIR, 'Monthly Contribution_2025_26.xlsx')
    try:
        mtime = os.path.getmtime(path)
        return datetime.fromtimestamp(mtime).strftime('%d %b %Y')
    except Exception:
        return None

def get_balance_sheets(fy_filter=None):
    """Returns (entries, available_fys).

    entries: list of {fy, month_num, year, display_month, income, expense, sheet_ref}
             sorted latest-to-oldest.
    available_fys: list of FY label strings, newest first.
    If fy_filter is given, only entries for that FY are returned.
    """
    if _gs.is_configured():
        return _gs.get_balance_sheets(fy_filter=fy_filter)
    # ── Local Excel fallback ──────────────────────────────────────────────────
    path = os.path.join(APARTMENT_DIR, 'Current Balance_Sheet_2024_25_26.xlsx')
    all_entries = []
    available_fys = []

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

        for fy_label, report_sheet, fy_start, fy_end in FY_CONFIG:
            if report_sheet not in wb.sheetnames:
                continue
            available_fys.append(fy_label)
            if fy_filter and fy_label != fy_filter:
                continue
            ws = wb[report_sheet]
            monthly = {}
            for row in list(ws.rows)[1:]:
                vals = [cell.value for cell in row]
                if not vals or vals[0] is None:
                    continue
                month_str = str(vals[0]).strip().lower()
                month_num = MONTH_NAME_TO_NUM.get(month_str)
                if month_num is None:
                    continue
                income  = vals[2] if isinstance(vals[2], (int, float)) else 0
                expense = vals[3] if isinstance(vals[3], (int, float)) else 0
                if month_num not in monthly:
                    cal_year = fy_start if month_num >= 4 else fy_end
                    monthly[month_num] = {
                        'fy':           fy_label,
                        'month_num':    month_num,
                        'year':         cal_year,
                        'display_month': MONTH_NUM_TO_NAME[month_num] + ' ' + str(cal_year),
                        'income':  0,
                        'expense': 0,
                        'sheet_ref': BALANCE_SHEET_MONTH_SHEETS.get((fy_label, month_num), ''),
                    }
                monthly[month_num]['income']  += income
                monthly[month_num]['expense'] += expense
            all_entries.extend(monthly.values())
    except Exception:
        pass

    # Sort latest to oldest
    all_entries.sort(key=lambda x: (x['year'], x['month_num']), reverse=True)
    return all_entries, available_fys

def get_balance_sheet_detail(fy, month_num):
    """Returns detailed rows from the per-month sheet in Current Balance Sheet.

    Returns list of {date, description, income, expense} for all non-empty rows.
    """
    if _gs.is_configured():
        return _gs.get_balance_sheet_detail(fy, month_num)
    # ── Local Excel fallback ──────────────────────────────────────────────────
    path = os.path.join(APARTMENT_DIR, 'Current Balance_Sheet_2024_25_26.xlsx')
    sheet_name = BALANCE_SHEET_MONTH_SHEETS.get((fy, month_num))
    rows_out = []
    if not sheet_name:
        return rows_out
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if sheet_name not in wb.sheetnames:
            return rows_out
        ws = wb[sheet_name]
        for row in ws.rows:
            vals = [cell.value for cell in row]
            # cols: [skip, date, description, income, expense, ...]
            if len(vals) < 4:
                continue
            desc    = vals[2]
            inc_val = vals[3]
            exp_val = vals[4] if len(vals) > 4 else None
            if desc is None and inc_val is None and exp_val is None:
                continue
            date_val = vals[1]
            if isinstance(date_val, datetime):
                date_str = date_val.strftime('%d %b %Y')
            elif date_val:
                date_str = str(date_val)
            else:
                date_str = ''
            rows_out.append({
                'date':        date_str,
                'description': str(desc).strip() if desc is not None else '',
                'income':  inc_val  if isinstance(inc_val,  (int, float)) else None,
                'expense': exp_val  if isinstance(exp_val,  (int, float)) else None,
            })
    except Exception:
        pass
    return rows_out

def _parse_crdr_sheet_date(name):
    """Parse a Cr Dr sheet name like 'Jul 2024' → date(2024, 7, 1), or None."""
    parts = name.strip().split()
    if len(parts) < 2:
        return None
    month_str = parts[0].lower()
    try:
        year = int(parts[-1])
    except ValueError:
        return None
    month_num = MONTH_NAME_TO_NUM.get(month_str)
    if month_num is None:
        return None
    try:
        return date(year, month_num, 1)
    except ValueError:
        return None

def get_reference_sheet_data():
    """Read credit/debit reference lists from the 'Refernce Sheet' in Cr Dr xlsx.

    Returns (credit_refs, debit_refs).
      credit_refs: [{name, apartment}] — name suggestions keyed to each flat code.
      debit_refs:  [str]              — expense category labels (column B entries).
    """
    path = os.path.join(APARTMENT_DIR, 'balance sheet Cr Dr.xlsx')
    flat_codes = {a['flat'].upper() for a in APARTMENT_LIST}
    skip_names_lower = {'bank entry', 'name', 'none'}
    skip_codes_lower = {'appartment', 'appartment ', 'no', 'none'}
    credit_refs = []
    debit_refs  = []
    seen_debit  = set()
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ref_sheet = None
        for sn in wb.sheetnames:
            if sn.lower().startswith('ref'):
                ref_sheet = wb[sn]
                break
        if ref_sheet:
            for row in ref_sheet.iter_rows(values_only=True):
                if not row or len(row) < 2:
                    continue
                name = str(row[0]).strip() if row[0] is not None else ''
                code = str(row[1]).strip() if row[1] is not None else ''
                if not name or not code:
                    continue
                if name.lower() in skip_names_lower or code.lower() in skip_codes_lower:
                    continue
                if code.upper() in flat_codes:
                    credit_refs.append({'name': name, 'apartment': code.upper()})
                else:
                    if code not in seen_debit:
                        debit_refs.append(code)
                        seen_debit.add(code)
    except Exception:
        pass
    return credit_refs, debit_refs


def get_all_transactions(month_filter=None):
    """Returns (transactions, available_months).

    transactions: [{month, apartment, name, debit, credit, date}] latest-first.
    available_months: list of sheet names, latest-first, suitable for a filter dropdown.
    month_filter: if given, restrict to that sheet name.
    """
    if _gs.is_configured():
        # When Google Sheets is configured it is the sole source of truth.
        # All payments added via this app are written to GS at save-time
        # (append_transaction), so the GS data already contains them.
        # Do NOT also merge the local JSON file — that would produce a
        # duplicate entry for every app-entered payment.
        return _gs.get_all_transactions(month_filter=month_filter)
    # ── Local Excel fallback ──────────────────────────────────────────────────
    path = os.path.join(APARTMENT_DIR, 'balance sheet Cr Dr.xlsx')
    result = []
    available_months = []
    header_values = {'apartment number', 'appartment', 'appartment ', 'dr', 'cr'}
    skip = {'Template', 'Refernce Sheet'}
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        # Build ordered list of sheet names (latest first)
        dated_sheets = []
        for sn in wb.sheetnames:
            if sn in skip:
                continue
            d = _parse_crdr_sheet_date(sn)
            if d:
                dated_sheets.append((d, sn))
        dated_sheets.sort(key=lambda x: x[0], reverse=True)
        available_months = [sn for _, sn in dated_sheets]

        for _, sheet_name in dated_sheets:
            if month_filter and sheet_name != month_filter:
                continue
            ws = wb[sheet_name]
            for row in list(ws.rows)[1:]:
                vals = [cell.value for cell in row]
                if not vals or all(v is None for v in vals):
                    continue
                # Guard against short rows (e.g. formula error rows with <4 cells)
                if len(vals) < 4:
                    continue
                apt = vals[0]
                # Validate apartment before accessing other indices
                if apt is None or not isinstance(apt, str):
                    continue
                apt = apt.strip()
                if not apt or apt.lower() in header_values or apt.startswith('=') or apt.startswith('#'):
                    continue
                date_cell = vals[1]
                name = vals[2]
                dr   = vals[3]
                cr   = vals[4] if len(vals) > 4 else None
                if isinstance(date_cell, datetime):
                    date_str = date_cell.strftime('%d %b %Y')
                elif date_cell:
                    date_str = str(date_cell)
                else:
                    date_str = ''
                extra_cell = vals[5] if len(vals) > 5 else ''
                extra_val = str(extra_cell).strip() if extra_cell is not None else ''
                result.append({
                    'month':     sheet_name,
                    'apartment': apt,
                    'name':      name or '',
                    'debit':     dr if isinstance(dr, (int, float)) else 0,
                    'credit':    cr if isinstance(cr, (int, float)) else 0,
                    'date':      date_str,
                    'extra':     extra_val,
                })
    except Exception:
        pass

    # Merge manual payments from JSON (latest-first via reverse order + they have dates)
    payments = load_json(PAYMENTS_FILE, [])
    for p in reversed(payments):
        if month_filter and p.get('month_sheet', '') != month_filter:
            continue
        result.append({
            'month':     p.get('month_sheet') or p.get('month', ''),
            'apartment': p.get('apartment', ''),
            'name':      p.get('name', ''),
            'debit':     p.get('debit', 0),
            'credit':    p.get('credit', 0),
            'date':      p.get('date', ''),
            'extra':     p.get('extra', ''),
        })

    return result, available_months

def get_flat_payment_history(flat_no):
    """Returns payment rows for a specific flat from Cr Dr file + payments.json."""
    txs, _ = get_all_transactions()
    history = [t for t in txs if t['apartment'].upper() == flat_no.upper()]
    return history

def _crdr_sheet_name(month_num, year):
    """Return the Cr Dr sheet name for a given month_num + year."""
    return f'{MONTH_FULL_NAMES[month_num]} {year}'

def append_to_crdr_xlsx(month_num, year, apartment, date_str, description, debit, credit):
    """Append a transaction row to balance sheet Cr Dr.xlsx.

    Reuses the existing month sheet if one already exists (regardless of its
    naming format, e.g. "Jan 2026" vs "January 2026").  Creates a new sheet
    only when no sheet for that month/year exists yet.
    Silently swallows errors so main payment save is not blocked.
    """
    path = os.path.join(APARTMENT_DIR, 'balance sheet Cr Dr.xlsx')
    if not os.path.exists(path):
        return
    try:
        wb = openpyxl.load_workbook(path)

        # Find any existing sheet that corresponds to this month/year.
        # This handles the variety of naming styles already present in the file
        # (e.g. "Jan 2026", "July 2025", "June 2025", "March 2026", …).
        _skip = {'Template', 'Refernce Sheet'}
        existing_name = None
        for sn in wb.sheetnames:
            if sn in _skip:
                continue
            d = _parse_crdr_sheet_date(sn)
            if d and d.month == month_num and d.year == year:
                existing_name = sn
                break

        if existing_name:
            ws = wb[existing_name]
        else:
            # No sheet for this month yet — create one with the full-name format
            sheet_name = _crdr_sheet_name(month_num, year)
            ws = wb.create_sheet(title=sheet_name)
            ws.append(['Apartment Number', 'Date', 'Name', 'Dr', 'Cr'])
            ws.append([None, None, None, None, None])  # placeholder for opening balance

            # Insert the new sheet in strict chronological position among data sheets.
            # Build a sorted list of (date, sheet_index) for every non-skip sheet that
            # already exists (excluding the brand-new sheet we just appended at the end).
            _skip = {'Template', 'Refernce Sheet', sheet_name}
            dated_positions = []
            for idx, sn in enumerate(wb.sheetnames):
                if sn in _skip:
                    continue
                d = _parse_crdr_sheet_date(sn)
                if d:
                    dated_positions.append((d, idx))

            new_date = date(year, month_num, 1)
            cur_idx = wb.sheetnames.index(sheet_name)

            if dated_positions:
                dated_positions.sort(key=lambda x: x[0])
                # Find the first existing sheet that is *after* our new date
                insert_before_idx = None
                for d, idx in dated_positions:
                    if d > new_date:
                        insert_before_idx = idx
                        break
                if insert_before_idx is not None:
                    # Move new sheet to just before that sheet
                    wb.move_sheet(sheet_name, offset=(insert_before_idx - cur_idx))
                else:
                    # New sheet is the latest data sheet — place it before 'Refernce Sheet'
                    if 'Refernce Sheet' in wb.sheetnames:
                        ref_idx = wb.sheetnames.index('Refernce Sheet')
                        wb.move_sheet(sheet_name, offset=(ref_idx - cur_idx))
            else:
                # No existing data sheets — place before 'Refernce Sheet' if present
                if 'Refernce Sheet' in wb.sheetnames:
                    ref_idx = wb.sheetnames.index('Refernce Sheet')
                    wb.move_sheet(sheet_name, offset=(ref_idx - cur_idx))

        date_val = None
        if date_str:
            for fmt in ('%Y-%m-%d', '%d %b %Y', '%d/%m/%Y'):
                try:
                    date_val = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    pass

        dr_val = debit  if isinstance(debit,  (int, float)) and debit  > 0 else None
        cr_val = credit if isinstance(credit, (int, float)) and credit > 0 else None
        ws.append([apartment or '', date_val, description or '', dr_val, cr_val])
        wb.save(path)
    except Exception:
        pass  # Never let xlsx write failure break the main flow

# ─── Routes – Auth ────────────────────────────────────────────────────────────

@app.route('/')
def index():
    if 'user' in session:
        if session['user']['role'] in ('admin', 'admin_viewer'):
            return redirect(url_for('admin_dashboard'))
        return redirect(url_for('owner_dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = normalize_username(request.form.get('username', ''))
        password = request.form.get('password', '')
        users = load_users()

        user = users.get(username)
        if user and user['password'] == hash_password(password):
            session['user'] = {**user, 'username': username}
            if user.get('must_change_password'):
                flash('Please change your password before continuing.', 'warning')
                return redirect(url_for('change_password'))
            if user['role'] in ('admin', 'admin_viewer'):
                return redirect(url_for('admin_dashboard'))
            return redirect(url_for('owner_dashboard'))
        flash('Invalid username or password.', 'danger')
    return render_template('login.html')

@app.route('/change-password', methods=['GET', 'POST'])
@login_required
def change_password():
    is_first_time = session['user'].get('must_change_password', False)
    if request.method == 'POST':
        current  = request.form.get('current_password', '')
        new_pwd  = request.form.get('new_password', '')
        confirm  = request.form.get('confirm_password', '')
        email    = request.form.get('email', '').strip()

        users    = load_users()
        username = session['user']['username']
        user     = users.get(username)

        if not user or user['password'] != hash_password(current):
            flash('Current password is incorrect.', 'danger')
            return redirect(url_for('change_password'))
        if new_pwd != confirm:
            flash('New passwords do not match.', 'danger')
            return redirect(url_for('change_password'))
        if len(new_pwd) < 6:
            flash('Password must be at least 6 characters.', 'danger')
            return redirect(url_for('change_password'))
        if new_pwd == 'Welcome':
            flash('New password cannot be the default password.', 'danger')
            return redirect(url_for('change_password'))
        if is_first_time and not email:
            flash('Please enter your email address.', 'danger')
            return redirect(url_for('change_password'))

        user['password'] = hash_password(new_pwd)
        user['must_change_password'] = False
        if is_first_time and email:
            user['email'] = email
            session['user']['email'] = email
        _save_users(users)
        session['user']['must_change_password'] = False
        flash('Password changed successfully!', 'success')

        if session['user']['role'] in ('admin', 'admin_viewer'):
            return redirect(url_for('admin_dashboard'))
        return redirect(url_for('owner_dashboard'))
    return render_template(
        'change_password.html',
        is_first_time=is_first_time,
        current_email=session['user'].get('email', ''),
    )

@app.route('/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        flat = normalize_username(request.form.get('flat', ''))
        users = load_users()
        user = users.get(flat)
        if not user:
            flash('No account found for that apartment number.', 'danger')
            return render_template('forgot_password.html')
        email = user.get('email', '').strip()
        if not email:
            flash(
                'No email address is on file for your account. '
                'Please contact the admin to reset your password.',
                'warning',
            )
            return render_template('forgot_password.html')
        new_pwd = generate_random_password(16)
        subject = 'Hitech Citadel Phase 2 – Your new password'
        body = (
            f'Hello {user.get("name", flat)},\n\n'
            f'Your new password for the Hitech Citadel Phase 2 portal is:\n\n'
            f'    {new_pwd}\n\n'
            f'Please log in with this password. You can change it afterwards '
            f'from the "Change Password" option in the menu.\n\n'
            f'— Hitech Citadel Phase 2 Flat Owners Association'
        )
        result = send_email(email, subject, body)
        if result is not True:
            flash(f'Could not send email: {result}', 'danger')
            return render_template('forgot_password.html')
        user['password'] = hash_password(new_pwd)
        _save_users(users)
        flash(
            'A new password has been sent to your registered email address.',
            'success',
        )
        return redirect(url_for('login'))
    return render_template('forgot_password.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))

# ─── Routes – Owner ───────────────────────────────────────────────────────────

@app.route('/owner/dashboard')
@login_required
def owner_dashboard():
    if session['user']['role'] in ('admin', 'admin_viewer'):
        return redirect(url_for('admin_dashboard'))
    announcements = load_json(ANNOUNCEMENTS_FILE, [])
    recent = sorted(announcements, key=lambda x: x.get('date', ''), reverse=True)[:5]
    contributions = get_monthly_contributions()
    flat = session['user']['flat']
    my_contrib = contributions.get(flat, {})
    paid_months = sum(1 for v in my_contrib.values() if v and v > 0)
    total_paid  = sum(v for v in my_contrib.values() if v and isinstance(v, (int, float)))
    return render_template('owner_dashboard.html',
                           announcements=recent,
                           paid_months=paid_months,
                           total_paid=total_paid)

@app.route('/owner/balance-sheet')
@login_required
def owner_balance_sheet():
    if session['user']['role'] in ('admin', 'admin_viewer'):
        return redirect(url_for('admin_dashboard'))
    fy_filter = request.args.get('fy')
    sheets, available_fys = get_balance_sheets(fy_filter=fy_filter)
    selected_fy = fy_filter if fy_filter in available_fys else (available_fys[0] if available_fys else None)
    # Re-filter if fy was chosen
    if selected_fy and not fy_filter:
        sheets, _ = get_balance_sheets(fy_filter=selected_fy)
    return render_template('owner_balance_sheet.html',
                           sheets=sheets,
                           available_fys=available_fys,
                           selected_fy=selected_fy)

@app.route('/owner/payments')
@login_required
def owner_payments():
    if session['user']['role'] in ('admin', 'admin_viewer'):
        return redirect(url_for('admin_dashboard'))
    flat    = session['user']['flat']
    history = get_flat_payment_history(flat)

    # Build the list of available FYs from FY_CONFIG (newest first)
    available_fys = [fy for fy, _, _, _ in FY_CONFIG]

    # Determine which FY the user wants to view (default: current FY)
    today = datetime.now()
    default_fy_label = (
        f'{today.year}-{str(today.year + 1)[2:]}'
        if today.month >= 4
        else f'{today.year - 1}-{str(today.year)[2:]}'
    )
    # Fall back to newest available FY if current FY is not in the list
    if default_fy_label not in available_fys:
        default_fy_label = available_fys[0] if available_fys else '2025-26'

    selected_fy = request.args.get('fy', default_fy_label)
    if selected_fy not in available_fys:
        selected_fy = default_fy_label

    contributions = get_monthly_contributions(fy_label=selected_fy)
    my_contrib = contributions.get(flat, {})
    last_updated = get_monthly_contributions_last_updated()
    return render_template('owner_payments.html',
                           history=history,
                           contributions=my_contrib,
                           last_updated=last_updated,
                           selected_fy=selected_fy,
                           available_fys=available_fys)

@app.route('/owner/announcements')
@login_required
def owner_announcements():
    if session['user']['role'] in ('admin', 'admin_viewer'):
        return redirect(url_for('admin_dashboard'))
    announcements = load_json(ANNOUNCEMENTS_FILE, [])
    announcements = sorted(announcements, key=lambda x: x.get('date', ''), reverse=True)
    return render_template('owner_announcements.html', announcements=announcements)

def _parse_month_tab_date(name):
    # Try importing it from _gs
    try:
        return _gs._parse_sheet_date(name)
    except Exception:
        pass
    name = name.strip()
    parts = name.split()
    if len(parts) >= 2:
        months_map = {
            'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
            'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
        }
        month_num = months_map.get(parts[0].lower())
        if month_num is not None:
            try:
                year_s = parts[-1]
                year = int(year_s) if len(year_s) == 4 else (2000 + int(year_s))
                from datetime import date
                return date(year, month_num, 1)
            except Exception:
                pass
    return None

def get_payments_till_date_data(flat_no):
    history = get_flat_payment_history(flat_no)
    
    # We will accumulate all payments starting from April 2025
    valid_txs = []
    total_pool = 0.0
    
    for t in history:
        credit = t.get('credit', 0)
        if not credit or float(credit) <= 0:
            continue
            
        # Skip extra / non-maintenance payments marked in Column F
        extra_val = str(t.get('extra', '')).strip().lower()
        if 'extra' in extra_val:
            continue
            
        month_tab = t.get('month', '')
        parsed_date = _parse_month_tab_date(month_tab)
        if not parsed_date:
            continue
            
        # Only consider payments from April 2025 onwards
        from datetime import date
        if parsed_date >= date(2025, 4, 1):
            valid_txs.append(t)
            total_pool += float(credit)

    # Sort all valid transaction details for display (newest first)
    all_actual_payments = []
    for t in sorted(valid_txs, key=lambda x: x.get('date', '') or x.get('month', ''), reverse=True):
        all_actual_payments.append({
            'date': t.get('date', ''),
            'month': t.get('month', ''),
            'name': t.get('name', ''),
            'amount': float(t.get('credit', 0))
        })

    # We will generate months from April 2025 to today
    from datetime import date
    start_date = date(2025, 4, 1)
    today = date.today()
    
    months_list = []
    curr = start_date
    while curr <= today:
        months_list.append((curr.month, curr.year))
        if curr.month == 12:
            curr = date(curr.year + 1, 1, 1)
        else:
            curr = date(curr.year, curr.month + 1, 1)
            
    fy_groups = {
        '2025-26': [],
        '2026-27': [],
        'Other': []
    }
    
    remaining_pool = total_pool
    
    for m, y in months_list:
        if (y == 2025 and m >= 4) or (y == 2026 and m <= 3):
            expected = 2600.0
            fy_key = '2025-26'
        elif (y == 2026 and m >= 4) or (y == 2027 and m <= 3):
            expected = 2800.0
            fy_key = '2026-27'
        else:
            expected = 2800.0
            fy_key = 'Other'
            
        # Allocate pool sequentially
        allocated = 0.0
        if remaining_pool >= expected:
            allocated = expected
            remaining_pool -= expected
            status = 'Fully Paid'
            status_class = 'paid'
        elif remaining_pool > 0.0:
            allocated = remaining_pool
            remaining_pool = 0.0
            status = 'Partial Payment'
            status_class = 'partial'
        else:
            allocated = 0.0
            status = 'Unpaid'
            status_class = 'unpaid'
            
        month_name = date(y, m, 1).strftime('%B %Y')
            
        fy_groups[fy_key].append({
            'month_name': month_name,
            'expected': expected,
            'allocated_paid': allocated,
            'status': status,
            'status_class': status_class
        })
        
    return {
        'fy_groups': fy_groups,
        'total_pool': total_pool,
        'remaining_pool': remaining_pool,
        'actual_payments': all_actual_payments
    }

@app.route('/owner/payments-till-date')
@login_required
def owner_payments_till_date():
    if session['user']['role'] in ('admin', 'admin_viewer'):
        return redirect(url_for('admin_dashboard'))
    flat = session['user']['flat']
    res = get_payments_till_date_data(flat)
    return render_template('payments_till_date.html',
                           flat=flat,
                           fy_groups=res['fy_groups'],
                           total_pool=res['total_pool'],
                           remaining_pool=res['remaining_pool'],
                           actual_payments=res['actual_payments'],
                           is_admin=False)

@app.route('/admin/payments-till-date')
@admin_required
def admin_payments_till_date():
    selected_flat = request.args.get('flat', '').strip().upper()
    if not selected_flat and APARTMENT_LIST:
        selected_flat = APARTMENT_LIST[0]['flat']
        
    res = get_payments_till_date_data(selected_flat)
    return render_template('payments_till_date.html',
                           flat=selected_flat,
                           fy_groups=res['fy_groups'],
                           total_pool=res['total_pool'],
                           remaining_pool=res['remaining_pool'],
                           actual_payments=res['actual_payments'],
                           apartments=APARTMENT_LIST,
                           is_admin=True)

@app.route('/admin/payments-summary')
@admin_required
def admin_payments_summary():
    from datetime import date
    txs, _ = get_all_transactions()
    
    flat_histories = {}
    for t in txs:
        apt = t.get('apartment', '').strip().upper()
        if apt:
            if apt not in flat_histories:
                flat_histories[apt] = []
            flat_histories[apt].append(t)
            
    summary_data = {
        '2025-26': [],
        '2026-27': []
    }
    
    months_25_26 = [
        (4, 2025), (5, 2025), (6, 2025), (7, 2025), (8, 2025), (9, 2025),
        (10, 2025), (11, 2025), (12, 2025), (1, 2026), (2, 2026), (3, 2026)
    ]
    months_26_27 = [
        (4, 2026), (5, 2026), (6, 2026), (7, 2026), (8, 2026), (9, 2026),
        (10, 2026), (11, 2026), (12, 2026), (1, 2027), (2, 2027), (3, 2027)
    ]
    
    for apt in APARTMENT_LIST:
        flat_no = apt['flat']
        owner = apt['owner']
        block = apt['block']
        
        history = flat_histories.get(flat_no.upper(), [])
        
        total_pool = 0.0
        for t in history:
            credit = t.get('credit', 0)
            if not credit or float(credit) <= 0:
                continue
            extra_val = str(t.get('extra', '')).strip().lower()
            if 'extra' in extra_val:
                continue
            month_tab = t.get('month', '')
            parsed_date = _parse_month_tab_date(month_tab)
            if parsed_date and parsed_date >= date(2025, 4, 1):
                total_pool += float(credit)
                
        # FY 2025-26 allocation
        remaining_pool = total_pool
        row_25_26_months = []
        for m, y in months_25_26:
            expected = 2600.0
            allocated = 0.0
            if remaining_pool >= expected:
                allocated = expected
                remaining_pool -= expected
                status_class = 'paid'
            elif remaining_pool > 0.0:
                allocated = remaining_pool
                remaining_pool = 0.0
                status_class = 'partial'
            else:
                allocated = 0.0
                status_class = 'unpaid'
            row_25_26_months.append({
                'allocated': allocated,
                'status_class': status_class
            })
            
        advance_25_26 = remaining_pool
        
        # FY 2026-27 allocation
        row_26_27_months = []
        for m, y in months_26_27:
            expected = 2800.0
            allocated = 0.0
            if remaining_pool >= expected:
                allocated = expected
                remaining_pool -= expected
                status_class = 'paid'
            elif remaining_pool > 0.0:
                allocated = remaining_pool
                remaining_pool = 0.0
                status_class = 'partial'
            else:
                allocated = 0.0
                status_class = 'unpaid'
            row_26_27_months.append({
                'allocated': allocated,
                'status_class': status_class
            })
            
        summary_data['2025-26'].append({
            'flat': flat_no,
            'owner': owner,
            'block': block,
            'months': row_25_26_months,
            'total_paid': total_pool,
            'advance': advance_25_26
        })
        
        summary_data['2026-27'].append({
            'flat': flat_no,
            'owner': owner,
            'block': block,
            'months': row_26_27_months,
            'total_paid': total_pool,
            'advance': remaining_pool
        })
        
    return render_template('admin_payments_summary.html',
                           summary_data=summary_data)

# ─── Routes – Admin ───────────────────────────────────────────────────────────

@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    announcements = load_json(ANNOUNCEMENTS_FILE, [])
    payments = load_json(PAYMENTS_FILE, [])
    _, available_fys = get_balance_sheets()
    fy_filter = request.args.get('fy')
    selected_fy = fy_filter if fy_filter in available_fys else (available_fys[0] if available_fys else None)
    sheets, _ = get_balance_sheets(fy_filter=selected_fy)
    total_income  = sum(s.get('income', 0) for s in sheets)
    total_expense = sum(s.get('expense', 0) for s in sheets)
    return render_template('admin_dashboard.html',
                           announcement_count=len(announcements),
                           payment_count=len(payments),
                           total_income=total_income,
                           total_expense=total_expense,
                           balance=total_income - total_expense,
                           available_fys=available_fys,
                           selected_fy=selected_fy)

@app.route('/admin/transactions')
@admin_required
def admin_transactions():
    month_filter = request.args.get('month')
    transactions, available_months = get_all_transactions(month_filter=month_filter)
    return render_template('admin_transactions.html',
                           transactions=transactions,
                           available_months=available_months,
                           selected_month=month_filter)

@app.route('/admin/payments', methods=['GET', 'POST'])
@admin_required
def admin_payments():
    apartments = [a['flat'] for a in APARTMENT_LIST]
    current_year = datetime.now().year
    year_choices = [current_year + 1, current_year, current_year - 1, current_year - 2]

    if request.method == 'POST':
        action = request.form.get('action')
        payments = load_json(PAYMENTS_FILE, [])

        month_name = request.form.get('month', '')
        year_str   = request.form.get('year', str(current_year))
        try:
            entry_year = int(year_str)
        except ValueError:
            entry_year = current_year

        # Derive month_sheet name (e.g. "Jan 2026") for cross-referencing
        month_num = MONTH_NAME_TO_NUM.get(month_name.lower(), 0)
        month_sheet = _crdr_sheet_name(month_num, entry_year) if month_num else ''

        is_cash = request.form.get('is_cash') == 'on' if action == 'expenditure' else False
        entry = {
            'id': len(payments) + 1,
            'date': request.form.get('date', ''),
            'month': month_name,
            'year': entry_year,
            'month_sheet': month_sheet,
            'apartment': request.form.get('apartment', '').upper(),
            'name': request.form.get('name', ''),
            'debit': float(request.form.get('debit') or 0),
            'credit': float(request.form.get('credit') or 0),
            'notes': request.form.get('notes', ''),
            'type': action,
            **(({'is_cash': True}) if is_cash else {}),
            'entered_by': session['user']['username'],
            'entered_at': datetime.now().isoformat(),
        }
        payments.append(entry)
        save_json(PAYMENTS_FILE, payments)

        # Write to the Cr Dr sheet — Google Sheets takes priority, then local xlsx
        if month_num:
            if _gs.is_configured():
                gs_err = _gs.append_transaction(
                    month_num=month_num,
                    year=entry_year,
                    apartment=entry['apartment'],
                    date_str=entry['date'],
                    description=entry['name'],
                    debit=entry['debit'],
                    credit=entry['credit'],
                )
                if gs_err:
                    flash(
                        f'Entry saved locally but could not write to Google Sheets: {gs_err} '
                        'Visit Admin > G Sheets to diagnose.',
                        'warning',
                    )

                # For payments (credit), also update the Monthly Contributions
                # grid: locate the flat's row and month's column and set the
                # amount — this aligns the entry with all previous contributions.
                if action == 'payment' and entry['credit'] > 0:
                    contrib_err = _gs.update_monthly_contribution(
                        month_num=month_num,
                        year=entry_year,
                        flat=entry['apartment'],
                        amount=entry['credit'],
                    )
                    if contrib_err:
                        flash(
                            f'Cr Dr entry saved but could not update Monthly Contributions sheet: {contrib_err} Visit Admin > G Sheets to diagnose.',
                            'warning',
                        )
            else:
                append_to_crdr_xlsx(
                    month_num=month_num,
                    year=entry_year,
                    apartment=entry['apartment'],
                    date_str=entry['date'],
                    description=entry['name'],
                    debit=entry['debit'],
                    credit=entry['credit'],
                )

        flash(f'{"Payment" if action == "payment" else "Expenditure"} recorded successfully.', 'success')
        return redirect(url_for('admin_payments'))

    payments = load_json(PAYMENTS_FILE, [])
    credit_refs, debit_refs = get_reference_sheet_data()
    # Pre-group credit refs by apartment for efficient JS rendering
    credit_refs_by_apt = {}
    for r in credit_refs:
        credit_refs_by_apt.setdefault(r['apartment'], []).append(r['name'])
    return render_template('admin_payments.html',
                           apartments=apartments,
                           payments=payments,
                           year_choices=year_choices,
                           current_year=current_year,
                           credit_refs=credit_refs,
                           credit_refs_by_apt=credit_refs_by_apt,
                           debit_refs=debit_refs)

@app.route('/admin/payments/delete/<int:payment_id>', methods=['POST'])
@admin_required
def delete_payment(payment_id):
    payments = load_json(PAYMENTS_FILE, [])

    # Capture the entry *before* removing it so we can sync the deletion to GS.
    target = next((p for p in payments if p.get('id') == payment_id), None)

    payments = [p for p in payments if p.get('id') != payment_id]
    save_json(PAYMENTS_FILE, payments)

    # Sync deletion to Google Sheets when configured
    if target and _gs.is_configured():
        month_name = target.get('month', '')
        month_num  = MONTH_NAME_TO_NUM.get(month_name.lower(), 0)
        entry_year = target.get('year', 0)

        if month_num and entry_year:
            # Remove the row from the Cr Dr sheet
            gs_err = _gs.delete_transaction(
                month_num=month_num,
                year=entry_year,
                apartment=target.get('apartment', ''),
                date_str=target.get('date', ''),
                description=target.get('name', ''),
                debit=target.get('debit', 0),
                credit=target.get('credit', 0),
            )
            if gs_err:
                flash(
                    f'Entry deleted locally but could not remove from Google Sheets Cr Dr: {gs_err} '
                    'Visit Admin \u2192 G Sheets to diagnose.',
                    'warning',
                )

            # If it was a payment (credit), also clear the Monthly Contributions cell
            if target.get('type') == 'payment' and target.get('credit', 0) > 0:
                contrib_err = _gs.clear_monthly_contribution(
                    month_num=month_num,
                    year=entry_year,
                    flat=target.get('apartment', ''),
                )
                if contrib_err:
                    flash(
                        f'Cr Dr row removed but could not clear Monthly Contributions: {contrib_err} '
                        'Visit Admin \u2192 G Sheets to diagnose.',
                        'warning',
                    )

    flash('Entry deleted.', 'success')
    return redirect(url_for('admin_payments'))

@app.route('/admin/announcements', methods=['GET', 'POST'])
@admin_required
def admin_announcements():
    if request.method == 'POST':
        action = request.form.get('action')
        announcements = load_json(ANNOUNCEMENTS_FILE, [])

        if action == 'delete':
            ann_id = int(request.form.get('ann_id', -1))
            announcements = [a for a in announcements if a.get('id') != ann_id]
            save_json(ANNOUNCEMENTS_FILE, announcements)
            flash('Announcement deleted.', 'success')
        else:
            title   = request.form.get('title', '').strip()
            content = request.form.get('content', '').strip()
            if title and content:
                attachment = None
                file = request.files.get('attachment')
                if file and file.filename:
                    # Prefix with timestamp to avoid overwriting files with the same name
                    ts = datetime.now().strftime('%Y%m%d%H%M%S')
                    filename = f"{ts}_{secure_filename(file.filename)}"
                    file.save(os.path.join(ANNOUNCEMENT_UPLOAD_FOLDER, filename))
                    attachment = filename
                announcements.append({
                    'id': int(datetime.now().timestamp()),
                    'title': title,
                    'content': content,
                    'date': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'author': session['user']['name'],
                    'attachment': attachment,
                })
                save_json(ANNOUNCEMENTS_FILE, announcements)
                flash('Announcement posted.', 'success')
            else:
                flash('Title and content are required.', 'danger')
        return redirect(url_for('admin_announcements'))

    announcements = load_json(ANNOUNCEMENTS_FILE, [])
    announcements = sorted(announcements, key=lambda x: x.get('date', ''), reverse=True)
    return render_template('admin_announcements.html', announcements=announcements)

@app.route('/announcements/attachment/<filename>')
@login_required
def announcement_attachment(filename):
    return send_from_directory(ANNOUNCEMENT_UPLOAD_FOLDER, secure_filename(filename))

@app.route('/admin/upload', methods=['GET', 'POST'])
@admin_required
def admin_upload():
    uploaded_files = []
    for f in os.listdir(UPLOAD_FOLDER):
        if f.endswith(('.xlsx', '.xls')):
            uploaded_files.append(f)

    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file selected.', 'danger')
            return redirect(url_for('admin_upload'))
        file = request.files['file']
        if file.filename == '':
            flash('No file selected.', 'danger')
            return redirect(url_for('admin_upload'))
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(UPLOAD_FOLDER, filename))
            flash(f'File "{filename}" uploaded successfully.', 'success')
            return redirect(url_for('admin_upload'))
        flash('Only .xlsx and .xls files are allowed.', 'danger')

    return render_template('admin_upload.html', uploaded_files=uploaded_files)

@app.route('/admin/balance-sheet')
@admin_required
def admin_balance_sheet():
    fy_filter = request.args.get('fy')
    sheets, available_fys = get_balance_sheets(fy_filter=fy_filter)
    selected_fy = fy_filter if fy_filter in available_fys else (available_fys[0] if available_fys else None)
    if selected_fy and not fy_filter:
        sheets, _ = get_balance_sheets(fy_filter=selected_fy)
    return render_template('admin_balance_sheet.html',
                           sheets=sheets,
                           available_fys=available_fys,
                           selected_fy=selected_fy)

@app.route('/balance-sheet/detail/<fy>/<int:month_num>')
@login_required
def balance_sheet_detail(fy, month_num):
    rows = get_balance_sheet_detail(fy, month_num)
    month_name = MONTH_NUM_TO_NAME.get(month_num, '')
    cal_year = None
    for fc in FY_CONFIG:
        if fc[0] == fy:
            cal_year = fc[2] if month_num >= 4 else fc[3]
            break
    display_month = f'{month_name} {cal_year}' if cal_year else month_name
    return render_template('balance_sheet_detail.html',
                           rows=rows,
                           display_month=display_month,
                           fy=fy)

@app.route('/admin/apartments')
@admin_required
def admin_apartments():
    users = load_users()
    contributions = get_monthly_contributions()
    apt_data = []
    for apt in APARTMENT_LIST:
        flat = apt['flat']
        user = users.get(flat, {})
        contrib = contributions.get(flat, {})
        total_paid = sum(v for v in contrib.values() if v and isinstance(v, (int, float)))
        apt_data.append({
            'flat': flat,
            'owner': apt['owner'],
            'block': apt['block'],
            'total_paid': total_paid,
            'must_change_password': user.get('must_change_password', True),
        })
    return render_template('admin_apartments.html', apartments=apt_data)

@app.route('/admin/reset-password/<flat>', methods=['POST'])
@admin_required
def admin_reset_password(flat):
    users = load_users()
    key = flat if flat != 'admin' else 'admin'
    if key in users:
        users[key]['password'] = hash_password('Welcome')
        users[key]['must_change_password'] = True
        _save_users(users)
        flash(f'Password reset to "Welcome" for {flat}.', 'success')
    else:
        flash('User not found.', 'danger')
    return redirect(url_for('admin_apartments'))

@app.route('/admin/google-sheets-status')
@admin_required
def admin_google_sheets_status():
    """Show the current Google Sheets connection status and env-var configuration."""
    configured = _gs.is_configured()
    env_vars = {
        'GOOGLE_SHEETS_CREDENTIALS_FILE': os.environ.get('GOOGLE_SHEETS_CREDENTIALS_FILE', ''),
        'GOOGLE_SHEET_USERS_ID':          os.environ.get('GOOGLE_SHEET_USERS_ID', ''),
        'GOOGLE_SHEET_CONTRIBUTIONS_ID':  os.environ.get('GOOGLE_SHEET_CONTRIBUTIONS_ID', ''),
        'GOOGLE_SHEET_BALANCE_ID':        os.environ.get('GOOGLE_SHEET_BALANCE_ID', ''),
        'GOOGLE_SHEET_CRDR_ID':           os.environ.get('GOOGLE_SHEET_CRDR_ID', ''),
    }
    diagnostics = _gs.get_diagnostics() if configured else None
    return render_template('admin_google_sheets_status.html',
                           configured=configured,
                           env_vars=env_vars,
                           diagnostics=diagnostics)


@app.route('/admin/google-sheets-diagnose')
@admin_required
def admin_google_sheets_diagnose():
    """JSON endpoint — returns live diagnostic data for each spreadsheet."""
    return jsonify(_gs.get_diagnostics())

@app.context_processor
def inject_globals():
    return {
        'now': datetime.now(),
        'google_sheets_active': _gs.is_configured(),
    }

# Initialise data when the module is loaded (works for both direct run and WSGI)
with app.app_context():
    init_data()

if __name__ == '__main__':
    debug_mode = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=debug_mode, host='0.0.0.0', port=port)
