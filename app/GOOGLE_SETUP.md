# Google Sheets Integration – Step-by-Step Setup Guide

This guide explains how to move the Apartment Management app from reading
local Excel files to reading live **Google Sheets** stored in your Google
Drive.

Once set up the app will continue to work exactly as before – the same
pages, the same data – but every number will come from your Google Sheets
instead of the `.xlsx` files on disk.  You can edit the sheets in Google
Drive and the website will reflect the changes immediately.

---

## What you will need

| Item | Where to get it |
|------|----------------|
| A Google account | Already have one (Gmail / Workspace) |
| The three Google Sheets (see Step 3) | You create them in Google Drive |
| A Service Account JSON key file | Created in Google Cloud Console (Step 2) |

---

## Step 1 – Create a Google Cloud Project

1. Open [https://console.cloud.google.com/](https://console.cloud.google.com/) and sign in.
2. Click the project selector at the top-left → **New Project**.
3. Give it a name like `apartment-mgmt` and click **Create**.
4. Make sure the new project is selected in the top-left dropdown.

---

## Step 2 – Enable APIs and Create a Service Account

### 2a – Enable the required APIs

1. In the left sidebar go to **APIs & Services → Library**.
2. Search for **Google Sheets API** → click it → click **Enable**.
3. Search for **Google Drive API** → click it → click **Enable**.

### 2b – Create a Service Account

1. Go to **APIs & Services → Credentials**.
2. Click **+ Create Credentials → Service account**.
3. Name it `apartment-app` (or anything you like) → click **Create and
   Continue** → skip optional role/user fields → click **Done**.

### 2c – Download the JSON Key

1. On the Credentials page, click the service account you just created.
2. Go to the **Keys** tab → **Add Key → Create new key**.
3. Choose **JSON** → click **Create**.
4. A `.json` file is downloaded to your computer – keep it safe.
   **Do NOT commit this file to GitHub.**  It is a secret password.
5. Copy the file to a safe directory on the server that runs the app,
   for example `/home/you/apartment_creds.json`.

> **Security tip:** The `.gitignore` in this repo already excludes files
> matching `*service_account*.json`, `*credentials*.json`, and `*-key.json`.
> Rename your file to match one of those patterns, or store it outside the
> repository folder.

---

## Step 3 – Prepare your Google Sheets

The app expects three separate Google Sheets, each mirroring the layout of
the corresponding Excel file.

### Option A – Upload your existing Excel files (easiest)

1. Go to [https://drive.google.com/](https://drive.google.com/).
2. Click **+ New → File upload** and upload each of the three `.xlsx` files:
   - `Monthly Contribution_2025_26.xlsx`
   - `Current Balance_Sheet_2024_25_26.xlsx`
   - `balance sheet Cr Dr.xlsx`
3. After uploading, **right-click each file → Open with → Google Sheets**.
   This opens the file in the Google Sheets editor.
4. **⚠️ Important — convert the file:** Inside Google Sheets, click
   **File → Save as Google Sheets**.  This creates a *new* native Google
   Sheets document.  You **must** do this step — just "opening" an Excel file
   in Sheets is not enough.  Write operations (saving new transactions) will
   fail with `[400]: This operation is not supported for this document` if you
   skip this step.
5. The URL of the new native file will now look like:
   `https://docs.google.com/spreadsheets/d/`**`1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms`**`/edit`
   Copy the **Spreadsheet ID** from the *new file's URL* (not the original
   uploaded Excel file).
6. You can delete the original `.xlsx` file from Drive once you have copied
   its data into the native Google Sheets document.

> **How to tell if your file is native Google Sheets:**  
> Open the Admin → G Sheets status page (`/admin/google-sheets-status`).
> If any spreadsheet shows an orange **"Excel format — read-only"** warning,
> that file needs to be converted (File → Save as Google Sheets).

### Option B – Create new Google Sheets manually

Create three new Google Sheets in Drive and replicate the column layout from
the Excel files (see the `app.py` comments for the exact expected structure).

---

## Step 4 – Share each Spreadsheet with the Service Account

The app connects to Google Sheets **as the service account**, not as your
personal Google account.  You must give the service account permission to
read (and write) the sheets.

1. Open each of the three Google Sheets.
2. Click the **Share** button (top-right).
3. In the "Add people" box paste the service account e-mail address.
   It looks like:  `apartment-app@apartment-mgmt.iam.gserviceaccount.com`
   (Find it in the JSON key file under the `"client_email"` field, or on the
   Credentials page in Google Cloud Console.)
4. Set the role to **Editor** (so the app can also write new transactions).
5. Un-tick "Notify people" → click **Share**.

Repeat for all three spreadsheets.

---

## Step 5 – Set Environment Variables

The app needs to know:

1. **Where** the JSON key file is on your computer.
2. **Which** Google Sheets spreadsheets to open.

These pieces of information are passed to the app as **environment
variables** — special named settings that live outside the code.

| Variable | What to put |
|----------|-------------|
| `GOOGLE_SHEETS_CREDENTIALS_FILE` | The **full path** to the JSON key file you downloaded in Step 2c |
| `GOOGLE_SHEET_USERS_ID`          | The **Spreadsheet ID** of the *User Credentials* sheet (see Step 5U below) |
| `GOOGLE_SHEET_CONTRIBUTIONS_ID`  | The **Spreadsheet ID** of *Monthly Contribution_2025_26* |
| `GOOGLE_SHEET_BALANCE_ID`        | The **Spreadsheet ID** of *Current Balance_Sheet_2024_25_26* |
| `GOOGLE_SHEET_CRDR_ID`           | The **Spreadsheet ID** of *balance sheet Cr Dr* |

> **`GOOGLE_SHEET_USERS_ID` is strongly recommended.**  When set, the app
> stores all usernames and password hashes in a Google Sheet instead of (or
> in addition to) the local `data/users.json` file.  Passwords are stored as
> SHA-256 hashes — **plain-text passwords are never written to the sheet.**
> Even the admin cannot see or recover passwords from the sheet.

---

### Step 5U – Create the User Credentials Google Sheet

1. Go to [https://drive.google.com/](https://drive.google.com/) and click
   **+ New → Google Sheets → Blank spreadsheet**.
2. Give it a name like `Apartment Users`.
3. **Share it with the service account** (same as Step 4):
   - Click **Share** (top-right).
   - Paste the service account e-mail (find it in the JSON key file under
     `"client_email"`).
   - Set role to **Editor** → click **Share**.
4. Copy the **Spreadsheet ID** from the URL (see below).
5. Leave the sheet completely blank — the app will create the header row and
   populate it with all users the first time any password is changed or reset.

> **Sheet layout (created automatically by the app):**
>
> | username | password_hash | role | name | flat | block | must_change_password |
> |----------|---------------|------|------|------|-------|----------------------|
> | admin | `sha256hash` | admin | Administrator | | | TRUE |
> | RG1 | `sha256hash` | owner | Ashok Kumar | RG1 | Rajkumar | TRUE |
> | … | … | … | … | … | … | … |
>
> The `password_hash` column contains a SHA-256 digest, not a plain-text
> password.  It is intentionally unreadable so that nobody — including the
> admin — can see a user's password from the spreadsheet.

---

### How to find your Spreadsheet ID

Open each Google Sheet in your browser.  The URL will look like this:

```
https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/edit
```

The long string between `/d/` and `/edit` is the **Spreadsheet ID**.  Copy
it — it is different for every sheet.

---

### Option A – Edit the startup script you already use (easiest ✅)

The startup scripts (`start.bat`, `start.ps1`, `start.sh`) already contain
a ready-to-fill Google Sheets block near the top.  You only need to:

**Step 5A-1 – Open the script in a text editor.**

- Windows: right-click `start.bat` → *Edit* (or open with Notepad).
- Windows PowerShell users: right-click `start.ps1` → *Edit* (or open with
  Notepad / VS Code).
- Linux / macOS: open `start.sh` in any text editor.

**Step 5A-2 – Find the Google Sheets block.** It looks like this in
`start.bat`:

```bat
REM set GOOGLE_SHEETS_CREDENTIALS_FILE=C:\Users\YourName\Downloads\apartment-mgmt-abc123.json
REM set GOOGLE_SHEET_CONTRIBUTIONS_ID=PASTE_SPREADSHEET_ID_FOR_MONTHLY_CONTRIBUTION
REM set GOOGLE_SHEET_BALANCE_ID=PASTE_SPREADSHEET_ID_FOR_CURRENT_BALANCE_SHEET
REM set GOOGLE_SHEET_CRDR_ID=PASTE_SPREADSHEET_ID_FOR_BALANCE_SHEET_CR_DR
REM set GOOGLE_SHEET_USERS_ID=PASTE_SPREADSHEET_ID_FOR_USER_CREDENTIALS
```

**Step 5A-3 – Remove the `REM ` from the start of each of the five lines**
(that turns them from comments into active commands).

**Step 5A-4 – Replace the placeholder text** with your real values.

For example, if your JSON key file is at
`C:\Users\Kiran\Downloads\apartment-mgmt-abc123.json`, your three financial
Spreadsheet IDs are `AAAA`, `BBBB`, `CCCC`, and your Users sheet ID is
`DDDD`, the five lines should look like this:

```bat
set GOOGLE_SHEETS_CREDENTIALS_FILE=C:\Users\Kiran\Downloads\apartment-mgmt-abc123.json
set GOOGLE_SHEET_CONTRIBUTIONS_ID=AAAA
set GOOGLE_SHEET_BALANCE_ID=BBBB
set GOOGLE_SHEET_CRDR_ID=CCCC
set GOOGLE_SHEET_USERS_ID=DDDD
```

**Step 5A-5 – Save the file.**  Double-click `start.bat` (or run `start.sh`)
to start the app.  Go to Step 6.

> **Windows PowerShell users:** remove the `# ` (hash-space) from the start of each
> `$env:` line instead of `REM `.
>
> **Linux / macOS users:** remove the `# ` from each `export` line.

---

### Option B – Use a `.env` file (recommended for always-on servers)

1. Find the file named `.env.example` in the project folder.
2. Make a copy of it and name the copy `.env` (no ".example").
3. Open `.env` in a text editor and replace each `PASTE_...` placeholder
   with your real value.
4. The file will look like this when filled in:
   ```
   GOOGLE_SHEETS_CREDENTIALS_FILE=C:\Users\Kiran\Downloads\apartment-mgmt-abc123.json
   GOOGLE_SHEET_USERS_ID=DDDD
   GOOGLE_SHEET_CONTRIBUTIONS_ID=AAAA
   GOOGLE_SHEET_BALANCE_ID=BBBB
   GOOGLE_SHEET_CRDR_ID=CCCC
   ```
5. Install `python-dotenv` so the app can read it:
   ```
   pip install python-dotenv
   ```
6. Add these two lines near the top of `app.py` (just after the existing
   `import os` line):
   ```python
   from dotenv import load_dotenv
   load_dotenv()
   ```
7. Save `app.py` and start the app.

> **Security:** `.env` is already in `.gitignore`.  It will never be
> accidentally pushed to GitHub.

---

### Option C – Set them manually in the terminal (for testing only)

**Windows Command Prompt** (type these before `python app.py`):

```bat
set GOOGLE_SHEETS_CREDENTIALS_FILE=C:\Users\Kiran\Downloads\apartment-mgmt-abc123.json
set GOOGLE_SHEET_USERS_ID=DDDD
set GOOGLE_SHEET_CONTRIBUTIONS_ID=AAAA
set GOOGLE_SHEET_BALANCE_ID=BBBB
set GOOGLE_SHEET_CRDR_ID=CCCC
python app.py
```

**Windows PowerShell** (type these before `python app.py`):

```powershell
$env:GOOGLE_SHEETS_CREDENTIALS_FILE = "C:\Users\Kiran\Downloads\apartment-mgmt-abc123.json"
$env:GOOGLE_SHEET_USERS_ID          = "DDDD"
$env:GOOGLE_SHEET_CONTRIBUTIONS_ID  = "AAAA"
$env:GOOGLE_SHEET_BALANCE_ID        = "BBBB"
$env:GOOGLE_SHEET_CRDR_ID           = "CCCC"
python app.py
```

**Linux / macOS terminal** (type these before `python app.py`):

```bash
export GOOGLE_SHEETS_CREDENTIALS_FILE="/home/kiran/apartment-mgmt-abc123.json"
export GOOGLE_SHEET_USERS_ID="DDDD"
export GOOGLE_SHEET_CONTRIBUTIONS_ID="AAAA"
export GOOGLE_SHEET_BALANCE_ID="BBBB"
export GOOGLE_SHEET_CRDR_ID="CCCC"
python app.py
```

> **Note:** With Option C, the variables are lost when you close the terminal
> window.  Use Option A or B for a permanent setup.

---

## Step 6 – Install the new dependencies

```bash
pip install -r requirements.txt
```

This adds `gspread` and `google-auth` (the libraries used to talk to Google
Sheets).

---

## Step 7 – Start the app and verify

1. Start the app as usual (`python app.py` or `start.bat`).
2. Log in as **admin**.
3. Go to **G Sheets** in the top navigation bar (or visit
   `/admin/google-sheets-status`).
4. You should see "Connected" with all environment variables showing
   as **Set**.  `GOOGLE_SHEET_USERS_ID` may show "Missing" if you skipped
   Step 5U — the app will continue to work using the local JSON file in that
   case, but setting it is strongly recommended.

If any variable shows **Missing**, go back to Step 5 and check the value.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| "Not connected" even though all vars are set | JSON key file path is wrong | Check `GOOGLE_SHEETS_CREDENTIALS_FILE` is an absolute path and the file exists |
| Data shows blank / empty tables | Spreadsheet not shared with service account | Re-do Step 4 |
| `gspread.exceptions.SpreadsheetNotFound` in logs | Wrong Spreadsheet ID | Re-check the ID in the sheet URL |
| `google.auth.exceptions.TransportError` | No internet access from the server | Ensure the server can reach `googleapis.com` |
| App still shows local Excel data | `GOOGLE_SHEETS_CREDENTIALS_FILE` env-var not loaded | Restart the app after setting the variable |
| `[400]: This operation is not supported for this document` | Spreadsheet is still in Excel (.xlsx) format in Drive — write operations are blocked | Open the file in Google Sheets → **File → Save as Google Sheets** → update the Spreadsheet ID in your startup script → restart the app |

---

## Reverting to local Excel files

Simply unset (or remove) the `GOOGLE_SHEETS_CREDENTIALS_FILE` environment
variable and restart the app.  It will fall back to the local `.xlsx` files
automatically – no code changes needed.

---

## Sheet layout reference

The Google Sheets must keep the same column layout as the original Excel
files.  Here is a quick reference:

### Monthly Contribution_2025_26 → Sheet tab: `Sheet1`

| Col index | Content |
|-----------|---------|
| 4 (E) | Flat number (e.g. `RS3`) |
| 6–17 (G–R) | Monthly amounts: Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec, Jan, Feb, Mar |

### Current Balance_Sheet_2024_25_26 → Report tabs: `25-26 Report`, `2024-25 Report`

| Col index | Content |
|-----------|---------|
| 0 (A) | Month name (e.g. `Apr`, `January`) |
| 2 (C) | Total income |
| 3 (D) | Total expense |

Detail tabs (e.g. `Jan 2026`, `Feb 2026`):

| Col index | Content |
|-----------|---------|
| 1 (B) | Date |
| 2 (C) | Description |
| 3 (D) | Income |
| 4 (E) | Expense |

### balance sheet Cr Dr → one tab per month (e.g. `Jul 2024`, `Jan 2026`)

| Col index | Content |
|-----------|---------|
| 0 (A) | Apartment number |
| 1 (B) | Date |
| 2 (C) | Name / description |
| 3 (D) | Dr (debit) |
| 4 (E) | Cr (credit) |

Keep a tab named **`Template`** (used when a new month tab is created
automatically) and a tab named **`Refernce Sheet`** (kept at the end).

> **Note:** `Refernce Sheet` is intentionally misspelled to exactly match the
> original Excel file's tab name.  The app uses this exact string – do not
> rename that tab.
