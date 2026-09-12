# Hitech Citadel Phase 2 – Apartment Management System

A web-based apartment management application for the Hitech Citadel Phase 2 Flat Owners Association.
Data is stored in the existing Excel files (`.xlsx`) — no database required.

---

## 🏗️ Architecture & Implementation Summary

The project has evolved into a robust, deployment-ready apartment management platform with the following key components and integrations:

### 🔌 Dual-Backend Data Architecture
The application runs dynamically with one of two data persistence modes (with automatic fallback):
* **Local Excel Files (`.xlsx`)**: Reads/writes directly to the files in the [Apartment/](file:///c:/Users/mskir/Appartment/Apartment/Apartment) folder.
* **Google Sheets API Backend**: Activated when Google Cloud credentials and Spreadsheet IDs are configured as environment variables (see [GOOGLE_SETUP.md](GOOGLE_SETUP.md)). The application uses the `gspread` client library to read/write live Google Sheets directly, rendering local files obsolete for cloud-hosted environments.

### 🛡️ Security & Authentication
* **Password Hashing**: User credentials (passwords) are securely hashed before being stored in `users.json` or the `users` tab of Google Sheets, replacing previous plain-text storage.
* **First-Login Setup**: Users are forced to set a new password and provide their email on their first login.
* **Forgot Password Flow**: Admin-configurable password recovery sending reset links/codes using three robust strategy levels:
  1. **Gmail API via Service Account** (decoupled from Sheet credentials).
  2. **SMTP Relay** with Port 465 (SSL) or Port 587 (STARTTLS) using App Passwords.
  3. **SendGrid HTTP API** fallback (essential for platforms like Render free tier that block outbound SMTP ports).

### ✍️ Payment & Expense Recording (Reference Sheet Integration)
* **Resident Name-to-Flat Autofill**: Admin payment forms automatically search and link resident names with their flat numbers.
* **Reference Dropdown Lists**: Payment and expenditure entries pull pre-grouped credit and debit references (e.g., specific expense categories), minimizing manual entry errors.
* **Cash Expense Support**: Dedicated logging of cash-based transactions directly into the ledger/balance sheets.
* **Month-End Sheet Creation**: Automatic creation of clean, month-header-safe worksheets for new financial months without copying corrupted templates.

### 🧮 Monthly Balance Sheet Generator
* **Unified Command-Line Generator**: Added `generate_monthly_balance_sheet.py` which dynamically reads credits (contributions) and debits (cheque and cash expenses) from the CR/DR spreadsheet, calculates opening balances and cash carry-forwards from the previous month, and builds a fully styled, formatted balance sheet in the Balance Sheet spreadsheet.
* **Automatic Summary Row & Double-Count Filtering**: Credits are strictly matched against the 30 valid flat numbers (`_VALID_FLATS`), and empty rows or bottom-most total summary rows are automatically ignored.
* **Dynamic Cash Withdrawal & Bank Interest Support**: Automatically detects cash withdrawals (even if misspelled as "Withrawal") and bank interest credits, dynamically linking them to the correct sections of the generated sheet.
* **Whitespace-Insensitive Sheet Mapping**: Web dashboard matching collapses spaces dynamically, allowing matching of sheets like `April  2026` or `April 2026` seamlessly.
* **Summary Fallback Scanning**: When summary reports tabs (like `26-27 Report`) are missing, the web dashboard automatically parses monthly totals from the detailed tabs in real-time.

### 🔍 Apartment Payment History & Status Finder
* **Command-Line Payment Lookup**: Added `find_apartment_payments.py` which retrieves the entire historical record of maintenance payments for a given flat directly from the Google Sheets CR/DR transaction database.
* **Aggregated Metrics**: Displays chronological payment dates, paid-by names, amounts, total amount paid, total number of payments, and the list of unique months paid (e.g. `python find_apartment_payments.py KG4`).
* **Payments Till Date Dashboard**: Integrated a new webpage option directly into the navigation bar (`Payments Till Date` for both owners and admins).
  * Grouped and split by **FY 2025-26** (expected: Rs. 2,600/month) and **FY 2026-27** (expected: Rs. 2,800/month).
  * Color-coded status tiles indicating **Fully Paid** (green), **Partial Payment** (yellow/orange, $\ge$ Rs. 2,000), and **Unpaid** (red, $<$ Rs. 2,000).
  * Detailed transaction audit table with dates, resident names, and individual payment logs.
  * For admins, includes a searchable dropdown selector to view status for any flat.
  * Optimized layout with responsive grid sizes to prevent badge text wrapping on mobile screens.
* **Payments Master Summary Matrix**: Excluded from the main navigation bar but accessible via the direct link `/admin/payments-summary` (admin-only).
  * Presents a side-by-side color-coded grid of all 30 flats and their month-by-month allocated payments for both FY 2025-26 and FY 2026-27.
  * Includes a real-time javascript filter search bar for flat names, owners, or blocks.
  * Shows grand total paid and prepayment/advance balances for each row.
  * Mobile-optimized layout freezing the **Flat** column on the left via sticky CSS positioning to allow seamless horizontal scrolling on phone screens.
  * Includes a dedicated 2-page print toolbar with buttons to print both FYs on 2 clean pages (FY 2025-26 on Page 1, FY 2026-27 on Page 2) or print either FY individually.
* **Read-Only Admin Account**: A special account is configured for viewing purposes:
  * **Username**: `admin1`
  * **Password**: `CheckThis2002`
  * **Role**: `admin_viewer`
  * Has access to all admin dashboards, matrix summaries, balance sheets, and transaction logs, but any state-changing operations (POST requests to admin routes like adding/deleting payments, posting announcements, uploading files, or resetting passwords) are automatically intercepted and blocked with an alert.

### ☁️ Cloud Deployments
In addition to running locally via double-clickable startup scripts (`start.bat`, `start.ps1`, `start.sh`), the project contains configuration files ready for deployment on:
* **Render**: via `render.yaml`
* **Railway**: via `railway.toml` and `Procfile`
* **Fly.io**: via `fly.toml` and `Dockerfile`
* **GitHub Actions**: automated CI pipeline and Docker deploy triggers.

---

## 🚀 How to Run

### Prerequisites – Install Python 3

You need **Python 3.8 or newer** installed on your computer.

| OS | Download / Install |
|---|---|
| **Windows** | Download from [python.org/downloads](https://www.python.org/downloads/) — during installation tick **"Add Python to PATH"** |
| **macOS** | `brew install python` or download from [python.org/downloads](https://www.python.org/downloads/) |
| **Linux (Ubuntu/Debian)** | `sudo apt update && sudo apt install python3 python3-pip` |

Verify Python is installed by opening a terminal and running:
```
python --version
```
You should see `Python 3.x.x`.

---

### Option A – Double-click to start (easiest)

| Your OS | Script to double-click |
|---|---|
| **Windows** (Explorer) | `start.bat` |
| **Windows** (PowerShell) | Right-click `start.ps1` → *Run with PowerShell* |
| **macOS / Linux** | Run `./start.sh` in a terminal |

> **macOS / Linux first-time only:** make the script executable:
> ```bash
> chmod +x start.sh
> ./start.sh
> ```

The script will automatically:
1. Check that Python 3 is installed
2. Install the required packages (`Flask`, `openpyxl`)
3. Start the web server
4. Open your browser to **http://localhost:5000**

---

### Option B – Manual steps (any OS)

Follow all four steps below in order.  **Steps 2 and 3** only need to be done the very first time.

---

#### Step 1 – Open a terminal **inside** the project folder

> This is the most common stumbling block.  The terminal must be opened *in the same folder that contains `app.py`*.

**Windows (Command Prompt)**
1. Open **File Explorer** and browse to the folder that contains `app.py` (e.g. `C:\Users\YourName\Downloads\Apartment`).
2. Click in the **address bar** at the top of File Explorer so the path is highlighted.
3. Type `cmd` and press **Enter**.
   A black Command Prompt window opens, already in the right folder.

**Windows (PowerShell)**
1. Open **File Explorer** and browse to the folder that contains `app.py`.
2. Hold **Shift** and **right-click** an empty area in the folder.
3. Choose **"Open PowerShell window here"** (or **"Open in Terminal"** on Windows 11).

**macOS**
1. Open **Finder** and browse to the folder that contains `app.py`.
2. Right-click (or Ctrl-click) the folder → **"New Terminal at Folder"**.
   *(If that option is missing: go to **System Preferences → Keyboard → Shortcuts → Services** and enable "New Terminal at Folder".)*

**Linux**
1. Open your file manager, navigate to the folder with `app.py`.
2. Right-click → **"Open Terminal Here"** (exact wording depends on your file manager).
3. Alternatively, open any terminal and type:
   ```bash
   cd /path/to/Apartment
   ```

You can verify you are in the right folder by typing:
```
# Windows (cmd/PowerShell)
dir app.py

# macOS / Linux
ls app.py
```
You should see `app.py` listed.

---

#### Step 2 – Install Python (first time only)

If `python --version` prints a version number you can skip this step.

| OS | Command / Download |
|---|---|
| **Windows** | Download from [python.org/downloads](https://www.python.org/downloads/) — tick **"Add Python to PATH"** during installation |
| **macOS** | `brew install python` *or* download from [python.org/downloads](https://www.python.org/downloads/) |
| **Linux (Ubuntu/Debian)** | `sudo apt update && sudo apt install python3 python3-pip` |

---

#### Step 3 – Install dependencies (first time only)

In the terminal you opened in Step 1, run **one** of the following commands.  Try them in order until one works:

```bash
# Primary command (works on most systems)
pip install -r requirements.txt
```

If you see `pip: command not found` or `'pip' is not recognized`:

```bash
# Alternative 1 – use pip3 explicitly
pip3 install -r requirements.txt
```

```bash
# Alternative 2 – invoke pip via Python directly
python -m pip install -r requirements.txt
```

```bash
# Alternative 3 – on Linux/macOS where python3 is the default
python3 -m pip install -r requirements.txt
```

A successful install ends with a line like:
```
Successfully installed Flask-... openpyxl-... Werkzeug-...
```

> **Tip – virtual environment (recommended for developers)**
> A virtual environment keeps these packages isolated from other Python projects:
> ```bash
> # Create it once
> python -m venv venv
>
> # Activate it – Windows cmd
> venv\Scripts\activate.bat
>
> # Activate it – Windows PowerShell
> venv\Scripts\Activate.ps1
>
> # Activate it – macOS / Linux
> source venv/bin/activate
>
> # Then install as usual
> pip install -r requirements.txt
> ```
> Run the activation command every time you open a new terminal before starting the app.

---

#### Step 4 – Start the app

```bash
# Windows
python app.py

# macOS / Linux (if 'python' is not found, use python3)
python3 app.py
```

Open your browser and go to **http://localhost:5000**.

Press **Ctrl + C** in the terminal to stop the server.

---

## 🔑 Default Login Credentials

| Role | Username | Password |
|---|---|---|
| **Admin** | `admin` | `Welcome` |
| Flat RG1 – Ashok Kumar | `RG1` | `Welcome` |
| Flat RG2 – Lakshminarayana Rao | `RG2` | `Welcome` |
| Flat RG3 – B R Vamshi Krishna | `RG3` | `Welcome` |
| Flat RF1 – Venkatesh Reddy | `RF1` | `Welcome` |
| Flat RF2 – Sunitha Sudarshan | `RF2` | `Welcome` |
| Flat RF3 – Dharma Kumar | `RF3` | `Welcome` |
| Flat RS1 – D H Sohan | `RS1` | `Welcome` |
| Flat RS2 – Aravind Shetty P N | `RS2` | `Welcome` |
| Flat RS3 – Kiran M S | `RS3` | `Welcome` |
| Flat VG1 – Mukesh Gupta | `VG1` | `Welcome` |
| Flat VG2 – Shreyas R Kulkarni | `VG2` | `Welcome` |
| Flat VG3 – K G Phaniraj | `VG3` | `Welcome` |
| Flat VF1 – Pushpa Badmi | `VF1` | `Welcome` |
| Flat VF2 – B R Ramesh | `VF2` | `Welcome` |
| Flat VF3 – Manu M | `VF3` | `Welcome` |
| Flat VS1 – Prakash H S | `VS1` | `Welcome` |
| Flat VS2 – Mahesh Kumar J S | `VS2` | `Welcome` |
| Flat VS3 – Sai Prasad V | `VS3` | `Welcome` |
| Flat KG1 – Raghavendra B H | `KG1` | `Welcome` |
| Flat KG2 – Devipriya Ramesh | `KG2` | `Welcome` |
| Flat KG3 – Usha B N | `KG3` | `Welcome` |
| Flat KG4 – Sreekanth Bilihalli | `KG4` | `Welcome` |
| Flat KF1 – Suman S Nayak | `KF1` | `Welcome` |
| Flat KF2 – Varun K Murthy | `KF2` | `Welcome` |
| Flat KF3 – K G Krishna Murthy | `KF3` | `Welcome` |
| Flat KF4 – Aneesh V Naik | `KF4` | `Welcome` |
| Flat KS1 – Ramesh Joshi | `KS1` | `Welcome` |
| Flat KS2 – Venkataramana S | `KS2` | `Welcome` |
| Flat KS3 – Shylaja Manjunath | `KS3` | `Welcome` |
| Flat KS4 – Srikanth S N | `KS4` | `Welcome` |

> ⚠️ Every user is **forced to change their password** the first time they log in.

---

## ✨ Features

### Apartment Owners (30 users)
- Login with apartment number and default password `Welcome`
- Forced password change on first login
- View monthly balance sheet (income & expenditure)
- View personal payment history and monthly contribution status (FY selector for 2025–26 / 2026–27)
- View announcements and notifications

### Admin (1 user)
- Dashboard with total income, expenses, and net balance
- Enter payments received from residents
- Enter expenditures
- View all apartment transactions (with live search)
- Post and delete announcements
- Upload new monthly balance sheet files
- View and reset passwords for all apartments

---

## 📁 Data Files (in `Apartment/` folder)

| File | Description |
|---|---|
| `Monthly Contribution_2025_26.xlsx` | Apartment names, blocks, and FY 2025–26 monthly contributions |
| `balance sheet Cr Dr.xlsx` | Transaction history (credits/debits) per month |
| `Current Balance_Sheet_2024_25_26.xlsx` | Monthly income & expenditure summary |

---

## ☁️ Deploy to the Cloud via GitHub (free, public URL)

Deploy the app to the internet so everyone in your building can access it from any device — no laptop needs to stay on.

### Recommended: Render.com (free tier)

Render gives you a public HTTPS URL (e.g. `https://apartment-mgmt.onrender.com`) that auto-updates every time you push to GitHub.

> ⚠️ **Important:** The cloud server has no persistent disk, so local `.xlsx` data files will be wiped on every deploy.  
> **You must set up Google Sheets integration** (see [GOOGLE_SETUP.md](GOOGLE_SETUP.md)) before deploying to the cloud. Once Google Sheets is active all data is stored there — not on the server.

---

#### Step 1 – Push your repo to GitHub

Make sure all your latest changes are committed and pushed to GitHub:

```bash
git add .
git commit -m "Ready for deployment"
git push
```

---

#### Step 2 – Create a free Render account

Go to **https://render.com** → **Sign Up** → choose **Sign up with GitHub**.  
This links Render to your GitHub account so it can deploy your repo automatically.

---

#### Step 3 – Create a new Web Service

1. In the Render dashboard click **New +** → **Web Service**.
2. Choose **Connect a repository** and select your `Apartment` repo.
3. Render will detect the `render.yaml` file and fill in the settings automatically.  
   If Render asks you to fill in the settings manually, use these values:

   | Field | Value |
   |---|---|
   | **Runtime** | Python 3 |
   | **Build Command** | `pip install -r requirements.txt` |
   | **Start Command** | `gunicorn app:app` |
   | **Instance Type** | Free |

   > 🚀 **Start Command** — if Render ever asks "how do you start your app?", the answer is:
   > ```
   > gunicorn app:app
   > ```
   > Copy and paste it exactly as shown above (no quotes, no extra spaces).

4. Click **Create Web Service**.

---

#### Step 4 – Add environment variables

Your app needs **five** settings (called *environment variables*) to run securely on the cloud.  
You add them one-by-one in the Render dashboard — no code editing required.

**Quick overview — all 5 credentials you will add in this step:**

| Sub-step | Variable name | What it is |
|---|---|---|
| 4-B | `SECRET_KEY` | A random password Flask uses to protect login sessions |
| 4-C | `GOOGLE_SHEETS_CREDENTIALS_JSON` | The full contents of the Google service-account JSON key file |
| 4-D | `GOOGLE_SHEET_CONTRIBUTIONS_ID` | The Spreadsheet ID of your *Monthly Contribution* Google Sheet |
| 4-E | `GOOGLE_SHEET_BALANCE_ID` | The Spreadsheet ID of your *Current Balance Sheet* Google Sheet |
| 4-F | `GOOGLE_SHEET_CRDR_ID` | The Spreadsheet ID of your *Balance Sheet Cr Dr* Google Sheet |

Follow sub-steps 4-A → 4-G in order. All five must be set before the app will work.

---

##### 4-A  Open the Environment tab

1. After creating the Web Service (Step 3) you will be on the service detail page.
2. Click the **Environment** tab in the row of tabs below the service name.
   You will see a section called **Environment Variables** with an **Add Environment Variable** button.

---

##### 4-B  Add `SECRET_KEY`

This is a secret password Flask uses to protect login sessions.  
It can be any long random string — nobody needs to read or remember it.

**How to generate one** (pick *one* of the options below):

- **Option 1 – Online generator:** go to [https://djecrety.ir/](https://djecrety.ir/) and copy the generated string.
- **Option 2 – Python (if installed on your computer):**  
  Open a terminal / Command Prompt and run:
  ```
  python -c "import secrets; print(secrets.token_hex(32))"
  ```
  Copy the output (a 64-character string like `3f8a2d...`).
- **Option 3 – Any password manager's "generate random password" feature** using at least 32 characters.

**Steps in Render:**

1. Click **Add Environment Variable**.
2. In the **Key** box type: `SECRET_KEY`
3. In the **Value** box paste your random string.
4. Click the ✓ (tick) or press **Enter** to save the row.

---

##### 4-C  Add `GOOGLE_SHEETS_CREDENTIALS_JSON`

**What this is:** When you set up Google Sheets integration (GOOGLE_SETUP.md Step 2), Google gave you a `.json` file called a *service-account key*. It is a secret password that lets the app talk to your Google Sheets. On Render there is no file system you can upload to, so instead you paste the *entire contents* of that file as a plain-text environment variable value.

---

**Part 1 – Find the JSON key file on your computer**

The file was automatically downloaded to your computer when you clicked "Create new key" in Google Cloud Console. It is almost certainly in your **Downloads** folder.

- **Windows:** Open File Explorer → click **Downloads** in the left sidebar.  
  Look for a file with a name like `apartment-mgmt-abc123-1a2b3c4d5e6f.json`.  
  (It starts with your project name and ends in `.json`.)
- **macOS:** Open Finder → click **Downloads** in the left sidebar.  
  The file has the same naming pattern.
- **Linux:** Check `~/Downloads/` in your file manager or terminal.

> 🔍 **Can't find it?** In Google Cloud Console go to  
> **APIs & Services → Credentials** → click the service account name →  
> **Keys** tab → **Add Key → Create new key → JSON → Create**.  
> A new `.json` file will download immediately.

---

**Part 2 – Open the file and copy its entire contents**

You need to open the file in a **plain-text editor** (not Word, not Google Docs — just plain text).

**Windows – Notepad:**
1. In File Explorer, right-click the `.json` file.
2. Choose **Open with** → **Notepad**.  
   (If Notepad is not listed, click *Choose another app* and select Notepad.)
3. The file opens and looks like this (yours will have real values):
   ```
   {
     "type": "service_account",
     "project_id": "apartment-mgmt",
     "private_key_id": "abc123...",
     "private_key": "-----BEGIN RSA PRIVATE KEY-----\nMIIE...\n-----END RSA PRIVATE KEY-----\n",
     "client_email": "apartment-app@apartment-mgmt.iam.gserviceaccount.com",
     "client_id": "1234567890",
     "auth_uri": "https://accounts.google.com/o/oauth2/auth",
     "token_uri": "https://oauth2.googleapis.com/token",
     ...
   }
   ```
4. Click anywhere inside the text area, then press **Ctrl + A** (selects everything).
5. Press **Ctrl + C** (copies it).

**macOS – TextEdit:**
1. In Finder, right-click the `.json` file → **Open With** → **TextEdit**.
2. If TextEdit opens in rich-text mode (you see a formatting toolbar), go to  
   **Format** menu → click **Make Plain Text**, then confirm.
3. Press **Cmd + A** (selects everything), then **Cmd + C** (copies it).

**Any computer – VS Code (if installed):**
1. Drag the `.json` file onto the VS Code window, or go to  
   **File → Open File** and select it.
2. Press **Ctrl + A** / **Cmd + A**, then **Ctrl + C** / **Cmd + C**.

> ✅ **You must copy from the very first `{` all the way to the very last `}`.  
> Do not copy just part of it. Do not add any extra spaces, quotes, or line breaks around it.**

---

**Part 3 – Paste into Render**

1. Go back to your browser where the **Render dashboard** is open.
2. Click **Add Environment Variable**.
3. In the **Key** box, type exactly:  
   ```
   GOOGLE_SHEETS_CREDENTIALS_JSON
   ```
4. Click inside the **Value** box.  
   > 💡 The Value box in Render is a large text area — it can hold many lines of text. You do not need to compress it to one line. Multi-line JSON is fine.
5. Press **Ctrl + V** (Windows/Linux) or **Cmd + V** (macOS) to paste.
6. **Verify the paste looks right:**
   - The first character should be `{`
   - The last character should be `}`
   - You should be able to see lines like `"type": "service_account"` and `"client_email": "..."` in the middle
   - The value box will show many lines — that is normal and expected
7. Click the **✓ tick** (or press **Enter**) to save the variable.

---

**Common mistakes for 4-C**

| Mistake | How to spot it | Fix |
|---|---|---|
| Copied only part of the file | Value starts or ends mid-sentence | Re-open the file, Ctrl+A, Ctrl+C, paste again |
| Copied the file *name* instead of the file *contents* | Value is just `apartment-mgmt-abc123.json` | Open the file in a text editor first, then copy the text inside |
| File was not yet downloaded | No `.json` file in Downloads | Follow the "Can't find it?" tip above to download a new key |
| Pasted into the Key box instead of the Value box | The Key box shows `{` | Click the Value box and paste there instead |

---

> ✅ **4-C is done!**  
> You have added 2 of 5 credentials (`SECRET_KEY` + `GOOGLE_SHEETS_CREDENTIALS_JSON`).  
> Three more to go — all much simpler (just copy a short ID from a URL):
> - **4-D** → `GOOGLE_SHEET_CONTRIBUTIONS_ID`
> - **4-E** → `GOOGLE_SHEET_BALANCE_ID`
> - **4-F** → `GOOGLE_SHEET_CRDR_ID`
>
> Continue reading below ↓

---

##### 4-D  Add `GOOGLE_SHEET_CONTRIBUTIONS_ID`

This is the Spreadsheet ID of your *Monthly Contribution* sheet.

**How to find the Spreadsheet ID:**

1. Open the *Monthly Contribution* Google Sheet in your browser.
2. Look at the URL in the address bar. It looks like:
   ```
   https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/edit
   ```
3. The **Spreadsheet ID** is the long string between `/d/` and `/edit`.  
   In the example above it is: `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms`

**Steps in Render:**

1. Click **Add Environment Variable**.
2. **Key:** `GOOGLE_SHEET_CONTRIBUTIONS_ID`
3. **Value:** paste the Spreadsheet ID you copied from the URL.
4. Click ✓ to save.

---

##### 4-E  Add `GOOGLE_SHEET_BALANCE_ID`

Same process as 4-D, but for the *Current Balance Sheet* spreadsheet.

1. Open the *Current Balance Sheet* Google Sheet.
2. Copy the Spreadsheet ID from the URL (the part between `/d/` and `/edit`).
3. In Render: **Add Environment Variable** → **Key:** `GOOGLE_SHEET_BALANCE_ID` → paste ID → ✓.

---

##### 4-F  Add `GOOGLE_SHEET_CRDR_ID`

Same process as 4-D, but for the *Balance Sheet Cr Dr* spreadsheet.

1. Open the *Balance Sheet Cr Dr* Google Sheet.
2. Copy the Spreadsheet ID from the URL.
3. In Render: **Add Environment Variable** → **Key:** `GOOGLE_SHEET_CRDR_ID` → paste ID → ✓.

---

##### 4-G  Final check

After adding all five variables, the Environment tab should show:

| Key | Status |
|---|---|
| `SECRET_KEY` | ✅ set |
| `GOOGLE_SHEETS_CREDENTIALS_JSON` | ✅ set |
| `GOOGLE_SHEET_CONTRIBUTIONS_ID` | ✅ set |
| `GOOGLE_SHEET_BALANCE_ID` | ✅ set |
| `GOOGLE_SHEET_CRDR_ID` | ✅ set |

Once all five are set, click **Save Changes** at the bottom of the page and move on to Step 5.

> 💡 **Tip:** If you made a mistake in a value, click the pencil ✏️ icon next to the variable to edit it.

---

#### Step 5 – Deploy

Click **Save Changes** (or **Manual Deploy → Deploy latest commit**).  
Render will build and start the app — watch the log at the bottom of the page.  
When you see `Listening on 0.0.0.0:10000` the app is live.

Your public URL is shown at the top of the Render service page, e.g.:  
**`https://apartment-mgmt.onrender.com`**

---

#### Auto-deploy on every push

Once set up, Render automatically re-deploys whenever you push to the `main` branch of your GitHub repo — no manual steps needed.

---

### Alternative: Railway.app (free tier — zero extra config)

Railway auto-detects Python and reads the `Procfile` that is already in this repo — no extra config files needed.

#### Step 1 – Push your repo to GitHub

Same as the Render guide above (skip if already done).

#### Step 2 – Create a free Railway account

Go to **https://railway.app** → **Start a New Project** → **Sign in with GitHub**.

#### Step 3 – Deploy from GitHub

1. Click **New Project** → **Deploy from GitHub repo**.
2. Choose your `Apartment` repo.
3. Railway detects Python automatically and starts building.  
   The start command (`gunicorn app:app`) is read from the `Procfile`.

#### Step 4 – Add the same environment variables

In the Railway project dashboard click the **service card** → **Variables** tab → **New Variable**.  
Add the identical five variables you would set on Render:

| Variable name | Value |
|---|---|
| `GOOGLE_SHEETS_CREDENTIALS_JSON` | Entire contents of your service-account JSON file |
| `GOOGLE_SHEET_CONTRIBUTIONS_ID` | Spreadsheet ID of your Contributions sheet |
| `GOOGLE_SHEET_BALANCE_ID` | Spreadsheet ID of your Balance sheet |
| `GOOGLE_SHEET_CRDR_ID` | Spreadsheet ID of your Cr/Dr sheet |
| `SECRET_KEY` | Any long random string (e.g. `openssl rand -hex 32`) |

After saving the variables Railway re-deploys automatically.  
Your public URL looks like `https://apartment-mgmt-production.up.railway.app`.

> ⚠️ Railway free tier gives **$5 of credit per month** (roughly enough to run one small always-on app).  
> The service does **not** sleep between requests (unlike Render free tier), so your pages load instantly.

---

### Alternative: Fly.io (always-on free tier — Docker-based)

Fly.io provides three always-on shared VMs for free (256 MB RAM each) with no credit-card required.  
This repo includes a `Dockerfile` and `fly.toml` that are ready to use.

#### Prerequisites

Install the `flyctl` command-line tool:

```bash
# macOS / Linux
curl -L https://fly.io/install.sh | sh

# Windows (PowerShell)
iwr https://fly.io/install.ps1 -useb | iex
```

#### Step 1 – Sign up and log in

```bash
fly auth signup   # or: fly auth login
```

#### Step 2 – Launch the app (first time only)

Run this from inside the project folder:

```bash
fly launch --no-deploy --name apartment-mgmt
```

Fly reads `fly.toml` automatically.  
When asked *"Would you like to copy its configuration to the new app?"* answer **Yes**.

#### Step 3 – Set secrets (environment variables)

```bash
fly secrets set \
  SECRET_KEY="paste-a-long-random-string-here" \
  GOOGLE_SHEETS_CREDENTIALS_JSON='{"type":"service_account",...}' \
  GOOGLE_SHEET_CONTRIBUTIONS_ID="your-sheet-id" \
  GOOGLE_SHEET_BALANCE_ID="your-sheet-id" \
  GOOGLE_SHEET_CRDR_ID="your-sheet-id"
```

#### Step 4 – Deploy

```bash
fly deploy
```

Fly builds the Docker image, pushes it, and starts your app.  
Your public URL looks like `https://apartment-mgmt.fly.dev`.

#### Auto-deploy on every push (optional)

Add this GitHub Actions workflow to auto-deploy on every push to `main`:

1. Get a deploy token: `fly tokens create deploy -x 999999h`
2. Add it as a GitHub secret named `FLY_API_TOKEN`.
3. Add a workflow file `.github/workflows/fly-deploy.yml`:

```yaml
name: Deploy to Fly.io
on:
  push:
    branches: [main]
jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: superfly/flyctl-actions/setup-flyctl@master
      - run: fly deploy --remote-only
        env:
          FLY_API_TOKEN: ${{ secrets.FLY_API_TOKEN }}
```

---

### Quick comparison of free platforms

| Platform | Cold start | Free tier limit | Auto-deploy | Needs extra config |
|---|---|---|---|---|
| **Render** | ~30 s (sleeps after 15 min) | 750 h/month | ✅ GitHub | `render.yaml` |
| **Railway** | Instant (no sleep) | $5 credit/month | ✅ GitHub | none (reads `Procfile`) |
| **Fly.io** | Instant (no sleep) | 3 shared VMs | ✅ GitHub Actions | `fly.toml` + `Dockerfile` |

All three options serve the same app — choose whichever feels easiest.

---

## ⚙️ Optional Configuration

| Environment Variable | Default | Description |
|---|---|---|
| `PORT` | `5000` | Port to run the server on |
| `SECRET_KEY` | `apartment-mgmt-secret-2025` | Flask session secret (change for production) |
| `FLASK_DEBUG` | `false` | Set to `true` for debug mode (development only) |

To change the port on Linux/macOS:
```bash
PORT=8080 python app.py
```

---

## ❓ Troubleshooting

| Problem | Fix |
|---|---|
| `python: command not found` | Use `python3` instead of `python`, or reinstall Python with "Add to PATH" ticked |
| `pip: command not found` | Run `python -m pip install -r requirements.txt` instead — see Step 3 above for all alternatives |
| `pip3: command not found` | Run `python3 -m pip install -r requirements.txt` |
| `No module named pip` | Run `python -m ensurepip --upgrade` then retry the pip install |
| `requirements.txt: No such file or directory` | Your terminal is in the wrong folder — see Step 1 above to open a terminal inside the project folder |
| `Address already in use` | Another app is using port 5000. Run `PORT=5001 python app.py` |
| Browser shows "This site can't be reached" | Make sure the terminal is still running (server was not stopped) |
| Forgot password | Ask admin to reset it via the **Apartments** page → **Reset** button |

---

## 📖 Rule: Keep this README Updated

Whenever a new feature, script, configuration, database change, or architectural evolution is introduced into the codebase, the developer or AI assistant MUST immediately update this `README.md` to document the change. This ensures the README acts as an accurate, living, and complete reference of the project at all times.

