# Hitech Citadel Phase 2 — AI Assistant Guidelines (GEMINI.md)

Welcome! This file provides standard operational guidelines and project-specific rules that you MUST follow when working on the Hitech Citadel Phase 2 Flat Owners Association codebase.

---

## 📋 Rule 1: Living Documentation (README.md)
Whenever you modify existing features, add new scripts, introduce configuration options, or make architectural changes, you **MUST** update [README.md](file:///c:/Users/mskir/Appartment/Apartment/README.md) to document these changes. The README must remain a complete, up-to-date, and accurate reference for this repository at all times.

---

## 🧮 Rule 2: Monthly Balance Sheet Generation
When generating monthly balance sheet worksheets in Google Sheets, follow these rules:
1. **Use the Unified Script**: Always use [generate_monthly_balance_sheet.py](file:///c:/Users/mskir/Appartment/Apartment/generate_monthly_balance_sheet.py) (e.g. `python generate_monthly_balance_sheet.py --month May --year 2026`). Do not create or reuse deprecated, single-month scripts.
2. **Sum Credits via Flat Filtering**: Ensure that credits (contributions) are summed strictly for apartment rows that belong to the 30 active flat units (listed in `_VALID_FLATS` in the generator script). This prevents summary/total cells at the bottom of the CR/DR worksheets from being double-counted.
3. **Flexible Cash Withdrawal & Bank Interest Parsing**:
   * Cash withdrawals must be detected by checking both the apartment category and description columns for `"WITHDRAWAL"` or the misspelled `"WITHRAWAL"`. When found, update the cheque description label to `"Cash Withdrawal"` and dynamically pass the amount to the Cash Withdrawal row of the Cash Expense section.
   * Bank interest credits must be parsed by searching for rows categorized under `"INTEREST"` in the apartment number column, and mapped directly to the Bank Interest row of the Collection section.

---

## 🔌 Rule 3: Web Dashboard Sheet Matching & Fallbacks
When interacting with worksheets in Google Sheets or local Excel files via [google_sheets.py](file:///c:/Users/mskir/Appartment/Apartment/google_sheets.py):
1. **Space-Collapsed Matching**: Always use space-collapsed, whitespace-insensitive matching (such as collapsing consecutive spaces into a single space) when searching for sheet titles. This ensures tabs like `April  2026` (double space) match seamlessly when searching for `April 2026`.
2. **Report Tab Fallback**: Maintain the summary detail fallback logic. If a report summary tab (such as `26-27 Report`) is missing or not yet updated, the backend should dynamically parse the income/expense totals from the detailed month tabs so the web dashboard selector remains fully populated and functional.
