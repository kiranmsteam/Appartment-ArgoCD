"""
find_apartment_payments.py

A command-line script to check the payment history and find how many months of
payments are done for a given apartment flat.

Usage:
  python find_apartment_payments.py KG4
"""

import os
import sys
import argparse

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import google_sheets as gs

def main():
    parser = argparse.ArgumentParser(description="Find payment history for a specific apartment flat.")
    parser.add_argument("apartment", help="Apartment number (e.g., KG4, VG3, RF2)")
    args = parser.parse_args()

    apt_no = args.apartment.strip().upper()

    # Load environment variables from .env
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("export "):
                    line = line[7:]
                parts = line.strip().split("=", 1)
                if len(parts) == 2:
                    key, val = parts
                    val = val.strip('"').strip("'")
                    os.environ[key.strip()] = val.strip()

    if not gs.is_configured():
        print("Error: Google Sheets API is not configured. Please check your .env file.")
        sys.exit(1)

    print(f"Connecting to Google Sheets and fetching payments for flat {apt_no} ...")
    try:
        history = gs.get_flat_payment_history(apt_no)
    except Exception as e:
        print(f"Error fetching history: {e}")
        sys.exit(1)

    if not history:
        print(f"\nNo payment history found for flat {apt_no}.")
        return

    # Filter credits (payments made by the flat owner)
    payments = [t for t in history if t.get('credit') and float(t['credit']) > 0]

    if not payments:
        print(f"\nNo payments found for flat {apt_no}.")
        return

    print("\n" + "=" * 70)
    print(f"  Payment History for Flat: {apt_no}")
    print("=" * 70)
    print(f"{'Date':<12} | {'Paid By (Name)':<25} | {'Amount':<14} | {'Month Tab':<12}")
    print("-" * 70)

    total_paid = 0.0
    paid_months = set()

    # Sort payments chronologically by parseable month date or as-is
    # Let's sort using reverse chronological order (newest first)
    for p in sorted(payments, key=lambda x: x.get('date', ''), reverse=True):
        date_str = p.get('date', '').strip()
        name = p.get('name', 'Resident').strip()
        amount = float(p.get('credit', 0.0))
        month_ref = p.get('month', '').strip()

        total_paid += amount
        if month_ref:
            paid_months.add(month_ref)

        print(f"{date_str:<12} | {name:<25} | Rs. {amount:<9,.2f} | {month_ref:<12}")

    print("-" * 70)
    print(f"Total Amount Paid:   Rs. {total_paid:,.2f}")
    print(f"Total Payments:      {len(payments)}")
    print(f"Unique Months Paid:  {len(paid_months)} month(s)")
    print(f"List of Paid Months: {', '.join(sorted(list(paid_months)))}")
    print("=" * 70)

if __name__ == "__main__":
    main()
