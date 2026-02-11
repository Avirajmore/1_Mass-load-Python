import pandas as pd
import pyperclip
import time
import os

# ---------- CONFIG ----------
DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")
CSV_FILE_PATH = os.path.expanduser("~/Downloads/Hashi oppty.csv")
COLUMN_NAME = "ACCOUNTID"
NEW_FILE_NAME = os.path.expanduser("~/Downloads/Account export.csv")
WAIT_TIME = 10                       # seconds
# ----------------------------

# Read CSV
df = pd.read_csv(CSV_FILE_PATH)

# Extract unique ACCOUNTID values
unique_ids = (
    df[COLUMN_NAME]
    .dropna()
    .astype(str)
    .unique()
)

# Format for SQL IN clause
formatted_ids = ",".join([f"'{i}'" for i in unique_ids])

# Build query
query = f"""
SELECT AccountNumber, Id
FROM Account
WHERE AccountNumber IN ({formatted_ids})
"""

# Copy query to clipboard
pyperclip.copy(query)

print("✅ Query copied to clipboard:\n")
print(query)

# ---------------- WAIT ----------------
print(f"\n⏳ Waiting for {WAIT_TIME} seconds...")
time.sleep(WAIT_TIME)

# ---------------- FILE RENAME ----------------
print("\n🔍 Looking for bulkQuery_result_ CSV file...")

matching_files = [
    f for f in os.listdir(DOWNLOAD_FOLDER)
    if f.lower().endswith(".csv") and "bulkquery_result_" in f.lower()
]

if not matching_files:
    print("❌ No matching bulkQuery_result_ CSV file found.")
else:
    # Pick latest modified file
    latest_file = max(
        matching_files,
        key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_FOLDER, f))
    )

    old_path = os.path.join(DOWNLOAD_FOLDER, latest_file)
    new_path = os.path.join(DOWNLOAD_FOLDER, NEW_FILE_NAME)

    os.rename(old_path, new_path)

    print(f"✅ Renamed file:\n{latest_file} → {NEW_FILE_NAME}")

import pandas as pd

# ---------- CONFIG ----------
HASHI_FILE = os.path.expanduser("~/Downloads/Hashi oppty.csv")
ACCOUNT_EXPORT_FILE = os.path.expanduser("~/Downloads/Account export.csv")
OUTPUT_FILE = os.path.expanduser("~/Downloads/Accounts to import.xlsx")

HASHI_COLUMN = "ACCOUNTID"
ACCOUNT_COLUMN = "AccountNumber"
# ----------------------------

# Read CSV files
hashi_df = pd.read_csv(HASHI_FILE)
account_df = pd.read_csv(ACCOUNT_EXPORT_FILE)

# Standardize case and clean data
hashi_ids = (
    hashi_df[HASHI_COLUMN]
    .dropna()
    .astype(str)
    .str.strip()
    .str.upper()
    .unique()
)

account_numbers = (
    account_df[ACCOUNT_COLUMN]
    .dropna()
    .astype(str)
    .str.strip()
    .str.upper()
    .unique()
)

# Find ACCOUNTIDs not present in Account export
missing_accounts = sorted(set(hashi_ids) - set(account_numbers))

# Create output DataFrame
output_df = pd.DataFrame(
    missing_accounts,
    columns=[HASHI_COLUMN]
)

# Write to Excel
output_df.to_excel(OUTPUT_FILE, index=False)

print("✅ Comparison completed")
print(f"📄 Missing accounts saved to: {OUTPUT_FILE}")
print(f"🔢 Total missing accounts: {len(missing_accounts)}")
