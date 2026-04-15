import os
import pyperclip
import pandas as pd
from datetime import datetime

def show_title(title):

    line_width = 100
    line = "=" * line_width
    print(f"\n{line}")
    print(title.center(line_width))
    print(f"{line}\n")

# Display the title for the folder creation and file movement process
title = "📂 CONFLUENT LOAD 📂"
show_title(title)

# ---------- CONFIG ----------
DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")
CSV_FILE_PATH = os.path.expanduser("~/Downloads/Confluent oppty.csv")
COLUMN_NAME = "ACCOUNTID"
NEW_FILE_NAME = os.path.expanduser("~/Downloads/Account export.csv")
# ----------------------------

def get_day_suffix(day):
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


# Folder name
folder_name = "Confluent Load"

# Full path
folder_path = os.path.join(DOWNLOAD_FOLDER, folder_name)

# Create main folder
os.makedirs(folder_path, exist_ok=True)

# ✅ Create subfolders
main_files_folder = os.path.join(folder_path, "Main Files")
unimportant_folder = os.path.join(folder_path, "Unimportant")

os.makedirs(main_files_folder, exist_ok=True)
os.makedirs(unimportant_folder, exist_ok=True)


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

# ---------------- WAIT ----------------
choice = input("\n✅ Account Query is copied, Paste this Query in the WorkBench and download the csv file. \n\nOnce done, type 'Y' !")

if choice.lower() == 'y':

    # ---------------- FILE RENAME ----------------

    matching_files = [
        f for f in os.listdir(DOWNLOAD_FOLDER)
        if f.lower().endswith(".csv") and "bulkquery_result_" in f.lower()
    ]

    if not matching_files:
        print("\n❌ No matching bulkQuery_result_ CSV file found.\n")
    else:
        # Pick latest modified file
        latest_file = max(
            matching_files,
            key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_FOLDER, f))
        )

        old_path = os.path.join(DOWNLOAD_FOLDER, latest_file)
        new_path = os.path.join(DOWNLOAD_FOLDER, NEW_FILE_NAME)

        os.rename(old_path, new_path)

        # ---------- CONFIG ----------
        HASHI_FILE = os.path.expanduser("~/Downloads/Confluent oppty.csv")
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

        print(f"\n📄 Missing accounts saved to: Accounts to import.xlsx")
        print(f"\n🔢 Total missing accounts: {len(missing_accounts)}\n")
else:
    print("\nSkipped")

def move_file(file_name, source_dir, destination_dir):
    import os
    import shutil

    src = os.path.join(source_dir, file_name)
    dst = os.path.join(destination_dir, file_name)

    if os.path.exists(src):
        shutil.move(src, dst)
    else:
        print(f"File not found: {file_name} (skipping)")

# Move files
for file in ["confluent oppty.csv"]:
    move_file(file, DOWNLOAD_FOLDER, main_files_folder)

for file in ["Account export.csv", "Accounts to import.xlsx"]:
    move_file(file, DOWNLOAD_FOLDER, unimportant_folder)

print(f"\n🚨 CHECK IF YOU NEED TO IMPORT ACCOUNTS, AFTER THAT LOAD THE OPPTY FILE")

title = "Step 1 Done"
show_title(title)