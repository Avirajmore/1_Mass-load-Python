import glob
import os
import pandas as pd
import pyperclip
from tkinter import filedialog, Tk
import shutil
# ---------- PRODUCT ----------
DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")
product_df = pd.read_csv(os.path.expanduser("~/Downloads/Hashi Load/Duplicate Files/product_Record_Mismatch.csv"))

product_delete = (
    product_df[product_df["_merge"] == "Not_in_ISCED"][["Id"]]
    .rename(columns={"Id": "Delete Product"})
)

product_delete.to_csv(os.path.expanduser("~/Downloads/Hashi Load/Duplicate Files/PreDelete_Product.csv"), index=False)

print("✅ Delete_Product.csv created")


# ---------- OPPORTUNITY ----------
oppty_df = pd.read_csv(os.path.expanduser("~/Downloads/Hashi Load/Duplicate Files/oppty_Record_Mismatch.csv"))

oppty_delete = (
    oppty_df[oppty_df["_merge"] == "Not_in_ISCED"][["Id"]]
    .rename(columns={"Id": "Delete Oppty"})
)

oppty_delete.to_csv(os.path.expanduser("~/Downloads/Hashi Load/Duplicate Files/PreDelete_Oppty.csv"), index=False)

print("✅ Delete_Oppty.csv created")



oppty_file_path = os.path.expanduser("~/Downloads/Hashi Load/Duplicate Files/PreDelete_Oppty.csv")

product_file_path = os.path.expanduser("~/Downloads/Hashi Load/Duplicate Files/PreDelete_Product.csv")

# Read CSV
df = pd.read_csv(oppty_file_path)

# Column name
column_name = "Delete Oppty"

# Convert values to quoted string list
values = df[column_name].dropna().astype(str)
formatted_values = ",".join(f"'{v}'" for v in values)

# Final query
query = f"""
SELECT Id,Source_ID__c
FROM Opportunity
WHERE Source_ID__c IN ({formatted_values})
"""

# Copy to clipboard
pyperclip.copy(query.strip())

print("✅ Oppty query copied to clipboard:\n")


choice = input("Do you want to Proceed further? (y/n)")

if choice == 'y':
    print("\n🔍 Looking for bulkQuery_result_ CSV file...")
    DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")                # change if needed (e.g. Downloads)
    NEW_FILE_NAME = "DELETE OPPTY.csv"
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


    # Read CSV
    df = pd.read_csv(product_file_path)

    # Column name
    column_name = "Delete Product"

    # Convert values to quoted string list
    values = df[column_name].dropna().astype(str)
    formatted_values = ",".join(f"'{v}'" for v in values)

    # Final query
    query = f"select Id,Lineitem_Legacy_Id__c from OpportunityLineitem where Lineitem_Legacy_Id__c in ({formatted_values})"

    # Copy to clipboard
    pyperclip.copy(query.strip())

    print("✅ Product query copied to clipboard:\n")

    choice = input("Do you want to Proceed further? (y/n)")

    if choice == 'y':
        print("\n🔍 Looking for bulkQuery_result_ CSV file...")
        DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")                # change if needed (e.g. Downloads)
        NEW_FILE_NAME = "DELETE PRODUCT.csv"
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

shutil.move(os.path.expanduser("~/Downloads/DELETE OPPTY.csv"), os.path.expanduser("~/Downloads/Hashi Load/Main Files/DELETE OPPTY.csv"))
shutil.move(os.path.expanduser("~/Downloads/DELETE PRODUCT.csv"), os.path.expanduser("~/Downloads/Hashi Load/Main Files/DELETE PRODUCT.csv"))


summary_folder = os.path.expanduser("~/Downloads/Hashi Load")   # change if needed

# CSV file paths (UPDATE THESE)
oppty_file = os.path.expanduser("~/Downloads/Hashi Load/Main Files/DELETE OPPTY.csv")
product_file = os.path.expanduser("~/Downloads/Hashi Load/Main Files/DELETE PRODUCT.csv")

summary_files = glob.glob(os.path.join(summary_folder, "SUMMARY FILE - HASHI PROD (*.xlsx)"))

if not summary_files:
    print("❌ No summary file found!")
    exit()

latest_summary = max(summary_files, key=os.path.getmtime)
print(f"✅ Using summary file: {latest_summary}")

# ============================
# Step 3: Read CSV files
# ============================
df_oppty = pd.read_csv(oppty_file)
df_product = pd.read_csv(product_file)

# ============================
# Step 4: Write to Excel
# ============================
with pd.ExcelWriter(latest_summary, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    
    df_oppty.to_excel(writer, sheet_name="DELETE OPPTY", index=False)
    df_product.to_excel(writer, sheet_name="DELETE PRODUCT", index=False)

print("✅ Delete Data successfully copied to summary file!")

import os
from datetime import datetime

def get_day_suffix(day):
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")

# Get today's date
now = datetime.now()
day = now.day
month = now.strftime("%B")

suffix = get_day_suffix(day)

# Final folder name
new_folder_name = f"Hashi Load ({day}{suffix} {month})"

# Path where your folder exists
base_path = os.path.expanduser("~/Downloads")

# Old folder name (change this)
old_folder_name = "Hashi Load"

old_path = os.path.join(base_path, old_folder_name)
new_path = os.path.join(base_path, new_folder_name)

# Rename folder
if os.path.exists(old_path):
    os.rename(old_path, new_path)
    print(f"Folder renamed to: {new_folder_name}")
else:
    print("Folder not found!")