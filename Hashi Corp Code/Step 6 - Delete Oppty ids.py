import os
import pandas as pd
import pyperclip
from tkinter import filedialog, Tk

# ---------- PRODUCT ----------
DOWNLOAD_FOLDER = os.path.expanduser("~/Downloads")
product_df = pd.read_csv(os.path.expanduser("~/Downloads/product_Record_Mismatch.csv"))

product_delete = (
    product_df[product_df["_merge"] == "Not_in_ISCED"][["Id"]]
    .rename(columns={"Id": "Delete Product"})
)

product_delete.to_csv(os.path.expanduser("~/Downloads/PreDelete_Product.csv"), index=False)

print("✅ Delete_Product.csv created")


# ---------- OPPORTUNITY ----------
oppty_df = pd.read_csv(os.path.expanduser("~/Downloads/oppty_Record_Mismatch.csv"))

oppty_delete = (
    oppty_df[oppty_df["_merge"] == "Not_in_ISCED"][["Id"]]
    .rename(columns={"Id": "Delete Oppty"})
)

oppty_delete.to_csv(os.path.expanduser("~/Downloads/PreDelete_Oppty.csv"), index=False)

print("✅ Delete_Oppty.csv created")



oppty_file_path = os.path.expanduser("~/Downloads/PreDelete_Oppty.csv")

product_file_path = os.path.expanduser("~/Downloads/PreDelete_Product.csv")

# Read CSV
df = pd.read_csv(oppty_file_path)

# Column name
column_name = "Delete Oppty"

# Convert values to quoted string list
values = df[column_name].dropna().astype(str)
formatted_values = ",".join(f"'{v}'" for v in values)

# Final query
query = f"""
SELECT Id
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
    query = f"select Id from OpportunityLineitem where Lineitem_Legacy_Id__c in ({formatted_values})"

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
