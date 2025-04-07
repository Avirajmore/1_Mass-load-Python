
'''
This code removes rows from the final product file where opportunities failed to load due to duplication in ISC.

Steps:

1) Select the Final Product file to clean duplicate failed opportunities.
2) Select the Oppty Error file to identify failed duplicates.
3) The code finds failed Legacy_Oppty_IDs and removes corresponding rows in the Product file.
4) The removed rows are saved in a separate file for later use in the summary.

'''

# ===========================================

# ===========================================

import pandas as pd
import tkinter as tk
from tkinter import filedialog
import sys

# Output header
print("=" * 100)
print(" " * 33 + " REMOVING DUPLIATE OPPORTUNITY FROM EACH FILE ")
print("=" * 100)

# ===========================================
# Select the Files
# ===========================================

# Function to select a CSV file
def select_file(title):
    while True:
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(title=title, filetypes=[("CSV files", "*.csv")])

        if file_path:
            return file_path
        else:
            retry = input(f"\n    ❌ No file selected for {title}. Do you want to select a file? (yes/no): ").strip().lower()
            if retry == "no":
                print(f"\n       ❗️ Skipping {title}.")
                return None  # Skip the file selection

# Select product file
print("\n\n🔍 Step 1: Select the Final CSV files")
product_csv_path = select_file("Select the Product CSV file")
if not product_csv_path:
    print("\n    ❗️ No Product CSV file selected. Skipping this step.")
else:
    print(f"\n    ✅ File Selected: {product_csv_path.split('/')[-1]}")


# Select team file
team_csv_path = select_file("Select the Team CSV file")
if not team_csv_path:
    print("\n    ❗️ No Product CSV file selected. Skipping this step.")
else:
    print(f"\n    ✅ File Selected: {team_csv_path.split('/')[-1]}")

# Select tag file
tags_csv_path = select_file("Select the Tags CSV file")
if not tags_csv_path:
    print("\n    ❗️ No Product CSV file selected. Skipping this step.")
else:
    print(f"\n    ✅ File Selected: {tags_csv_path.split('/')[-1]}")

# Select reporting file
codes_csv_path = select_file("Select the Codes CSV file")
if not codes_csv_path:
    print("\n    ❗️ No Product CSV file selected. Skipping this step.")
else:
    print(f"\n    ✅ File Selected: {codes_csv_path.split('/')[-1]}")


# Select Oppty Error file
print("\n\n🔍 Step 2: Select the Oppty Error file")
oppty_error_path = select_file("Select the Opportunity Error CSV file")
print(f"\n    ✅ File Selected: {oppty_error_path.split('/')[-1]}")

# ===========================================
# Removed Rows file path
# ===========================================

# Define the path where removed rows should be saved
base_path = '/'.join(oppty_error_path.split('/')[:-2]) 

removed_rows_product = base_path + "/Removed Rows/Removed_Rows - Product (Duplicate Opportunity).csv"
removed_rows_team    = base_path + "/Removed Rows/Removed_Rows - Team member (Duplicate Opportunity).csv"
removed_rows_tags    = base_path + "/Removed Rows/Removed_Rows - Tags (Duplicate Opportunity).csv"
removed_rows_codes   = base_path + "/Removed Rows/Removed_Rows - Codes (Duplicate Opportunity).csv"

# ===========================================
# Removing duplicate Oppty Products and saving those rows in differnt file
# ===========================================

if not product_csv_path:

    print("\n    ❗️ Required files are missing. Skipping this step.")

else:
    print("\n\n🔍 Step 3: Removing duplicate Oppty Products")

    # Load the first CSV file
    df_oppty = pd.read_csv(oppty_error_path)

    # Check if the required columns exist
    if "ERRORS" not in df_oppty.columns or "Legacy_Opportunity_Split_Id__c" not in df_oppty.columns:
        print("\n    ❌ Missing required columns in Opportunity Error CSV file. Exiting.")
        sys.exit()

    # Identify rows where "ERRORS" column contains the specific substring
    error_message_substring = "ERROR: duplicate value found: Legacy_Opportunity_Split_Id__c duplicates value on record with id"
    matching_rows = df_oppty[df_oppty["ERRORS"].str.contains(error_message_substring, na=False, case=True)]

    # Extract values from "Legacy_Opportunity_Split_Id__c" column
    ids_to_remove = set(matching_rows["Legacy_Opportunity_Split_Id__c"])

    # Print matching values
    if ids_to_remove:
        print("\n    🔍 Matching values found in both files (Legacy_Opportunity_Split_Id__c):")

    else:
        print("\n    ❗️ No matching values found. No rows will be removed.")
        sys.exit()

    # Load the second CSV file
    df_product = pd.read_csv(product_csv_path)

    # Check if required column exists in the second file
    if "Legacy_Opportunity_Split_Id__c" not in df_product.columns:
        print("\n    ❌ Missing required column 'Legacy_Opportunity_Split_Id__c' in Product CSV file. Exiting.")
        sys.exit()


    # Filter rows that should be removed (create a copy to avoid SettingWithCopyWarning)
    removed_rows = df_product[df_product["Legacy_Opportunity_Split_Id__c"].isin(ids_to_remove)].copy()

    # Add the "Reason" column
    removed_rows["Reason"] = "Duplicate Opportunity"


    # Add the "Reason" column to the removed rows and fill it with "Duplicate Opportunity"
    if not removed_rows.empty:
        removed_rows.loc[:, "Reason"] = "Duplicate Opportunity"
        removed_rows.to_csv(removed_rows_product, index=False)
        print("\n    ✅ Saved Removed rows with reason.")
    else:
        print("\n    ❗️ No matching rows found to remove.")

    # Keep only rows that are NOT in ids_to_remove
    df_product_filtered = df_product[~df_product["Legacy_Opportunity_Split_Id__c"].isin(ids_to_remove)]

    # Save the updated second CSV file (overwriting the original)
    df_product_filtered.to_csv(product_csv_path, index=False)
    print("\n    ✅ Updated product CSV file.")



# ===========================================
# Removing duplicate Oppty Team and saving those rows in differnt file
# ===========================================

if not team_csv_path:

    print("\n    ❗️ Required files are missing. Skipping this step.")

else:

    print("\n\n🔍 Step 4: Removing duplicate Oppty Team Member")

    # Load the first CSV file
    df_oppty = pd.read_csv(oppty_error_path)

    # Check if the required columns exist
    if "ERRORS" not in df_oppty.columns or "Legacy_Opportunity_Split_Id__c" not in df_oppty.columns:
        print("\n    ❌ Missing required columns in Opportunity Error CSV file. Exiting.")
        sys.exit()

    # Identify rows where "ERRORS" column contains the specific substring
    error_message_substring = "duplicate value found: Legacy_Opportunity_Split_Id__c duplicates value on record with id"
    matching_rows = df_oppty[df_oppty["ERRORS"].str.contains(error_message_substring, na=False, case=True)]

    # Extract values from "Legacy_Opportunity_Split_Id__c" column
    ids_to_remove = set(matching_rows["Legacy_Opportunity_Split_Id__c"])

    # Print matching values
    if ids_to_remove:
        print("\n    🔍 Matching values found in both files (OpportunityId):")

    else:
        print("\n    ❗️ No matching values found. No rows will be removed.")
        sys.exit()

    # Load the second CSV file
    df_team = pd.read_csv(team_csv_path)

    # Check if required column exists in the second file
    if "OpportunityId" not in df_team.columns:
        print("\n    ❌ Missing required column 'OpportunityId' in Team CSV file. Exiting.")
        sys.exit()

    # Filter rows that should be removed (create a copy to avoid SettingWithCopyWarning)
    removed_rows = df_team[df_team["OpportunityId"].isin(ids_to_remove)].copy()

    # Add the "Reason" column
    removed_rows["Reason"] = "Duplicate Opportunity"


    # Add the "Reason" column to the removed rows and fill it with "Duplicate Opportunity"
    if not removed_rows.empty:
        removed_rows.loc[:, "Reason"] = "Duplicate Opportunity"
        removed_rows.to_csv(removed_rows_team, index=False)
        print("\n    ✅ Saved Removed rows with reason.")
    else:
        print("\n    ❗️ No matching rows found to remove.")

    # Keep only rows that are NOT in ids_to_remove
    df_team_filtered = df_team[~df_team["OpportunityId"].isin(ids_to_remove)]

    # Save the updated second CSV file (overwriting the original)
    df_team_filtered.to_csv(team_csv_path, index=False)
    print("\n    ✅ Updated product CSV file.")

# ===========================================
# Removing duplicate Oppty tags and saving those rows in differnt file
# ===========================================


if not tags_csv_path:

    print("\n    ❗️ Required files are missing. Skipping this step.")

else:
    print("\n\n🔍 Step 5: Removing duplicate Oppty Tags")

    # Load the first CSV file
    df_oppty = pd.read_csv(oppty_error_path)

    # Check if the required columns exist
    if "ERRORS" not in df_oppty.columns or "Legacy_Opportunity_Split_Id__c" not in df_oppty.columns:
        print("\n    ❌ Missing required columns in Opportunity Error CSV file. Exiting.")
        sys.exit()

    # Identify rows where "ERROR" column contains the specific substring
    error_message_substring = "duplicate value found: Legacy_Opportunity_Split_Id__c duplicates value on record with id"
    matching_rows = df_oppty[df_oppty["ERRORS"].str.contains(error_message_substring, na=False, case=True)]

    # Extract values from "Legacy_Opportunity_Split_Id__c" column
    ids_to_remove = set(matching_rows["Legacy_Opportunity_Split_Id__c"])

    # Print matching values
    if ids_to_remove:
        print("\n    🔍 Matching values found in both files (opportunityid):")
    else:
        print("\n    ❗️ No matching values found. No rows will be removed.")
        sys.exit()

    # Load the second CSV file
    df_tags = pd.read_csv(tags_csv_path)

    # Check if required column exists in the second file
    if "opportunityid" not in df_tags.columns:
        print("\n    ❌ Missing required column 'opportunityid' in Tags CSV file. Exiting.")
        sys.exit()


    # Filter rows that should be removed (create a copy to avoid SettingWithCopyWarning)
    removed_rows = df_tags[df_tags["opportunityid"].isin(ids_to_remove)].copy()

    # Add the "Reason" column
    removed_rows["Reason"] = "Duplicate Opportunity"


    # Add the "Reason" column to the removed rows and fill it with "Duplicate Opportunity"
    if not removed_rows.empty:
        removed_rows.loc[:, "Reason"] = "Duplicate Opportunity"
        removed_rows.to_csv(removed_rows_tags, index=False)
        print("\n    ✅ Saved Removed rows with reason.")
    else:
        print("\n    ❗️ No matching rows found to remove.")

    # Keep only rows that are NOT in ids_to_remove
    df_tags_filtered = df_tags[~df_tags["opportunityid"].isin(ids_to_remove)]

    # Save the updated second CSV file (overwriting the original)
    df_tags_filtered.to_csv(tags_csv_path, index=False)
    print("\n    ✅ Updated product CSV file.")



# ===========================================
# Removing duplicate Oppty tags and saving those rows in differnt file
# ===========================================

print("\n\n🔍 Step 6: Removing duplicate Oppty Codes")

if not codes_csv_path:

    print("\n    ❗️ Required files are missing. Skipping this step.")

else:
    # Load the first CSV file
    df_oppty = pd.read_csv(oppty_error_path)

    # Check if the required columns exist
    if "ERRORS" not in df_oppty.columns or "Legacy_Opportunity_Split_Id__c" not in df_oppty.columns:
        print("\n    ❌ Missing required columns in Opportunity Error CSV file. Exiting.")
        sys.exit()

    # Identify rows where "ERROR" column contains the specific substring
    error_message_substring = "duplicate value found: Legacy_Opportunity_Split_Id__c duplicates value on record with id"
    matching_rows = df_oppty[df_oppty["ERRORS"].str.contains(error_message_substring, na=False, case=True)]

    # Extract values from "opportunityid" column
    ids_to_remove = set(matching_rows["Legacy_Opportunity_Split_Id__c"])

    # Print matching values
    if ids_to_remove:
        print("\n    🔍 Matching values found in both files (opportunityid):")

    else:
        print("\n    ❗️ No matching values found. No rows will be removed.")
        sys.exit()

    # Load the second CSV file
    df_codes = pd.read_csv(codes_csv_path)

    # Check if required column exists in the second file
    if "opportunityid" not in df_codes.columns:
        print("\n    ❌ Missing required column 'opportunityid' in Codes CSV file. Exiting.")
        sys.exit()


    # Filter rows that should be removed (create a copy to avoid SettingWithCopyWarning)
    removed_rows = df_codes[df_codes["opportunityid"].isin(ids_to_remove)].copy()

    # Add the "Reason" column
    removed_rows["Reason"] = "Duplicate Opportunity"


    # Add the "Reason" column to the removed rows and fill it with "Duplicate Opportunity"
    if not removed_rows.empty:
        removed_rows.loc[:, "Reason"] = "Duplicate Opportunity"
        removed_rows.to_csv(removed_rows_codes, index=False)
        print("\n    ✅ Saved Removed rows with reason.")
    else:
        print("\n    ❗️ No matching rows found to remove.")

    # Keep only rows that are NOT in ids_to_remove
    df_codes_filtered = df_codes[~df_codes["opportunityid"].isin(ids_to_remove)]

    # Save the updated second CSV file (overwriting the original)
    df_codes_filtered.to_csv(codes_csv_path, index=False)
    print("\n    ✅ Updated product CSV file.")

