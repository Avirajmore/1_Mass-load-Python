'''
Use the below code to prepare the contact file from already Prepared Copy file

Steps:- 

1) You will need to select a Copy file which is already processed, not the raw file
2) It will add exsisiting column to the sheet which will make sure if the give oppties are present in main oppty file or not!
3) Later you create a final csv load file 
'''



import os
import sys
import tkinter as tk
from tkinter import filedialog

print("\n🔍 Step 1: Selecting a file...")
# Set the directory to search for Excel files
directory = "/Users/avirajmore/Downloads"

# Create a hidden root window (used for file dialog)
root = tk.Tk()
root.withdraw()

# Ask the user to select an Excel file
file_path = filedialog.askopenfilename(
    initialdir=directory,
    title="📄 Select an Excel file",
    filetypes=[("Excel files", "*.xlsx")]
)

# Print the selected file path
if file_path:
    filename = str(file_path.split('/')[-1])
    print(f"\n    ✅ File selected: '{filename}'.")
else:
    print("\n    ❌ No file selected. Exiting the program. ❌")
    sys.exit()

print("\n🔍 Step 2: Checking if the file exists...")

if os.path.exists(file_path):
    filename = str(file_path.split('/')[-1])
    print(f"\n    ✅ File '{filename}' exists at the specified path. ✅")
else:
    print("\n    ❌ Error: File does not exist or the path is invalid. ❌\n")
    sys.exit()  # Stops further execution of the program

# ===================================================================================
# ===================================================================================

import pandas as pd
import sys

print("\n🔍 Step 3: Verifying opportunities in the 'Opportunity' sheet...")

opportunity_sheet_name = 'Opportunity'
contact_sheet_name = 'Contact Roles'

try:
    # Load the sheets into DataFrames
    all_sheets = pd.read_excel(file_path, sheet_name=None)  # Load all sheets into a dictionary
    sheet_names = [sheet.lower() for sheet in all_sheets.keys()]  # Convert sheet names to lowercase

    # Check if the required sheets exist (case-insensitive)
    if opportunity_sheet_name.lower() not in sheet_names:
        print(f"\n    ❌ Sheet '{opportunity_sheet_name}' not found. ❌")
        sys.exit()
    if contact_sheet_name.lower() not in sheet_names:
        print(f"\n    ❌ Sheet '{contact_sheet_name}' not found. ❌")
        sys.exit()

    # Load the relevant sheets into DataFrames (case-insensitive)
    opportunity_df = all_sheets[list(all_sheets.keys())[sheet_names.index(opportunity_sheet_name.lower())]]
    contact_df = all_sheets[list(all_sheets.keys())[sheet_names.index(contact_sheet_name.lower())]]

    # Validate the required columns (case-insensitive)
    opportunity_columns = [col.lower() for col in opportunity_df.columns]
    product_columns = [col.lower() for col in contact_df.columns]

    if 'opportunity_legacy_id__c'.lower() not in opportunity_columns:
        print(f"\n    ❌ Column 'opportunity_legacy_id__c' not found in the '{opportunity_sheet_name}' sheet. ❌")
        sys.exit()
    elif 'opportunityid'.lower() not in product_columns:
        print(f"\n    ❌ Column 'opportunityid' not found in the '{contact_sheet_name}' sheet. ❌")
        sys.exit()

    # Perform the comparison
    contact_df['Existing'] = contact_df['opportunityid'].isin(opportunity_df['opportunity_legacy_id__c'])

    # Calculate the number of false values
    false_count = (~contact_df['Existing']).sum()

    # Save the updated data back to the Excel file
    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
        contact_df.to_excel(writer, sheet_name=contact_sheet_name, index=False)

    # Success message with false count
    print(f"\n    ✅ Verification completed. 'Existing' column has been added to the '{contact_sheet_name}' sheet. ✅")
    print(f"\n    ❗️ Number of False values in 'Existing' column: {false_count}")

except FileNotFoundError:
    # Handle file not found
    print(f"\n    ❌ Error: File not found. ❌")
    sys.exit()
except Exception as e:
    # Handle any unexpected errors
    print(f"\n    ❌ Error: An unexpected error occurred. Details: {e} ❌")
    sys.exit()

# =================================

import pandas as pd

# Define the file path
# file_path = '/Users/avirajmore/Downloads/Mass_load_template_CHEVA - IAPP 2024 copy.xlsx'

# Read the "Contact Roles" sheet from the Excel file
df = pd.read_excel(file_path, sheet_name='Contact Roles')

# Print out the column names to confirm the correct name
# print(df.columns)

# Assuming the correct column name is found, apply the transformation
# Adjust 'contactid' to match the correct column name from the printout
if 'contactid' in df.columns:
    df['contactid'] = df['contactid'].apply(lambda x: str(int(x)))
else:
    print("Column 'contactid' not found!")

# Use ExcelWriter to write the changes back to the same file, replacing the existing "Contact Roles" sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Contact Roles', index=False)
