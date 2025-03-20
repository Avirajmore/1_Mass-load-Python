'''
Use the below code to prepare the contact file from already Prepared Copy file

Steps:- 

1) You will need to select a Copy file which is already processed, not the raw file
2) It will add exsisiting column to the sheet which will make sure if the give oppties are present in main oppty file or not!
3) Later you create a final csv load file 
'''

import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils.exceptions import SheetTitleException



print("\n🔍 Step 1: Selecting a Copy file...")
# Set the directory to search for Excel files
directory = "/Users/avirajmore/Downloads"

# Create a hidden root window (used for file dialog)
root = tk.Tk()
root.withdraw()

# Ask the user to select an Excel file
file_path = filedialog.askopenfilename(
    initialdir=directory,
    title="📄 Select an Excel file",
    # filetypes=[("Excel files", "*.xlsx")]
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


print("\n🔍Step 4: Creating Final Contact file")


predefined_columns_Reportingcode = ['existing']

# File paths
# file_path = 'your_file_path_path.xlsx'  # Uncomment and set your file path
sheet_name = 'Contact Roles'
name_of_file= "/"+filename.strip('_Copy.xlsx')
destination_folder = "/".join(file_path.split('/')[0:9]) + "/Final iteration files/"+name_of_file
output_file = destination_folder + name_of_file+'_Contact.csv' # Path for the processed CSV
removed_rows_file = destination_folder +'/Removed Rows/Removed_Rows - Contact.csv' # Path for removed rows CSV
# Initialize variables
deleted_columns = []
rows_dropped = 0

try:
    # Step 1: Read the Excel file into a DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Step 2: Convert all column names to lowercase
    df.columns = df.columns.str.lower()

    # Step 3: Check if 'existing' column exists
    if 'existing' not in df.columns:
        raise ValueError(f"\n    ❌ Column 'existing' not found in the DataFrame from sheet '{sheet_name}'. Please check your input data.")

    # Step 4: Remove rows where "existing" == False and save them to another file
    removed_rows = df[df['existing'].astype(str).str.lower() != 'true'].copy()
    df = df[df['existing'].astype(str).str.lower() == 'true']
    rows_dropped = len(removed_rows)

    if rows_dropped > 0:
        removed_rows['Reason'] = "Opportunity Missing From Main sheet"
        removed_rows.drop(columns=[col for col in predefined_columns_Reportingcode if col in removed_rows.columns], inplace=True)
        removed_rows.to_csv(removed_rows_file, index=False)


    # Step 5: Remove predefined columns from the main data
    predefined_to_delete = [col for col in predefined_columns_Reportingcode if col in df.columns]
    if predefined_to_delete:
        df.drop(columns=predefined_to_delete, inplace=True)
        deleted_columns.extend(predefined_to_delete)

    # Step 6: Create UI for selecting additional columns to delete
    root = Tk()
    root.title("Select Columns to Delete")

    # Set window size
    root.geometry("500x600")
    root.resizable(False, False)

    # Scrollable UI
    canvas = Canvas(root)
    scrollbar = Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    frame = Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")

    # Checkbox setup
    checkboxes = {}
    for column in df.columns:
        var = IntVar()
        checkboxes[column] = var
        checkbutton = Checkbutton(frame, text=column, variable=var, font=('Helvetica', 12), anchor="w", padx=10)
        checkbutton.pack(anchor="w", pady=5)

    # Submit button
    submit_button = Button(frame, text="Submit", command=root.quit, font=('Helvetica', 12, 'bold'), relief='flat', padx=20, pady=10)
    submit_button.pack(pady=20)

    # Run the UI
    frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
    root.mainloop()
    root.destroy()

    # Step 7: Process user-selected columns for deletion
    user_selected_columns = [col for col, var in checkboxes.items() if var.get() == 1]
    if user_selected_columns:
        df.drop(columns=user_selected_columns, inplace=True)
        deleted_columns.extend(user_selected_columns)
        print("\n    ✅ Additional columns deleted:")
        for col in user_selected_columns:
            print(f"\n        🔸 {col}")
    else:
        print("\n    ✅ No additional columns selected for deletion.")

    # Step 8: Clean up the DataFrame
    df.dropna(axis=0, how='all', inplace=True)  # Remove rows with all blank values
    df.dropna(axis=1, how='all', inplace=True)  # Remove columns with all blank values
    df.drop_duplicates(inplace=True)  # Remove duplicate rows

    # Step 9: Save the processed DataFrame
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    df.to_csv(output_file, index=False)

    # Summary of deletions
    print(f"\n    ❗️ Total rows removed where 'existing' == False: {rows_dropped}")

    # Final summary messages
    print("\n    ✅ Processed data saved to:")
    print(f"\n        📂 {"/".join(output_file.split("/")[-5:])}")

    if rows_dropped > 0:
        print("\n    ✅ Removed rows saved to:")
        print(f"\n        📂 {"/".join(removed_rows_file.split("/")[-5:])}")

except ValueError as ve:
    print(f"\n    ❌ ValueError: {ve}")
except Exception as e:
    print(f"\n    ❌ An error occurred: {e}")
