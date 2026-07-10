import os
import shutil
from tkinter import Tk, filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tabulate import tabulate

# -------- CONFIG --------
FOLDER_PATH = os.path.expanduser("~/Documents/Office docs/Massload Files/2026")   # <-- change this
# BASE_DIR must match the BASE_DIR in 3_Main File.py — used for auto-detecting the latest iteration folder
BASE_DIR = os.path.expanduser("~/Documents/Office Docs/Massload Files/2026")
# ------------------------


# =========================================================================================================================================
#  Auto-detect helpers for Q5
# =========================================================================================================================================

def _find_latest_iteration_folder(base_dir):
    """
    Scan BASE_DIR for the most recently modified iteration folder.
    An iteration folder is any direct subfolder of BASE_DIR that itself
    contains a 'Final iteration files' subdirectory.

    Returns the full path of the most recently modified matching folder,
    or None if nothing is found.
    """
    if not os.path.isdir(base_dir):
        return None

    candidates = []
    for name in os.listdir(base_dir):
        full = os.path.join(base_dir, name)
        if os.path.isdir(full) and os.path.isdir(os.path.join(full, "Final iteration files")):
            candidates.append((os.path.getmtime(full), full))

    if not candidates:
        return None

    candidates.sort(reverse=True)   # most recently modified first
    return candidates[0][1]


def _find_summary_file(iteration_folder):
    """
    Walk every subfolder of 'Final iteration files' inside the given
    iteration folder and return the most recently modified *_summary file.xlsx.
    Returns None if none is found.
    """
    final_dir = os.path.join(iteration_folder, "Final iteration files")
    if not os.path.isdir(final_dir):
        return None

    matches = []
    for root, dirs, files in os.walk(final_dir):
        for f in files:
            if f.endswith("_summary file.xlsx"):
                full = os.path.join(root, f)
                matches.append((os.path.getmtime(full), full))

    if not matches:
        return None

    matches.sort(reverse=True)
    return matches[0][1]


def _auto_detect_paths():
    """
    Try to auto-detect the iteration folder and summary file.
    Shows the detected paths and asks the user to confirm.

    Returns summary_file_path (str) or None if the user declines or
    nothing is found.
    """
    print("\n\n🔍 Auto-detecting latest iteration folder...")

    iteration_folder = _find_latest_iteration_folder(BASE_DIR)
    if not iteration_folder:
        print(f"\n    ❗️ No iteration folder found in: {BASE_DIR}")
        print("       Falling back to manual selection.\n")
        return None

    summary_file = _find_summary_file(iteration_folder)
    if not summary_file:
        print(f"\n    ❗️ No summary file found inside: {iteration_folder}")
        print("       Falling back to manual selection.\n")
        return None

    # Show detected paths
    short_iter   = os.path.sep.join(iteration_folder.split(os.path.sep)[-3:])
    short_sum    = os.path.sep.join(summary_file.split(os.path.sep)[-4:])
    success_error_path = os.path.join(os.path.dirname(summary_file), "Success and error files")

    print(f"\n    📂 Detected iteration folder : …/{short_iter}")
    print(f"    📄 Detected summary file     : …/{short_sum}")
    print(f"    📁 Success & error files in  : …/Success and error files")

    while True:
        confirm = input("\n    👉 Is this the correct folder? (yes / no — 'no' opens manual picker): ").strip().lower()
        if confirm == "yes":
            return summary_file
        elif confirm == "no":
            print("\n    Opening manual file picker...")
            return None
        else:
            print("    ❗️ Please enter 'yes' or 'no'.")

# -------- RENAME RULES --------
RENAME_RULES = {
    "opportunity insert": {
        "success": "opptysuccess",
        "error": "opptyerror"
    },
    "opportunity product insert": {
        "success": "productsuccess",
        "error": "producterror"
    },
    "feed item insert": {
        "success": "feedsuccess",
        "error": "feederror"
    },
    "opportunity strategy insert":{
        "success":"tagssuccess",
        "error":"tagserror"
    },
    "opportunity code insert":{
        "success":"codessuccess",
        "error":"codeserror"
    },
    "opportunity tags insert":{
        "success":"tagssuccess",
        "error":"tagserror"
    },
    "team member insert":{
        "success":"teamsuccess",
        "error":"teamerror"
    }
}
# -------------------------------

print("🔍 Scanning folder for CSV files...\n")

for file in os.listdir(FOLDER_PATH):

    if not file.lower().endswith(".csv"):
        continue

    file_lower = file.lower()
    new_name = None

    # Find matching rule
    for pattern, outcomes in RENAME_RULES.items():
        if pattern in file_lower:
            for status, base_name in outcomes.items():
                if status in file_lower:
                    new_name = f"{base_name}.csv"
                    break
        if new_name:
            break

    if new_name:
        old_path = os.path.join(FOLDER_PATH, file)
        new_path = os.path.join(FOLDER_PATH, new_name)

        # Prevent overwrite
        counter = 1
        base, ext = os.path.splitext(new_name)
        while os.path.exists(new_path):
            new_path = os.path.join(FOLDER_PATH, f"{base}_{counter}{ext}")
            counter += 1

        os.rename(old_path, new_path)
        print(f"✅ Renamed: {file} → {os.path.basename(new_path)}")

    else:
        print(f"⚠️ Skipped (no rule matched): {file}")

print("\n🎯 Renaming completed!")

# =======================================================================================================
line_width = 100
line = "=" * line_width
title = "📝 Success And Error Files EXECUTION 📝"
print(f"\n{line}")
print(title.center(line_width))
print(f"{line}\n")
# =======================================================================================================

# print("\n🔔 Is this the FINAL summary file for this batch?")
# print("1) No (Intermediate run)")
# print("2) Yes (Final run - MPP must be completed)")

# while True:
#     final_run = input("\nEnter choice (yes/no): ").strip()
#     if final_run in ["yes", "no"]:
#         break
#     print("Invalid input. Please enter yes or no.")

# # ============================================================
# # MPP Mandatory Confirmation (Final Run Only)
# # ============================================================

# if final_run == "yes":
#     print("\n⚠️ FINAL SUMMARY CHECK ⚠️")
    
#     while True:
#         mpp_check = input(
#             "\nHave all MPP post-load steps been completed in ISC? (yes/no): "
#         ).strip().lower()

#         if mpp_check == "yes":
#             print("\n✅ MPP confirmation recorded. Proceeding with final summary...")
#             break
#         elif mpp_check == "no":
#             print("\n❌ Please complete the MPP steps before generating the final summary.")
#             exit()
#         else:
#             print("Invalid input. Please enter 'yes' or 'no'.")


# ─────────────────────────────────────────────────────────────────────────────
# Step 1 — Select the success/error CSVs from Downloads and move them.
#           The CSV picker stays manual (files come from Salesforce).
#           But the destination folder is now auto-detected (Q5).
# ─────────────────────────────────────────────────────────────────────────────
print("\n\n🔍 Step 1: Select success/error CSV files from Downloads.")

# ── Auto-detect summary file first (used to derive the destination folder) ──
summary_file_path = _auto_detect_paths()

# ── Fall back to manual picker if auto-detect declined or failed ──
if not summary_file_path:
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    summary_file_path = filedialog.askopenfilename(
        title="📂 Select the Summary File for this iteration"
    )
    if not summary_file_path:
        print("\n\n    ❌ No summary file selected. Operation canceled.")
        exit()

print(f"\n    ✅ Summary file:")
shortened_path = "/".join(summary_file_path.split("/")[-5:])
print(f"\n         📂 {shortened_path}")

# Derive the Success and error files destination automatically from the summary file
_summary_parent     = os.path.dirname(summary_file_path)
destination_folder  = os.path.join(_summary_parent, "Success and error files")
os.makedirs(destination_folder, exist_ok=True)
print(f"\n    📁 Destination: …/{'/'.join(destination_folder.split('/')[-4:])}")

# ── Now pick the CSVs from Downloads (always manual — they come from Salesforce) ──
try:
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    csv_files = root.tk.splitlist(filedialog.askopenfilenames(
        title="📂 Select Success/Error CSV files from Downloads",
        filetypes=[("CSV Files", "*.csv")],
        initialdir=os.path.expanduser("~/Downloads")
    ))

    if not csv_files:
        print("\n    ❗️ No CSV files selected. Skipping file move.")
    else:
        success_files = []
        failed_files  = []

        for file in csv_files:
            try:
                shutil.move(file, destination_folder)
                success_files.append(os.path.basename(file))
            except Exception:
                failed_files.append(os.path.basename(file))

        if success_files:
            print("\n    ✅ Successfully moved files:")
            for f in success_files:
                print(f"\n        📄 {f}")

        if failed_files:
            print("\n    ❌ Failed to move files:")
            for f in failed_files:
                print(f"\n        ❗️ {f}")

except Exception as e:
    print(f"Error: {e}")

# =======================================================================================================
# Step 2 — summary file already selected above (auto-detected or manual).
print("\n\n🔍 Step 2: Summary file already selected ✅")

print(f"\n    ✅ Selected file:")
shortened_path = "/".join(summary_file_path.split("/")[-5:])
print(f"\n         📂 {shortened_path}")


# Taking the file name automatically from Summary file name
file_name = summary_file_path.split("/")[-1].split("_summary file.xlsx")[0]



# Define the CSV files and their corresponding sheet names
files_and_sheets = {
    "opptysuccess.csv"  : "Opportunity Success",
    "opptyerror.csv"    : "Opportunity Failures",
    "productsuccess.csv": "Opportunity Product Success",
    "producterror.csv"  : "Opportunity Product Failures",
    "teamsuccess.csv"   : "Team Member Success",
    "teamerror.csv"     : "Team Member Failure",
    "codessuccess.csv"  : "Reporting Code Success",
    "codeserror.csv"    : "Reporting Code Failure",
    "tagssuccess.csv"   : "Tag Success",
    "tagserror.csv"     : "Tag Failure",
    "feedsuccess.csv"   : "Feed Item Success",
    "feederror.csv"     : "Feed Item Failure",
    "contactsuccess.csv": "Contact Success",
    "contacterror.csv"  : "Contact Failure"
}

# Define Success and error folder path

half_path = '/'.join(summary_file_path.split('/')[:-1]) #Automatically take the path from Summary file selected above 
# for example 
# if summary file path is '/Users/avirajmore/Documents/Office Docs/Massload Files/2025/Demo/demo/Final iteration files/TSS Massload - Batch 133-10012025/Mass_load_template_CHEVA - IAPP 2024_summary file.xlsx'
# Half path will be:- '/Users/avirajmore/Documents/Office Docs/Massload Files/2025/Demo/demo/Final iteration files/TSS Massload - Batch 133-10012025

Success_error_path = half_path + '/Success and error files'
# For example:- 
# '/Users/avirajmore/Documents/Office Docs/Massload Files/2025/Demo/demo/Final iteration files/TSS Massload - Batch 133-10012025' + '/Success and error files'

# Check if the Excel file exists
if os.path.exists(summary_file_path):
    workbook = load_workbook(summary_file_path)
else:
    print(f"\n    ❌ Error: Excel file '{summary_file_path}' not found. Exiting...")
    exit()

# Check if any sheet already has data
data_exists = any(
    sheet.max_row > 1 for sheet in workbook.worksheets if sheet.title in files_and_sheets.values()
)

# Ask user once if they want to continue, keep asking until valid input
if data_exists:
    while True:
        user_response = input(
            "\n        ❗️ Some sheets already have data. Do you want to continue and append data to all sheets? (yes/no): "
        ).strip().lower()

        if user_response == 'yes':
            break
        elif user_response == 'no':
            print("\n    ❌ Operation canceled by the user.")
            exit()
        else:
            print("\n           ❌ Invalid choice. Please enter 'yes' or 'no'.")

# Start processing each CSV file
# Process each CSV file
processed_files = []
skipped_files = []

for csv_filename, sheet_name in files_and_sheets.items():
    csv_file_path = os.path.join(Success_error_path, csv_filename)

    # Check if the CSV file exists
    if not os.path.exists(csv_file_path):
        skipped_files.append((csv_filename, "File not found"))
        continue

    # Load the CSV data
    csv_data = pd.read_csv(csv_file_path)

    # Create the sheet if it doesn't exist
    if sheet_name not in workbook.sheetnames:
        workbook.create_sheet(sheet_name)

    # Overwrite the CSV data in the sheet
    sheet = workbook[sheet_name]

    # Clear existing sheet content
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            cell.value = None
    # Clear all rows in the sheet
    sheet.delete_rows(1, sheet.max_row)

    # Write the new CSV data
    for row in dataframe_to_rows(csv_data, index=False, header=True):
        sheet.append(row)

    processed_files.append(csv_filename)

# Save the workbook after all modifications
workbook.save(summary_file_path)


print(f"\n    ✅ {len(processed_files)} CSV file(s) successfully written to the Excel file.")

print(f"\n    ⭕️ {len(skipped_files)} file(s) were not processed.")
for i, (file, reason) in enumerate(skipped_files, 1):
    print(f"\n        ❗️ {file}")



# ==============================================================
# Step 3 :- Copy the data from Removed rows to summary file
# ==============================================================

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
print("\n\n🔍 Step 3: Copy the data from Removed rows to summary file")

# Define the CSV-to-Sheet mapping
files_and_sheets = {

    "Removed_Rows - Oppty.csv"                                  : "Opportunity Failures",

    "Removed_Rows - Product.csv"                                : "Opportunity Product Failures",
    
    "Removed_Rows - Product (Duplicate Opportunity).csv"        : "Opportunity Product Failures",
    
    "Removed_Rows - Team.csv"                                   : "Team Member Failure",
    
    "Removed_Rows - Team member (Duplicate Opportunity).csv"    : "Team Member Failure",
    
    "Removed_Rows - ReportingCodes.csv"                         : "Reporting Code Failure",
    
    "Removed_Rows - Codes (Duplicate Opportunity).csv"          : "Reporting Code Failure",
    
    "Removed_Rows - Tags.csv"                                   : "Tag Failure",
    
    "Removed_Rows - Tags (Duplicate Opportunity).csv"           : "Tag Failure",
    
    "Removed_Rows - Feed Item.csv"                              : "Feed Item Failure",
    
    "Removed_Rows - Contact.csv"                                : "Contact Failure"
}

# Define column mapping (CSV column name -> Summary sheet column name)
column_mapping = {
    "Reason": "ERRORS",  
    "ERROR": "ERRORS"
}

# Define the folder path where the CSV files are stored
Remove_rows_path = half_path + '/Removed Rows'

# Ask user for the summary file path
# summary_file_path = input("Enter the full path of the summary Excel file: ").strip()

# Check if the Excel file exists
if not os.path.exists(summary_file_path):
    print(f"\n    ❌ Error: Excel file '{summary_file_path}' not found. Exiting...")
    exit()

# Load the workbook
workbook = load_workbook(summary_file_path)

# Process each CSV file
processed_files = []
skipped_files = []

for csv_filename, sheet_name in files_and_sheets.items():
    csv_file_path = os.path.join(Remove_rows_path, csv_filename)

    # Check if the CSV file exists
    if not os.path.exists(csv_file_path):
        skipped_files.append((csv_filename, "File not found"))
        continue

    # Load the CSV data
    csv_data = pd.read_csv(csv_file_path)

    # Apply column mapping
    csv_data.rename(columns=column_mapping, inplace=True)

    # Check if the sheet exists in the workbook
    if sheet_name not in workbook.sheetnames:
        workbook.create_sheet(sheet_name)

    # Get the sheet
    sheet = workbook[sheet_name]

    # Get the header of the sheet (summary file)
    sheet_header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
    
    # ✅ Rename sheet headers if needed (e.g., ERROR → ERRORS)
    if sheet_header:
        sheet_header = list(sheet_header)
        for col_idx, col_name in enumerate(sheet_header):
            if col_name in column_mapping:
                new_name = column_mapping[col_name]
                sheet.cell(row=1, column=col_idx + 1, value=new_name)
                sheet_header[col_idx] = new_name  # Update local list
    
    # If the sheet is empty, populate it with CSV headers
    if not sheet_header or all(h is None for h in sheet_header):
        sheet_header = list(csv_data.columns)  # Convert pandas Index to a list
        for col_idx, header in enumerate(sheet_header, start=1):
            sheet.cell(row=1, column=col_idx, value=header)  # Write headers to the first row

    # Reorder CSV columns to match the sheet's columns
    csv_data = csv_data.reindex(columns=sheet_header, fill_value=None)

    # Append data row by row (excluding headers from CSV)
    for row in csv_data.itertuples(index=False):
        sheet.append(row)

    processed_files.append(csv_filename)

# Save the workbook after all modifications
workbook.save(summary_file_path)

# Print Summary
print(f"\n    ✅ {len(processed_files)} CSV file(s) successfully appended to the Excel file.")
if skipped_files:
    print(f"\n    ⭕ {len(skipped_files)} file(s) were not processed:")
    for i, (file, reason) in enumerate(skipped_files, 1):
        print(f"\n        {i}. {file}")


# ==============================================================
# Step 4 :- Mention the count of Success and Error files
# ==============================================================

import os
from openpyxl import load_workbook

print("\n\n🔍 Step 4: Mention the count of Success and Error files")

# Define the mapping of sheets to cells in the "Summary" sheet
sheet_to_cell_mapping = {
    "Opportunity Success": "E5",
    "Opportunity Failures": "F5",
    "Opportunity Product Success": "E6",
    "Opportunity Product Failures": "F6",
    "Team Member Success": "E7",
    "Team Member Failure": "F7",
    "Reporting Code Success": "E8",
    "Reporting Code Failure": "F8",
    "Tag Success": "E9",
    "Tag Failure": "F9",
    "Contact Success":"E10",
    "Contact Failure":"F10",
    "Feed Item Success":"E11",
    "Feed Item Failure":"F11"
}

# Define the cell ranges for summation in the "Summary" sheet
sum_mappings = {
    ("E5", "F5"): "G5",
    ("E6", "F6"): "G6",
    ("E7", "F7"): "G7",
    ("E8", "F8"): "G8",
    ("E9", "F9"): "G9",
    ("E10", "F10"): "G10",
    ("E11", "F11"): "G11",

}

# Ask user for the summary file path
# summary_file_path = input("Enter the full path of the summary Excel file: ").strip()

# Check if the file exists
if not os.path.exists(summary_file_path):
    print(f"\n    ❌ Error: Summary file '{summary_file_path}' not found. Exiting...")
    exit()

# Load the workbook
workbook = load_workbook(summary_file_path)

# Check if the "Summary" sheet exists
if "Summary" not in workbook.sheetnames:
    print(f"\n    ❌ Error: 'Summary' sheet not found in the workbook. Exiting...")
    exit()

# Get the "Summary" sheet
summary_sheet = workbook["Summary"]

# List to collect names of missing sheets
missing_sheets = []

# Process each sheet and count rows
for sheet_name, cell_address in sheet_to_cell_mapping.items():
    if sheet_name not in workbook.sheetnames:
        missing_sheets.append(sheet_name)  # Collect missing sheet names
        continue

    # Get the sheet
    sheet = workbook[sheet_name]

    # Count the number of rows (excluding the header)
    row_count = sum(1 for _ in sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True) if any(_))

    # Write the row count to the specified cell in the "Summary" sheet
    summary_sheet[cell_address] = row_count

# Perform the summation in the "Summary" sheet
for (cell_1, cell_2), result_cell in sum_mappings.items():
    value_1 = summary_sheet[cell_1].value or 0
    value_2 = summary_sheet[cell_2].value or 0
    summary_sheet[result_cell] = value_1 + value_2

# Save the updated workbook
workbook.save(summary_file_path)

# Output the result
if missing_sheets:
    print(f"\n    ❗️ Skipped as the sheet was not found")
    for i, sheet in enumerate(missing_sheets, 1):
        print(f"\n        {i}. {sheet}")

print(f"\n    ✅ Row counts and summations updated successfully in the 'Summary' sheet of :")
shortened_path = "/".join(summary_file_path.split('/')[-5:])
print(f"\n         📂 {shortened_path}")


# ============================================================
# Show the summary data as table
# ============================================================
from collections import defaultdict
from tabulate import tabulate

# Group data by high-level categories
categories = defaultdict(lambda: {"Success": 0, "Failures": 0, "Total": 0})

# # Assuming `sheet_to_cell_mapping` is already defined as provided
# sheet_to_cell_mapping = {
#     "Opportunity Success": "E5",
#     "Opportunity Failures": "F5",
#     "Opportunity Product Success": "E6",
#     "Opportunity Product Failures": "F6",
#     "Team Member Success": "E7",
#     "Team Member Failure": "F7",
#     "Reporting Code Success": "E8",
#     "Reporting Code Failure": "F8",
#     "Tag Success": "E9",
#     "Tag Failure": "F9"
# }

# Example summary_sheet (assuming it's already loaded)
# summary_sheet = your_loaded_summary_sheet_object

for sheet_name, cell_address in sheet_to_cell_mapping.items():
    # Extract high-level category from the sheet name (e.g., remove "Success"/"Failures")
    category = " ".join(sheet_name.split()[:-1])
    
    # Identify Success or Failures based on the sheet name
    if "Success" in sheet_name:
        value = summary_sheet[cell_address].value or 0
        categories[category]["Success"] += value
    elif "Failure" in sheet_name:
        value = summary_sheet[cell_address].value or 0
        categories[category]["Failures"] += value

# Calculate the Total for each category
for category, counts in categories.items():
    counts["Total"] = counts["Success"] + counts["Failures"]

# Prepare the table data without the "Comments" column
table_data = [
    [category, counts["Success"], counts["Failures"], counts["Total"]]
    for category, counts in categories.items()
]

# Define table headers with bold formatting (headers are bold)
headers = ["Summary", "Success", "Failures", "Total"]

# Print the table with enhanced formatting
print("\n\n🔍 Aggregated Counts from the 'Summary' Sheet:\n")
table = tabulate(table_data, headers=headers, tablefmt="fancy_grid", numalign="right", stralign="center")
# Bold headers
table = table.replace(headers[0], f"\033[1m{headers[0]}\033[0m")
table = table.replace(headers[1], f"\033[1m{headers[1]}\033[0m")
table = table.replace(headers[2], f"\033[1m{headers[2]}\033[0m")
table = table.replace(headers[3], f"\033[1m{headers[3]}\033[0m")

print(table)

print("\n    ✅ Simulation Complete. 'Summary' Sheet Updated Successfully.\n")

# ============================================================
# ============================================================

print("\n")
title =  "📝 Success And Error Files Completed 📝"
print(f"{line}\n")
print(title.center(line_width))
print(f"{line}\n")
print("\n")