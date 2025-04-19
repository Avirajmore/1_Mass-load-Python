import pandas as pd
import pyperclip
import os
import re

# Folder paths
folder_path = '/Users/avirajmore/Downloads/'
accounts_csv_path = '/Users/avirajmore/Downloads/accounts.csv'
valid_output_path = os.path.expanduser("~/Downloads/Accounts_to_be_imported.xlsx")
invalid_output_path = os.path.expanduser("~/Downloads/Invalid_Accounts.xlsx")

# Load AccountNumber list from CSV
account_list_df = pd.read_csv(accounts_csv_path)
account_numbers_set = set(account_list_df['AccountNumber'].astype(str).str.strip())

# Regular expression for country codes like -US, -KA etc.
country_code_pattern = r'-[A-Za-z]{2,3}$'

# Summary stats
total_valid = 0
total_invalid = 0
files_with_no_matches = []

# Helper to process accountid values
def process_value(value):
    if isinstance(value, str) and value.startswith('DC'):
        return value.split('-')[0]
    return value

# Read existing valid/invalid files if they exist
if os.path.exists(valid_output_path):
    valid_df = pd.read_excel(valid_output_path)
else:
    valid_df = pd.DataFrame(columns=['Accounts'])

if os.path.exists(invalid_output_path):
    invalid_df = pd.read_excel(invalid_output_path)
else:
    invalid_df = pd.DataFrame(columns=['Invalid Accounts'])

print("\n✅ Files To be Processed")
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        try:
            df = pd.read_excel(file_path, sheet_name='Opportunity')
        except Exception as e:
            continue

        if 'accountid' not in df.columns:
            continue
        print(f"\n    📁 {file_name}")

# Process each Excel file in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        try:
            df = pd.read_excel(file_path, sheet_name='Opportunity')
        except Exception as e:
            continue

        if 'accountid' not in df.columns:
            continue

        # Clean and process accountid
        df['accountid'] = df['accountid'].astype(str).str.replace(r'\s+', '', regex=True).str.strip()
        df['accountid'] = df['accountid'].apply(process_value)

        # Find values not in CSV
        not_in_csv = df[~df['accountid'].isin(account_numbers_set)]['accountid'].dropna().unique()

        file_valid = []
        file_invalid = []

        for value in not_in_csv:
            value = str(value).strip()
            if not (value.lower().startswith('db') or value.lower().startswith('dc')):
                file_invalid.append(value)
            elif value.lower().startswith('db'):
                if not re.search(country_code_pattern, value):
                    file_invalid.append(value)
                else:
                    file_valid.append(value)
            else:
                file_valid.append(value)

        # Track if file had no matches
        if not file_valid and not file_invalid:
            files_with_no_matches.append(file_name)

        # Append to cumulative DataFrames
        if file_valid:
            valid_df = pd.concat([valid_df, pd.DataFrame(file_valid, columns=['Accounts'])], ignore_index=True)
            total_valid += len(file_valid)
        if file_invalid:
            invalid_df = pd.concat([invalid_df, pd.DataFrame(file_invalid, columns=['Invalid Accounts'])], ignore_index=True)
            total_invalid += len(file_invalid)

# Save final valid and invalid files
valid_df.drop_duplicates().to_excel(valid_output_path, index=False)
invalid_df.drop_duplicates().to_excel(invalid_output_path, index=False)

# Final summary
print("\n✅ Processing Complete!")
print(f"\n    ❗️ Total Accounts to Be Imported: {total_valid}")
print(f"\n    ❗️ Invalid Accounts: {total_invalid}")
if files_with_no_matches:
    print("\n✅ Files with No Missing Accounts:")
    for fname in files_with_no_matches:
        print(f"\n    📄 {fname}")
else:
    print("\n    ✅ All files had some valid or invalid accounts.\n")

