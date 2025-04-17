import shutil
import tkinter as tk
from tkinter import ttk
import re
import os
import pandas as pd
import pyperclip
import itertools
from openpyxl import load_workbook

# ==================================================
print(f"\n{'='*120}\n{' ' * 30} 📝  Extract Data and create Queries 📝 {' ' * 30}\n{'='*120}\n")
# ==================================================


# Directory where extracted files will be saved
folder_path = 'Extracted_Files'

# Create the folder if it doesn't exist
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# Path where source Excel files are located
DIR_PATH = os.path.expanduser("~/Downloads")

# Path for the consolidated extracted data
output_file = "Extracted_Files/Extracted_data.xlsx"

print("\n🔍 Select Files to Process")
# Path to your folder
folder_path = '/Users/avirajmore/Downloads'

# Get only Excel files in the folder
files = [f for f in os.listdir(folder_path)
         if os.path.isfile(os.path.join(folder_path, f)) and f.lower().endswith(('.xlsx', '.xls'))]

# Store selected file names
selected_files = []

# Tkinter window
root = tk.Tk()
root.title("Select Excel Files")
root.geometry("400x500")

# Scrollable frame
canvas = tk.Canvas(root)
scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

# Create a dictionary to hold file checkboxes
check_vars = {}

def toggle_all():
    state = select_all_var.get()
    for var in check_vars.values():
        var.set(state)

def submit_selection():
    global selected_files
    selected_files = [file for file, var in check_vars.items() if var.get()]
    print("Selected Excel Files:")
    for file in selected_files:
        print(file)
    root.destroy()

# Select All checkbox
select_all_var = tk.BooleanVar()
select_all_cb = ttk.Checkbutton(scrollable_frame, text="Select All", variable=select_all_var, command=toggle_all)
select_all_cb.pack(anchor='w', pady=(10, 5), padx=10)

# Individual file checkboxes
for file in files:
    var = tk.BooleanVar()
    cb = ttk.Checkbutton(scrollable_frame, text=file, variable=var)
    cb.pack(anchor='w', padx=20)
    check_vars[file] = var

# Submit button
submit_btn = ttk.Button(root, text="Submit", command=submit_selection)
submit_btn.pack(pady=10)

# Pack canvas and scrollbar
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Start GUI
root.mainloop()

# Now selected_files contains all the checked Excel file names


print("\n🔍 Step 1: Extract Data from Files:")
print("\n   ✅ Files Processed:")

# Loop through each file in the source directory
for file in selected_files:
    if file.endswith(".xlsx"):
        file_path = os.path.join(DIR_PATH, file)

        # Load all sheets from the current Excel file
        xls = pd.read_excel(file_path, sheet_name=None,engine='openpyxl')

        # Extract and modify data from the 'Opportunity_product' sheet
        if 'Opportunity_product' in xls:
            df = xls['Opportunity_product']

            # Concatenate 'Product' and 'product_type' columns into a new 'product_family' column
            df['product_family'] = df['Product'].astype(str) + '-' + df['product_type'].astype(str)

            # Define how columns should be extracted and grouped from various sheets
            sheet_config = {
                "Opportunity": {
                    "columns": ["accountid", "ownerid", "created_by", "currency_code", "opportunity_legacy_id_c"],
                    "output": {
                        "Accounts": ["accountid"],
                        "Email_id": ["ownerid", "created_by"],
                        "Currency": ["currency_code"],
                        "Legacy_Ids": ["opportunity_legacy_id_c"]
                    }
                },
                "Opportunity_product": {
                    "columns": ["product_family"],
                    "output": {
                        "Product": ["product_family"]
                    }
                },
                "Opportunity_Team ": {
                    "columns": ["Email"],
                    "output": {
                        "Email_id": ["Email"]
                    }
                },
                "Reporting_codes": {
                    "columns": ["reporting_codes", "Tags"],
                    "output": {
                        "Strategy": ["reporting_codes", "Tags"]
                    }
                }
            }

            collected_data = {}

            # Extract data from each sheet based on config
            for sheet_key, config in sheet_config.items():
                if sheet_key not in xls:
                    print(f"Sheet '{sheet_key}' not found.")
                    continue

                df = xls[sheet_key]
                df.columns = df.columns.str.strip() # Clean column names

                for out_sheet, needed_cols in config["output"].items():
                    valid_cols = [col for col in needed_cols if col in df.columns]
                    if not valid_cols:
                        continue

                    # Stack and combine columns into the specified output sheet structure
                    if out_sheet == "Email_id":
                        stacked = pd.concat([df[col].dropna().astype(str) for col in valid_cols], ignore_index=True).to_frame(name="Email_id")
                        collected_data.setdefault(out_sheet, []).append(stacked)

                    elif out_sheet == "Strategy":
                        stacked = pd.concat([df[col].dropna().astype(str) for col in valid_cols], ignore_index=True).to_frame(name="Strategy")
                        collected_data.setdefault(out_sheet, []).append(stacked)

                    else:
                        subset = df[valid_cols].dropna(how='all')
                        if not subset.empty:
                            collected_data.setdefault(out_sheet, []).append(subset)

        # Append or create a new extracted data Excel file
        file_exists = os.path.exists(output_file)

        with pd.ExcelWriter(output_file, engine="openpyxl", mode="a" if file_exists else "w", if_sheet_exists="overlay" if file_exists else None) as writer:
            for sheet_name, dfs in collected_data.items():
                combined = pd.concat(dfs, ignore_index=True)

                # If sheet exists, append data; else create new sheet
                if file_exists:
                    try:
                        existing = pd.read_excel(output_file, sheet_name=sheet_name)
                        combined = pd.concat([existing, combined], ignore_index=True)
                    except Exception:
                        pass  # Sheet doesn't exist, just write new

                combined.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"\n       📁 {file}")

print(f"\n   ✅ Data Extracted and stored in {output_file}:")

# ============================
# Trim whitespace in all cells and column names
# ============================

print("\n🔍 Step 2: Cleaning data")

def trim_columns_all_sheets(extract_file_path):
    # Read all sheets into a dictionary
    sheets = pd.read_excel(extract_file_path, sheet_name=None, dtype=str)  # Read all as strings to ensure trimming works
    
    trimmed_sheets = {}

    for sheet_name, df in sheets.items():
        new_df = pd.DataFrame()
        
        for column in df.columns:
            new_column_name = column.strip()  # Also trim column names
            if column.lower().strip() == 'accountid':
                new_df[new_column_name] = df[column].astype(str).str.replace(r'\s+', '', regex=True).str.strip()
            else:
                new_df[new_column_name] = df[column].astype(str).str.strip()
        
        trimmed_sheets[sheet_name] = new_df
    
    # Save the trimmed data back to a new Excel file
    with pd.ExcelWriter(extract_file_path, engine='openpyxl') as writer:
        for sheet_name, df in trimmed_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# Usage
extract_file_path = 'Extracted_Files/Extracted_data.xlsx'

trim_columns_all_sheets(extract_file_path)

print("\n   ✅ Trimmed Data")

# ============================
# Expand comma-separated values into separate rows via cartesian product
# ============================

# Load the Excel file
sheets_to_process = ['Email_id', 'Strategy']

# Read all sheets
all_sheets = pd.read_excel(extract_file_path, sheet_name=None)

# Process the Email_id and Strategy sheets
for sheet_name in sheets_to_process:
    if sheet_name in all_sheets:
        df = all_sheets[sheet_name]
        new_rows = []

        for _, row in df.iterrows():
            # Split values in each cell by comma, strip whitespace
            split_row = [str(cell).split(',') if isinstance(cell, str) and ',' in cell else [cell] for cell in row]
            split_row = [[val.strip() for val in values] for values in split_row]

            # Create cartesian product (every combination from split lists)
            for combination in itertools.product(*split_row):
                new_rows.append(combination)

        # Create a new DataFrame from the expanded rows
        all_sheets[sheet_name] = pd.DataFrame(new_rows, columns=df.columns)
    else:
        print(f"Sheet '{sheet_name}' not found in the Excel file.")

# Write all sheets back to the same Excel file
with pd.ExcelWriter(extract_file_path, engine='openpyxl', mode='w') as writer:
    for sheet_name, df in all_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\n   ✅ Separated Values based on Comma")

# ============================
# Process Account IDs - strip extra parts from values starting with 'DC'
# ============================

accountid_column = 'accountid'
new_column_name = 'accountid'

# Load the specific sheet into a DataFrame
df = pd.read_excel(extract_file_path, sheet_name="Accounts")

# Define a function to process the values
def process_value(value):
    if isinstance(value, str) and value.startswith('DC'):
        return value.split('-')[0]
    return value

# Apply the function to the accountid column and store results in the new column
df[new_column_name] = df[accountid_column].apply(process_value)

# Save the updated DataFrame back to the Excel file
with pd.ExcelWriter(extract_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Accounts', index=False)

print("\n   ✅ Formated Account values to the correct Format")

# ============================
# Remove duplicates from each sheet based on the first column
# ============================

# Load all sheets as a dict of DataFrames
xls = pd.read_excel(extract_file_path, sheet_name=None)

# Remove duplicates in each sheet
for sheet_name, df in xls.items():
    # Check if there's at least one column
    if not df.empty:
        first_col = df.columns[0]
        df = df.drop_duplicates(subset=first_col)
        xls[sheet_name] = df

# Save back to the same file
with pd.ExcelWriter(extract_file_path, engine='openpyxl', mode='w') as writer:
    for sheet_name, df in xls.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\n   ✅ Removed Duplicates")

# ============================
# Concatenate each cell value with quotes and a comma for SQL formatting
# ============================

print("\n🔍 Step 3: Create file with concatenated Data")
xls = pd.ExcelFile(extract_file_path)

# Define the output file path to avoid overwriting the original
output_file_path = 'Extracted_Files/Concatenated_excel_file.xlsx'

# Create an Excel writer to save the modified data
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # Iterate through all sheets
    for sheet_name in xls.sheet_names:
        # Load the sheet into a DataFrame
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Iterate through each column in the DataFrame
        for column in df.columns:
            # Concatenate each value with single quotes and a comma
            df[column] = df[column].apply(lambda x: f"'{x}',")
        
        # Save the modified DataFrame to the new Excel file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\n   ✅ Concated values saved in {output_file_path}")

# ============================
# Function to generate a query file from a single sheet column
# ============================
print("\n🔍 Step 4: Generating Queries")
def generate_query_from_sheet(extract_file_path, sheet_name, column_name, query_template, output_txt_base):
    # Read specific sheet
    df = pd.read_excel(extract_file_path, sheet_name=sheet_name, dtype=str)

    # Drop empty values and clean
    values = df[column_name].dropna().astype(str).str.strip().tolist()

    # Query parts length
    base_query_length = len(query_template.replace("{values}", ""))
    max_query_length = 90000

    # Estimate max value chunk size per query
    value_strings = [f"'{val}'" for val in values]

    chunks = []
    current_chunk = []
    current_length = base_query_length

    for val in value_strings:
        val_length = len(val) + 1  # for comma and newline
        if current_length + val_length > max_query_length:
            # Save current chunk
            chunks.append(current_chunk)
            # Start new chunk
            current_chunk = [val]
            current_length = base_query_length + val_length
        else:
            current_chunk.append(val)
            current_length += val_length

    if current_chunk:
        chunks.append(current_chunk)

    # Ensure output folder exists
    os.makedirs(os.path.dirname(output_txt_base), exist_ok=True)

    # Write each query to separate file
    for idx, chunk in enumerate(chunks, start=1):
        formatted_values = ",\n".join(chunk)
        query = query_template.replace("{values}", formatted_values)
        if len(chunks) == 1:
                output_txt = output_txt_base
        else:
            output_txt = output_txt_base.replace(".txt", f"_part{idx}.txt")
        with open(output_txt, "w") as file:
            file.write(query)

    print(f"\n       📁 Generated {len(chunks)} query file(s) for {output_txt_base.split("/")[-1]}.")


# Queries generation (various sheets & templates)

account_output_txt = 'Extracted_Files/Queries/1_Account_query.txt'
account_query = "SELECT AccountNumber, id FROM Account WHERE AccountNumber IN ({values})"
generate_query_from_sheet(extract_file_path, 'Accounts', 'accountid', account_query,account_output_txt)

Email_output_txt = 'Extracted_Files/Queries/2_Userid_query.txt'
userid_query = "select email,id,Profile.Name,isactive from user where email in ({values}) and Profile.Name != 'IBM Partner Community Login User' and IsActive = true"
generate_query_from_sheet(extract_file_path, 'Email_id', 'Email_id', userid_query,Email_output_txt)

strategy_output_txt = 'Extracted_Files/Queries/4_Strategy_query.txt'
strategy_query = "Select id,name,Record_Type_Name__c from Strategy__c where name in ({values})"
generate_query_from_sheet(extract_file_path, 'Strategy', 'Strategy', strategy_query,strategy_output_txt)

legacy_output_txt = 'Extracted_Files/Queries/5_Legacy_query.txt'
legacy_query = "SELECT Opportunity_Legacy_Id__c, Id,Name,Owned_By_Name__c,OwnerId FROM Opportunity WHERE Opportunity_Legacy_Id__c IN ({values})"
generate_query_from_sheet(extract_file_path, 'Legacy_Ids', 'opportunity_legacy_id_c', legacy_query,legacy_output_txt)

# ============================
# Function to generate query using two sheets and columns
# ============================

def generate_query_from_two_sheets(extract_file_path, sheet1, column1, sheet2, column2, query_template, output_txt):
    # Read both sheets
    df1 = pd.read_excel(extract_file_path, sheet_name=sheet1, dtype=str)
    df2 = pd.read_excel(extract_file_path, sheet_name=sheet2, dtype=str)

    # Clean and dropna
    values1 = df1[column1].dropna().astype(str).str.strip().tolist()
    values2 = df2[column2].dropna().astype(str).str.strip().tolist()

    # Format values for SQL IN clause
    formatted_values1 = ",\n".join([f"'{val}'" for val in values1])
    formatted_values2 = ",\n".join([f"'{val}'" for val in values2])

    # Replace placeholders in query template
    query = query_template.replace("{product}", formatted_values1).replace("{currency}", formatted_values2)

    # Ensure output folder exists
    os.makedirs(os.path.dirname(output_txt), exist_ok=True)

    # Write query to file
    with open(output_txt, "w") as file:
        file.write(query)


# Product & currency query generation
sheet1 = 'Product'
column1 = 'product_family'
sheet2 = 'Currency'
column2 = 'currency_code'
product_query = "SELECT Product2.Product_Code_Family__c, CurrencyIsoCode, id, isactive FROM PricebookEntry WHERE Product2.Product_Code_Family__c IN ({product}) AND CurrencyIsoCode IN ({currency})"
product_output_txt = 'Extracted_Files/Queries/3_PricebookEntry_query.txt'

generate_query_from_two_sheets(extract_file_path, sheet1, column1, sheet2, column2, product_query, product_output_txt)

print("\n   ✅ Queries Generated")

# ============================
# Function to find & rename 'bulkQuery' CSV file
# ============================

def rename_and_move_bulkquery_file(new_name, DIR_PATH):
    """
    Searches the downloads folder for a file with 'bulkQuery' in the name and 
    renames/moves it to the designated CSV directory using the provided new name.
    """
    
    for filename in os.listdir(DIR_PATH):
        if "bulkQuery" in filename and filename.endswith(".csv"):
            old_path = os.path.join(DIR_PATH, filename)
            new_path = os.path.join(DIR_PATH, new_name)
            shutil.move(old_path, new_path)
            return True  # Successful rename and move
    return False  # No matching file found

# ==================================================
print(f"\n{'='*120}\n{' ' * 30} 📝 Files Processed 📝 {' ' * 30}\n{'='*120}\n")
# ==================================================


DIR_PATH = "Extracted_Files"

while True:
    Copy_query = input("\n📝 Do you want to proceed With Copying the Queries? (yes/no): ").strip().lower()

    if Copy_query == 'yes':
        query_dir = os.path.join(DIR_PATH, "Queries")
        query_files = sorted([f for f in os.listdir(query_dir) if f.endswith(".txt")])  # sorted alphabetically

        if not query_files:
            print("\n❌ No query files found.")
            break

        print("\n📑 Available Query Files:")
        for idx, file_name in enumerate(query_files, start=1):
            print(f"\n   {idx}) {file_name}")

        while True:
            choose = input('\n📝 Enter the number of the query you want to copy or type "s" to Skip: ').strip()

            if choose.lower() == 's':
                print("\n🚫 Skipping")
                break_outer_loop = True
                break

            if choose.isdigit() and 1 <= int(choose) <= len(query_files):
                selected_file = query_files[int(choose) - 1]
                with open(os.path.join(query_dir, selected_file), "r") as f:
                    query = f.read()

                print(f"\n✅ {selected_file} copied to clipboard")
                pyperclip.copy(query)

            else:
                print("\n❗️ Invalid Selection")
                continue

    elif Copy_query == 'no':
        print("\n🚫 Skipping\n")
        break

    else:
        print("\n❗️ Invalid Choice")

# ======================

while True:

    vlookup = input("\n📝 Do you want to proceed With Vlookup?(yes/no): ").strip().lower()

    if vlookup == 'yes':

        # Load the Excel sheet with account IDs
        extract_file_path = 'Extracted_Files/Extracted_data.xlsx'  # replace with your Excel file path
        accounts_df = pd.read_excel(extract_file_path, sheet_name='Accounts')

        # Load the CSV file with AccountNumber and Id
        csv_file_path = os.path.expanduser("~/Downloads")+'/accounts.csv'  # replace with your CSV file path
        lookup_df = pd.read_csv(csv_file_path)

        # Convert both to lowercase for case-insensitive matching
        accounts_df['accountid_lower'] = accounts_df['accountid'].astype(str).str.lower()
        lookup_df['AccountNumber_lower'] = lookup_df['AccountNumber'].astype(str).str.lower()

        # Perform the left join like VLOOKUP
        merged_df = pd.merge(
            accounts_df,
            lookup_df[['AccountNumber_lower', 'Id']],
            left_on='accountid_lower',
            right_on='AccountNumber_lower',
            how='left'
        )

        # Fill NaN values with 'Not found in ISC'
        merged_df['Id'] = merged_df['Id'].fillna('Not found in ISC')

        # Optional: drop helper columns if not needed
        merged_df.drop(columns=['accountid_lower', 'AccountNumber_lower'], inplace=True)

        # Save to a new Excel file
        output_file_path = 'Extracted_Files/Accounts_vlookup.xlsx'
        merged_df.to_excel(output_file_path, index=False)

        print(f"\n    ✅ VLOOKUP completed for Accounts.")

        try:

            # Load the Excel file's "Strategy" sheet
            extract_file_path = 'Extracted_Files/Extracted_data.xlsx'  # replace with your Excel file path
            strategy_df = pd.read_excel(extract_file_path, sheet_name='Strategy')

            # Load the tags CSV file
            tags_csv_path = os.path.expanduser("~/Downloads")+ '/tags.csv'  # replace with your tags CSV file path
            tags_df = pd.read_csv(tags_csv_path)

            # Convert both columns to lowercase for case-insensitive matching
            strategy_df['Strategy_lower'] = strategy_df['Strategy'].astype(str).str.lower()
            tags_df['Name_lower'] = tags_df['Name'].astype(str).str.lower()

            # Perform left join (like VLOOKUP)
            merged_strategy_df = pd.merge(
                strategy_df,
                tags_df[['Name_lower', 'Id']],
                left_on='Strategy_lower',
                right_on='Name_lower',
                how='left'
            )

            # Fill NaN values with 'Not found in ISC'
            merged_strategy_df['Id'] = merged_strategy_df['Id'].fillna('Not found in ISC')

            # Drop helper columns if you don't want them in final output
            merged_strategy_df.drop(columns=['Strategy_lower', 'Name_lower'], inplace=True)

            # Save to a new Excel file
            output_file_path = 'Extracted_Files/tags_vlookup.xlsx'
            merged_strategy_df.to_excel(output_file_path, index=False)

            print(f"\n    ✅ VLOOKUP for Strategy completed.")
        except FileNotFoundError:
            print(f"\n    ⚠️ tags.csv not found — skipping tags VLOOKUP.")
        except Exception as e:
            print(f"\n    ⚠️ Error during tags VLOOKUP: {e}")

        # Load the Excel file
        excel_file_path = 'Extracted_Files/Accounts_vlookup.xlsx'  # replace with your file path
        df = pd.read_excel(excel_file_path)

        # Filter rows where Id is 'Not found in ISC'
        not_found_df = df[df['Id'] == 'Not found in ISC']

        # Select only the 'accountid' column
        accountids_not_found = not_found_df[['accountid']]

        # Save to a new Excel file
        output_file_path = os.path.expanduser("~/Downloads")+'/Accounts_to_be_imported.xlsx'
        accountids_not_found.to_excel(output_file_path, index=False)

        print(f"\n    ✅ Missing accountids saved to {output_file_path}")
        try:
            # Load the Excel file
            excel_file_path = 'Extracted_Files/tags_vlookup.xlsx'  # replace with your file path
            df = pd.read_excel(excel_file_path)

            # Filter rows where Id is 'Not found in ISC'
            not_found_df = df[df['Id'] == 'Not found in ISC']

            # Select only the 'accountid' column
            accountids_not_found = not_found_df[['Strategy']]

            # Save to a new Excel file
            output_file_path = os.path.expanduser("~/Downloads")+'/tags_Missing.xlsx'
            accountids_not_found.to_excel(output_file_path, index=False)

            print(f"\n    ✅ Missing tags saved to {output_file_path}")
        except FileNotFoundError:
            print(f"\n    ⚠️ tags_vlookup.xlsx not found — skipping missing tags export.")
        except Exception as e:
            print(f"\n    ⚠️ Error while handling tags_vlookup.xlsx: {e}")


        # Load the original Excel file
        file_path = os.path.expanduser("~/Downloads/Accounts_to_be_imported.xlsx")
        df = pd.read_excel(file_path)
        # Initialize lists to store valid and invalid values
        invalid_values = []
        valid_values = []

        if not df.empty:
            # Define a regular expression for country codes (e.g., "-US", "-KA", etc.)
            country_code_pattern = r'-[A-Za-z]{2,3}$'

            # Check each value in the column (assuming the values are in the first column)
            for value in df.iloc[:, 0]:
                value = str(value).strip()  # Ensure it's a string and remove leading/trailing spaces
                if not (value.lower().startswith('db') or value.lower().startswith('dc')):
                    # If it doesn't start with DB or DC (case insensitive), it's invalid
                    invalid_values.append(value)
                elif value.lower().startswith('db'):
                    # If it starts with DB or db, it should have a country code
                    if not re.search(country_code_pattern, value):
                        invalid_values.append(value)
                    else:
                        valid_values.append(value)  # Add to valid values list
                else:
                    valid_values.append(value)  # Add to valid values list for DC or other valid entries

        # Count invalid values
        invalid_count = len(invalid_values)

        # If there are invalid values, write them to a new Excel file
        if invalid_count > 0:
            invalid_df = pd.DataFrame(invalid_values, columns=['Invalid Accounts'])
            invalid_df.to_excel(os.path.expanduser("~/Downloads/Invalid_Accounts.xlsx"), index=False)


        # Update the original dataframe with only valid values
        valid_df = pd.DataFrame(valid_values, columns=['Accounts'])

        # Save the updated dataframe back to the original file
        valid_df.to_excel(file_path, index=False)

        print('\n   🔚 End Of Script\n') 
        break

    elif vlookup == 'no' :
        print('\n   🚫 Skipping ') 
        break

    else:
        print('\n   ❗️ Invalid Choice')
