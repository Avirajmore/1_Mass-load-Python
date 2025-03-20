import sys
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import pandas as pd
from openpyxl import load_workbook

print("=" * 100)
print(" " * 33 + "📂 Vlookup Opearation  📂")
print("=" * 100)

print("\n\n🔍 Step 1: Select an Excel File to Process")
directory = "/Users/avirajmore/Downloads"

# Create a hidden root window (used for file dialog)
root = tk.Tk()
root.withdraw()

# Ask the user to select a file (Excel or CSV)
excel_path = filedialog.askopenfilename(
    initialdir=directory,
    title="📄 Select an Excel file",
    filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
)

# Print the selected file path
if excel_path:
    filename = str(excel_path.split('/')[-1])
    print(f"\n    ✅ Excel file selected: '{filename}'.")
else:
    print("\n    ❌ No file selected. Exiting the program. ❌")
    sys.exit()

print("\n\n🔍 Step 2: Select a CSV File to Process")
# Create a hidden root window (used for file dialog)
root = tk.Tk()
root.withdraw()

# Ask the user to select a file (Excel or CSV)
csv_path = filedialog.askopenfilename(
    initialdir=directory,
    title="📄 Select a CSV file",
    filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
)

# Print the selected file path
if csv_path:
    filename = str(csv_path.split('/')[-1])
    print(f"\n    ✅ CSV file selected: '{filename}'.")
else:
    print("\n    ❌ No file selected. Exiting the program. ❌")
    sys.exit()


print("\n\n🔍 Step 3: Select a column from the sheet as a Lookup Key")
# Load Excel File
# excel_path = input("Enter the path of the Excel file: ")
xls = pd.ExcelFile(excel_path)
print("\n    📋 Available sheets:")
for i, sheet in enumerate(xls.sheet_names, 1):
    print(f"\n       {i}: {sheet}")
sheet_index = int(input("\n    👉 Enter the number corresponding to the sheet: ")) - 1
sheet_name = xls.sheet_names[sheet_index]

df_excel = xls.parse(sheet_name)
print("\n    🔸 Columns in selected sheet:")
for i, col in enumerate(df_excel.columns, 1):
    print(f"\n       {i}: {col}")
lookup_index = int(input("\n    👉 Enter the number corresponding to the lookup column: ")) - 1
lookup_column = df_excel.columns[lookup_index]


print("\n\n🔍 Step 4: Select a column from the CSV ")
# Load CSV File
# csv_path = input("Enter the path of the CSV file: ")
df_csv = pd.read_csv(csv_path)
print("\n    🔸 Columns in CSV file:")
for i, col in enumerate(df_csv.columns, 1):
    print(f"\n       {i}: {col}")
csv_lookup_index = int(input("\n    👉 Enter the number corresponding to the lookup column in CSV: ")) - 1
csv_lookup_column = df_csv.columns[csv_lookup_index]
csv_result_index = int(input("\n    👉 Enter the number corresponding to the result column in CSV: ")) - 1
csv_result_column = df_csv.columns[csv_result_index]

# Normalize case for VLOOKUP
df_excel[lookup_column] = df_excel[lookup_column].astype(str).str.lower()
df_csv[csv_lookup_column] = df_csv[csv_lookup_column].astype(str).str.lower()

# Perform VLOOKUP
lookup_dict = df_csv.set_index(csv_lookup_column)[csv_result_column].to_dict()
df_excel[f"{csv_result_column}_Lookup"] = df_excel[lookup_column].map(lambda x: lookup_dict.get(x, f"Not found - {x}"))

# Save results back to the same Excel file and sheet
book = load_workbook(excel_path)
with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_excel.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\n    ✅ Updated Excel file saved in the same sheet: {sheet_name}")
