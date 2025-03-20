
import pandas as pd
import os
from tabulate import tabulate

folder_path = '/Users/avirajmore/Downloads'

file_paths = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith(('.xlsx', '.xls'))]
if not file_paths:
    print("No Excel files found in the folder.")

for file_path in file_paths:
    print(f"\n\n")
    print("=" * 100)
    print(f"\n📂 {file_path} 📂\n")
    print("=" * 100)
    def check_all_sheets(file_path):
        """
        Checks all sheets in an Excel file to determine if they contain any data beyond headers.
        """
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # Check if the sheet is empty or only contains headers
                if df.empty or df.dropna(how='all').shape[0] == 0:
                    print(f"\n❌ The sheet '{sheet_name}' is empty or contains only headers.\n")
                    print(tabulate(df.iloc[:4, :4], headers='keys', tablefmt='fancy_grid', showindex=False))
                else:
                    print(f"\n✅ The sheet '{sheet_name}' contains data. Here are the first 4 rows and 4 columns:\n")
                    print(tabulate(df.iloc[:4, :4], headers='keys', tablefmt='fancy_grid', showindex=False))
        except Exception as e:
            print(f"\n⚠️ Error processing the Excel file: {e}\n")

    check_all_sheets(file_path)
