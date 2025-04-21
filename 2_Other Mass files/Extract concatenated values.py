'''
Description: Code to Extract Concatenated Values from a Column

Functionality Overview:

1) Prompt the user to select an Excel file.
2) Display the list of available sheets in the selected file and ask the user to choose one.
3) Display the column names from the selected sheet and prompt the user to pick a column.
4) Extract all values from the selected column and save them to a .txt file.
'''

import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def process_file(file_path):
    try:
        filename = os.path.basename(file_path)
        print(f"\n    ✅ File selected: '{filename}'.")
    except Exception as e:
        print(f"Error selecting file: {e}")
        return

    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        if file_extension == ".csv":
            df = pd.read_csv(file_path)
            sheet_name = "CSV File"
        elif file_extension in [".xlsx", ".xls"]:
            xls = pd.ExcelFile(file_path)
            print("\nAvailable sheets in the file:")
            for idx, sheet in enumerate(xls.sheet_names):
                print(f"    {idx + 1}. {sheet}")
            
            sheet_index = int(input("\nSelect a sheet number: ")) - 1
            sheet_name = xls.sheet_names[sheet_index]
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            print("Unsupported file format!")
            return
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    try:
        print("\nAvailable columns in the sheet:")
        for idx, col in enumerate(df.columns):
            print(f"    {idx + 1}. {col}")
        column_indices = input("\nEnter column numbers separated by commas: ")
        selected_columns = [df.columns[int(i) - 1] for i in column_indices.split(",")]
    except Exception as e:
        print(f"Error selecting columns: {e}")
        return
    
    output_file = "transformed_data.txt"
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            for col in selected_columns:
                f.write(f"\n    {col}:\n")
                transformed_values = df[col].astype(str).str.strip().unique()
                formatted_values = [f"'{val}'," for val in transformed_values]
                f.write("\n".join(formatted_values) + "\n\n")
        print(f"\nTransformed data saved successfully as '{output_file}'")
    except Exception as e:
        print(f"Error writing to file: {e}")
    
    try:
        with open(output_file, 'r', encoding='utf-8') as infile:
            lines = [line.rstrip() for line in infile if line.strip()]
        with open(output_file, 'w', encoding='utf-8') as outfile:
            outfile.write("\n".join(lines))
    except Exception as e:
        print(f"Error cleaning output file: {e}")
    
    def remove_last_char_from_last_line(extract_file):
        try:
            with open(extract_file, 'r') as file:
                lines = file.readlines()
            if lines:
                lines[-1] = lines[-1][:-1]
                with open(extract_file, 'w') as file:
                    file.writelines(lines)
        except Exception as e:
            print(f"Error modifying last line: {e}")
    
    try:
        remove_last_char_from_last_line(output_file)
    except Exception as e:
        print(f"Error in last character removal: {e}")
    
    try:
        with open(output_file, 'r', encoding='utf-8') as infile:
            lines = infile.readlines()[1:]
        with open(output_file, 'w', encoding='utf-8') as outfile:
            outfile.writelines(lines)
    except Exception as e:
        print(f"Error trimming first line: {e}")

print("\n\n🔍 Step 1: Select a File to Process")

# Set the directory to search for Excel files
directory = os.path.expanduser("~/Downloads")

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
    process_file(file_path)
else:
    print("\n    ❌ No file selected. Exiting the program. ❌")
    sys.exit()
