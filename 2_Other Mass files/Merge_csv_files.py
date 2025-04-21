'''
Description: Code to Merge Multiple CSV Files into a Single File

Functionality Overview:

1) Prompt the user to select multiple CSV files to be merged.
2) Ask the user to provide a name for the merged CSV file.
3) Combine all selected CSV files and save the merged file to the Downloads folder.
'''

import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

# Function to select multiple CSV files
def select_files(title):
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title=title, filetypes=[("CSV files", "*.csv")])

    if file_paths:
        return file_paths
    else:
        print("\n❌ No files selected.")
        return None

# Function to merge multiple CSV files
def merge_csv_files(file_paths, output_file):
    merged_df = pd.DataFrame()  # Empty DataFrame to start with

    for file_path in file_paths:
        df = pd.read_csv(file_path)
        merged_df = pd.concat([merged_df, df], ignore_index=True)

    merged_df.to_csv(output_file, index=False)
    print(f"\n   ✅ Files merged successfully into: {output_file}")

print("\n🔍 Select the csv files to merge")
# Main flow
file_paths = select_files("Select CSV files to merge")

print("\n🔍 Name of new Csv File")
if file_paths:
    output_name = input("\n📄 Enter the name for the merged CSV file (without .csv extension): ").strip()
    if not output_name:
        output_name = "Merged_file"  # Default name if left blank

    output_file = os.path.expanduser(f"~/Downloads/{output_name}.csv")

    merge_csv_files(file_paths, output_file)
else:
    print("\n   ⚠️ No files selected. Exiting.")
