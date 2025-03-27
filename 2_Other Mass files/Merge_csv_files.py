import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

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

file1 = select_file("Select the First CSV file")
file2 = select_file("Select the Second CSV file")
output_file = os.path.expanduser("~/Downloads/Merged_file.csv")

def merge_csv_files(file1, file2, output_file):
    # Read both CSV files
    df1 = pd.read_csv(file1)
    df2 = pd.read_csv(file2)

    # Concatenate DataFrames while keeping the header only once
    merged_df = pd.concat([df1, df2], ignore_index=True)

    # Save the merged DataFrame to a new CSV file
    merged_df.to_csv(output_file, index=False)

    print(f"Files merged successfully into {output_file}")

# Usage
merge_csv_files(file1, file2,output_file)
