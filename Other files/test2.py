
import pandas as pd
import tkinter as tk
from tkinter import ttk
from pandastable import Table


file_path = "/Users/avirajmore/Downloads/0913_SecuringAI Massload.xlsx"
sheet_name = "Opportunity"  # Replace with your sheet name

def is_sheet_empty(file_path, sheet_name):
    """
    Checks if a given sheet in an Excel file contains any data beyond headers.
    Returns:
        - True, None if the sheet is empty or has only headers.
        - False, DataFrame (first 4 rows) if the sheet contains data.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Check if the sheet is empty or only contains headers
        if df.empty or df.dropna(how='all').shape[0] == 0:
            return True, None  # Sheet is empty or has only headers
        
        return False, df.head(4)  # Sheet contains data, return first 4 rows
    except Exception as e:
        print(f"Error reading sheet '{sheet_name}': {e}")
        return None, None

def show_data_in_ui(dataframe):
    """Displays the DataFrame in a GUI window."""
    root = tk.Tk()
    root.title("Excel Sheet Data Preview")
    
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True)
    
    table = Table(frame, dataframe=dataframe)
    table.show()
    
    root.mainloop()


is_empty, preview = is_sheet_empty(file_path, sheet_name)

if is_empty:
    print(f"\n:open_file_folder: The sheet '{sheet_name}' is empty or contains only headers.\n")
elif is_empty is None:
    print("\n:warning: Could not process the sheet due to an error.\n")
else:
    print(f"\n:white_check_mark: The sheet '{sheet_name}' contains data. Opening UI for preview...\n")
    show_data_in_ui(preview)