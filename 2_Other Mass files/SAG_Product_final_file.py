'''
Use the below code to process and create final file of only Product sheet of the sag file
'''

import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import messagebox, filedialog


# ======================================================================
# Select the file to process
# ======================================================================

print("\n\n🔍 Step 1: Select the file to Process")

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
    filename = str(file_path.split('/')[-1])
    print(f"\n    ✅ File selected: '{filename}'.")
else:
    print("\n    ❌ No file selected. Exiting the program. ❌")
    sys.exit()

# ======================================================================
# Main code to create the final file
# ======================================================================

print("\n\n🔍 CREATING PRODUCT FILE")
# Step 1: Define the input Excel file path (commented out as requested)
# file_path = 'your_file_path_here.xlsx'  # Replace with your actual file path
sheet_name = 'Opportunity_product'

# Define the predefined columns to delete
predefined_columns_product = [
    'product', 'product_type', 'Product_Family__c', 
    'opportunity currency', 'practise_multiple country', 
    'quantity.1', 'concatenated product family', 'concatenated currency'
]

# Initialize lists to track deleted columns and dropped rows
deleted_columns = []  # To store names of columns deleted
rows_dropped = 0  # To count total rows dropped
rows_dropped_existing_false = 0  # Track rows where 'existing' == False

# Initialize a DataFrame to store removed rows where 'existing' == False
removed_rows_df = pd.DataFrame()

# Try to read data from the "Opportunity_product" sheet into a DataFrame
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # # Step 2: Remove all rows where 'existing' == False and store them in removed_rows_df
    # initial_row_count = len(df)
    # removed_rows_df = df[df['existing'] == False].copy()
    # rows_dropped_existing_false = len(removed_rows_df)  # Track the number of rows removed
    # df = df[df['existing'] == True]

    # Print the count of rows removed where 'existing' == False
    # print(f"\n    ❗️ Rows removed where 'existing' == False: {rows_dropped_existing_false}")

    # Add a "Reason" column to the removed rows to specify why they were removed
    # removed_rows_df['Reason'] = "Opportunity Missing From Main sheet"

    # Step 3: Remove predefined columns from both the main DataFrame and removed rows DataFrame
    columns_to_delete_predefined = [col for col in predefined_columns_product if col in df.columns]
    if columns_to_delete_predefined:
        df.drop(columns=columns_to_delete_predefined, inplace=True)
        removed_rows_df.drop(columns=columns_to_delete_predefined, inplace=True, errors='ignore')
        deleted_columns.extend(columns_to_delete_predefined)

    # Step 4: Set up a graphical interface (GUI) to select columns to delete
    root = Tk()
    root.title("Select Columns to Delete")

    # Set window size and make it fixed
    root.geometry("500x600")
    root.resizable(False, False)

    # Scrollbar setup
    canvas = Canvas(root)
    scrollbar = Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    frame = Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")

    # Dictionary to hold IntVar for each checkbox (for column selection)
    checkboxes = {}

    # Step 5: Add checkboxes for each column in the DataFrame
    for column in df.columns:
        var = IntVar()
        checkboxes[column] = var
        checkbutton = Checkbutton(frame, text=column, variable=var, font=('Helvetica', 12), anchor="w", padx=10)
        checkbutton.pack(anchor="w", pady=5)

    # Create a frame for the submit button and place it at the top right
    button_frame = Frame(root)
    submit_button = Button(button_frame, text="Submit", command=root.quit, 
                        font=('Helvetica', 12, 'bold'), relief='flat', padx=20, pady=10)
    submit_button.pack(side="right")

    # Place the button frame in the grid to ensure it stays at the top right
    button_frame.pack(anchor="ne", padx=20, pady=10)  # 'ne' positions it top-right

    # Update the scroll region to fit all elements
    frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    # Run the Tkinter main loop
    root.mainloop()
    root.destroy()

    # Step 6: Process the selected columns to delete after user submits
    columns_to_delete_from_user = [col for col, var in checkboxes.items() if var.get() == 1]

    if columns_to_delete_from_user:
        # Remove the selected columns from the main DataFrame (df)
        df.drop(columns=columns_to_delete_from_user, inplace=True)
        # Also remove the same columns from the removed rows DataFrame (removed_rows_df)
        removed_rows_df.drop(columns=columns_to_delete_from_user, inplace=True, errors='ignore')
        deleted_columns.extend(columns_to_delete_from_user)
        print("\n    ✅ Additional columns deleted:")
        for col in columns_to_delete_from_user:
            print(f"\n        🔸 {col}")
    else:
        print("\n    ✅ No additional columns selected for deletion.")


    # Step 7: Remove any rows that contain only blank values in the main DataFrame
    df.dropna(axis=0, how='all', inplace=True)

    # Step 8: Remove any columns that contain only blank values in the main DataFrame
    df.dropna(axis=1, how='all', inplace=True)

    # Step 9: Remove any duplicate rows based on all columns in the main DataFrame
    # df.drop_duplicates(inplace=True)

    # Step 10: Define the output CSV file path (updated with new values)
    # output_file = output + "/" + opportunity_product  # Path for the processed CSV
    output_file = os.path.expanduser("~/Downloads/SAG Product Load file.csv")  # Path for the processed CSV

    # Step 11: Check if the directory exists before saving
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)  # Create the directory if it doesn't exist

    # Step 12: Save the processed DataFrame to the specified CSV file
    df.to_csv(output_file, index=False)
    print("\n    ✅ Processed data saved to")
    print(f"\n        📂 {"/".join(output_file.split("/")[-5:])}")

except Exception as e:
    print(f"\n    ❌ An error occurred: {e}")
