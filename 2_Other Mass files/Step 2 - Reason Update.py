
'''
Step 2:- To add a Reason column after the masss load is done
'''


import pandas as pd
import os

# Specify the file path of your CSV file
file_path = os.path.expanduser("~/Downloads/reason.csv")  # Replace with the path to your CSV file

# Load the CSV file into a pandas DataFrame
try:
    df = pd.read_csv(file_path)
except FileNotFoundError:
    print(f"Error: The file '{file_path}' was not found.")
    exit()

# Check if the column 'Opportunity.Reason__c' exists
if "Opportunity.Reason__c" in df.columns:
    # Fill NaN values with an empty string, then append "ADDLIC"
    df["Opportunity.Reason__c"] = df["Opportunity.Reason__c"].fillna("").astype(str) + "ADDLIC"

    # Save the modified DataFrame back to the same CSV file
    df.to_csv(file_path, index=False)
    print(f"File '{file_path}' updated successfully.")
else:
    print("Error: The column 'Opportunity.Reason__c' does not exist in the file.")
