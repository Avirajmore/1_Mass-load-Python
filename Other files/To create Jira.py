import pandas as pd

# Function to process the text file and create an Excel file
def create_excel_from_txt(input_file, output_file):
    data = []

    # Read the input text file
    with open(input_file, 'r') as file:
        for line in file:
            # Ignore empty lines
            if line.strip():
                # Split the line by tab
                parts = line.strip().split('\t')
                if len(parts) == 2:
                    ticket_number, request = parts
                    data.append([ticket_number, request, "User Ticket"])

    # Create a DataFrame
    df = pd.DataFrame(data, columns=['Ticket Number', 'Request', 'Source'])

    # Write DataFrame to Excel
    df.to_excel(output_file, index=False)

    print(f"Excel file created successfully: {output_file}")

# File paths (update these as needed)
input_file = "Other files/input.txt"  # Input text file
output_file = '/Users/avirajmore/Downloads/jira.xlsx'  # Output Excel file

create_excel_from_txt(input_file, output_file)