# ===========================================
# Define the File path
# ===========================================
input_file = "Other files/input.txt"
output_file = "Other files/output.txt"

# ===========================================
# To read the content of input file
# ===========================================
with open(input_file , 'r') as file:
     a = file.readlines()

# ===========================================
# To erase all the content in the output file
# ===========================================
with open(output_file, "w") as file:
    pass  # This leaves the file empty

# ===========================================
# Main code
# ===========================================
lst= []

for i in a:
    i = i.strip()  # Remove leading/trailing whitespace including blank lines
    if not i:      # Skip empty lines
        continue
    
    if 'https:' in i:                               # To identify if it is URL or just an id
        id = i.split('/')[-2]
        print (id)
        lst.append(id)        
        with open (output_file, 'a') as file:        # Save the ids in ouput file
            file.write(f"{id}\n")

    else:
        id = i.split("\n")[0]        
        print(id)        
        lst.append(id)        
        with open (output_file, 'a') as file:          # Save the ids in ouput file
            file.write(f"{id}\n")

print("\n\n")
concatenated_values = str(tuple(lst))
Stripped_values  = concatenated_values.strip("()")
print (Stripped_values)

with open (output_file, 'a') as file:
        file.write(f"\n{Stripped_values}")
# Write tuple to a output file

import csv
import re

def create_csv_with_owner_id(txt_file_path, csv_file_path):
    # Ask for the owner ID
    owner_id = input("Please enter the owner ID: ")

    # Read the IDs from the text file
    try:
        with open(txt_file_path, 'r') as txt_file:
            ids = txt_file.readlines()
    except FileNotFoundError:
        print(f"Error: The file {txt_file_path} does not exist.")
        return
    
    # Process IDs: remove extra whitespace and filter out invalid lines
    valid_ids = []
    for id in ids:
        id = id.strip()  # Remove leading/trailing whitespace
        
        # Skip empty lines or lines containing special characters
        if not id or not re.match(r'^[a-zA-Z0-9]+$', id):  # Only alphanumeric IDs are valid
            continue
        
        valid_ids.append(id)

    # Create and write to the CSV file
    try:
        with open(csv_file_path, mode='w', newline='') as csv_file:
            writer = csv.writer(csv_file)
            # Write the header
            writer.writerow(['ID', 'OwnerID'])
            
            # Write the data
            for id in valid_ids:
                writer.writerow([id, owner_id])

        print(f"CSV file created successfully at {csv_file_path}")

    except Exception as e:
        print(f"Error writing to CSV file: {e}")

# Example usage
txt_file_path = "Other files/output.txt"
output = input('What should be the file name?')
csv_file_path = '/Users/avirajmore/Downloads/'+output+'.csv'

create_csv_with_owner_id(txt_file_path, csv_file_path)
