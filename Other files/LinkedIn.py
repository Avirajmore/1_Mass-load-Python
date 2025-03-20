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
    else:
        id = i.split("\n")[0]        
        print(id)        
        lst.append(id)        
        with open (output_file, 'a') as file:          # Save the ids in ouput file
            file.write(f"{id},\n")

print("\n\n")
concatenated_values = str(tuple(lst))
Stripped_values  = concatenated_values.strip("()")
print (Stripped_values)

with open (output_file, 'a') as file:
        file.write(f"\n{Stripped_values}")
# Write tuple to a output file
