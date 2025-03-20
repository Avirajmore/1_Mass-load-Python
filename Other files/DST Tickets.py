# ===========================================
# Define the File path
# ===========================================
input_file = "1_Mass load Python/Other files/input.txt"
output_file = "1_Mass load Python/Other files/output.txt"

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
    
    if i:                               
        print (i)
        lst.append(id)        
        with open (output_file, 'a') as file:        # Save the ids in ouput file
            file.write(f"{i}\n")