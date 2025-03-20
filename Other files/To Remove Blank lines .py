input_file = "Other files/input.txt"
# Read all non-empty lines
with open(input_file, 'r', encoding='utf-8') as infile:
    lines = [line.rstrip() for line in infile if line.strip()]  # Strip trailing newlines

# Write back to the same file, ensuring proper newlines
with open(input_file, 'w', encoding='utf-8') as outfile:
    outfile.write("\n".join(lines))  # Join lines with newlines, ensuring no extra blank lines

print(f"Blank lines removed. Changes saved in {input_file}.")

