import openpyxl
import re

# Define the paths to the text and Excel files
text_file_path = "PyWhatKit_DB.txt"
excel_file_path = "Sheet.xlsx"

# Load the Excel workbook
wb = openpyxl.load_workbook(excel_file_path)
sheet = wb.active

# Function to extract phone numbers from text file
def extract_phone_numbers(text):
    phone_numbers = re.findall(r'Phone Number: \+(\d+)', text)
    return phone_numbers

# Read the text file and extract phone numbers
with open(text_file_path, 'r') as text_file:
    text_data = text_file.read()

phone_numbers_to_delete = extract_phone_numbers(text_data)

# Iterate through the Excel file and remove matching phone numbers
for row in sheet.iter_rows(min_row=2, max_col=6):
    cell_value = row[5].value  # Assuming Phone Number is in column F (index 5)
    if cell_value is not None and str(cell_value) in phone_numbers_to_delete:
        row[5].value = None  # Set the cell value to None to delete the number

# Save the modified Excel file
wb.save(excel_file_path)

print("Phone numbers deleted from the Excel file.")

