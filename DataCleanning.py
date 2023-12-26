import pandas as pd
import os
from datetime import datetime

# Function to clean phone numbers
def clean_phone_number(number):
    # Remove non-numeric characters
    number = ''.join(filter(str.isdigit, str(number)))

    # Remove leading '0's or '00's
    while number.startswith(('0', '00')):
        number = number[1:]

    # UAE number cleaning
    if number.startswith(('971', '00971', '05', '5')):
        number = '9715' + number[-8:]

    return number

# Replace 'Sheet.xlsx' with your actual file path
file_path = 'Sheet.xlsx'

# Read the Excel file
excel_data = pd.read_excel(file_path, sheet_name=None)

# Data Cleaning and Formatting
for sheet_name, data in excel_data.items():
    # Cleaning operations for 'Date', 'Start Date', and 'Phone Number' columns
    date_columns = ['Date', 'Start Date']
    phone_column = 'Phone Number'

    # Convert specified columns to datetime with date only and fill missing values with empty string
    for col in date_columns:
        if col in data.columns:
            data[col] = pd.to_datetime(data[col], errors='coerce').dt.strftime('%Y-%m-%d')
            data[col].fillna('', inplace=True)

    # Clean 'Phone Number' column
    if phone_column in data.columns:
        data[phone_column] = data[phone_column].apply(clean_phone_number)

    # Modify 'Location' column to create the output file name
    output_file_name = f"{data['Location'].iloc[0]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    # Save the original data to 'OldSheet.xlsx' with location and datetime in the filename
    output_directory = 'OutputFiles'
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    old_sheet_path = f'{output_directory}/{output_file_name}'
    data.to_excel(old_sheet_path, sheet_name=sheet_name, index=False)

    # Overwrite the original 'Sheet.xlsx' file
    if sheet_name == 'Sheet1':  # Adjust if you have specific sheets to overwrite
        data.to_excel(file_path, sheet_name=sheet_name, index=False)

    # Print the first few rows of each sheet to verify changes (for demonstration purposes)
    print(f"Sheet Name: {sheet_name}")
    print(data.head())
