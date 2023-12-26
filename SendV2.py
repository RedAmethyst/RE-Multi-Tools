import pandas as pd
import pywhatkit as pwk
import re
import time
import pytz
from datetime import datetime

# Function to format phone numbers by adding '+' sign if missing
def format_phone_number(number):
    return '+' + number if not number.startswith('+') else number

# Function to extract names using regular expressions
def extract_name(name):
    match = re.findall(r'\b[A-Z][a-z]*\b', name)
    return ' '.join(match[:2]) if len(match) >= 2 else name.split()[0]

# Function to personalize greetings based on time of day
def time_based_greeting():
    tz = pytz.timezone('Asia/Dubai')
    current_hour = datetime.now(tz).hour
    if current_hour < 12:
        return "Good morning"
    elif current_hour < 18:
        return "Good afternoon"
    else:
        return "Good evening"

# Function to calculate months left until the end date
def months_left(end_date):
    today = datetime.today()
    return (end_date.year - today.year) * 12 + end_date.month - today.month

# Function to load data from Excel file
def load_data(file_path):
    return pd.read_excel(file_path)

# Function to generate phone number and message pairs
def generate_phone_number_message_pairs(data):
    phone_number_message_pairs = []

    for _, row in data.iterrows():
        phone_number = format_phone_number(str(row['Phone Number']))
        name = extract_name(row['Name'])
        consultant_name = "Abdelhalim"  # Replace with your name
        location = row['Location']
        unit_number = row['Unit Number']
        sub_location = row['Sub Location']
        rental_end = row['End Date']

        # Check for NaT in End Date and convert to datetime if necessary
        if pd.isna(rental_end):
            end_date_formatted = "an unknown date"
            months_to_end = None
        else:
            if isinstance(rental_end, str):
                try:
                    rental_end = datetime.strptime(rental_end, "%Y-%m-%d")  # Adjust the format as per your data
                except ValueError:
                    print(f"Invalid date format for {rental_end}")
                    continue  # Skip this row or handle it as you see fit
            end_date_formatted = rental_end.strftime("%d %B %Y")
            months_to_end = months_left(rental_end)

        # Handle NaN or missing values in 'Feedback' column
        category = str(row['Feedback']).lower() if pd.notna(row['Feedback']) else "unknown"

        message = "{greeting} " + name + ", I am " + consultant_name + ", property consultant at " + location + ". "

        if category == "rented":
            if rental_end < datetime.now():
                # New greeting for past end date
                message += "Your rent contract for the property " + str(unit_number) + " has ended in " + end_date_formatted + ". Did you renew or you are still looking for a tenant?"
            elif months_to_end is not None:
                if months_to_end <= 3:
                    message += "Your rental term for property " + str(unit_number) + " in " + sub_location + " is nearing its end on " + end_date_formatted + ". Are you considering renewing the lease, or would you like assistance in finding a new tenant?"
                elif 3 < months_to_end <= 6:
                    message += "As the rental term for your property " + str(unit_number) + " in " + sub_location + " is due on " + end_date_formatted + ", have you thought about renewing the lease with a rent adjustment to match the market rate, or perhaps considering selling?"
                elif 6 < months_to_end <= 12:
                    message += "With the rental term for your property " + str(unit_number) + " in " + sub_location + " ending on " + end_date_formatted + ", are you considering selling it? If so, have you issued a notice of vacating to your tenant? I can assist with these processes."
        elif category == "sold":
            message += "I hope this message finds you well. I would like to discuss some new opportunities with you."
        elif category == "unknown":
            message += "I am reaching out regarding your property " + str(unit_number) + " in " + sub_location + ". What are your current plans for it, is it for Rent or Sale?"

        phone_number_message_pairs.append((phone_number, message))

    return phone_number_message_pairs

# Function to create a new Excel file with a timestamp
def create_new_excel_file(original_file):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    new_file_name = f"{original_file.split('.')[0]}_{timestamp}.xlsx"
    return new_file_name

# Functions for updating progress
def update_progress(file_path, phone_number):
    with open(file_path, 'a') as file:
        file.write(phone_number + '\n')

def read_processed_numbers(file_path):
    try:
        with open(file_path, 'r') as file:
            return set(file.read().splitlines())
    except FileNotFoundError:
        return set()

# Function to update Excel files after sending a message
def update_excel_files(original_file, new_file, row_index):
    df = pd.read_excel(original_file)
    new_df = pd.read_excel(new_file)

    # Copy the row to the new file
    new_row = df.iloc[[row_index]]
    new_df = pd.concat([new_df, new_row], ignore_index=True)

    # Delete the row from the original file
    df = df.drop(df.index[row_index])

    # Save both files
    new_df.to_excel(new_file, index=False)
    df.to_excel(original_file, index=False)

# Function to read the list of sent phone numbers from PyWhatKit_DB.txt
def read_existing_numbers(file_path):
    existing_numbers = set()
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            phone_numbers = re.findall(r'Phone Number: \+(\d+)', content)
            existing_numbers.update(phone_numbers)
    except FileNotFoundError:
        pass  # The file doesn't exist, which is okay for the first run
    return existing_numbers

# Function to send messages to phone numbers in batches
def send_messages_in_batches(phone_number_message_pairs, original_file, output_file, progress_file):
    for index, (phone_number, message) in enumerate(phone_number_message_pairs):
        if phone_number in read_processed_numbers(progress_file):
            continue

        greeting = time_based_greeting()
        message = message.format(greeting=greeting)

        try:
            # Send the WhatsApp message
            pwk.sendwhatmsg_instantly(phone_number, message, 10, True, 5)
            print(f"Message sent to {phone_number} successfully.")
            
            # Update Excel files and progress
            update_excel_files(original_file, output_file, index)
            update_progress(progress_file, phone_number)

        except Exception as e:
            print(f"Error sending message to {phone_number}: {e}")

        time.sleep(10)

# Main execution
if __name__ == "__main__":
    input_file = 'Sheet.xlsx'
    output_file = create_new_excel_file(input_file)
    progress_file = 'progress.txt'

    # Initialize the new Excel file
    pd.DataFrame().to_excel(output_file, index=False)

    # Load data and filter out processed numbers
    data = load_data(input_file)
    processed_numbers = read_processed_numbers(progress_file)
    data = data[~data['Phone Number'].isin(processed_numbers)]

    # Generate phone number and message pairs
    phone_number_message_pairs = generate_phone_number_message_pairs(data)

    # Send messages in batches
    send_messages_in_batches(phone_number_message_pairs, input_file, output_file, progress_file)