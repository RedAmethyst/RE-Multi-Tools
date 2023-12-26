import pandas as pd
from datetime import datetime
from plyer import notification
import tkinter as tk
from tkinter import filedialog, Label, Entry, Button
from tkinter.messagebox import showinfo

def notify_from_excel(file_path, start_date, end_date):
    df = pd.read_excel(file_path)
    df['End Date'] = pd.to_datetime(df['End Date'])
    filtered_df = df[(df['End Date'] >= start_date) & (df['End Date'] <= end_date)]

    for _, row in filtered_df.iterrows():
        # Truncate message to 256 characters
        message = row['Messages']
        if len(message) > 256:
            message = message[:253] + '...'

        notification.notify(
            title='Real Estate Notification',
            message=message,
            timeout=10
        )

def select_file():
    file_path = filedialog.askopenfilename()
    file_label.config(text=file_path)
    return file_path

def submit():
    file_path = file_label.cget("text")
    start_date = datetime.strptime(start_date_entry.get(), '%Y-%m-%d')
    end_date = datetime.strptime(end_date_entry.get(), '%Y-%m-%d')
    notify_from_excel(file_path, start_date, end_date)
    showinfo("Success", "Notifications sent!")

# Create the main window
window = tk.Tk()
window.title("Real Estate Notification System")

# Create and place widgets
select_file_button = Button(window, text="Select Excel File", command=select_file)
select_file_button.pack()

file_label = Label(window, text="No file selected")
file_label.pack()

Label(window, text="Start Date (YYYY-MM-DD):").pack()
start_date_entry = Entry(window)
start_date_entry.pack()

Label(window, text="End Date (YYYY-MM-DD):").pack()
end_date_entry = Entry(window)
end_date_entry.pack()

submit_button = Button(window, text="Submit", command=submit)
submit_button.pack()

# Run the event loop
window.mainloop()
