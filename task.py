import tkinter as tk
from tkinter import ttk
import sqlite3
from tkcalendar import Calendar
from datetime import datetime

# Initialize the main window
root = tk.Tk()
root.title("Matter Management System")
root.geometry("500x500")

# Connect to SQLite database
conn = sqlite3.connect('matters.db')
cursor = conn.cursor()
cursor.execute('''CREATE TABLE IF NOT EXISTS matters (
                    id INTEGER PRIMARY KEY,
                    matter_no TEXT,
                    serial_no TEXT,
                    matter_name TEXT,
                    start_date TEXT,
                    due_date TEXT,
                    reminder INTEGER)''')

# Function to submit data
def submit():
    matter_no = entry_matter_no.get()
    serial_no = entry_serial_no.get()
    matter_name = entry_matter_name.get()
    start_date = datetime.now().strftime("%Y-%m-%d")
    due_date = cal.get_date()
    reminder = reminder_var.get()
    
    # Insert into database
    cursor.execute("INSERT INTO matters (matter_no, serial_no, matter_name, start_date, due_date, reminder) VALUES (?, ?, ?, ?, ?, ?)",
                   (matter_no, serial_no, matter_name, start_date, due_date, reminder))
    conn.commit()

    # Clear fields and update display
    entry_matter_no.delete(0, tk.END)
    entry_serial_no.delete(0, tk.END)
    entry_matter_name.delete(0, tk.END)
    update_display()

# Function to update display
def update_display():
    for widget in display_frame.winfo_children():
        widget.destroy()
    cursor.execute("SELECT * FROM matters")
    rows = cursor.fetchall()
    for row in rows:
        tk.Label(display_frame, text=row).pack()

# Input fields
matter_label=tk.Label(root, text="Matter No.").pack()
entry_matter_no = tk.Entry(root).pack()
serial_label=tk.Label(root, text="Serial No.").pack()
entry_serial_no = tk.Entry(root).pack()
label_name=tk.Label(root, text="Matter Name").pack()
entry_matter_name = tk.Entry(root).pack()
start_date_label=tk.Label(root, text="start_date").pack()
start_date = tk.Entry(root).pack()
due_date_label=tk.Label(root, text="due_date").pack()
due_date = tk.Entry(root).pack()

# Calendar for due date
cal = Calendar(root, selectmode='day', year=2025, month=8, day=5)
cal.pack()

# Reminder checkbox
reminder_var = tk.IntVar()
reminder_check = tk.Checkbutton(root, text="Reminder", variable=reminder_var)
reminder_check.pack()

# Submit button
submit_button = tk.Button(root, text="Submit", command=submit)
submit_button.pack()

# Display Frame
display_frame = tk.Frame(root)
display_frame.pack()

update_display()

root.mainloop()