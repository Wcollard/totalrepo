import os
import sqlite3
from tkinter import Tk, Label, Entry, Button, StringVar, Listbox, END
from tkcalendar import DateEntry
from datetime import datetime

# Initialize the database
conn = sqlite3.connect('client_data.db')
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS clients
             (filename text, date_due text, matter_number text)''')
conn.commit()

def create_folder():
    filename = filename_var.get()
    date_due = date_due_var.get()
    matter_number = matter_number_var.get()

    # Store the data in the database
    c.execute("INSERT INTO clients VALUES (?, ?, ?)", (filename, date_due, matter_number))
    conn.commit()
    
    # Create the directory
    base_dir = os.path.expanduser("~/Documents/clients")
    folder_name = f"{filename}_{matter_number}_{date_due}"
    folder_path = os.path.join(base_dir, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    print(f"Folder created at: {folder_path} 📂")

    # Refresh the folder list display
    update_folder_list()

def update_folder_list():
    # Clear the current list
    folder_listbox.delete(0, END)
    
    # Fetch and sort data by date
    c.execute("SELECT * FROM clients ORDER BY date_due ASC")
    rows = c.fetchall()
    
    for row in rows:
        folder_name = f"{row[0]}_{row[2]}_{row[1]}"
        folder_listbox.insert(END, folder_name)

# Set up the GUI
root = Tk()
root.title("Client Folder Creator")

filename_var = StringVar()
date_due_var = StringVar()
matter_number_var = StringVar()

Label(root, text="Filename").grid(row=0, column=0)
Entry(root, textvariable=filename_var).grid(row=0, column=1)

Label(root, text="Matter Number").grid(row=1, column=0)
Entry(root, textvariable=matter_number_var).grid(row=1, column=1)

Label(root, text="Date Due").grid(row=2, column=0)
Entry(root, textvariable=date_due_var).grid(row=2, column=1)

def update_date_entry(*args):
    date_due_var.set(cal.get_date())

cal = DateEntry(root, selectmode='day', year=datetime.now().year, month=datetime.now().month, day=datetime.now().day)
cal.grid(row=3, column=1)
cal.bind("<<DateEntrySelected>>", update_date_entry)

Button(root, text="Create Folder", command=create_folder).grid(row=4, column=0, columnspan=2)

# Listbox for displaying folders
Label(root, text="Folders Sorted by Date Due").grid(row=5, column=0, columnspan=2)
folder_listbox = Listbox(root, width=50)
folder_listbox.grid(row=6, column=0, columnspan=2)

# Initial population of the folder list
update_folder_list()

root.mainloop()