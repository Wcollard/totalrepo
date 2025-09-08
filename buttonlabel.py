import tkinter as tk
import sqlite3

def save_data():
    matter_no = matter_entry.get()
    serial_no = serial_entry.get()
    # Insert into database
    # ...

def search_data():
    # Search database logic
    # ...

# Create database connection
# ...

root = tk.Tk()
root.title("Entry Manager")

tk.Label(root, text="Matter No.").grid(row=0, column=0)
matter_entry = tk.Entry(root, width=50)
matter_entry.grid(row=0, column=1)

tk.Label(root, text="Serial No.").grid(row=1, column=0)
serial_entry = tk.Entry(root, width=50)
serial_entry.grid(row=1, column=1)

save_button = tk.Button(root, text="Save", command=save_data)
save_button.grid(row=2, column=0)

search_button = tk.Button(root, text="Search", command=search_data)
search_button.grid(row=2, column=1)

root.mainloop()