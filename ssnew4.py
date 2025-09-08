import tkinter as tk
from tkinter import messagebox, filedialog
import requests
import pandas as pd
from datetime import datetime
import os

def extract_names(data_list):
    """Helper function to extract names from list of dictionaries"""
    if not data_list:
        return ""
    if isinstance(data_list, list):
        names = []
        for item in data_list:
            if isinstance(item, dict):
                names.append(item.get('name', ''))
            else:
                names.append(str(item))
        return ", ".join(names)
    return str(data_list)

def search_patent():
    try:
        patent_number = entry.get().strip()
        if not patent_number:
            messagebox.showwarning("Warning", "Please enter a patent number! 📝")
            return

        api_url = f"https://serpapi.com/search?engine=google_patents&q={patent_number}&api_key=7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
        params = {
            "engine": "google_patents_details",
            "patent_id": f"patent/{patent_number}/en"
        }
        
        # Show loading message
        status_label.config(text="Searching... ⌛")
        root.update()

        resp = requests.get(api_url, params=params)
        data = resp.json()
        
        # Extract fields with proper handling
        patent_info = {
            "patent": f"{patent_number}",
            "title": str(data.get("title", "")),
            "pdf": str(data.get("pdf", "")),
            "inventors": extract_names(data.get("inventors", [])),
            "assignees": extract_names(data.get("assignees", [])),
            "publication_date": str(data.get("publication_date", "")),
            "abstract": str(data.get("abstract", "")),
            "description_link": str(data.get("description_link", "")),
            "claims": str(data.get("claims", "")),
            "external_links": str(data.get("external_links", ""))
        }

        # Let user choose save location
        default_filename = f"patent_data_{patent_number}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not file_path:
            status_label.config(text="Export cancelled")
            return

        # Create DataFrame and export
        df = pd.DataFrame([patent_info])
        print("Data before writing to Excel:", df.head())
        df.to_excel(file_path, index=False)
        
        status_label.config(text="Export successful! ✨")
        messagebox.showinfo("Success", f"Patent data exported to {os.path.basename(file_path)}! 🎉")

    except requests.exceptions.RequestException:
        status_label.config(text="Network error! 🌐")
        messagebox.showerror("Error", "Failed to connect to the server. Please check your internet connection.")
    except Exception as e:
        status_label.config(text="Error occurred! ⚠️")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main window
root = tk.Tk()
root.title("Patent Search & Export 📑")
root.geometry("500x300")

# Create and style the widgets
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill='both')

label = tk.Label(frame, text="Enter Patent Number:", font=("Arial", 12))
label.pack(pady=10)

entry = tk.Entry(frame, font=("Arial", 11))
entry.pack(pady=10)

search_button = tk.Button(
    frame,
    text="Search & Export",
    command=search_patent,
    font=("Arial", 11),
    bg="#4CAF50",
    fg="white",
    pady=5
)
search_button.pack(pady=10)
# Add status label
status_label = tk.Label(frame, text="Ready", font=("Arial", 10), fg="#666666")
status_label.pack(pady=10)

root.mainloop()