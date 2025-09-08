import tkinter as tk
from tkinter import messagebox
import requests
import pandas as pd
from openpyxl import Workbook
import os
from datetime import datetime
from tkinter import filedialog


def extract_names(data_list):
    """Helper function to extract names from list of dictionaries"""
    if not data_list:
        return ""
    if isinstance(data_list, list):
        # If items are dictionaries, extract the 'name' field
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
        api_url = f"https://serpapi.com/search?engine=google_patents&q={patent_number}&api_key=YOUR_API_KEY"
        params = {
            "engine": "google_patents_details",
            "patent_id": f"patent/{patent_number}/en"
        }
        resp = requests.get(api_url, params=params)
        data = resp.json()
        
        # Extract fields with proper handling of complex data
        patent_info = {
            "title": str(data.get("title", "")),
            "pdf": str(data.get("pdf_link", "")),  # Changed to pdf_link
            "inventors": extract_names(data.get("inventors", [])),
            "assignees": extract_names(data.get("assignees", [])),
            "publication_date": str(data.get("publication_date", "")),
            "abstract": str(data.get("abstract", "")),
            "description_link": str(data.get("description_link", "")),
            "claims": str(data.get("claims", ""))
        }
'''        
        # Create DataFrame and export to Excel
        df = pd.DataFrame([patent_info])
        df.to_excel('patent_data.xlsx', index=False)
        messagebox.showinfo("Success", "Patent data exported to patent_data.xlsx! 🎉")
        
        # Print the data for debugging
        print("Extracted Data:", patent_info)
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")'''

# First, let user choose where to save the file
        default_filename = f"patent_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not file_path:  # If user cancels the file dialog
            return
            
        # Rest of your code remains the same until the Excel export part
        
        # Modified Excel export part
        df = pd.DataFrame([patent_info])
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Patent data exported to {os.path.basename(file_path)}! 🎉")
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

        # Print the full error for debugging
        print(f"Full error: {e}")

root = tk.Tk()
root.title("Patent Search")
root.geometry("400x200")
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill='both')

label = tk.Label(frame, text="Enter Patent Number:", font=("Arial", 12))
label.pack(pady=10)

entry = tk.Entry(frame, font=("Arial", 11))
entry.pack(pady=10)

search_button = tk.Button(frame, 
                         text="Search & Export", 
                         command=search_patent,
                         font=("Arial", 11),
                         bg="#4CAF50",
                         fg="white",
                         pady=5)
search_button.pack(pady=10)

root.mainloop()