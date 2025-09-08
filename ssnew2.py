import tkinter as tk
from tkinter import messagebox
import requests
import pandas as pd
from openpyxl import Workbook

def search_patent():
    try:
        patent_number = entry.get().strip()
        api_url = f"https://serpapi.com/search?engine=google_patents&q={patent_number}&api_key=7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
        params = {
            "engine": "google_patents_details",
            "patent_id": f"patent/{patent_number}/en"
        }
        resp = requests.get(api_url, params=params)
        data = resp.json()
        
        # Extract fields (make sure keys match API response)
        patent_info = {
            "title": data.get("title", ""),
            "pdf": data.get("pdf", ""),
            "inventors": ", ".join(data.get("inventors", [])),
            "assignees": ", ".join(data.get("assignees", [])),
            "publication_date": data.get("publication_date", ""),
            "abstract": data.get("abstract", ""),
            "description_link": data.get("description_link", ""),
            "claims": data.get("claims", "")
        }
        
        # Export to Excel
        wb = Workbook()
        ws = wb.active
        ws.append(list(patent_info.keys()))  # Header
        ws.append(list(patent_info.values()))  # Data
        wb.save('patent_data.xlsx')
        messagebox.showinfo("Success", "Patent data exported to patent_data.xlsx! 🎉")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

root = tk.Tk()
root.title("Patent Search")
root.geometry("400x200")
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill='both')

label = tk.Label(frame, text="Enter Patent Number:", font=("Arial", 12))
label.pack(pady=10)

entry = tk.Entry(frame, font=("Arial", 11))
entry.pack(pady=10)

search_button = tk.Button(frame, text="Search & Export", command=search_patent, font=("Arial", 11), bg="#4CAF50", fg="white", pady=5)
search_button.pack(pady=10)

root.mainloop()