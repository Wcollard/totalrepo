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
        # Get all patent numbers (one per line)
        patent_numbers = entry.get("1.0", tk.END).strip().split('\n')
        if not patent_numbers:
            messagebox.showwarning("Warning", "Please enter patent numbers! 📝")
            return

        all_patent_info = []
        for patent_number in patent_numbers:
            # [Previous API call code remains the same]
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
                "patent_no": f"{patent_number}",
                "title": str(data.get("title", "")),
                "pdf": str(data.get("pdf", "")),
                "inventors": extract_names(data.get("inventors", [])),
                "assignees": extract_names(data.get("assignees", [])),
                "publication_date": str(data.get("publication_date", "")),
                "abstract": str(data.get("abstract", "")),
                "description_link": str(data.get("description_link", "")),
                "claims": str(data.get("claims", "")),
            #   "external_links": str(data.get("external_links", ""))
        }
            # Add the patent info to the list
            all_patent_info.append(patent_info)

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

# Create DataFrame with all patents
        df = pd.DataFrame(all_patent_info)
        print("Data before writing to Excel:", df.head())
        # Configure Excel writer for better formatting
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            
            # Adjust column widths and text wrapping

            worksheet.set_column('A:A', 30)  # Title
            worksheet.set_column('B:B', 20)  # PDF
            worksheet.set_column('C:D', 25)  # Inventors & Assignees
            worksheet.set_column('E:E', 15)  # Publication date
            worksheet.set_column('F:F', 50, {'text_wrap': True})  # Abstract
            worksheet.set_column('G:G', 30)  # Description link
            worksheet.set_column('H:H', 40, {'text_wrap': True})  # Claims
            worksheet.set_column('I:I', 30)  # External links
 
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

#entry = tk.Entry(frame, font=("Arial", 11))
#entry.pack(pady=10)
# Replace the entry widget
entry = tk.Text(frame, font=("Arial", 11), height=5)
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
#an error occurred "dict" object has no attribute "_get_xf_index"
