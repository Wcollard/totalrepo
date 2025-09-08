import tkinter as tk
from tkinter import messagebox
import requests
import pandas as pd

def search_patent():
    try:
        patent_number = entry.get().strip()  # Remove any whitespace
        api_url = f"https://serpapi.com/search?engine=google_patents&q={patent_number}&api_key=7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
        resp = requests.get(api_url)
        
        if resp.status_code == 200:
            data = resp.json()
            # The organic_results is a list, we typically want the first result
            if 'organic_results' in data and len(data['organic_results']) > 0:
                result = data['organic_results'][0]  # Get the first result
                
                patent_data = {
                    "Patent Number": patent_number,
                    "Title": result.get("title", "N/A"),
                    "Abstract": result.get("description", "N/A"),  # Note: 'description' instead of 'abstract'
                    "Publication Date": result.get("publication_info", {}).get("publication_date", "N/A"),
                    "Assignee": result.get("assignee", "N/A"),
                    "Inventor": result.get("inventors", "N/A")
                }
                
                # Export to Excel
                df = pd.DataFrame([patent_data])
                df.to_excel("patent_data.xlsx", index=False)
                messagebox.showinfo("Success", "Patent data exported to Excel!")
            else:
                messagebox.showerror("Error", "No patent results found.")
        else:
            messagebox.showerror("Error", "API request failed.")
            
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main window
root = tk.Tk()
root.title("Patent Search")
root.geometry("400x200")  # Set window size

# Create and style the widgets
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