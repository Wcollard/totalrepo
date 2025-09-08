import tkinter as tk
from tkinter import messagebox
import requests
import pandas as pd

def search_patent():
    patent_number = entry.get().strip()
    if not patent_number:
        messagebox.showerror("Error", "Please enter a patent number.")
        return

    api_url = f"https://serpapi.com/search?engine=google_patents&patent_number={patent_number}&api_key=7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
    try:
        resp = requests.get(api_url)
        if resp.status_code == 200:
            data = resp.json()
            if 'patent_result' in data:
                result = data['patent_result']
                # Extracting fields with safe defaults
                title = result.get("title", "N/A")
                abstract = result.get("abstract", result.get("description", "N/A"))
                pub_date = result.get("publication_date", "N/A")
                assignee = result.get("assignee", {}).get("name", "N/A") if isinstance(result.get("assignee"), dict) else result.get("assignee", "N/A")
                inventors = ', '.join(result.get("inventors", [])) if isinstance(result.get("inventors"), list) else result.get("inventors", "N/A")
                pdf_link = result.get("pdf", "N/A")  # SerpApi provides a 'pdf' field[3](https://serpapi.com/blog/export-patent-details-from-google-patents-to-csv-using-python/)
                
                patent_data = {
                    "Patent Number": patent_number,
                    "Title": title,
                    "Abstract": abstract,
                    "Publication Date": pub_date,
                    "Assignee": assignee,
                    "Inventor": inventors,
                    "PDF Link": pdf_link
                }
                
                df = pd.DataFrame([patent_data])
                df.to_excel("patent_data.xlsx", index=False)
                messagebox.showinfo("Success", f"Patent data exported to Excel!\nPDF Link: {pdf_link}")
            else:
                messagebox.showerror("Error", "Patent details not found.")
        else:
            messagebox.showerror("Error", f"API request failed. Status: {resp.status_code}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

root = tk.Tk()
root.title("Patent Search")
root.geometry("420x230")

frame = tk.Frame(root, padx=16, pady=16)
frame.pack(expand=True, fill='both')

label = tk.Label(frame, text="Enter Patent Number:", font=("Arial", 12))
label.pack(pady=8)

entry = tk.Entry(frame, font=("Arial", 11))
entry.pack(pady=8)

search_button = tk.Button(frame, 
                         text="Search & Export", 
                         command=search_patent,
                         font=("Arial", 11),
                         bg="#4CAF50",
                         fg="white",
                         pady=5)
search_button.pack(pady=12)

root.mainloop()