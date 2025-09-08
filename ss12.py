#7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d
import tkinter as tk
from tkinter import messagebox
import requests
import pandas as pd

def search_patent():
    patent_number = entry.get()
    api_url = f"https://serpapi.com/search?engine=google_patents&q={patent_number}&api_key=7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
    resp = requests.get(api_url)
    if resp.status_code == 200:
        data = resp.json()
        # Extract needed fields
        result = data['organic_results']
        patent_data = {
#            "Patent Number": result.get("patent_number"),
            "Title": result.get("title"),
            "Abstract": result.get("abstract"),
            "Publication Date": result.get("publication_date"),
            "Assignee": result.get("assignee"),
            "Inventor": result.get("inventor"),
        }
        # Export to Excel
        df = pd.DataFrame([patent_data])
        df.to_excel("patent_data.xlsx", index=False)
        messagebox.showinfo("Success", "Patent data exported to Excel!")
    else:
        messagebox.showerror("Error", "Patent not found or API issue.")

root = tk.Tk()
root.title("Patent Search")
tk.Label(root, text="Enter Patent Number:").pack()
entry = tk.Entry(root)
entry.pack()
tk.Button(root, text="Search & Export", command=search_patent).pack()
root.mainloop()