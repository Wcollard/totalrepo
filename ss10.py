import tkinter as tk
from tkinter import messagebox
import requests
import pandas as pd

API_KEY = '7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d'  # Replace with your actual API key
API_URL = 'https://www.searchapi.io/api/v1/search?engine=google_patents_details'

def fetch_patent_data(patent_number):
    params = {
        "api_key": API_KEY,
        "patent_number": patent_number.strip()
    }
    try:
        response = requests.get(API_URL, params=params)
        data = response.json()
        # Adapt these keys to actual API response!
        return {
            'Patent Number': patent_number,
            'Publication Date': data.get('publication_date', ''),
            'Inventor': ', '.join(data.get('inventors', [])),
            'Assignee': data.get('assignee', ''),
            'Abstract': data.get('abstract', ''),
            'Link': data.get('link', '')
        }
    except Exception as e:
        return {
            'Patent Number': patent_number,
            'Publication Date': 'Error',
            'Inventor': 'Error',
            'Assignee': 'Error',
            'Abstract': str(e),
            'Link': ''
        }

def search_and_export():
    patent_numbers = txt_patents.get("1.0", tk.END).strip().split('\n')
    results = []
    for num in patent_numbers:
        if num.strip():
            results.append(fetch_patent_data(num))
    df = pd.DataFrame(results)
    df.to_excel('patent_results.xlsx', index=False)
    messagebox.showinfo("Export Finished", "Results exported to patent_results.xlsx 🎉")

root = tk.Tk()
root.title("Google Patents Fetcher")

tk.Label(root, text="Enter Patent Numbers (one per line):").pack()
txt_patents = tk.Text(root, height=10, width=40)
txt_patents.pack()

btn_export = tk.Button(root, text="Fetch & Export to Excel", command=search_and_export)
btn_export.pack(pady=10)

root.mainloop()