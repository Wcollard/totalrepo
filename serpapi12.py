import tkinter as tk
from tkinter import ttk
import requests
import pandas as pd

# --- 1. Tkinter GUI for multiline input ---
def fetch_patents():
    # Get all patent numbers as a list
    patent_numbers = text_input.get("1.0", tk.END).strip().split('\n')
    data = []
    for patent in patent_numbers:
        # --- 2. Query SerpAPI ---
        params = {
            "engine": "google_patents",
            "q": patent,
            "api_key": "YOUR_SERPAPI_KEY"
        }
        response = requests.get("https://serpapi.com/search", params=params)
        result = response.json()
        # --- 3. Parse and collect results (handle rare missing/variant fields) ---
        if 'patents_results' in result and result['patents_results']:
            r = result['patents_results']
            data.append({
                "patent_number": r.get('patent_number', ''),
                "Inventor": r.get('inventor', ''),
                "Publication_Date": r.get('publication_date', ''),
                "Title": r.get('title', ''),
                "Abstract": r.get('abstract', ''),
                "PDF_Link": r.get('pdf', ''),
                "External_Links": r.get('link', '')
            })
        else:
            # Handle rare case: patent not found or API structure change
            data.append({
                "patent_number": patent,
                "Inventor": 'Not found',
                "Publication_Date": '',
                "Title": '',
                "Abstract": '',
                "PDF_Link": '',
                "External_Links": ''
            })
    # --- 4. Pandas DataFrame ---
    df = pd.DataFrame(data)
    print(df)

# Tkinter setup
root = tk.Tk()
root.title("Patent Data Fetcher")
ttk.Label(root, text="Enter Patent Numbers (one per line):").pack()
text_input = tk.Text(root, height=10, width=40)
text_input.pack()
ttk.Button(root, text="Fetch Patent Data", command=fetch_patents).pack()
root.mainloop()