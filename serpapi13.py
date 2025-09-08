from serpapi import GoogleSearch
import tkinter as tk
from tkinter import ttk
import requests
import pandas as pd
def fetch_patents():
    # Get all patent numbers as a list
    patent_numbers = text_input.get("1.0", tk.END).strip().split('\n')
    data = []
    for patent in patent_numbers:
        # --- 2. Query SerpAPI ---
        params = {
          "engine": "google_patents_details",
          "patent_id": f"patent/{patent}/en",
          "api_key": "7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
}

        search = GoogleSearch(params)
        results = search.get_dict()

        print (results)

# Tkinter setup
root = tk.Tk()
root.title("Patent Data Fetcher")
ttk.Label(root, text="Enter Patent Numbers (one per line):").pack()
text_input = tk.Text(root, height=10, width=40)
text_input.pack()
ttk.Button(root, text="Fetch Patent Data", command=fetch_patents).pack()
root.mainloop()