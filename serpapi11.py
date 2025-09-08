import os
import tkinter as tk
from tkinter import messagebox
import pandas as pd
from serpapi import GoogleSearch

API_KEY = os.getenv("7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d") or "YOUR_API_KEY"  # Replace or use env variable

def fetch_patent_data(patent_number):
    params = {
        "engine": "google_patents",
        "patent_id": f'patent/{patent_number}/en',
        "api_key": API_KEY,
    }
    search = GoogleSearch(params)
    results = search.get_dict()
    data = results.get("organic_results", [])
    if not data:
        return None

    patent = data  # Take the most relevant result
    # Extracting fields, using get() for safety
    return {
        "patent No": patent.get("patent_id"),
        "title": patent.get("title"),
        "pdf": patent.get("pdf"),
        "inventors": patent.get("inventor"),
        "assignees": patent.get("assignee"),
        "publication_date": patent.get("grant_date"),
        "abstract": patent.get("snippet"),
        "description_link": patent.get("link"),
        "claims": patent.get("claims", ""),  # claims might not be present in all results
        "external_links": patent.get("link"),
    }

def export_to_excel(data_list):
    df = pd.DataFrame(data_list)
    with pd.ExcelWriter("patents.xlsx", engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']

        # Set column widths
        worksheet.set_column('B:B', 30)  # title
        worksheet.set_column('G:G', 70)  # abstract
        worksheet.set_column('E:E', 30)  # assignees

        # Autofit abstract row height (XlsxWriter sets to default, but you can wrap text)
        abstract_col = df.columns.get_loc("abstract")
        for row_idx, val in enumerate(df["abstract"], start=1):  # header is row 0
            worksheet.write(row_idx, abstract_col, val)
            worksheet.set_row(row_idx, None, None, {'text_wrap': True})

def on_submit():
    patent_numbers = entry.get().split(",")
    data = []
    for num in map(str.strip, patent_numbers):
        patent_info = fetch_patent_data(num)
        if patent_info:
            data.append(patent_info)
        else:
            messagebox.showwarning("Not Found", f"No data found for: {num}")
    if data:
        export_to_excel(data)
        messagebox.showinfo("Exported", "Patent data exported to patents.xlsx!")

root = tk.Tk()
root.title("Patent Fetcher")
entry = tk.Entry(root, width=50)
entry.pack(padx=10, pady=10)
entry.insert(0, "Enter comma-separated patent numbers")
submit_btn = tk.Button(root, text="Fetch & Export", command=on_submit)
submit_btn.pack(padx=10, pady=10)
root.mainloop()