import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from serpapi import GoogleSearch
import os

def fetch_patent_data(patent_number, api_key):
    params = {
        "engine": "google_patents_details",
        "patent_id": f"patent/US11734097B1/en",
        "api_key": api_key
    }
    search = GoogleSearch(params)
    try:
        result = search.get_dict()
        patent = result.get("patent", {})
    except Exception as e:
        print(f"Error fetching data: {e}")  # Print any errors encountered
        
    return {
        "patent No": patent.get("patent_number"),
        "title": patent.get("title"),
        "pdf": patent.get("pdf", ""),
        "inventors": ", ".join(patent.get("inventors", [])),
        "assignees": ", ".join(patent.get("assignees", [])),
        "publication_date": patent.get("publication_date"),
        "abstract": patent.get("abstract"),
        "description_link": patent.get("link", ""),
        "claims": patent.get("claims", ""),
        "external_links": ", ".join(patent.get("external_links", [])),
    }

def on_fetch():
#    api_key = os.environ.get("7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d")
    api_key= "7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867duv "
    numbers = entry.get("1.0", tk.END).replace('\n', ',').split(',')
    numbers = [n.strip() for n in numbers if n.strip()]
    results = []
    for num in numbers:
        try:
            data = fetch_patent_data(num, api_key)
            results.append(data)
            print(f"Fetched data for {num}: {data}")  # Debugging line
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch {num}: {e}")
    if results:
        savepath = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if savepath:
            write_to_excel(results, savepath)
            messagebox.showinfo("Success", f"Exported {len(results)} records to Excel!")

def write_to_excel(data, filepath):
    df = pd.DataFrame(data)
    print(df.head())  # Debugging line to check DataFrame contents
    
    with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Patents')
        workbook  = writer.book
        worksheet = writer.sheets['Patents']
        
        # Set column widths
        worksheet.set_column('A:A', 20)  # Patent No
        worksheet.set_column('B:B', 30)  # Title
        worksheet.set_column('C:C', 20)  # PDF link
        worksheet.set_column('D:D', 20)  # Inventors
        worksheet.set_column('E:E', 30)  # Assignees
        worksheet.set_column('F:F', 20)  # Publication Date
        worksheet.set_column('G:G', 70)  # Abstract
        worksheet.set_column('H:H', 20)  # Description Link
        worksheet.set_column('I:I', 20)  # Claims
        worksheet.set_column('J:J', 20)  # External Links

        # Format the abstract column to auto-fit row height
        for idx, row in enumerate(data, start=1):
            abstract = row["abstract"]
            if abstract:
                worksheet.write_string(idx, 6, abstract)
                worksheet.set_row(idx, None, None, {'text_wrap': True})
root = tk.Tk()
root.title("Patent Data Fetcher 📄")
tk.Label(root, text="Enter patent numbers (comma or newline separated):").pack()
entry = tk.Text(root, height=5, width=50)
entry.pack()
tk.Button(root, text="Fetch and Export", command=on_fetch).pack(pady=8)
root.mainloop()