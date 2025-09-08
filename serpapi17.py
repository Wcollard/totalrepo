from serpapi import GoogleSearch
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import xlsxwriter
import pandas as pd

def fetch_patents():
    patent_numbers = text_input.get("1.0", tk.END).strip().split('\n')
    data = []
    for patent in patent_numbers:
        params = {
            "engine": "google_patents_details",
            "patent_id": f"patent/{patent}/en",
            "api_key": "7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"  # <-- Insert your key!
        }
        search = GoogleSearch(params)
        results = search.get_dict()
        
        # Extract desired fields
        patent_info = {
            "patent": f"{patent}",
            "title": str(results.get("title", "")),
            "inventors": str(results.get("inventors", [])),
            "assignees": str(results.get("assignees", [])),
            "publication_date": str(results.get("publication_date", "")),
            "abstract": str(results.get("abstract", "")),
            "description_link": str(results.get("description_link", "")),
            "claims": str(results.get("claims", "")),
            "pdf": str(results.get("pdf", "")),
            "google_url": f"https://patents.google.com/patent/{patent}/en",
            "espacenet_url": f"https://worldwide.espacenet.com/patent/search?q={patent}",
            "external_links": str(results.get("external_links", ""))
        }
        data.append(patent_info)  # Add this line to collect all patent data
        print (patent_info)
    if data:
        export_to_excel(data)
        messagebox.showinfo("Success", "Patent data exported to patents.xlsx! 🎉")

def export_to_excel(data):
    df = pd.DataFrame(data)
    with pd.ExcelWriter("patents.xlsx", engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Define formats 🎨
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter'
        })

        cell_format = workbook.add_format({
            'text_wrap': True,
            'border': 1,
            'align': 'left',
            'valign': 'top'
        })

        url_format = workbook.add_format({
            'font_color': 'blue',
            'underline': True,
            'text_wrap': True
        })

        # Apply formats
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Set column widths and formats
        worksheet.set_column('A:A', 15, cell_format)  # patent number
        worksheet.set_column('B:B', 35, cell_format)  # title
        worksheet.set_column('C:C', 25, cell_format)  # inventors
        worksheet.set_column('D:D', 25, cell_format)  # assignees
        worksheet.set_column('E:E', 15, cell_format)  # publication date
        worksheet.set_column('F:F', 70, cell_format)  # abstract
        worksheet.set_column('G:L', 30, url_format)   # links

        # Set row heights
        worksheet.set_row(0, 30)  # header row
        for row in range(1, len(data) + 1):
            worksheet.set_row(row, 60)

        # Freeze top row
        worksheet.freeze_panes(1, 0)

# Tkinter setup
root = tk.Tk()
root.title("Patent Data Fetcher")
ttk.Label(root, text="Enter Patent Numbers (one per line):").pack()
text_input = tk.Text(root, height=10, width=40)
text_input.pack()
ttk.Button(root, text="Fetch Patent Data", command=fetch_patents).pack()
root.mainloop()

