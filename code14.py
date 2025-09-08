import tkinter as tk
from tkinter import messagebox
import xlsxwriter
from datetime import datetime
import requests
from bs4 import BeautifulSoup
'''
def get_abstract(google_url):
    try:
        response = requests.get(google_url, timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            # Abstract is in <meta name="description" content="...">
            meta = soup.find('meta', attrs={'name': 'description'})
            if meta:
                return meta.get('content', '').strip()
        return "Abstract not found"
    except Exception as e:
        return f"Error: {e}"
'''
def search_patent(patent_number):
    try:
        patent_number = text_input.get().strip()
        if not patent_number:
            messagebox.showwarning("Warning", "Please enter a patent number! 📝")
            return

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
            "patent": f"{patent_number}",
            "title": str(data.get("title", "")),
            "pdf": str(data.get("pdf", "")),
            "inventors": extract_names(data.get("inventors", [])),
            "assignees": extract_names(data.get("assignees", [])),
            "publication_date": str(data.get("publication_date", "")),
            "abstract": str(data.get("abstract", "")),
            "description_link": str(data.get("description_link", "")),
            "claims": str(data.get("claims", "")),
            "external_links": str(data.get("external_links", ""))
        }

        return patent_info
    except Exception as e:
        print(f"Error: {e}")
        return None

    

def export_to_excel():
    patent_numbers = text_input.get("1.0", tk.END).strip().split("\n")
    if not patent_numbers or not any(patent_numbers):
        messagebox.showwarning("No Input", "Please enter at least one patent or publication number.")
        return

    for number in patent_numbers:
        cleaned_number = number.strip()
        if not cleaned_number:
            continue

        
        
        # Call search_patent for each patent number
        patent_info = search_patent(cleaned_number)
        if not patent_info:
            continue

        abstract = patent_info["abstract"]

        # Proceed with writing to Excel as before...    

    # Create timestamped filename
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"patent_export_{now}.xlsx"

    # Create workbook and worksheet
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Add formats
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'border': 1
    })

    abstract_format = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1
    })

    link_format = workbook.add_format({
        'underline': True,
        'color': 'blue',
        'border': 1
    })

    # Set column widths
    column_widths = [25, 25, 25, 25, 70, 50]
    for col, width in enumerate(column_widths):
        worksheet.set_column(col, col, width)

    # Write headers
    headers = ["Ref No.", "Google Link", "Espacenet Link", "USPTO Link", "ABSTRACT", "NOTES"]
    worksheet.write_row(0, 0, headers, header_format)

    # Write data
    row = 1
    for number in patent_numbers:
        cleaned_number = number.strip()
        if not cleaned_number:
            continue

        google_url = f"https://patents.google.com/patent/{cleaned_number}"
        espacenet_url = f"https://worldwide.espacenet.com/patent/search?q={cleaned_number}"
        uspto_number = cleaned_number.replace("US", "")
        uspto_url = f"https://ppubs.uspto.gov/pubwebapp/external.html?q={uspto_number}.pn."

      #  abstract = get_abstract(google_url)
        abstract = patent_info["abstract"]

        # Write data with appropriate formats
        worksheet.write(row, 0, cleaned_number)
        worksheet.write_url(row, 1, google_url, link_format, string=cleaned_number)
        worksheet.write_url(row, 2, espacenet_url, link_format, string=cleaned_number)
        worksheet.write_url(row, 3, uspto_url, link_format, string=cleaned_number)
        worksheet.write(row, 4, abstract, abstract_format)

        # Set row height based on abstract content
        text_lines = len(abstract) // 70 + abstract.count('\n') + 1  # Rough estimate
        row_height = min(text_lines * 15, 409)  # 409 is Excel's maximum row height
        worksheet.set_row(row, row_height)

        row += 1

    workbook.close()
    messagebox.showinfo("Export Successful", f"Data exported to {filename}")
# Tkinter UI
root = tk.Tk()
root.title("Patent Exporter")

tk.Label(root, text="Enter patent/publication numbers (one per line):").pack(pady=5)
text_input = tk.Text(root, height=15, width=45)
text_input.pack(padx=10)

tk.Button(root, text="Export to Excel", command=export_to_excel).pack(pady=10)

root.mainloop()