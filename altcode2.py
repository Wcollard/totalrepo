import tkinter as tk
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup
import xlsxwriter

def get_patent_details(patent_number):
    google_url = f"https://patents.google.com/patent/{patent_number}/en"
    try:
        response = requests.get(google_url, timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')

            # Extract details
            title_tag = soup.find('meta', attrs={'name': 'DC.title'})
            title = title_tag.get('content', '').strip() if title_tag else "Title not found"

            inventor_tag = soup.find('meta', attrs={'name': 'DC.contributor'})
            inventor = inventor_tag.get('content', '').strip() if inventor_tag else "Inventor not found"

            assignee_tag = soup.find('meta', attrs={'name': 'DC.assignee'})
            assignee = assignee_tag.get('content', '').strip() if assignee_tag else "Assignee not found"

            pub_date_tag = soup.find('meta', attrs={'name': 'DC.date'})
            publication_date = pub_date_tag.get('content', '').strip() if pub_date_tag else "Publication date not found"

            abstract = "Abstract not found"
            meta = soup.find('meta', attrs={'name': 'DC.description'})
            if meta and meta.get('content'):
                abstract = meta['content'].strip()

            return title, inventor, assignee, publication_date, abstract
        return None
    except Exception as e:
        return None

def export_to_excel(data):
    workbook = xlsxwriter.Workbook('patent_details.xlsx')
    worksheet = workbook.add_worksheet()

    # Write headers
    headers = ['Title', 'Inventor', 'Assignee', 'Publication Date', 'Abstract']
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    # Write data
    for row_num, details in enumerate(data, start=1):
        for col_num, value in enumerate(details):
            worksheet.write(row_num, col_num, value)

    # Format the abstract column
    abstract_format = workbook.add_format({'text_wrap': True})
    worksheet.set_column('E:E', 50, abstract_format)

    # Adjust row heights for abstracts
    for row_num in range(1, len(data) + 1):
        worksheet.set_row(row_num, None, abstract_format)

    workbook.close()

def on_export():
    patent_numbers = entry.get().split(',')
    data = []
    for patent_number in patent_numbers:
        details = get_patent_details(patent_number.strip())
        if details:
            data.append(details)
        else:
            messagebox.showerror("Error", f"Failed to retrieve details for {patent_number.strip()}")
            return

    export_to_excel(data)
    messagebox.showinfo("Success", "Patent details exported to 'patent_details.xlsx' successfully!")

# Tkinter setup
root = tk.Tk()
root.title("Patent Details Exporter")

label = tk.Label(root, text="Enter Patent Numbers (comma separated):")
label.pack()

entry = tk.Entry(root, width=50)
entry.pack()

button = tk.Button(root, text="Export to Excel", command=on_export)
button.pack()

root.mainloop()