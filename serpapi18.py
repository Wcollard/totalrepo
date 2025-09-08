import sqlite3
from serpapi import GoogleSearch
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

def create_database():
    conn = sqlite3.connect('patents.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS patents
                 (patent TEXT PRIMARY KEY, title TEXT, inventors TEXT, 
                  assignees TEXT, publication_date TEXT, abstract TEXT,
                  description_link TEXT, claims TEXT, pdf TEXT,
                  google_url TEXT, espacenet_url TEXT, external_links TEXT)''')
    conn.commit()
    conn.close()

def check_patent_exists(patent_number):
    conn = sqlite3.connect('patents.db')
    c = conn.cursor()
    c.execute('SELECT * FROM patents WHERE patent = ?', (patent_number,))
    result = c.fetchone()
    conn.close()
    return result is not None

def save_to_database(patent_info):
    conn = sqlite3.connect('patents.db')
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO patents VALUES 
                 (?,?,?,?,?,?,?,?,?,?,?,?)''', 
                 (patent_info['patent'], patent_info['title'], 
                  patent_info['inventors'], patent_info['assignees'],
                  patent_info['publication_date'], patent_info['abstract'],
                  patent_info['description_link'], patent_info['claims'],
                  patent_info['pdf'], patent_info['google_url'],
                  patent_info['espacenet_url'], patent_info['external_links']))
    conn.commit()
    conn.close()
def fetch_patents():
    create_database()
    patent_numbers = text_input.get("1.0", tk.END).strip().split('\n')
    data = []
    
    for patent in patent_numbers:
        try:
            if check_patent_exists(patent):
                # Fetch from database
                conn = sqlite3.connect('patents.db')
                c = conn.cursor()
                c.execute('SELECT * FROM patents WHERE patent = ?', (patent,))
                result = c.fetchone()
                conn.close()
                
                patent_info = {
                    "patent": result[0], "title": result[1],
                    "inventors": result[2], "assignees": result[3],
                    "publication_date": result[4], "abstract": result[5],
                    "description_link": result[6], "claims": result[7],
                    "pdf": result[8], "google_url": result[9],
                    "espacenet_url": result[10], "external_links": result[11]
                }
            else:
                # Fetch from API
                params = {
                    "engine": "google_patents_details",
                    "patent_id": f"patent/{patent}/en",
                    "api_key": "7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867duv"
                }
                
                try:
                    search = GoogleSearch(params)
                    results = search.get_dict()
                    
                    patent_info = {
                        "patent": f"{patent}",
                        "title": results.get("title", "No title available"),
                        "inventors": str(results.get("inventors", [])),
                        "assignees": str(results.get("assignees", [])),
                        "publication_date": results.get("publication_date", "No date available"),
                        "abstract": results.get("abstract", "No abstract available"),
                        "description_link": results.get("description_link", "No link available"),
                        "claims": results.get("claims", "No claims available"),
                        "pdf": results.get("pdf", "No PDF available"),
                        "google_url": f"https://patents.google.com/patent/{patent}/en",
                        "espacenet_url": f"https://worldwide.espacenet.com/patent/search?q={patent}",
                        "external_links": str(results.get("external_links", []))
                    }
                    save_to_database(patent_info)
                except Exception as e:
                    messagebox.showerror("API Error", f"Error fetching patent {patent}: {str(e)}")
                    continue
                    
            data.append(patent_info)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing patent {patent}: {str(e)}")
            continue
    
    if data:
        export_to_excel(data)
        messagebox.showinfo("Success", "Patent data exported to patents.xlsx! 🎉")

# [Your existing export_to_excel function remains the same]
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

