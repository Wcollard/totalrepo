from serpapi import GoogleSearch
import tkinter as tk
from tkinter import ttk

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
'''
        # Extract required fields, with safe fallback for missing data
        patent_number = results.get('patent_id', patent)
        inventors = ", ".join([inv.get('name', '') for inv in results.get('inventor', [])]) if 'inventor' in results else ''
        publication_date = results.get('publication_date', '')
        title = results.get('title', '')
        abstract = results.get('abstract', '')
        pdf_link = results.get('pdf', '')
        #external_links = ", ".join(results.get('external_links', [])) if 'external_links' in results else ''

        # Display in console
        print(f"Patent Number: {patent_number}")
        print(f"Inventor(s): {inventors}")
        print(f"Publication Date: {publication_date}")
        print(f"Title: {title}")
        print(f"Abstract: {abstract}")
        print(f"PDF Link: {pdf_link}")
        #print(f"External Links: {external_links}")
        print("="*40)

        # Store for further processing if needed
        data.append({
            "patent_number": patent_number,
            "Inventor": inventors,
            "Publication_date": publication_date,
            "Title": title,
            "Abstract": abstract,
            "pdf_link": pdf_link,
            #"external_links": external_links
        })
'''

    # Extract desired fields
            patent_info = {
                "patent": f"{patent_number}",
                "title": str(data.get("title", "")),
                "inventors": extract_names(data.get("inventors", [])),
                "assignees": extract_names(data.get("assignees", [])),
                "publication_date": str(data.get("publication_date", "")),
                "abstract": str(data.get("abstract", "")),
                "description_link": str(data.get("description_link", "")),
                "claims": str(data.get("claims", "")),
                "pdf": str(data.get("pdf", "")),
                "google_url" : f"https://patents.google.com/patent/{patent_number}/en",
                "espacenet_url" : f"https://worldwide.espacenet.com/patent/search?q={patent_number}",
                "external_links": str(data.get("external_links", ""))
            }



# Create an Excel workbook and worksheet
workbook = xlsxwriter.Workbook('PatentDetails.xlsx')
worksheet = workbook.add_worksheet()

#Export data
df = pd.DataFrame([patent_info])
print("Data before writing to Excel:", df.head())
df.to_excel(file_path, index=False)
        
status_label.config(text="Export successful! ✨")
messagebox.showinfo("Success", f"Patent data exported to {os.path.basename(file_path)}! 🎉")

    



# Tkinter setup
root = tk.Tk()
root.title("Patent Data Fetcher")
ttk.Label(root, text="Enter Patent Numbers (one per line):").pack()
text_input = tk.Text(root, height=10, width=40)
text_input.pack()
ttk.Button(root, text="Fetch Patent Data", command=fetch_patents).pack()
root.mainloop()