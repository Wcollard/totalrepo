import xlsxwriter
from serpapi import GoogleSearch

# Set up your API parameters
params = {
    "engine": "google_patents_details",
    "patent_id": "patent/US11734097B1/en",
    "api_key": "7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
}

# Fetch results from GoogleSearch
search = GoogleSearch(params)
results = search.get_dict()

# Extract desired fields
title = results.get("title", "N/A")
pdf = results.get("pdf", "N/A")
assignees = results.get("assignees", [])
publication_date = results.get("publication_date", "N/A")
abstract = results.get("abstract", "N/A")
inventors = results.get("inventors", "N/A")
external_links = results.get("external_links", [])

# Create an Excel workbook and worksheet
workbook = xlsxwriter.Workbook('PatentDetails.xlsx')
worksheet = workbook.add_worksheet()

# Define headers and data
headers = ['Title', 'PDF', 'Assignees', 'Publication Date', 'Abstract', 'Inventors', 'External Links']
data = [title, pdf, assignees, publication_date, abstract, inventors, external_links]

# Write headers to the first row
for col_num, header in enumerate(headers):
    worksheet.write(0, col_num, header)

# Write data to the second row
for col_num, item in enumerate(data):
    worksheet.write(1, col_num, item)

# Close the workbook
workbook.close()



# Create the main window
root = tk.Tk()
root.title("Patent Search & Export 📑")
root.geometry("500x300")

# Create and style the widgets
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill='both')

label = tk.Label(frame, text="Enter Patent Number:", font=("Arial", 12))
label.pack(pady=10)

entry = tk.Entry(frame, font=("Arial", 11))
entry.pack(pady=10)

search_button = tk.Button(
    frame,
    text="Search & Export",
#    command=search_patent,
    font=("Arial", 11),
    bg="#4CAF50",
    fg="white",
    pady=5
)
search_button.pack(pady=10)
# Add status label
status_label = tk.Label(frame, text="Ready", font=("Arial", 10), fg="#666666")
status_label.pack(pady=10)

root.mainloop()