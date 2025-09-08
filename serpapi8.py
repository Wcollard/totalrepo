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