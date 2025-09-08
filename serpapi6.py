import tkinter as tk
from tkinter import messagebox
from serpapi import GoogleSearch

params = {
  "engine": "google_patents_details",
  "patent_id": "patent/US11734097B1/en",
  "api_key": "7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
}

search = GoogleSearch(params)
results = search.get_dict()

# Extract desired fields
title = results.get("title", "N/A")
pdf = results.get("pdf", "N/A")
assignees = results.get("assignees", [])
publication_date = results.get("publication_date", "N/A")
abstract= results.get("abstract", "N/A")
inventors= results.get("inventors", "N/A")
external_links=results.get("external_links", [])

print("Title:", title)
print("PDF:", pdf)
print("Assignees:", assignees)
print("Publication Date:", publication_date)
print("Abstract:", abstract)
print("Inventors:", inventors)
print("external_links:", external_links)

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