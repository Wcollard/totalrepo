import tkinter as tk
from tkinter import messagebox
import requests
import pandas as pd
from datetime import datetime

def search_patent():
    try:
        patent_number = entry.get().strip()
        # Using Google Patents API endpoint
        api_url = f"https://serpapi.com/search?engine=google_patents&patent_number={patent_number}&api_key=7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
        resp = requests.get(api_url)
        
        if resp.status_code == 200:
            data = resp.json()
            if 'patent_result' in data:  # Changed to patent_result
                result = data['patent_result']
                
                # Extracting inventors (usually comes as a list)
                inventors = ', '.join(result.get('inventors', [])) if result.get('inventors') else 'N/A'
                
                # Extracting abstract (might be in different locations)
                abstract = result.get('abstract', result.get('description', 'N/A'))
                
                # Extracting and formatting publication date
                pub_date = result.get('publication_date', 'N/A')
                if pub_date != 'N/A':
                    try:
                        pub_date = datetime.strptime(pub_date, '%Y-%m-%d').strftime('%Y-%m-%d')
                    except:
                        pass

                patent_data = {
                    "Patent Number": patent_number,
                    "Title": result.get("title", "N/A"),
                    "Abstract": abstract,
                    "Publication Date": pub_date,
                    "Assignee": result.get("assignee", {}).get("name", "N/A"),
                    "Inventor": inventors
                }
                
                # Export to Excel with proper column formatting
                df = pd.DataFrame([patent_data])
                with pd.ExcelWriter('patent_data.xlsx', engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                    # Adjust column widths
                    worksheet = writer.sheets['Sheet1']
                    for idx, col in enumerate(df.columns):
                        worksheet.column_dimensions[chr(65+idx)].width = 30

                messagebox.showinfo("Success", "Patent data exported to Excel!")
            else:
                messagebox.showerror("Error", "Patent details not found")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
# Create the main window
root = tk.Tk()
root.title("Patent Search")
root.geometry("400x200")  # Set window size

# Create and style the widgets
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill='both')

label = tk.Label(frame, text="Enter Patent Number:", font=("Arial", 12))
label.pack(pady=10)

entry = tk.Entry(frame, font=("Arial", 11))
entry.pack(pady=10)

search_button = tk.Button(frame, 
                         text="Search & Export", 
                         command=search_patent,
                         font=("Arial", 11),
                         bg="#4CAF50",
                         fg="white",
                         pady=5)
search_button.pack(pady=10)

root.mainloop()