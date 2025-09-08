import tkinter as tk
from tkinter import filedialog, messagebox
import docx
import re
import pandas as pd
from openpyxl import Workbook

def create_gui():
    window = tk.Tk()
    window.title("Word Document Part Extractor 🔍")
    window.geometry("500x300")
    
    def browse_file():
        filename = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=(("Word files", "*.docx"), ("All files", "*.*"))
        )
        entry.delete(0, tk.END)
        entry.insert(0, filename)

    def search_document():
        filename = entry.get()
        if not filename:
            messagebox.showerror("Error", "Please select a file first! 😅")
            return
            
        try:
            doc = docx.Document(filename)
            parts_list = []
            
            for paragraph in doc.paragraphs:
                words = paragraph.text.split()
                for i in range(len(words)-1):
                    if words[i].isalpha() and words[i+1].isdigit() and len(words[i+1]) >= 2:
                        parts_list.append([words[i+1], words[i]])

            if parts_list:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="Save Excel File"
                )
                
                if save_path:
                    df = pd.DataFrame(parts_list, columns=['Part Number', 'Description'])
                    df.to_excel(save_path, index=False)
                    result_label.config(text="Success! Excel file saved 🎉")
            else:
                result_label.config(text="No parts found in the document 😕")
                
        except Exception as e:
            result_label.config(text=f"Error: {str(e)} ❌")

    # GUI Elements with improved styling
    frame = tk.Frame(window, padx=20, pady=20)
    frame.pack(expand=True, fill='both')

    label = tk.Label(frame, text="Select Word document:", font=('Arial', 12))
    label.pack(pady=10)

    entry = tk.Entry(frame, width=50)
    entry.pack(pady=5)

    browse_button = tk.Button(frame, text="Browse 📂", command=browse_file)
    browse_button.pack(pady=5)

    search_button = tk.Button(frame, text="Search & Export 🔍", command=search_document)
    search_button.pack(pady=20)

    result_label = tk.Label(frame, text="", font=('Arial', 10))
    result_label.pack(pady=10)

    window.mainloop()

if __name__ == "__main__":
    create_gui()