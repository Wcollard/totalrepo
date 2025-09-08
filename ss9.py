import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import requests
import pandas as pd
from datetime import datetime
import threading
import json
import re

class PatentExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Patents Data Extractor")
        self.root.geometry("800x600")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Google Patents Data Extractor", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="Enter patent numbers (one per line):")
        instructions.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # Patent numbers input
        self.patent_input = scrolledtext.ScrolledText(main_frame, height=8, width=60)
        self.patent_input.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(0, 10))
        
        # Extract button
        self.extract_btn = ttk.Button(button_frame, text="Extract Patent Data", 
                                     command=self.start_extraction)
        self.extract_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Export button
        self.export_btn = ttk.Button(button_frame, text="Export to Excel", 
                                    command=self.export_to_excel, state='disabled')
        self.export_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear button
        self.clear_btn = ttk.Button(button_frame, text="Clear All", 
                                   command=self.clear_all)
        self.clear_btn.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=5, column=0, columnspan=2, sticky=tk.W)
        
        # Results treeview
        self.create_results_tree(main_frame)
        
        # Data storage
        self.patent_data = []
        
    def create_results_tree(self, parent):
        # Results frame
        results_frame = ttk.LabelFrame(parent, text="Results", padding="5")
        results_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                          pady=(10, 0))
        
        # Configure grid weights for results frame
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        parent.rowconfigure(6, weight=1)
        
        # Create treeview
        columns = ('Patent Number', 'Title', 'Inventors', 'Publication Date', 'Assignee')
        self.tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=10)
        
        # Define headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, minwidth=100)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
    
    def get_patent_data(self, patent_number):
        """Fetch patent data using Google Patents Public API"""
        try:
            # Clean patent number
            patent_number = patent_number.strip()
            if not patent_number:
                return None
            
            # Try Google Patents Public API first
            url = f"https://patents.googleapis.com/v1/patents/{patent_number}"
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                data = response.json()
                return self.parse_google_api_response(data, patent_number)
            else:
                # Fallback to scraping approach
                return self.scrape_patent_data(patent_number)
                
        except Exception as e:
            print(f"Error fetching patent {patent_number}: {str(e)}")
            return {
                'patent_number': patent_number,
                'title': f'Error: {str(e)}',
                'abstract': '',
                'inventors': '',
                'publication_date': '',
                'assignee': ''
            }
    
    def parse_google_api_response(self, data, patent_number):
        """Parse Google Patents API response"""
        try:
            title = data.get('title', '')
            abstract = data.get('abstract', '')
            
            # Parse inventors
            inventors = []
            if 'inventor' in data:
                for inv in data['inventor']:
                    if isinstance(inv, dict) and 'name' in inv:
                        inventors.append(inv['name'])
                    elif isinstance(inv, str):
                        inventors.append(inv)
            
            # Parse publication date
            pub_date = ''
            if 'publicationDate' in data:
                pub_date = data['publicationDate']
            elif 'filingDate' in data:
                pub_date = data['filingDate']
            
            # Parse assignee
            assignees = []
            if 'assignee' in data:
                for ass in data['assignee']:
                    if isinstance(ass, dict) and 'name' in ass:
                        assignees.append(ass['name'])
                    elif isinstance(ass, str):
                        assignees.append(ass)
            
            return {
                'patent_number': patent_number,
                'title': title,
                'abstract': abstract,
                'inventors': '; '.join(inventors),
                'publication_date': pub_date,
                'assignee': '; '.join(assignees)
            }
            
        except Exception as e:
            return {
                'patent_number': patent_number,
                'title': f'Parse Error: {str(e)}',
                'abstract': '',
                'inventors': '',
                'publication_date': '',
                'assignee': ''
            }
    
    def scrape_patent_data(self, patent_number):
        """Fallback method to scrape patent data from Google Patents website"""
        try:
            url = f"https://patents.google.com/patent/{patent_number}/en"
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                html = response.text
                
                # Extract title
                title_match = re.search(r'<title[^>]*>([^<]+)</title>', html)
                title = title_match.group(1) if title_match else ''
                title = title.replace(' - Google Patents', '').strip()
                
                # Extract abstract (simplified)
                abstract_match = re.search(r'<div[^>]*class="[^"]*abstract[^"]*"[^>]*>.*?<div[^>]*>([^<]+)', html, re.DOTALL)
                abstract = abstract_match.group(1) if abstract_match else ''
                
                return {
                    'patent_number': patent_number,
                    'title': title,
                    'abstract': abstract,
                    'inventors': 'Data extraction limited',
                    'publication_date': '',
                    'assignee': 'Data extraction limited'
                }
            else:
                return {
                    'patent_number': patent_number,
                    'title': f'HTTP Error: {response.status_code}',
                    'abstract': '',
                    'inventors': '',
                    'publication_date': '',
                    'assignee': ''
                }
                
        except Exception as e:
            return {
                'patent_number': patent_number,
                'title': f'Scraping Error: {str(e)}',
                'abstract': '',
                'inventors': '',
                'publication_date': '',
                'assignee': ''
            }
    
    def start_extraction(self):
        """Start the patent data extraction process"""
        patent_numbers = self.patent_input.get(1.0, tk.END).strip().split('\n')
        patent_numbers = [num.strip() for num in patent_numbers if num.strip()]
        
        if not patent_numbers:
            messagebox.showwarning("No Input", "Please enter at least one patent number.")
            return
        
        # Disable buttons and start progress
        self.extract_btn.config(state='disabled')
        self.export_btn.config(state='disabled')
        self.progress.start()
        self.status_label.config(text="Extracting patent data...")
        
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.patent_data = []
        
        # Start extraction in separate thread
        thread = threading.Thread(target=self.extract_patents, args=(patent_numbers,))
        thread.daemon = True
        thread.start()
    
    def extract_patents(self, patent_numbers):
        """Extract patent data in background thread"""
        for i, patent_num in enumerate(patent_numbers):
            try:
                self.root.after(0, self.update_status, 
                              f"Processing {patent_num} ({i+1}/{len(patent_numbers)})...")
                
                patent_data = self.get_patent_data(patent_num)
                if patent_data:
                    self.patent_data.append(patent_data)
                    # Update UI in main thread
                    self.root.after(0, self.add_result_to_tree, patent_data)
                
            except Exception as e:
                error_data = {
                    'patent_number': patent_num,
                    'title': f'Error: {str(e)}',
                    'abstract': '',
                    'inventors': '',
                    'publication_date': '',
                    'assignee': ''
                }
                self.patent_data.append(error_data)
                self.root.after(0, self.add_result_to_tree, error_data)
        
        # Extraction complete
        self.root.after(0, self.extraction_complete)
    
    def update_status(self, message):
        """Update status label"""
        self.status_label.config(text=message)
    
    def add_result_to_tree(self, patent_data):
        """Add result to treeview"""
        values = (
            patent_data['patent_number'],
            patent_data['title'][:50] + '...' if len(patent_data['title']) > 50 else patent_data['title'],
            patent_data['inventors'][:30] + '...' if len(patent_data['inventors']) > 30 else patent_data['inventors'],
            patent_data['publication_date'],
            patent_data['assignee'][:30] + '...' if len(patent_data['assignee']) > 30 else patent_data['assignee']
        )
        self.tree.insert('', tk.END, values=values)
    
    def extraction_complete(self):
        """Handle extraction completion"""
        self.progress.stop()
        self.extract_btn.config(state='normal')
        if self.patent_data:
            self.export_btn.config(state='normal')
            self.status_label.config(text=f"Extraction complete! {len(self.patent_data)} patents processed.")
        else:
            self.status_label.config(text="No data extracted.")
    
    def export_to_excel(self):
        """Export patent data to Excel file"""
        if not self.patent_data:
            messagebox.showwarning("No Data", "No patent data to export.")
            return
        
        try:
            # Ask for file location
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save patent data as..."
            )
            
            if filename:
                # Create DataFrame
                df = pd.DataFrame(self.patent_data)
                
                # Reorder columns
                column_order = ['patent_number', 'title', 'abstract', 'inventors', 
                               'publication_date', 'assignee']
                df = df[column_order]
                
                # Rename columns for better readability
                df.columns = ['Patent Number', 'Title', 'Abstract', 'Inventors', 
                             'Publication Date', 'Assignee']
                
                # Export to Excel
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Patent Data', index=False)
                    
                    # Auto-adjust column widths
                    worksheet = writer.sheets['Patent Data']
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                
                messagebox.showinfo("Export Complete", f"Data exported successfully to {filename}")
                self.status_label.config(text=f"Data exported to {filename}")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export data: {str(e)}")
    
    def clear_all(self):
        """Clear all data and inputs"""
        self.patent_input.delete(1.0, tk.END)
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.patent_data = []
        self.export_btn.config(state='disabled')
        self.status_label.config(text="Ready")

def main():
    root = tk.Tk()
    app = PatentExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()