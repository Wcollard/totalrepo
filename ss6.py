import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
from datetime import datetime
import threading

class EPOPatentScraper:
    def __init__(self, root):
        self.root = root
        self.root.title("European Patent Office Patent Scraper")
        self.root.geometry("800x700")
        
        # Main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="European Patent Office Patent Scraper", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Patent numbers input
        ttk.Label(main_frame, text="Patent Numbers:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Label(main_frame, text="(Enter one patent number per line)", 
                 font=('Arial', 9, 'italic')).grid(row=2, column=0, sticky=tk.W, pady=(0, 10))
        
        self.patent_text = scrolledtext.ScrolledText(main_frame, width=60, height=15)
        self.patent_text.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Example text
        example_text = """EP1234567A1
EP2345678B1
EP3456789A2"""
        self.patent_text.insert(tk.END, example_text)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Scrape button
        self.scrape_button = ttk.Button(button_frame, text="Scrape Patents", 
                                       command=self.start_scraping)
        self.scrape_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear button
        clear_button = ttk.Button(button_frame, text="Clear", 
                                 command=lambda: self.patent_text.delete(1.0, tk.END))
        clear_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Export button
        self.export_button = ttk.Button(button_frame, text="Export to Excel", 
                                       command=self.export_to_excel, state="disabled")
        self.export_button.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to scrape patents")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        # Results area
        ttk.Label(main_frame, text="Results:").grid(row=7, column=0, sticky=tk.W, pady=(10, 5))
        
        # Treeview for results
        columns = ('Patent Number', 'Title', 'Abstract', 'Publication Date', 'Inventor(s)', 'Link')
        self.tree = ttk.Treeview(main_frame, columns=columns, show='headings', height=8)
        
        # Define headings
        for col in columns:
            self.tree.heading(col, text=col)
            if col in ['Title', 'Abstract']:
                self.tree.column(col, width=200)
            elif col == 'Link':
                self.tree.column(col, width=300)
            else:
                self.tree.column(col, width=120)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        scrollbar.grid(row=8, column=2, sticky=(tk.N, tk.S), pady=10)
        
        # Configure grid weights for resizing
        main_frame.rowconfigure(8, weight=1)
        
        # Data storage
        self.scraped_data = []
        
    def get_patent_data(self, patent_number):
        """Scrape patent data from European Patent Office"""
        try:
            # Clean patent number
            patent_number = patent_number.strip().upper()
            
            # EPO search URL
            base_url = "https://worldwide.espacenet.com/patent/search/family"
            search_url = f"{base_url}?q={patent_number}"
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            # First, search for the patent
            response = requests.get(search_url, headers=headers, timeout=30)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Try to find the direct patent link
            patent_link = None
            title = "Not found"
            abstract = "Not found"
            pub_date = "Not found"
            inventors = "Not found"
            
            # Look for patent results
            result_links = soup.find_all('a', href=True)
            for link in result_links:
                href = link.get('href', '')
                if 'publication' in href and patent_number.replace(' ', '') in href.upper():
                    patent_link = f"https://worldwide.espacenet.com{href}"
                    break
            
            if not patent_link:
                # Alternative URL construction
                patent_link = f"https://worldwide.espacenet.com/patent/search?q={patent_number}"
            
            # Try to get detailed patent page
            if patent_link and 'publication' in patent_link:
                detail_response = requests.get(patent_link, headers=headers, timeout=30)
                if detail_response.status_code == 200:
                    detail_soup = BeautifulSoup(detail_response.content, 'html.parser')
                    
                    # Extract title
                    title_elem = detail_soup.find('h1') or detail_soup.find('title')
                    if title_elem:
                        title = title_elem.get_text().strip()[:200]
                    
                    # Extract abstract
                    abstract_elem = detail_soup.find(text=re.compile('Abstract', re.I))
                    if abstract_elem:
                        abstract_parent = abstract_elem.parent
                        if abstract_parent:
                            abstract_text = abstract_parent.find_next('p') or abstract_parent.find_next('div')
                            if abstract_text:
                                abstract = abstract_text.get_text().strip()[:500]
                    
                    # Extract publication date
                    date_elem = detail_soup.find(text=re.compile('Publication date', re.I))
                    if date_elem:
                        date_parent = date_elem.parent
                        if date_parent:
                            date_text = date_parent.find_next()
                            if date_text:
                                pub_date = date_text.get_text().strip()
                    
                    # Extract inventors
                    inventor_elem = detail_soup.find(text=re.compile('Inventor', re.I))
                    if inventor_elem:
                        inventor_parent = inventor_elem.parent
                        if inventor_parent:
                            inventor_text = inventor_parent.find_next()
                            if inventor_text:
                                inventors = inventor_text.get_text().strip()[:200]
            
            return {
                'patent_number': patent_number,
                'title': title,
                'abstract': abstract,
                'publication_date': pub_date,
                'inventors': inventors,
                'link': patent_link or search_url
            }
            
        except requests.exceptions.RequestException as e:
            return {
                'patent_number': patent_number,
                'title': f"Error: {str(e)[:100]}",
                'abstract': "Request failed",
                'publication_date': "N/A",
                'inventors': "N/A",
                'link': f"https://worldwide.espacenet.com/patent/search?q={patent_number}"
            }
        except Exception as e:
            return {
                'patent_number': patent_number,
                'title': f"Error: {str(e)[:100]}",
                'abstract': "Parsing failed",
                'publication_date': "N/A",
                'inventors': "N/A",
                'link': f"https://worldwide.espacenet.com/patent/search?q={patent_number}"
            }
    
    def scrape_patents_thread(self, patent_numbers):
        """Thread function to scrape patents"""
        try:
            total_patents = len(patent_numbers)
            self.scraped_data = []
            
            for i, patent_number in enumerate(patent_numbers):
                if not patent_number.strip():
                    continue
                    
                self.root.after(0, lambda i=i, total=total_patents: 
                              self.status_label.config(text=f"Scraping patent {i+1} of {total}: {patent_number}"))
                
                patent_data = self.get_patent_data(patent_number)
                self.scraped_data.append(patent_data)
                
                # Update UI in main thread
                self.root.after(0, lambda data=patent_data: self.add_result_to_tree(data))
                
                # Add delay to be respectful to the server
                time.sleep(2)
            
            # Update UI when done
            self.root.after(0, self.scraping_complete)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Scraping failed: {str(e)}"))
            self.root.after(0, self.scraping_complete)
    
    def add_result_to_tree(self, patent_data):
        """Add a patent result to the treeview"""
        self.tree.insert('', 'end', values=(
            patent_data['patent_number'],
            patent_data['title'][:50] + "..." if len(patent_data['title']) > 50 else patent_data['title'],
            patent_data['abstract'][:50] + "..." if len(patent_data['abstract']) > 50 else patent_data['abstract'],
            patent_data['publication_date'],
            patent_data['inventors'][:30] + "..." if len(patent_data['inventors']) > 30 else patent_data['inventors'],
            patent_data['link']
        ))
    
    def start_scraping(self):
        """Start the scraping process"""
        patent_text = self.patent_text.get(1.0, tk.END).strip()
        if not patent_text:
            messagebox.showwarning("Warning", "Please enter at least one patent number")
            return
        
        patent_numbers = [line.strip() for line in patent_text.split('\n') if line.strip()]
        
        if not patent_numbers:
            messagebox.showwarning("Warning", "Please enter valid patent numbers")
            return
        
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Disable button and start progress
        self.scrape_button.config(state="disabled")
        self.export_button.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="Starting to scrape patents...")
        
        # Start scraping in a separate thread
        thread = threading.Thread(target=self.scrape_patents_thread, args=(patent_numbers,))
        thread.daemon = True
        thread.start()
    
    def scraping_complete(self):
        """Called when scraping is complete"""
        self.progress.stop()
        self.scrape_button.config(state="normal")
        self.export_button.config(state="normal")
        self.status_label.config(text=f"Scraping complete! Found {len(self.scraped_data)} patents.")
    
    def export_to_excel(self):
        """Export scraped data to Excel"""
        if not self.scraped_data:
            messagebox.showwarning("Warning", "No data to export. Please scrape patents first.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialname=f"patents_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                df = pd.DataFrame(self.scraped_data)
                df.to_excel(filename, index=False)
                messagebox.showinfo("Success", f"Data exported successfully to {filename}")
                self.status_label.config(text=f"Data exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")

def main():
    root = tk.Tk()
    app = EPOPatentScraper(root)
    root.mainloop()

if __name__ == "__main__":
    main()