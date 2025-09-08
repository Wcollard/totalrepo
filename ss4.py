import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog, ttk
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
from urllib.parse import quote
import threading

class PatentScraper:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Patents Scraper")
        self.root.geometry("800x600")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Patent numbers input
        ttk.Label(main_frame, text="Enter Patent Numbers (one per line):").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        self.patent_input = scrolledtext.ScrolledText(main_frame, height=10, width=50)
        self.patent_input.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
        self.scrape_button = ttk.Button(button_frame, text="Scrape Patents", command=self.start_scraping)
        self.scrape_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.export_button = ttk.Button(button_frame, text="Export to Excel", command=self.export_to_excel, state=tk.DISABLED)
        self.export_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_button = ttk.Button(button_frame, text="Clear All", command=self.clear_all)
        self.clear_button.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to scrape patents")
        self.status_label.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Results display
        ttk.Label(main_frame, text="Results:").grid(row=5, column=0, columnspan=2, sticky=tk.W)
        
        self.results_text = scrolledtext.ScrolledText(main_frame, height=15, width=70)
        self.results_text.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Data storage
        self.patent_data = []
        
        # Configure scrolledtext grid weight
        main_frame.rowconfigure(6, weight=2)
    
    def clean_patent_number(self, patent_num):
        """Clean and format patent number"""
        # Remove whitespace and common prefixes
        patent_num = patent_num.strip().upper()
        patent_num = re.sub(r'^(US|EP|WO|JP|CN|DE|FR|GB)', '', patent_num)
        # Remove non-alphanumeric characters except hyphens
        patent_num = re.sub(r'[^A-Z0-9\-]', '', patent_num)
        return patent_num
    
    def scrape_patent_data(self, patent_number):
        """Scrape patent data from Google Patents"""
        try:
            # Clean patent number
            clean_patent = self.clean_patent_number(patent_number)
            
            # Construct Google Patents URL
            url = f"https://patents.google.com/patent/{clean_patent}"
            
            # Set up headers to mimic a real browser
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5',
                'Accept-Encoding': 'gzip, deflate',
                'Connection': 'keep-alive',
            }
            
            # Make request
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Extract data with multiple selectors as fallback
            title = self.extract_title(soup)
            abstract = self.extract_abstract(soup)
            inventor = self.extract_inventor(soup)
            pub_date = self.extract_publication_date(soup)
            
            return {
                'Patent Number': patent_number,
                'Clean Patent Number': clean_patent,
                'Title': title,
                'Abstract': abstract,
                'Inventor': inventor,
                'Publication Date': pub_date,
                'Google Patents URL': url,
                'Status': 'Success'
            }
            
        except requests.RequestException as e:
            return {
                'Patent Number': patent_number,
                'Clean Patent Number': clean_patent if 'clean_patent' in locals() else patent_number,
                'Title': 'Error',
                'Abstract': f'Request error: {str(e)}',
                'Inventor': 'Error',
                'Publication Date': 'Error',
                'Google Patents URL': url if 'url' in locals() else 'N/A',
                'Status': 'Error'
            }
        except Exception as e:
            return {
                'Patent Number': patent_number,
                'Clean Patent Number': clean_patent if 'clean_patent' in locals() else patent_number,
                'Title': 'Error',
                'Abstract': f'Parsing error: {str(e)}',
                'Inventor': 'Error',
                'Publication Date': 'Error',
                'Google Patents URL': url if 'url' in locals() else 'N/A',
                'Status': 'Error'
            }
    
    def extract_title(self, soup):
        """Extract patent title with multiple fallback selectors"""
        selectors = [
            'meta[name="DC.title"]',
            'h1[data-proto="TITLE"]',
            'h1.style-scope.patent-text',
            'h1',
            '[data-proto="TITLE"]'
        ]
        
        for selector in selectors:
            element = soup.select_one(selector)
            if element:
                if element.name == 'meta':
                    return element.get('content', '').strip()
                else:
                    return element.get_text().strip()
        
        return 'Title not found'
    
    def extract_abstract(self, soup):
        """Extract patent abstract with multiple fallback selectors"""
        selectors = [
            'meta[name="DC.description"]',
            'div[data-proto="ABSTRACT"]',
            'section[data-proto="ABSTRACT"]',
            'div.abstract',
            'p.abstract'
        ]
        
        for selector in selectors:
            element = soup.select_one(selector)
            if element:
                if element.name == 'meta':
                    return element.get('content', '').strip()
                else:
                    return element.get_text().strip()
        
        # Try to find abstract in text content
        abstract_indicators = ['ABSTRACT', 'Abstract', 'FIELD OF THE INVENTION']
        for indicator in abstract_indicators:
            if indicator in soup.get_text():
                # This is a simplified approach - in practice, you might need more sophisticated parsing
                text = soup.get_text()
                start_idx = text.find(indicator)
                if start_idx != -1:
                    # Get text after the indicator (simplified)
                    abstract_text = text[start_idx:start_idx+500]  # Get first 500 chars
                    return abstract_text.strip()
        
        return 'Abstract not found'
    
    def extract_inventor(self, soup):
        """Extract patent inventor with multiple fallback selectors"""
        selectors = [
            'meta[name="DC.creator"]',
            'dd[data-proto="INVENTOR"]',
            'span[data-proto="INVENTOR"]',
            '.inventor',
            '[data-proto="INVENTOR"]'
        ]
        
        inventors = []
        
        for selector in selectors:
            elements = soup.select(selector)
            for element in elements:
                if element.name == 'meta':
                    content = element.get('content', '').strip()
                    if content:
                        inventors.append(content)
                else:
                    text = element.get_text().strip()
                    if text:
                        inventors.append(text)
        
        # Remove duplicates while preserving order
        inventors = list(dict.fromkeys(inventors))
        
        return '; '.join(inventors) if inventors else 'Inventor not found'
    
    def extract_publication_date(self, soup):
        """Extract patent publication date with multiple fallback selectors"""
        selectors = [
            'meta[name="DC.date"]',
            'time[data-proto="PUBLICATION_DATE"]',
            'dd[data-proto="PUBLICATION_DATE"]',
            '.publication-date',
            '[data-proto="PUBLICATION_DATE"]'
        ]
        
        for selector in selectors:
            element = soup.select_one(selector)
            if element:
                if element.name == 'meta':
                    return element.get('content', '').strip()
                elif element.name == 'time':
                    return element.get('datetime', element.get_text()).strip()
                else:
                    return element.get_text().strip()
        
        return 'Publication date not found'
    
    def update_status(self, message):
        """Update status label"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def update_progress(self, current, total):
        """Update progress bar"""
        progress_percent = (current / total) * 100
        self.progress.config(value=progress_percent)
        self.root.update_idletasks()
    
    def start_scraping(self):
        """Start scraping in a separate thread"""
        patent_numbers = self.patent_input.get("1.0", tk.END).strip().split('\n')
        patent_numbers = [num.strip() for num in patent_numbers if num.strip()]
        
        if not patent_numbers:
            messagebox.showwarning("Warning", "Please enter at least one patent number")
            return
        
        # Disable scrape button and clear previous results
        self.scrape_button.config(state=tk.DISABLED)
        self.export_button.config(state=tk.DISABLED)
        self.patent_data.clear()
        self.results_text.delete("1.0", tk.END)
        
        # Start scraping in thread to prevent UI freezing
        thread = threading.Thread(target=self.scrape_patents, args=(patent_numbers,))
        thread.daemon = True
        thread.start()
    
    def scrape_patents(self, patent_numbers):
        """Scrape patents in background thread"""
        total_patents = len(patent_numbers)
        self.progress.config(maximum=100)
        
        for i, patent_num in enumerate(patent_numbers):
            self.update_status(f"Scraping patent {i+1}/{total_patents}: {patent_num}")
            self.update_progress(i, total_patents)
            
            # Scrape patent data
            patent_info = self.scrape_patent_data(patent_num)
            self.patent_data.append(patent_info)
            
            # Update results display
            result_text = f"Patent: {patent_info['Patent Number']}\n"
            result_text += f"Title: {patent_info['Title'][:100]}{'...' if len(patent_info['Title']) > 100 else ''}\n"
            result_text += f"Status: {patent_info['Status']}\n"
            result_text += "-" * 50 + "\n"
            
            self.results_text.insert(tk.END, result_text)
            self.results_text.see(tk.END)
            
            # Add delay to be respectful to the server
            if i < total_patents - 1:  # Don't delay after the last patent
                time.sleep(2)
        
        # Final update
        self.update_progress(total_patents, total_patents)
        self.update_status(f"Completed scraping {total_patents} patents")
        
        # Re-enable buttons
        self.scrape_button.config(state=tk.NORMAL)
        self.export_button.config(state=tk.NORMAL)
        
        messagebox.showinfo("Complete", f"Finished scraping {total_patents} patents!")
    
    def export_to_excel(self):
        """Export scraped data to Excel"""
        if not self.patent_data:
            messagebox.showwarning("Warning", "No data to export. Please scrape patents first.")
            return
        
        # Ask user for file location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save patent data as..."
        )
        
        if file_path:
            try:
                # Create DataFrame
                df = pd.DataFrame(self.patent_data)
                
                # Reorder columns for better presentation
                column_order = [
                    'Patent Number', 'Title', 'Abstract', 'Inventor', 
                    'Publication Date', 'Google Patents URL', 'Status'
                ]
                df = df[[col for col in column_order if col in df.columns]]
                
                # Export to Excel
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                messagebox.showinfo("Success", f"Data exported successfully to {file_path}")
                self.update_status(f"Data exported to {file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def clear_all(self):
        """Clear all input and results"""
        self.patent_input.delete("1.0", tk.END)
        self.results_text.delete("1.0", tk.END)
        self.patent_data.clear()
        self.progress.config(value=0)
        self.update_status("Ready to scrape patents")
        self.export_button.config(state=tk.DISABLED)

def main():
    # Check for required packages
    try:
        import requests
        import bs4
        import pandas
        import openpyxl
    except ImportError as e:
        print("Missing required packages. Please install:")
        print("pip install requests beautifulsoup4 pandas openpyxl")
        return
    
    root = tk.Tk()
    app = PatentScraper(root)
    root.mainloop()

if __name__ == "__main__":
    main()