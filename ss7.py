import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import asyncio
import threading
from playwright.async_api import async_playwright
import pandas as pd
from datetime import datetime
import os

class PatentScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Patents Scraper")
        self.root.geometry("800x700")
        
        # Data storage
        self.scraped_data = []
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Google Patents Scraper", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Patent Number input
        ttk.Label(main_frame, text="Patent Number:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.patent_entry = ttk.Entry(main_frame, width=30)
        self.patent_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        
        # XPath input
        ttk.Label(main_frame, text="XPath Selector:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.xpath_entry = ttk.Entry(main_frame, width=30)
        self.xpath_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        
        # Section Name input (optional)
        ttk.Label(main_frame, text="Section Name (optional):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.section_entry = ttk.Entry(main_frame, width=30)
        self.section_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        # Add to queue button
        self.add_button = ttk.Button(buttons_frame, text="Add to Queue", 
                                   command=self.add_to_queue)
        self.add_button.pack(side=tk.LEFT, padx=5)
        
        # Scrape button
        self.scrape_button = ttk.Button(buttons_frame, text="Start Scraping", 
                                      command=self.start_scraping)
        self.scrape_button.pack(side=tk.LEFT, padx=5)
        
        # Export button
        self.export_button = ttk.Button(buttons_frame, text="Export to Excel", 
                                      command=self.export_to_excel)
        self.export_button.pack(side=tk.LEFT, padx=5)
        
        # Clear button
        self.clear_button = ttk.Button(buttons_frame, text="Clear All", 
                                     command=self.clear_all)
        self.clear_button.pack(side=tk.LEFT, padx=5)
        
        # Queue display
        ttk.Label(main_frame, text="Scraping Queue:", font=("Arial", 12, "bold")).grid(
            row=5, column=0, columnspan=3, sticky=tk.W, pady=(20, 5))
        
        # Queue listbox with scrollbar
        queue_frame = ttk.Frame(main_frame)
        queue_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        queue_frame.columnconfigure(0, weight=1)
        queue_frame.rowconfigure(0, weight=1)
        
        self.queue_listbox = tk.Listbox(queue_frame, height=6)
        self.queue_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        queue_scrollbar = ttk.Scrollbar(queue_frame, orient=tk.VERTICAL, 
                                       command=self.queue_listbox.yview)
        queue_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.queue_listbox.config(yscrollcommand=queue_scrollbar.set)
        
        # Results display
        ttk.Label(main_frame, text="Scraping Results:", font=("Arial", 12, "bold")).grid(
            row=7, column=0, columnspan=3, sticky=tk.W, pady=(20, 5))
        
        self.results_text = scrolledtext.ScrolledText(main_frame, height=15, width=80)
        self.results_text.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=400, mode='determinate')
        self.progress.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=10, column=0, columnspan=3, pady=5)
        
        # Configure grid weights for resizing
        main_frame.rowconfigure(8, weight=1)
        
        # Sample XPath examples
        self.add_sample_data()
    
    def add_sample_data(self):
        """Add some sample XPath examples to help users"""
        sample_text = """Sample XPath examples for Google Patents:
- Abstract: //section[@itemprop='abstract']//div[@class='abstract']
- Claims: //section[@aria-label='Claims']//div[@class='claims']
- Description: //section[@aria-label='Description']//div[@class='description']
- Patent Title: //h1[@id='title']
- Inventors: //dd[@itemprop='inventor']
- Publication Date: //time[@itemprop='publicationDate']
        """
        self.results_text.insert(tk.END, sample_text + "\n" + "="*80 + "\n\n")
    
    def add_to_queue(self):
        patent_num = self.patent_entry.get().strip()
        xpath = self.xpath_entry.get().strip()
        section_name = self.section_entry.get().strip() or "Scraped Content"
        
        if not patent_num or not xpath:
            messagebox.showerror("Error", "Please enter both Patent Number and XPath")
            return
        
        # Add to queue display
        queue_item = f"{patent_num} | {section_name} | {xpath[:50]}{'...' if len(xpath) > 50 else ''}"
        self.queue_listbox.insert(tk.END, queue_item)
        
        # Clear entries
        self.patent_entry.delete(0, tk.END)
        self.xpath_entry.delete(0, tk.END)
        self.section_entry.delete(0, tk.END)
        
        self.status_label.config(text=f"Added to queue. Total items: {self.queue_listbox.size()}")
    
    def start_scraping(self):
        if self.queue_listbox.size() == 0:
            messagebox.showwarning("Warning", "No items in queue to scrape")
            return
        
        # Disable buttons during scraping
        self.scrape_button.config(state='disabled')
        self.add_button.config(state='disabled')
        
        # Start scraping in a separate thread
        threading.Thread(target=self.run_scraping, daemon=True).start()
    
    def run_scraping(self):
        """Run the scraping process in a separate thread"""
        try:
            asyncio.run(self.scrape_patents())
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Scraping failed: {str(e)}"))
        finally:
            # Re-enable buttons
            self.root.after(0, lambda: self.scrape_button.config(state='normal'))
            self.root.after(0, lambda: self.add_button.config(state='normal'))
    
    async def scrape_patents(self):
        """Main scraping function"""
        queue_items = []
        for i in range(self.queue_listbox.size()):
            item = self.queue_listbox.get(i)
            parts = item.split(' | ')
            if len(parts) >= 3:
                patent_num = parts[0]
                section_name = parts[1]
                xpath = ' | '.join(parts[2:])  # Rejoin in case xpath contains |
                queue_items.append((patent_num, section_name, xpath))
        
        if not queue_items:
            return
        
        self.root.after(0, lambda: self.progress.config(maximum=len(queue_items), value=0))
        
        async with async_playwright() as p:
            # Launch browser
            browser = await p.chromium.launch(headless=True)
            context = await browser.new_context()
            page = await context.new_page()
            
            for i, (patent_num, section_name, xpath) in enumerate(queue_items):
                try:
                    self.root.after(0, lambda p=patent_num: self.status_label.config(
                        text=f"Scraping patent: {p}"))
                    
                    # Navigate to Google Patents
                    url = f"https://patents.google.com/patent/{patent_num}"
                    await page.goto(url, wait_until='networkidle')
                    
                    # Wait a bit for dynamic content to load
                    await page.wait_for_timeout(2000)
                    
                    # Try to find the element using the provided xpath
                    try:
                        element = await page.locator(f'xpath={xpath}').first
                        if await element.count() > 0:
                            content = await element.inner_text()
                            
                            # Store the scraped data
                            data_entry = {
                                'Patent Number': patent_num,
                                'Section Name': section_name,
                                'XPath': xpath,
                                'Content': content.strip(),
                                'Scraped At': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                'URL': url
                            }
                            self.scraped_data.append(data_entry)
                            
                            # Update results display
                            result_text = f"✓ {patent_num} - {section_name}\n"
                            result_text += f"Content: {content[:200]}{'...' if len(content) > 200 else ''}\n"
                            result_text += "-" * 80 + "\n"
                            
                            self.root.after(0, lambda t=result_text: self.results_text.insert(tk.END, t))
                            
                        else:
                            error_msg = f"✗ {patent_num} - Element not found with XPath: {xpath}\n"
                            error_msg += "-" * 80 + "\n"
                            self.root.after(0, lambda t=error_msg: self.results_text.insert(tk.END, t))
                            
                    except Exception as e:
                        error_msg = f"✗ {patent_num} - Error: {str(e)}\n"
                        error_msg += "-" * 80 + "\n"
                        self.root.after(0, lambda t=error_msg: self.results_text.insert(tk.END, t))
                
                except Exception as e:
                    error_msg = f"✗ {patent_num} - Failed to load page: {str(e)}\n"
                    error_msg += "-" * 80 + "\n"
                    self.root.after(0, lambda t=error_msg: self.results_text.insert(tk.END, t))
                
                # Update progress
                self.root.after(0, lambda v=i+1: self.progress.config(value=v))
                
                # Small delay between requests
                await page.wait_for_timeout(1000)
            
            await browser.close()
        
        # Clear queue after scraping
        self.root.after(0, self.queue_listbox.delete, 0, tk.END)
        self.root.after(0, lambda: self.status_label.config(
            text=f"Scraping complete. {len(self.scraped_data)} items scraped."))
    
    def export_to_excel(self):
        if not self.scraped_data:
            messagebox.showwarning("Warning", "No data to export")
            return
        
        # Ask user for file location
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save scraped data as..."
        )
        
        if filename:
            try:
                df = pd.DataFrame(self.scraped_data)
                df.to_excel(filename, index=False, engine='openpyxl')
                messagebox.showinfo("Success", f"Data exported successfully to:\n{filename}")
                self.status_label.config(text=f"Exported {len(self.scraped_data)} records to Excel")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data:\n{str(e)}")
    
    def clear_all(self):
        """Clear all data and reset the application"""
        self.queue_listbox.delete(0, tk.END)
        self.scraped_data.clear()
        self.results_text.delete(1.0, tk.END)
        self.add_sample_data()
        self.progress.config(value=0)
        self.status_label.config(text="Cleared all data")

def main():
    root = tk.Tk()
    app = PatentScraperApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()