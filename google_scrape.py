
#pip install playwright pandas openpyxl
#playwright install chromium

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from playwright.sync_api import sync_playwright
import re
import os
from datetime import datetime
import threading
import time

class PatentScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Patents Scraper")
        self.root.geometry("800x600")
        
        # Data storage
        self.patent_data = []
        
        # Setup UI
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Google Patents Data Scraper", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # URL input section
        url_frame = ttk.LabelFrame(main_frame, text="Patent URLs", padding="10")
        url_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(url_frame, text="Enter Google Patents URL:").grid(row=0, column=0, sticky=tk.W)
        
        self.url_entry = ttk.Entry(url_frame, width=60)
        self.url_entry.grid(row=1, column=0, padx=(0, 10), sticky=(tk.W, tk.E))
        
        self.add_url_btn = ttk.Button(url_frame, text="Add URL", command=self.add_url)
        self.add_url_btn.grid(row=1, column=1)
        
        # URL list
        self.url_listbox = tk.Listbox(url_frame, height=5)
        self.url_listbox.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Buttons for URL management
        btn_frame = ttk.Frame(url_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(5, 0))
        
        self.remove_url_btn = ttk.Button(btn_frame, text="Remove Selected", 
                                        command=self.remove_url)
        self.remove_url_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.clear_urls_btn = ttk.Button(btn_frame, text="Clear All", 
                                        command=self.clear_urls)
        self.clear_urls_btn.pack(side=tk.LEFT)
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.progress_var = tk.StringVar(value="Ready to scrape patents")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, length=400, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Action buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        self.scrape_btn = ttk.Button(action_frame, text="Start Scraping", 
                                    command=self.start_scraping, style="Accent.TButton")
        self.scrape_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.export_btn = ttk.Button(action_frame, text="Export to Excel", 
                                    command=self.export_to_excel, state="disabled")
        self.export_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.view_data_btn = ttk.Button(action_frame, text="View Data", 
                                       command=self.view_data, state="disabled")
        self.view_data_btn.pack(side=tk.LEFT)
        
        # Results display
        results_frame = ttk.LabelFrame(main_frame, text="Scraped Data Preview", padding="10")
        results_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Treeview for data display
        columns = ("Patent Number", "Title", "Assignee", "Inventor", "Publication Date")
        self.tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=8)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
        
        # Scrollbars for treeview
        v_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        url_frame.columnconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
    
    def add_url(self):
        url = self.url_entry.get().strip()
        if url:
            if "patents.google.com" in url:
                self.url_listbox.insert(tk.END, url)
                self.url_entry.delete(0, tk.END)
            else:
                messagebox.showerror("Error", "Please enter a valid Google Patents URL")
    
    def remove_url(self):
        selection = self.url_listbox.curselection()
        if selection:
            self.url_listbox.delete(selection[0])
    
    def clear_urls(self):
        self.url_listbox.delete(0, tk.END)
    
    def start_scraping(self):
        urls = [self.url_listbox.get(i) for i in range(self.url_listbox.size())]
        if not urls:
            messagebox.showerror("Error", "Please add at least one patent URL")
            return
        
        # Disable scraping button and start in separate thread
        self.scrape_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.view_data_btn.config(state="disabled")
        
        # Start scraping in a separate thread
        threading.Thread(target=self.scrape_patents, args=(urls,), daemon=True).start()
    
    def scrape_patents(self, urls):
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False)  # Set to True for headless mode
                page = browser.new_page()
                
                total_urls = len(urls)
                self.patent_data = []  # Reset data
                
                for i, url in enumerate(urls):
                    self.progress_var.set(f"Scraping patent {i+1} of {total_urls}...")
                    self.progress_bar['value'] = (i / total_urls) * 100
                    self.root.update()
                    
                    try:
                        # Navigate to patent page
                        page.goto(url, wait_until="networkidle")
                        time.sleep(2)  # Wait for page to fully load
                        
                        # Extract data using provided XPaths
                        data = self.extract_patent_data(page, url)
                        if data:
                            self.patent_data.append(data)
                            # Update treeview
                            self.root.after(0, self.update_treeview, data)
                    
                    except Exception as e:
                        print(f"Error scraping {url}: {str(e)}")
                        self.progress_var.set(f"Error with patent {i+1}: {str(e)}")
                        self.root.update()
                        time.sleep(2)
                
                browser.close()
                
                # Update UI when done
                self.root.after(0, self.scraping_completed)
                
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Scraping failed: {str(e)}"))
            self.root.after(0, self.scraping_completed)
    
    def extract_patent_data(self, page, url):
        data = {"URL": url}
        
        # XPaths as provided
        xpaths = {
            "Title": '//*[@id="title"]',
            "Abstract": '/html/body/search-app/search-result/search-ui/div/div/div/div/div/result-container/patent-result/div/div/div/div[1]/div[1]/section[1]/patent-text/div/section/abstract/div',
            "Publication Date": '/html/body/search-app/search-result/search-ui/div/div/div/div/div/result-container/patent-result/div/div/div/div[1]/div[2]/section/application-timeline/div/div[7]/div[1]',
            "Assignee": '/html/body/search-app/search-result/search-ui/div/div/div/div/div/result-container/patent-result/div/div/div/div[1]/div[2]/section/dl[1]/dd[2]',
            "Inventor": '//*[@id="link"]'
        }
        
        # Extract patent/publication number from URL
        patent_match = re.search(r'/patent/([^/]+)', url)
        data["Patent Number"] = patent_match.group(1) if patent_match else "Unknown"
        
        # Extract data using XPaths
        for field, xpath in xpaths.items():
            try:
                element = page.locator(xpath).first
                if element.is_visible():
                    text = element.inner_text().strip()
                    data[field] = text if text else "N/A"
                else:
                    data[field] = "N/A"
            except:
                data[field] = "N/A"
        
        # Try to click download button (optional)
        try:
            download_xpath = '/html/body/search-app/search-result/search-ui/div/div/div/div/div/result-container/patent-result/div/div/div/div[1]/div[2]/section/header/div/a'
            download_btn = page.locator(download_xpath).first
            if download_btn.is_visible():
                data["Download Available"] = "Yes"
                # Optionally click the download button
                # download_btn.click()
            else:
                data["Download Available"] = "No"
        except:
            data["Download Available"] = "No"
        
        return data
    
    def update_treeview(self, data):
        # Insert data into treeview
        values = (
            data.get("Patent Number", "N/A"),
            data.get("Title", "N/A")[:50] + "..." if len(data.get("Title", "")) > 50 else data.get("Title", "N/A"),
            data.get("Assignee", "N/A"),
            data.get("Inventor", "N/A"),
            data.get("Publication Date", "N/A")
        )
        self.tree.insert("", tk.END, values=values)
    
    def scraping_completed(self):
        self.progress_var.set(f"Completed! Scraped {len(self.patent_data)} patents.")
        self.progress_bar['value'] = 100
        
        # Re-enable buttons
        self.scrape_btn.config(state="normal")
        if self.patent_data:
            self.export_btn.config(state="normal")
            self.view_data_btn.config(state="normal")
    
    def export_to_excel(self):
        if not self.patent_data:
            messagebox.showerror("Error", "No data to export")
            return
        
        # Ask user for save location
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save patent data as..."
        )
        
        if filename:
            try:
                df = pd.DataFrame(self.patent_data)
                df.to_excel(filename, index=False)
                messagebox.showinfo("Success", f"Data exported successfully to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def view_data(self):
        if not self.patent_data:
            messagebox.showerror("Error", "No data to view")
            return
        
        # Create a new window to display all data
        view_window = tk.Toplevel(self.root)
        view_window.title("Patent Data Details")
        view_window.geometry("800x600")
        
        # Create text widget with scrollbar
        text_frame = ttk.Frame(view_window, padding="10")
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Display data
        for i, patent in enumerate(self.patent_data, 1):
            text_widget.insert(tk.END, f"=== PATENT {i} ===\n\n")
            for key, value in patent.items():
                text_widget.insert(tk.END, f"{key}: {value}\n")
            text_widget.insert(tk.END, "\n" + "="*50 + "\n\n")
        
        text_widget.config(state=tk.DISABLED)

def main():
    root = tk.Tk()
    app = PatentScraperApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()