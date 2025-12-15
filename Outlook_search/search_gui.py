"""
Outlook Search GUI - User interface for email searching
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from typing import Optional
import csv
from datetime import datetime
from outlook_searcher import OutlookSearcher


class SearchGUI:
    """
    Main GUI class for Outlook email search application.
    """
    
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook Email Search - Phase 1")
        self.root.geometry("1100x800")  # Wider to accommodate log panel
        
        # Initialize Outlook searcher with log callback
        self.searcher = OutlookSearcher(log_callback=self._log)
        self.search_results = []
        
        # Create GUI
        self._create_widgets()
        self._setup_layout()
        
        # Connect to Outlook
        self._connect_outlook()
        
    def _create_widgets(self):
        """Create all GUI widgets."""
        
        # Main container
        self.main_frame = ttk.Frame(self.root, padding="10")
        
        # Create left and right panels
        self.left_panel = ttk.Frame(self.main_frame)
        self.right_panel = ttk.LabelFrame(self.main_frame, text="Activity Log", padding="5")
        
        # TOP SECTION - Search Criteria (in left panel)
        self.search_frame = ttk.LabelFrame(self.left_panel, text="Search Criteria", padding="10")
        
        # From field
        ttk.Label(self.search_frame, text="From:").grid(row=0, column=0, sticky='w', pady=5)
        self.from_entry = ttk.Entry(self.search_frame, width=40)
        self.from_entry.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        self._add_tooltip(self.from_entry, "Enter partial email or name (e.g., 'john' or 'smith')")
        
        # Search Terms field
        ttk.Label(self.search_frame, text="Search Terms:").grid(row=1, column=0, sticky='w', pady=5)
        self.terms_entry = ttk.Entry(self.search_frame, width=40)
        self.terms_entry.grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        self._add_tooltip(self.terms_entry, "Enter search terms (all must be present)")
        
        # Search In checkboxes
        ttk.Label(self.search_frame, text="Search In:").grid(row=2, column=0, sticky='w', pady=5)
        self.search_in_frame = ttk.Frame(self.search_frame)
        self.search_in_frame.grid(row=2, column=1, sticky='w', pady=5, padx=5)
        
        self.search_subject_var = tk.BooleanVar(value=True)
        self.search_body_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(self.search_in_frame, text="Subject", variable=self.search_subject_var).pack(side='left', padx=5)
        ttk.Checkbutton(self.search_in_frame, text="Body", variable=self.search_body_var).pack(side='left', padx=5)
        
        # Date Range
        ttk.Label(self.search_frame, text="Date Range:").grid(row=3, column=0, sticky='w', pady=5)
        self.date_frame = ttk.Frame(self.search_frame)
        self.date_frame.grid(row=3, column=1, sticky='w', pady=5, padx=5)
        
        ttk.Label(self.date_frame, text="From:").pack(side='left', padx=2)
        self.date_from_entry = ttk.Entry(self.date_frame, width=12)
        self.date_from_entry.pack(side='left', padx=2)
        self._add_tooltip(self.date_from_entry, "Format: YYYY-MM-DD (e.g., 2024-01-01)")
        
        ttk.Label(self.date_frame, text="To:").pack(side='left', padx=2)
        self.date_to_entry = ttk.Entry(self.date_frame, width=12)
        self.date_to_entry.pack(side='left', padx=2)
        self._add_tooltip(self.date_to_entry, "Format: YYYY-MM-DD (e.g., 2024-12-31)")
        
        # Folder selection
        ttk.Label(self.search_frame, text="Folder:").grid(row=4, column=0, sticky='w', pady=5)
        self.folder_var = tk.StringVar(value="All Folders (including subfolders)")
        
        # Get dynamic folder list from Outlook
        folder_list = self.searcher.get_folders()
        
        self.folder_combo = ttk.Combobox(self.search_frame, textvariable=self.folder_var, 
                                         values=folder_list,
                                         state='readonly', width=37)
        self.folder_combo.grid(row=4, column=1, sticky='w', pady=5, padx=5)
        
        # Search and Clear buttons
        self.button_frame = ttk.Frame(self.search_frame)
        self.button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        self.search_button = ttk.Button(self.button_frame, text="Search", command=self._perform_search)
        self.search_button.pack(side='left', padx=5)
        
        self.clear_button = ttk.Button(self.button_frame, text="Clear", command=self._clear_form)
        self.clear_button.pack(side='left', padx=5)
        
        # Configure grid weights for search frame
        self.search_frame.columnconfigure(1, weight=1)
        
        # MIDDLE SECTION - Results (in left panel)
        self.results_frame = ttk.LabelFrame(self.left_panel, text="Search Results", padding="10")
        
        # Create Treeview
        columns = ('Folder', 'Date', 'From', 'Subject', 'Attachments')
        self.tree = ttk.Treeview(self.results_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        self.tree.heading('Folder', text='Folder')
        self.tree.heading('Date', text='Date')
        self.tree.heading('From', text='From')
        self.tree.heading('Subject', text='Subject')
        self.tree.heading('Attachments', text='Attachments')
        
        self.tree.column('Folder', width=120, minwidth=80)
        self.tree.column('Date', width=120, minwidth=120)
        self.tree.column('From', width=150, minwidth=100)
        self.tree.column('Subject', width=300, minwidth=200)
        self.tree.column('Attachments', width=100, minwidth=80)
        
        # Scrollbars
        vsb = ttk.Scrollbar(self.results_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # Configure grid weights for results frame
        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.rowconfigure(0, weight=1)
        
        # Status label
        self.status_label = ttk.Label(self.results_frame, text="Ready to search", relief='sunken')
        self.status_label.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(5, 0))
        
        # Bind double-click to open email
        self.tree.bind('<Double-1>', self._on_double_click)
        self.tree.bind('<<TreeviewSelect>>', self._on_selection_change)
        
        # BOTTOM SECTION - Actions (in left panel)
        self.actions_frame = ttk.LabelFrame(self.left_panel, text="Actions", padding="10")
        
        self.open_button = ttk.Button(self.actions_frame, text="Open Email", 
                                       command=self._open_selected_email, state='disabled')
        self.open_button.pack(side='left', padx=5)
        
        self.export_button = ttk.Button(self.actions_frame, text="Export CSV", 
                                        command=self._export_csv, state='disabled')
        self.export_button.pack(side='left', padx=5)
        
        # Search time status label (bottom left)
        self.status_label = ttk.Label(self.actions_frame, text="", 
                                       font=('Arial', 9), foreground='#666666')
        self.status_label.pack(side='left', padx=20)
        
        # Progress bar (initially hidden)
        self.progress = ttk.Progressbar(self.actions_frame, mode='indeterminate', length=200)
        
        # RIGHT PANEL - Activity Log
        self.log_text = scrolledtext.ScrolledText(
            self.right_panel,
            wrap=tk.WORD,
            width=40,
            height=35,
            font=('Consolas', 9),
            bg='#f0f0f0',
            fg='#000000'
        )
        self.log_text.pack(fill='both', expand=True)
        
        # Add clear log button
        self.clear_log_button = ttk.Button(self.right_panel, text="Clear Log", 
                                           command=self._clear_log)
        self.clear_log_button.pack(pady=5)
        
    def _setup_layout(self):
        """Setup the layout of all frames."""
        self.main_frame.pack(fill='both', expand=True)
        
        # Configure main_frame grid
        self.main_frame.columnconfigure(0, weight=3)  # Left panel gets more space
        self.main_frame.columnconfigure(1, weight=1)  # Right panel (log)
        self.main_frame.rowconfigure(0, weight=1)
        
        # Pack left panel contents
        self.left_panel.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        self.search_frame.pack(fill='x', pady=(0, 10))
        self.results_frame.pack(fill='both', expand=True, pady=(0, 10))
        self.actions_frame.pack(fill='x')
        
        # Pack right panel (log)
        self.right_panel.grid(row=0, column=1, sticky='nsew')
        
    def _add_tooltip(self, widget, text):
        """Add a simple tooltip to a widget."""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            label = ttk.Label(tooltip, text=text, background="#ffffe0", relief='solid', borderwidth=1)
            label.pack()
            widget.tooltip = tooltip
            
        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
                
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)
    
    def _log(self, message, level='INFO'):
        """Add a message to the activity log."""
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {level}: {message}\n"
        
        self.log_text.insert('end', log_entry)
        
        # Color code based on level
        if level == 'ERROR':
            # Find the last line and color it red
            self.log_text.tag_add('error', 'end-2l', 'end-1l')
            self.log_text.tag_config('error', foreground='red')
        elif level == 'SUCCESS':
            self.log_text.tag_add('success', 'end-2l', 'end-1l')
            self.log_text.tag_config('success', foreground='green')
        elif level == 'WARNING':
            self.log_text.tag_add('warning', 'end-2l', 'end-1l')
            self.log_text.tag_config('warning', foreground='orange')
        elif level == 'DEBUG':
            self.log_text.tag_add('debug', 'end-2l', 'end-1l')
            self.log_text.tag_config('debug', foreground='blue')
        
        # Auto-scroll to bottom
        self.log_text.see('end')
        self.root.update()
    
    def _clear_log(self):
        """Clear the activity log."""
        self.log_text.delete('1.0', 'end')
        self._log("Log cleared", "INFO")
    
    def _connect_outlook(self):
        """Connect to Outlook on startup."""
        self._log("Attempting to connect to Outlook...", "INFO")
        success, message = self.searcher.connect_to_outlook()
        if not success:
            self._log(f"Connection failed: {message}", "ERROR")
            messagebox.showerror("Connection Error", message)
            self.status_label.config(text="Failed to connect to Outlook")
        else:
            self._log(message, "SUCCESS")
            self.status_label.config(text="Connected to Outlook - Ready to search")
            # Update folder list
            folders = self.searcher.get_folders()
            if folders:
                self.folder_combo['values'] = folders
                self._log(f"Available folders: {', '.join(folders)}", "INFO")
    
    def _perform_search(self):
        """Execute the search based on form inputs."""
        self._log("=" * 40, "INFO")
        self._log("Starting new search...", "INFO")
        
        # Validate at least one search criterion
        if not any([
            self.from_entry.get().strip(),
            self.terms_entry.get().strip(),
            self.date_from_entry.get().strip(),
            self.date_to_entry.get().strip()
        ]):
            self._log("No search criteria provided", "WARNING")
            messagebox.showwarning("Search Criteria", 
                                   "Please enter at least one search criterion.")
            return
        
        # Validate date format if provided
        date_from = self.date_from_entry.get().strip()
        date_to = self.date_to_entry.get().strip()
        
        if date_from:
            try:
                datetime.strptime(date_from, '%Y-%m-%d')
                self._log(f"Date from: {date_from}", "DEBUG")
            except ValueError:
                self._log(f"Invalid date format: {date_from}", "ERROR")
                messagebox.showerror("Invalid Date", 
                                    "From date must be in YYYY-MM-DD format")
                return
        
        if date_to:
            try:
                datetime.strptime(date_to, '%Y-%m-%d')
                self._log(f"Date to: {date_to}", "DEBUG")
            except ValueError:
                self._log(f"Invalid date format: {date_to}", "ERROR")
                messagebox.showerror("Invalid Date", 
                                    "To date must be in YYYY-MM-DD format")
                return
        
        # Build criteria dictionary
        criteria = {
            'from_text': self.from_entry.get().strip(),
            'search_terms': self.terms_entry.get().strip(),
            'search_subject': self.search_subject_var.get(),
            'search_body': self.search_body_var.get(),
            'date_from': date_from,
            'date_to': date_to,
            'folder': self.folder_var.get()
        }
        
        # Log search criteria
        self._log("Search criteria:", "INFO")
        self._log(f"  From: '{criteria['from_text']}'", "DEBUG")
        self._log(f"  Search Terms: '{criteria['search_terms']}'", "DEBUG")
        self._log(f"  Search in Subject: {criteria['search_subject']}", "DEBUG")
        self._log(f"  Search in Body: {criteria['search_body']}", "DEBUG")
        self._log(f"  Date Range: {criteria['date_from']} to {criteria['date_to']}", "DEBUG")
        self._log(f"  Folder: {criteria['folder']}", "DEBUG")
        
        # Clear previous results
        self._clear_results()
        
        # Show progress
        self._log("Executing search...", "INFO")
        self.status_label.config(text="Searching...")
        self.search_button.config(state='disabled')
        self.progress.pack(side='left', padx=5)
        self.progress.start(10)
        self.root.update()
        
        import time
        search_start = time.time()
        
        try:
            # Perform search
            self._log("Calling searcher.search_emails()...", "DEBUG")
            self.search_results = self.searcher.search_emails(criteria)
            
            search_time = time.time() - search_start
            self._log(f"Search completed. Found {len(self.search_results)} emails", "SUCCESS")
            
            # Display results
            for email in self.search_results:
                self.tree.insert('', 'end', values=(
                    email.get('folder_name', 'Unknown'),
                    email['date'],
                    email['from'],
                    email['subject'],
                    email['attachments']
                ), tags=(email['entry_id'],))
            
            self._log(f"Displayed {len(self.search_results)} emails in results", "INFO")
            
            # Update status with search time
            count = len(self.search_results)
            status_text = f"⏱️ {search_time:.2f}s | Found {count} email{'s' if count != 1 else ''}"
            if count >= 500:
                status_text += " (limited to 500)"
                self._log("Results limited to 500 emails", "WARNING")
            self.status_label.config(text=status_text)
            
            # Enable export button if results found
            if count > 0:
                self.export_button.config(state='normal')
            
        except Exception as e:
            error_msg = f"An error occurred during search:\n{str(e)}"
            self._log(f"Search error: {str(e)}", "ERROR")
            messagebox.showerror("Search Error", error_msg)
            search_time = time.time() - search_start
            self.status_label.config(text=f"⏱️ {search_time:.2f}s | Search failed")
        finally:
            # Hide progress
            self.progress.stop()
            self.progress.pack_forget()
            self.search_button.config(state='normal')
            self._log("Search operation completed", "INFO")
    
    def _clear_results(self):
        """Clear the results tree."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.search_results = []
        self.export_button.config(state='disabled')
    
    def _clear_form(self):
        """Clear all search fields."""
        self.from_entry.delete(0, 'end')
        self.terms_entry.delete(0, 'end')
        self.date_from_entry.delete(0, 'end')
        self.date_to_entry.delete(0, 'end')
        self.search_subject_var.set(True)
        self.search_body_var.set(True)
        self.folder_var.set("Inbox")
        self._clear_results()
        self.status_label.config(text="Ready to search")
    
    def _on_selection_change(self, event):
        """Handle tree selection change."""
        selection = self.tree.selection()
        if selection:
            self.open_button.config(state='normal')
        else:
            self.open_button.config(state='disabled')
    
    def _on_double_click(self, event):
        """Handle double-click on tree item."""
        self._open_selected_email()
    
    def _open_selected_email(self):
        """Open the selected email in Outlook."""
        selection = self.tree.selection()
        if not selection:
            return
        
        # Get the EntryID from the item's tags
        item = selection[0]
        tags = self.tree.item(item, 'tags')
        if not tags:
            messagebox.showerror("Error", "Could not retrieve email ID")
            return
        
        entry_id = tags[0]
        success, message = self.searcher.open_email(entry_id)
        
        if not success:
            messagebox.showerror("Open Email Error", message)
    
    def _export_csv(self):
        """Export search results to CSV."""
        if not self.search_results:
            messagebox.showwarning("No Results", "No search results to export")
            return
        
        # Ask user for file location
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"outlook_search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        if not filename:
            return
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                # Write header
                writer.writerow(['Folder', 'Date', 'From', 'From Email', 'Subject', 'Attachment Count', 'Attachment Names'])
                
                # Write data
                for email in self.search_results:
                    writer.writerow([
                        email.get('folder_name', 'Unknown'),
                        email['date'],
                        email['from'],
                        email['from_email'],
                        email['subject'],
                        email['attachment_count'],
                        email['attachment_names']
                    ])
            
            messagebox.showinfo("Export Success", 
                              f"Exported {len(self.search_results)} emails to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export CSV:\n{str(e)}")
