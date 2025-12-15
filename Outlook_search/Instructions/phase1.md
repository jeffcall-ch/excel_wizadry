I want to build a Python desktop application for searching Outlook emails with advanced features. This is Phase 1 - let's build the foundation.

REQUIREMENTS:

1. TECHNOLOGY STACK:
   - Python with tkinter for GUI
   - win32com.client for Outlook integration
   - Use ttk for modern-looking widgets
   - Organize code with proper OOP structure

2. GUI LAYOUT:
   Create a main window (800x600) with these sections:

   TOP SECTION - Search Criteria:
   - "From:" label + Entry field (partial email/name match)
   - "Search Terms:" label + Entry field (will support operators in Phase 2)
   - "Search In:" Checkbuttons for [Subject] [Body] (both checked by default)
   - "Date Range:" Two date entry fields (From/To) with format YYYY-MM-DD
   - "Folder:" Dropdown with options: [Inbox, Sent Items, All Folders]
   - [Search] button (primary) and [Clear] button

   MIDDLE SECTION - Results:
   - Treeview widget with columns: [Date | From | Subject | Attachments]
   - Make Date column 120px, From 150px, Subject flexible width
   - Show attachment icon (📎) or count if email has attachments
   - Status label showing "Found X emails" or "Ready to search"

   BOTTOM SECTION - Actions:
   - [Open Email] button - opens selected email in Outlook
   - [Export CSV] button - exports results
   - Progress bar (initially hidden, shown during search)

3. OUTLOOK INTEGRATION:
   - Create OutlookSearcher class with methods:
     * connect_to_outlook() - establish COM connection
     * get_folders() - return list of available folders
     * search_emails(criteria_dict) - main search function
     * open_email(entry_id) - open email in Outlook
   
   - Handle Outlook not running/installed gracefully
   - Store email EntryID for opening later

4. BASIC SEARCH LOGIC (Phase 1 - simple version):
   - From field: case-insensitive partial match on SenderEmailAddress or SenderName
   - Search terms: simple AND logic (all terms must exist)
   - Search location: honor Subject/Body checkboxes
   - Date range: filter by ReceivedTime
   - Folder: search in selected folder only

5. CODE ORGANIZATION:
   Create three files:
   - outlook_searcher.py - OutlookSearcher class (Outlook COM logic)
   - search_gui.py - GUI class with all widgets
   - main.py - entry point, creates window and starts app

6. ERROR HANDLING:
   - Try/except for Outlook connection
   - Show user-friendly error messages in messagebox
   - Handle non-email items in folders gracefully

7. NICE-TO-HAVES:
   - Use grid layout for clean alignment
   - Add tooltips to explain features
   - Disable [Open Email] button if no selection
   - Double-click on result row to open email
   - Remember window size/position

Please create the complete code for Phase 1. Make it modular so we can easily add advanced features in Phase 2.
