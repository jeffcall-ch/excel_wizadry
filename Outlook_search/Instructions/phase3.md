This is Phase 3 - final polish and power user features.

CONTEXT: We have advanced search working from Phase 2. Now add:

1. SAVE/LOAD SEARCH PRESETS:
   - Add menu bar: File > [Save Search] [Load Search] [Manage Searches]
   - Store searches in JSON file: ~/.outlook_searcher/presets.json
   - Each preset stores:
     * Name (user-provided)
     * From field value
     * Search terms
     * All checkbox/radio states
     * Date range
     * Folder selection
   - Add dropdown in GUI: "Saved Searches:" [dropdown] [Load] [Delete]
   - Auto-populate form when preset selected

2. ADVANCED FILTERS UI:
   Add collapsible "Advanced Filters" section:
   - [☐] Unread only
   - [☐] Flagged/Important emails
   - [☐] Has categories/tags
   - [☐] To: field contains [Entry field]
   - [☐] CC: field contains [Entry field]
   - Size filter: [dropdown: Any, <100KB, 100KB-1MB, 1MB-10MB, >10MB]

3. IMPLEMENT ADVANCED FILTERS:
   - Check mailItem.UnRead property
   - Check mailItem.Importance (OlImportance.olImportanceHigh)
   - Check mailItem.FlagRequest for follow-up flags
   - Check mailItem.Categories for category names
   - Check mailItem.To and mailItem.CC fields
   - Check mailItem.Size for size filtering

4. RESULTS ENHANCEMENTS:
   - Add sorting: click column headers to sort
   - Add filtering: right-click column header for quick filters
   - Multi-select results (Ctrl+Click, Shift+Click)
   - Context menu: [Open] [Copy Subject] [Copy From Address] [Move to Folder]
   - Show results count: "Showing 50 of 234 results" with pagination

5. EXPORT FUNCTIONALITY:
   Enhance CSV export:
   - Include all metadata: Date, From, To, CC, Subject, Body preview, Attachments
   - Add export options dialog:
     * [☐] Include email body
     * [☐] Include headers
     * [☐] Save attachments to folder
   - Add "Export to Excel" option using openpyxl:
     * Formatted table with filters
     * Frozen header row
     * Auto-sized columns

6. BULK ACTIONS:
   Add buttons that work on selected emails:
   - [Mark as Read/Unread]
   - [Flag/Unflag]
   - [Move to Folder...] - shows folder picker
   - [Delete] - with confirmation
   - [Forward to...] - opens compose window

7. SEARCH HISTORY:
   - Keep last 10 searches in dropdown
   - Store in JSON file
   - Click to restore search parameters

8. KEYBOARD SHORTCUTS:
   - Ctrl+F: Focus search terms field
   - Ctrl+Enter: Execute search
   - Ctrl+E: Export results
   - Enter on result: Open email
   - Delete: Delete selected email (with confirm)
   - Ctrl+A: Select all results

9. PREFERENCES/SETTINGS:
   Add Settings dialog:
   - Default folder for searches
   - Results limit (default 500)
   - Date format preference
   - Theme: [Light/Dark] if using modern ttk themes
   - Auto-save search history (on/off)

10. STATISTICS:
    Add info panel showing:
    - Total emails in searched folders
    - Search execution time
    - Most common senders in results
    - Attachment statistics (if filtered)

11. ERROR HANDLING & LOGGING:
    - Add logging to file: ~/.outlook_searcher/app.log
    - Log search queries, errors, performance metrics
    - Add "Help" menu with:
      * About dialog (version info)
      * View logs
      * Report issue (opens email template)
      * Search syntax help (shows examples)

12. INSTALLER/PACKAGING:
    Provide instructions for:
    - Creating requirements.txt
    - PyInstaller command to create .exe
    - Creating desktop shortcut

Make the application production-ready with proper error handling, logging, and user-friendly messages. Add comments explaining complex logic.