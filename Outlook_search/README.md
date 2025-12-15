# Outlook Email Search - Phase 1

A Python desktop application for searching Outlook emails with advanced filtering capabilities.

## ⚠️ IMPORTANT - READ-ONLY APPLICATION
**This application is 100% READ-ONLY. It will NEVER modify, delete, move, or alter any Outlook data in any way.**

## Features (Phase 1)

- **Search by Sender**: Partial match on email address or name
- **Full-text Search**: Search in subject and/or body with AND logic
- **Date Range Filtering**: Filter emails by date range
- **Folder Selection**: Search in Inbox, Sent Items, or All Folders
- **Results Display**: View results in a sortable table with attachment indicators
- **Open in Outlook**: Double-click or use button to view emails in Outlook
- **Export to CSV**: Save search results for further analysis

## Installation

### Prerequisites
- Windows OS
- Microsoft Outlook installed and configured
- Python 3.7 or higher

### Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python main.py
```

## Usage

1. **Connect to Outlook**: The application automatically connects when launched
2. **Enter Search Criteria**:
   - **From**: Enter partial sender name or email (e.g., "john" or "@company.com")
   - **Search Terms**: Enter words to search for (all must be present)
   - **Search In**: Choose to search in Subject and/or Body
   - **Date Range**: Optional date filtering (format: YYYY-MM-DD)
   - **Folder**: Select which folder(s) to search
3. **Click Search**: Results will appear in the table below
4. **View Results**:
   - Double-click any email to open it in Outlook
   - Single-click to select, then use "Open Email" button
5. **Export**: Click "Export CSV" to save results to a file

## File Structure

```
Outlook_search/
├── main.py                 # Application entry point
├── search_gui.py          # GUI implementation
├── outlook_searcher.py    # Outlook COM interface (READ-ONLY)
├── requirements.txt       # Python dependencies
└── README.md             # This file
```

## Safety Features

- **No Write Operations**: The application only reads data from Outlook
- **No Email Modifications**: Cannot delete, move, or modify emails
- **No Send Capability**: Cannot send or forward emails
- **Display Only**: Opening emails only displays them in Outlook

## Search Tips

- Use partial names for broader results (e.g., "smith" matches "John Smith", "Jane Smithson")
- Leave fields empty to skip that filter
- Combine multiple criteria for more precise results
- Search results are limited to 500 emails for performance

## Troubleshooting

**"Failed to connect to Outlook"**
- Ensure Microsoft Outlook is installed
- Try opening Outlook manually first
- Check if Outlook is configured with an email account

**"No results found"**
- Try broadening your search criteria
- Check date range is not too restrictive
- Verify you're searching in the correct folder

**Slow searches**
- Searching in "All Folders" with body text can be slow
- Try limiting to specific folders
- Use date ranges to narrow the search

## Next Phases (Planned)

- **Phase 2**: Advanced operators (AND, OR, NOT), proximity search, attachment filtering
- **Phase 3**: Save search presets, bulk actions, advanced filters

## License

Internal tool - use at your own discretion.

## Version

Phase 1 - v1.0.0 - December 2025
