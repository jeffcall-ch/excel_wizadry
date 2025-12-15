# Outlook Advanced Email Search

A powerful VBA-based email search tool for Microsoft Outlook that leverages Windows Search index (AQS - Advanced Query Syntax) for fast, efficient searching across all your emails.

## 🎯 Features

- **Fast Search**: Uses Windows Search index (AQS) for instant results
- **Flexible Search Options**:
  - Search by sender (From), recipients (To/CC)
  - Search terms in Subject and/or Body
  - Date range filtering (dd.mm.yyyy format)
  - Attachment filters (has/no attachments, specific file types)
  - Folder scope (Current Folder, Inbox, Sent Items, All Folders, etc.)
- **Modern UI**: Compact, Outlook-inspired design (640x465px)
- **Quick Search**: Press Enter anywhere in form to start search
- **Result Capture**: Capture search results for further processing
- **Bulk Actions**: Mark as read, flag items, move to folder
- **Export Options**: Export results to CSV or Excel
- **Search Statistics**: Analyze captured results

## 📁 Files

| File | Description |
|------|-------------|
| `VBA_AQS_UserForm_SIMPLE.vba` | Main search form UI (optimized design) |
| `VBA_AQS_ThisOutlookSession.vba` | Backend search engine and utilities |
| `modSharedTypes.vba` | Shared data structures (SearchCriteria type) |
| `README.md` | This documentation |

## 🚀 Installation

### Step 1: Enable Developer Tab in Outlook

1. Open Outlook
2. Go to **File** → **Options** → **Customize Ribbon**
3. Check **Developer** in right column
4. Click **OK**

### Step 2: Open VBA Editor

1. Press **Alt + F11** or click **Developer** → **Visual Basic**
2. In the VBA editor, find **ThisOutlookSession** in Project Explorer (left panel)

### Step 3: Import Code Files

**Method A: Copy-Paste (Recommended)**

1. **ThisOutlookSession Module**:
   - Double-click **ThisOutlookSession** in Project Explorer
   - Open `VBA_AQS_ThisOutlookSession.vba` in a text editor
   - Select all content and copy
   - Paste into ThisOutlookSession code window (replace any existing code)

2. **Shared Types Module**:
   - Right-click on **Modules** folder → **Insert** → **Module**
   - Rename it to `modSharedTypes` (press F4, change name in Properties window)
   - Open `modSharedTypes.vba` in a text editor
   - Copy all content and paste into the module

3. **User Form**:
   - Right-click on **Forms** folder → **Insert** → **UserForm**
   - Delete the empty form that appears
   - Open `VBA_AQS_UserForm_SIMPLE.vba` in a text editor
   - Copy all content
   - Paste into a new code window in VBA editor (the form will auto-generate)
   - Make sure the form is named `frmEmailSearch` (F4 → Name property)

**Method B: Import Files**

1. In VBA Editor, go to **File** → **Import File**
2. Select `modSharedTypes.vba` → Opens as a new module
3. Select `VBA_AQS_UserForm_SIMPLE.vba` → Opens as a form
4. For ThisOutlookSession, copy-paste content (cannot be imported directly)

### Step 4: Save and Test

1. Press **Ctrl + S** to save
2. Close VBA Editor
3. Run the macro:
   - Press **Alt + F8**
   - Select **ShowEmailSearchForm**
   - Click **Run**

### Step 5: Create Quick Access Button (Optional)

1. In Outlook, click **Developer** → **Macros**
2. Select **ShowEmailSearchForm** → **Run** to test
3. Add to Quick Access Toolbar:
   - Right-click Quick Access Toolbar → **Customize Quick Access Toolbar**
   - Choose **Macros** from dropdown
   - Find **ShowEmailSearchForm**
   - Click **Add**
   - Click **OK**

## 📖 Usage Guide

### Basic Search

1. **Launch Form**: Click Quick Access button or press Alt+F8 → Run macro
2. **Enter Criteria**:
   - **From**: Sender name or email (e.g., "john.smith" or "john.smith@company.com")
   - **To**: Recipient name or email
   - **CC**: CC recipient name or email
   - **Search Terms**: Keywords to find (e.g., "invoice", "meeting notes")
   - **Search In**: Check "Subject" and/or "Body" to specify where to search
3. **Optional Filters**:
   - **Date From/To**: Use dd.mm.yyyy format (e.g., 15.12.2024)
   - **Attachments**: Select filter (Any, No attachments, Has attachments, specific types)
   - **Folder**: Choose search scope (Current Folder, Inbox, All Folders, etc.)
4. **Press Enter** or click **Search** to execute

### Search Tips

- **Wildcards**: Use `*` for partial matches (e.g., `wel*` finds "welcome", "wellness", "wellington")
- **Boolean Operators**: 
  - Use `AND` to require all terms: `invoice AND payment`
  - Use `OR` for alternatives: `invoice OR receipt`
- **Phrase Search**: Put quotes for exact phrases: `"quarterly report"`
- **Multiple Criteria**: Combine From/To/CC with search terms for precise filtering
- **Date Ranges**: Specify both dates to search a specific period

### Advanced Features

#### Capture Results

After search completes:
1. Review results in Outlook window
2. Click **Capture Results** button
3. Now you can use bulk actions and exports

#### Bulk Actions

Once results are captured:
- **Mark as Read**: Mark all results as read
- **Flag Items**: Flag all results for follow-up
- **Move Items**: Move all results to selected folder

#### Export Results

- **CSV Export**: Exports to Desktop as timestamped CSV file
- **Excel Export**: Creates formatted Excel workbook with AutoFilter
- Both include: Date, From, To, Subject, Size, Attachments, Unread, Flagged

#### Search Statistics

View analytics on captured results:
- Total count, unread count, flagged count
- Emails with attachments
- Total/average size
- Unique senders, most common sender

### Keyboard Shortcuts

- **Enter**: Start search (works from any field)
- **Esc**: Close form
- **Tab**: Navigate between fields
- **Alt + F11**: Open VBA Editor (to modify code)
- **Alt + F8**: Open Macros dialog

## 🔧 Configuration

### Customize Folder List

Edit the folder dropdown in `VBA_AQS_UserForm_SIMPLE.vba` (lines 399-406):

```vba
.AddItem "Current Folder"
.AddItem "Inbox"
.AddItem "Sent Items"
.AddItem "Deleted Items"
.AddItem "Drafts"
.AddItem "All Folders (including subfolders)"
.ListIndex = 5  ' Default to "All Folders"
```

Add your custom folders:
```vba
.AddItem "Project Alpha"  ' Custom folder name
```

### Change Default Folder

Modify line 406 in `VBA_AQS_UserForm_SIMPLE.vba`:
```vba
.ListIndex = 0  ' 0=Current Folder, 1=Inbox, 5=All Folders
```

### Adjust Form Size

Modify line 90 in `VBA_AQS_UserForm_SIMPLE.vba`:
```vba
Me.Width = 640   ' Form width in pixels
Me.Height = 465  ' Form height in pixels
```

### Customize Attachment Filters

Edit the attachment dropdown in `VBA_AQS_UserForm_SIMPLE.vba` (lines 412-420).

## 🐛 Troubleshooting

### Search Returns No Results

- **Check Windows Search Service**: Make sure Windows Search is running and indexing your mailbox
- **Verify Search Terms**: Try broader terms or remove wildcards
- **Check Date Format**: Use dd.mm.yyyy (e.g., 15.12.2024)
- **Review Debug Output**: Press Ctrl+G in VBA Editor to see Immediate Window with debug logs

### Form Doesn't Open

- **Verify Form Name**: Form must be named `frmEmailSearch`
- **Check for Errors**: Open VBA Editor (Alt+F11), look for red text indicating errors
- **Compile Code**: In VBA Editor, go to Debug → Compile VBAProject

### Search Opens New Window

This is expected behavior for "All Folders" scope. To search in current folder without opening new window, select a specific folder (Inbox, Sent Items, etc.) instead of "All Folders".

### Capture Results Fails

- **Wait for Search**: Ensure search has completed before capturing
- **Check Search Window**: The search window must remain open
- **Re-run Search**: If capture fails, run search again and try capturing immediately after it completes

### Macro Security Warning

1. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. Select **Macro Settings**
3. Choose **Notifications for all macros** (recommended) or **Enable all macros** (less secure)
4. Click **OK**, restart Outlook

## 📊 Technical Details

### Architecture

- **AQS (Advanced Query Syntax)**: Native Windows Search query language
- **Explorer.Search**: Outlook's built-in search method
- **Async Search**: Results appear in Outlook UI, not returned directly
- **EntryID Collection**: Captured results stored as EntryIDs for bulk operations

### Search Scope Options

| Scope | Value | Description |
|-------|-------|-------------|
| Current Folder Only | 0 | Search only the selected folder |
| Subfolders | 1 | Search folder and all subfolders |
| Current Store | 2 | Search entire mailbox |
| All Folders | 3 | Search all mailboxes/accounts |

### Date Format

- **Input**: dd.mm.yyyy (e.g., 15.12.2024)
- **AQS Format**: mm/dd/yyyy (automatically converted)
- **Validation**: ParseDDMMYYYY function validates and converts

### Attachment Tokens

Internal token format for attachment filters:
- `""` - No filter
- `"NOATT"` - No attachments
- `"HASATT"` - Has attachments (any)
- `"HASATT|PDF"` - Has PDF attachments
- `"HASATT|PDF|XLS|DOC"` - Has specific types

## 🔄 Version History

### v2.0.0 (2024-12-15)
- Complete refactor to AQS (Advanced Query Syntax)
- Optimized UI design (640x465px, fixed widths)
- Form-wide Enter key handling
- Current Folder search option
- Simplified search with From/To/CC + Search In checkboxes
- Enhanced folder resolution with fallbacks
- Comprehensive debug logging
- Bug fixes and performance improvements

## 📝 License

This is a personal utility script. Use and modify as needed for your organization.

## 🤝 Contributing

This is a personal project, but suggestions and improvements are welcome.

## 📧 Support

For issues or questions:
1. Check Troubleshooting section above
2. Review debug output in VBA Immediate Window (Ctrl+G)
3. Verify all installation steps were completed

## 🔗 Related Files

- **VBA_AQS_UserForm_SIMPLE.vba**: 641 lines - Main UI form with programmatic control generation
- **VBA_AQS_ThisOutlookSession.vba**: 1460 lines - Backend search engine, folder resolution, bulk actions, exports
- **modSharedTypes.vba**: Shared data structures used by both files

---

**Last Updated**: December 15, 2024  
**Version**: 2.0.0  
**Platform**: Microsoft Outlook (VBA)  
**Requirements**: Windows Search indexing enabled
