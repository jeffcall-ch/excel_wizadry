# Outlook VBA Email Search - Installation Guide

**This solution uses Outlook's native AdvancedSearch API for INSTANT results (3 seconds instead of 99 seconds)**

## Installation Steps

### 1. Enable Developer Tab in Outlook
1. Open Outlook
2. File → Options → Customize Ribbon
3. Check "Developer" on the right side
4. Click OK

### 2. Open VBA Editor
1. Press `Alt + F11` (or click Developer tab → Visual Basic)
2. You'll see the VBA editor window

### 3. Create the UserForm

**In VBA Editor:**
1. Insert → UserForm (creates UserForm1)
2. Right-click UserForm1 in Project Explorer → Properties
3. Set `(Name)` to `frmEmailSearch`
4. Set `Caption` to `Advanced Email Search`
5. Set `Width` to `400`
6. Set `Height` to `320`

**Add Controls to the Form:**

Use the Toolbox to add these controls (View → Toolbox if not visible):

| Control Type | Name | Caption/Text | Position (Left, Top, Width, Height) |
|-------------|------|--------------|-------------------------------------|
| Label | lblFrom | From (Sender): | 12, 12, 360, 15 |
| TextBox | txtFrom | | 12, 30, 360, 24 |
| Label | lblSubject | Subject/Keywords: | 12, 60, 360, 15 |
| TextBox | txtSubject | | 12, 78, 360, 24 |
| CheckBox | chkSearchBody | Search in message body (slower) | 12, 108, 360, 18 |
| Label | lblDateFrom | Date From: | 12, 132, 120, 15 |
| TextBox | txtDateFrom | | 12, 150, 120, 24 |
| Label | lblDateTo | Date To: | 150, 132, 120, 15 |
| TextBox | txtDateTo | | 150, 150, 120, 24 |
| Label | lblFolder | Folder: | 12, 180, 360, 15 |
| ComboBox | cboFolder | | 12, 198, 360, 24 |
| CommandButton | btnSearch | Search | 120, 240, 80, 30 |
| CommandButton | btnCancel | Cancel | 210, 240, 80, 30 |

### 4. Add Code to UserForm Module

**In VBA Editor:**
1. Double-click `frmEmailSearch` in Project Explorer
2. View → Code (or press F7)
3. Copy ALL code from `VBA_UserForm_Code.vba` and paste it

### 5. Add Code to ThisOutlookSession Module

**In VBA Editor:**
1. Double-click `ThisOutlookSession` in Project Explorer
2. Copy ALL code from `VBA_ThisOutlookSession_Code.vba` and paste it

### 6. Add a Quick Access Toolbar Button

**Back in Outlook (not VBA editor):**
1. Right-click the Quick Access Toolbar (top-left corner)
2. Customize Quick Access Toolbar
3. Choose commands from: **Macros**
4. Select `ShowEmailSearchForm`
5. Click **Add >>**
6. Click **Modify** → choose an icon (magnifying glass recommended)
7. Display Name: `Email Search`
8. Click OK

## Usage

1. Click the Email Search button in Quick Access Toolbar
2. Enter search criteria:
   - **From**: Sender name or email (partial match)
   - **Subject/Keywords**: Words to find in subject
   - **Search in body**: Check to search message body (slower)
   - **Date From/To**: Format: YYYY-MM-DD or MM/DD/YYYY
   - **Folder**: Select specific folder or "All Folders"
3. Click **Search**
4. Results appear in a Search Folder called "Email Search Results"
5. Double-click any result to open the email

## Features

✅ **Instant search** - Uses Windows Search index (3-5 seconds for all folders)  
✅ **Search all folders** - Recursively searches subfolders  
✅ **SMTP email detection** - Finds actual email addresses, not just display names  
✅ **Date range filtering** - Narrow results by date  
✅ **Results in native Outlook UI** - Full Outlook functionality (reply, forward, etc.)  
✅ **Persistent Search Folder** - Results stay until you search again  

## Troubleshooting

**"Condition is not valid" error:**
- Windows Search Service may not be running
- Open Services (Win+R → services.msc)
- Find "Windows Search" → Start it
- Set Startup Type to "Automatic"

**Search Folder not appearing:**
- Check "Search Folders" in Outlook navigation pane
- Folder is named "Email Search Results"

**Macro security warning:**
- File → Options → Trust Center → Trust Center Settings
- Macro Settings → Enable all macros (or sign your macros)
