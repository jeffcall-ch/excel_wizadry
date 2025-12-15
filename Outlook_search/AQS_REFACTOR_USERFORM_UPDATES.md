# UserForm Updates for AQS Refactor

## Overview

The search architecture has been completely refactored from DASL (`Application.AdvancedSearch`) to AQS (`Explorer.Search`). This requires adding a "Capture Results" button to the UserForm and updating event handlers.

---

## Required Form Controls

### **New Controls to Add:**

1. **btnCaptureResults** (CommandButton)
   - **Caption**: `Capture Results`
   - **Position**: Place next to or below the Search button
   - **Initial State**: `Enabled = False` (enable after search completes)
   
2. **lblStatus** (Label)
   - **Caption**: `` (empty initially)
   - **Position**: At bottom of form, full width
   - **AutoSize**: `False`
   - **Width**: Match form width
   - **ForeColor**: `&H00000000&` (black, will change dynamically)
   - **Font**: Regular, Size 9

---

## Button State Management

### **Initial State (Form Load):**
```vba
Private Sub UserForm_Initialize()
    btnCaptureResults.Enabled = False
    btnExport.Enabled = False
    btnBulkActions.Enabled = False
    btnStatistics.Enabled = False
    lblStatus.Caption = "Ready to search"
    lblStatus.ForeColor = vbBlack
End Sub
```

### **After Search Click:**
```vba
btnCaptureResults.Enabled = True
btnExport.Enabled = False
btnBulkActions.Enabled = False
btnStatistics.Enabled = False
lblStatus.Caption = "Search launched. Review results, then click 'Capture Results'."
lblStatus.ForeColor = vbBlue
```

### **After Capture Results:**
```vba
btnExport.Enabled = True
btnBulkActions.Enabled = True
btnStatistics.Enabled = True
lblStatus.Caption = "Captured X results. Export and bulk actions now available."
lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
```

---

## Updated Event Handlers

### **1. Search Button Click Handler**

Replace the existing `btnSearch_Click` handler with:

```vba
Private Sub btnSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim criteria As SearchCriteria
    
    ' Validate form inputs (reuse existing validation)
    If Not ValidateForm() Then Exit Sub
    
    ' Populate criteria from form controls (reuse existing code)
    criteria = GetSearchCriteriaFromForm()
    
    ' Execute search (NEW - calls ThisOutlookSession.ExecuteEmailSearch)
    ThisOutlookSession.ExecuteEmailSearch criteria
    
    ' Update UI state (NEW)
    lblStatus.Caption = "Search launched in Outlook. Review results, then click 'Capture Results'."
    lblStatus.ForeColor = vbBlue
    btnCaptureResults.Enabled = True
    btnExport.Enabled = False
    btnBulkActions.Enabled = False
    btnStatistics.Enabled = False
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error executing search: " & Err.Description, vbCritical, "Search Error"
    lblStatus.Caption = "Search failed: " & Err.Description
    lblStatus.ForeColor = vbRed
End Sub
```

### **2. NEW Capture Results Button Handler**

Add this new handler:

```vba
Private Sub btnCaptureResults_Click()
    On Error GoTo ErrorHandler
    
    Dim resultCount As Long
    Dim capturedResults As Collection
    
    ' Show progress
    lblStatus.Caption = "Capturing results..."
    lblStatus.ForeColor = vbBlue
    DoEvents
    
    ' Capture from search Explorer
    Set capturedResults = ThisOutlookSession.CaptureCurrentSearchResults()
    resultCount = capturedResults.Count
    
    If resultCount = 0 Then
        lblStatus.Caption = "No results captured. Search may still be running."
        lblStatus.ForeColor = vbRed
        MsgBox "No results were captured." & vbCrLf & vbCrLf & _
               "If the search is still running, wait a moment and try again.", _
               vbInformation, "No Results"
        Exit Sub
    End If
    
    ' Update UI state - Enable export and bulk action buttons
    lblStatus.Caption = "Captured " & resultCount & " results. Export and bulk actions now available."
    lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
    btnExport.Enabled = True
    btnBulkActions.Enabled = True
    btnStatistics.Enabled = True
    
    ' Optional: Show quick confirmation
    MsgBox "Captured " & resultCount & " email(s)." & vbCrLf & vbCrLf & _
           "You can now use Export or Bulk Actions.", vbInformation, "Capture Complete"
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Capture failed: " & Err.Description
    lblStatus.ForeColor = vbRed
    MsgBox "Error capturing results: " & Err.Description, vbCritical, "Capture Error"
End Sub
```

### **3. Updated Export Button Handler**

Update to check for captured results:

```vba
Private Sub btnExport_Click()
    On Error GoTo ErrorHandler
    
    Dim capturedResults As Collection
    Set capturedResults = ThisOutlookSession.GetCapturedResults()
    
    ' Validate captured results exist
    If capturedResults Is Nothing Or capturedResults.Count = 0 Then
        MsgBox "No results captured. Click 'Capture Results' first.", vbExclamation, "No Results"
        Exit Sub
    End If
    
    ' Show export options dialog
    Dim exportChoice As VbMsgBoxResult
    exportChoice = MsgBox("Export to CSV (Yes) or Excel (No)?", vbYesNoCancel + vbQuestion, "Export Format")
    
    Select Case exportChoice
        Case vbYes
            ' Export to CSV
            ThisOutlookSession.ExportSearchResultsToCSV
        Case vbNo
            ' Export to Excel
            ThisOutlookSession.ExportSearchResultsToExcel
        Case vbCancel
            Exit Sub
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during export: " & Err.Description, vbCritical, "Export Error"
End Sub
```

### **4. Updated Bulk Actions Button Handler**

Update to check for captured results:

```vba
Private Sub btnBulkActions_Click()
    On Error GoTo ErrorHandler
    
    Dim capturedResults As Collection
    Set capturedResults = ThisOutlookSession.GetCapturedResults()
    
    ' Validate captured results exist
    If capturedResults Is Nothing Or capturedResults.Count = 0 Then
        MsgBox "No results captured. Click 'Capture Results' first.", vbExclamation, "No Results"
        Exit Sub
    End If
    
    ' Show action menu
    Dim action As String
    action = InputBox("Choose bulk action:" & vbCrLf & vbCrLf & _
                     "1 = Mark as Read" & vbCrLf & _
                     "2 = Flag Items" & vbCrLf & _
                     "3 = Move Items" & vbCrLf & vbCrLf & _
                     "Enter number (1-3):", "Bulk Action Menu")
    
    Select Case action
        Case "1"
            ThisOutlookSession.BulkMarkAsRead
        Case "2"
            ThisOutlookSession.BulkFlagItems
        Case "3"
            ThisOutlookSession.BulkMoveItems
        Case Else
            MsgBox "Invalid choice or cancelled.", vbInformation
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during bulk action: " & Err.Description, vbCritical, "Bulk Action Error"
End Sub
```

### **5. Updated Statistics Button Handler**

Update to check for captured results:

```vba
Private Sub btnStatistics_Click()
    On Error GoTo ErrorHandler
    
    Dim capturedResults As Collection
    Set capturedResults = ThisOutlookSession.GetCapturedResults()
    
    ' Validate captured results exist
    If capturedResults Is Nothing Or capturedResults.Count = 0 Then
        MsgBox "No results captured. Click 'Capture Results' first.", vbExclamation, "No Results"
        Exit Sub
    End If
    
    ' Show statistics
    ThisOutlookSession.ShowSearchStatistics
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error calculating statistics: " & Err.Description, vbCritical, "Statistics Error"
End Sub
```

---

## Helper Functions (Keep Existing)

### **Keep These Functions Unchanged:**

1. **ValidateForm()** - Validates form inputs before search
2. **GetSearchCriteriaFromForm()** - Populates SearchCriteria from form controls
3. Any other existing helper functions

---

## Layout Recommendation

```
+------------------------------------------+
| From: [_____________]  To: [_____________] |
| CC: [_______________]  Subject: [________] |
| Date From: [______]  Date To: [__________] |
| [ ] Unread Only  [ ] Important Only       |
| Attachment Filter: [___________________]  |
| Size Filter: [_________________________]  |
|                                          |
| [Search] [Capture Results] [Clear]      |
|                                          |
| [Export] [Bulk Actions] [Statistics]     |
|                                          |
| Status: Ready to search                  |
+------------------------------------------+
```

---

## Testing Checklist

After making these changes:

### **Basic Flow Test:**
1. ✅ Fill in search criteria
2. ✅ Click "Search" → New Outlook window opens with results
3. ✅ Review results in Outlook UI
4. ✅ Click "Capture Results" → Confirmation shows count
5. ✅ Click "Export" → CSV/Excel file created
6. ✅ Click "Bulk Actions" → Items modified successfully
7. ✅ Click "Statistics" → Stats display correctly

### **Error Handling Test:**
1. ✅ Click "Capture Results" before Search → Error message shown
2. ✅ Click "Export" before Capture → Error message shown
3. ✅ Close search window, then Capture → Graceful error
4. ✅ Search with no results, then Capture → 0 results message

### **UI State Test:**
1. ✅ On form load: Capture/Export/Bulk/Stats all disabled
2. ✅ After Search: Capture enabled, others disabled
3. ✅ After Capture: All buttons enabled
4. ✅ Status label updates with correct colors

---

## Color Coding for Status Label

```vba
' Black - Neutral/Ready
lblStatus.ForeColor = vbBlack

' Blue - In Progress
lblStatus.ForeColor = vbBlue

' Green - Success
lblStatus.ForeColor = RGB(0, 128, 0)

' Red - Error
lblStatus.ForeColor = vbRed
```

---

## Migration Notes

### **What to Keep:**
- All existing form controls (text boxes, checkboxes, dropdowns)
- All validation logic
- GetSearchCriteriaFromForm() function
- Form appearance and layout

### **What to Add:**
- `btnCaptureResults` button
- `lblStatus` label
- New event handlers

### **What to Update:**
- `btnSearch_Click` - Add UI state management
- `btnExport_Click` - Add captured results validation
- `btnBulkActions_Click` - Add captured results validation
- `btnStatistics_Click` - Add captured results validation

### **What to Remove:**
- Any direct references to `Search.Results` (old DASL architecture)
- Any `Application_AdvancedSearchComplete` event handlers in form

---

## Sample Complete Form Code Snippet

```vba
' In frmEmailSearch UserForm

Option Explicit

Private Sub UserForm_Initialize()
    ' Initialize button states
    btnCaptureResults.Enabled = False
    btnExport.Enabled = False
    btnBulkActions.Enabled = False
    btnStatistics.Enabled = False
    lblStatus.Caption = "Ready to search"
    lblStatus.ForeColor = vbBlack
End Sub

Private Sub btnSearch_Click()
    On Error GoTo ErrorHandler
    Dim criteria As SearchCriteria
    
    If Not ValidateForm() Then Exit Sub
    criteria = GetSearchCriteriaFromForm()
    
    ThisOutlookSession.ExecuteEmailSearch criteria
    
    lblStatus.Caption = "Search launched. Review results, then click 'Capture Results'."
    lblStatus.ForeColor = vbBlue
    btnCaptureResults.Enabled = True
    btnExport.Enabled = False
    btnBulkActions.Enabled = False
    btnStatistics.Enabled = False
    
    Exit Sub
ErrorHandler:
    MsgBox "Search error: " & Err.Description, vbCritical
    lblStatus.Caption = "Search failed: " & Err.Description
    lblStatus.ForeColor = vbRed
End Sub

Private Sub btnCaptureResults_Click()
    On Error GoTo ErrorHandler
    Dim resultCount As Long
    Dim capturedResults As Collection
    
    lblStatus.Caption = "Capturing results..."
    lblStatus.ForeColor = vbBlue
    DoEvents
    
    Set capturedResults = ThisOutlookSession.CaptureCurrentSearchResults()
    resultCount = capturedResults.Count
    
    If resultCount = 0 Then
        lblStatus.Caption = "No results captured"
        lblStatus.ForeColor = vbRed
        MsgBox "No results captured. Search may still be running.", vbInformation
        Exit Sub
    End If
    
    lblStatus.Caption = "Captured " & resultCount & " results"
    lblStatus.ForeColor = RGB(0, 128, 0)
    btnExport.Enabled = True
    btnBulkActions.Enabled = True
    btnStatistics.Enabled = True
    
    Exit Sub
ErrorHandler:
    lblStatus.Caption = "Capture failed"
    lblStatus.ForeColor = vbRed
    MsgBox "Capture error: " & Err.Description, vbCritical
End Sub

' ... other handlers as documented above ...
```

---

## FAQ

**Q: Why can't we auto-capture results?**
A: `Explorer.Search` is fire-and-forget with no completion event. User needs to verify results first.

**Q: How long should users wait before capturing?**
A: Typically 1-5 seconds depending on result count. UI will show results appearing.

**Q: Can users continue working while search runs?**
A: Yes! Search runs in separate Outlook window. User can review, sort, preview emails before capturing.

**Q: What if user closes the search window?**
A: They'll get an error when clicking Capture. They can run the search again.

**Q: Does this work with large result sets (10,000+ emails)?**
A: Yes! GetTable efficiently captures any size. Capture may take 5-15 seconds for very large sets.
