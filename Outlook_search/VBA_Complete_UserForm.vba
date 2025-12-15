' ============================================
' Module: frmEmailSearch (UserForm Code)
' Purpose: Complete email search with ALL Phase 1-3 features
' Version: 1.1 - Refactored with enhanced error handling
' ============================================

Option Explicit

' Constants
Private Const MAX_HISTORY_ITEMS As Long = 10
Private Const MAX_HISTORY_TEXT_LENGTH As Long = 100

' Module-level variables
Private m_SearchHistory As Collection
Private m_LastSearchResults As Outlook.Search

' ============================================
' Form Initialization
' ============================================

Private Sub UserForm_Initialize()
    On Error Resume Next  ' Allow form to load even if controls don't exist yet
    
    ' Populate folder dropdown
    With Me.cboFolder
        .Clear
        .AddItem "All Folders (including subfolders)"
        .AddItem "---"
        .AddItem "Inbox"
        .AddItem "Sent Items"
        .AddItem "---"
        .AddItem "01 Projects"
        .AddItem "02 Office"
        .AddItem "99 Misc"
        .AddItem "Archive"
        .AddItem "Deleted Items"
        .AddItem "Drafts"
        .ListIndex = 0  ' Default to "All Folders"
    End With
    
    ' Populate attachment type dropdown
    With Me.cboAttachmentFilter
        .Clear
        .AddItem "Any"
        .AddItem "With Attachments"
        .AddItem "Without Attachments"
        .AddItem "PDF files"
        .AddItem "Excel files (.xls, .xlsx)"
        .AddItem "Word files (.doc, .docx)"
        .AddItem "Images (.jpg, .png, .gif)"
        .AddItem "PowerPoint files (.ppt, .pptx)"
        .AddItem "ZIP/Archives"
        .ListIndex = 0  ' Default to "Any"
    End With
    
    ' Populate size filter dropdown
    With Me.cboSizeFilter
        .Clear
        .AddItem "Any Size"
        .AddItem "< 100 KB"
        .AddItem "100 KB - 1 MB"
        .AddItem "1 MB - 10 MB"
        .AddItem "> 10 MB"
        .ListIndex = 0
    End With
    
    ' Set default values
    Me.chkSearchSubject.Value = True
    Me.chkSearchBody.Value = False
    Me.chkCaseSensitive.Value = False
    Me.chkUnreadOnly.Value = False
    Me.chkImportantOnly.Value = False
    Me.txtDateFrom.Text = ""
    Me.txtDateTo.Text = ""
    Me.txtTo.Text = ""
    Me.txtCc.Text = ""
    
    ' Load search history
    LoadSearchHistory
    
    ' Update status
    Me.lblStatus.Caption = "Ready - Use Ctrl+F to focus search, Ctrl+Enter to search, Esc to close"
    
    On Error GoTo 0
End Sub

' ============================================
' Search Button Click
' ============================================

Private Sub btnSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim criteria As SearchCriteria
    
    ' Validate that at least one search criterion is provided
    If Trim(Me.txtFrom.Text) = "" And _
       Trim(Me.txtSearchTerms.Text) = "" And _
       Trim(Me.txtDateFrom.Text) = "" And _
       Trim(Me.txtDateTo.Text) = "" And _
       Trim(Me.txtTo.Text) = "" And _
       Trim(Me.txtCc.Text) = "" And _
       Me.cboAttachmentFilter.ListIndex = 0 And _
       Me.cboSizeFilter.ListIndex = 0 And _
       Not Me.chkUnreadOnly.Value And _
       Not Me.chkImportantOnly.Value Then
        MsgBox "Please enter at least one search criterion.", vbExclamation, "Search Criteria Required"
        Exit Sub
    End If
    
    ' Validate date formats if provided
    If Trim(Me.txtDateFrom.Text) <> "" Then
        If Not IsDate(Me.txtDateFrom.Text) Then
            MsgBox "Please enter a valid date in the 'From Date' field (e.g., 2024-01-01).", vbExclamation, "Invalid Date"
            Me.txtDateFrom.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(Me.txtDateTo.Text) <> "" Then
        If Not IsDate(Me.txtDateTo.Text) Then
            MsgBox "Please enter a valid date in the 'To Date' field (e.g., 2024-12-31).", vbExclamation, "Invalid Date"
            Me.txtDateTo.SetFocus
            Exit Sub
        End If
    End If
    
    ' Build search criteria
    With criteria
        .FromAddress = Trim(Me.txtFrom.Text)
        .SearchTerms = Trim(Me.txtSearchTerms.Text)
        .SearchSubject = (Me.chkSearchSubject.Value = True)
        .SearchBody = (Me.chkSearchBody.Value = True)
        .DateFrom = Trim(Me.txtDateFrom.Text)
        .DateTo = Trim(Me.txtDateTo.Text)
        .FolderName = Me.cboFolder.Text
        .CaseSensitive = Me.chkCaseSensitive.Value
        .ToAddress = Trim(Me.txtTo.Text)
        .CcAddress = Trim(Me.txtCc.Text)
        .AttachmentFilter = Me.cboAttachmentFilter.Text
        .SizeFilter = Me.cboSizeFilter.Text
        .UnreadOnly = Me.chkUnreadOnly.Value
        .ImportantOnly = Me.chkImportantOnly.Value
    End With
    
    ' Update status
    Me.lblStatus.Caption = "Searching... (using Windows Search index)"
    DoEvents
    
    ' Save to search history
    SaveToSearchHistory criteria
    
    ' Execute the search
    ExecuteEmailSearch criteria
    
    ' Keep form open to show results and allow bulk actions
    Me.lblStatus.Caption = "Search completed - See Search Folder for results"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during search: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Search Error"
    Me.lblStatus.Caption = "Error occurred - Please try again"
End Sub

' ============================================
' Clear Button Click
' ============================================

Private Sub btnClear_Click()
    On Error GoTo ErrorHandler
    
    ' Clear all search fields
    Me.txtFrom.Text = ""
    Me.txtSearchTerms.Text = ""
    Me.txtTo.Text = ""
    Me.txtCc.Text = ""
    Me.txtDateFrom.Text = ""
    Me.txtDateTo.Text = ""
    Me.chkSearchSubject.Value = True
    Me.chkSearchBody.Value = False
    Me.chkCaseSensitive.Value = False
    Me.chkUnreadOnly.Value = False
    Me.chkImportantOnly.Value = False
    Me.cboFolder.ListIndex = 0
    Me.cboAttachmentFilter.ListIndex = 0
    Me.cboSizeFilter.ListIndex = 0
    Me.lblStatus.Caption = "Cleared - Ready for new search"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error clearing form: " & Err.Description, vbExclamation, "Clear Error"
End Sub

' ============================================
' Preset Management
' ============================================

Private Sub btnSavePreset_Click()
    On Error GoTo ErrorHandler
    
    Dim presetName As String
    Dim criteria As SearchCriteria
    
    ' Ask for preset name
    presetName = InputBox("Enter a name for this search preset:", "Save Search Preset", "My Search")
    If Trim(presetName) = "" Then Exit Sub
    
    ' Build criteria
    With criteria
        .FromAddress = Trim(Me.txtFrom.Text)
        .SearchTerms = Trim(Me.txtSearchTerms.Text)
        .SearchSubject = Me.chkSearchSubject.Value
        .SearchBody = Me.chkSearchBody.Value
        .DateFrom = Trim(Me.txtDateFrom.Text)
        .DateTo = Trim(Me.txtDateTo.Text)
        .FolderName = Me.cboFolder.Text
        .CaseSensitive = Me.chkCaseSensitive.Value
        .ToAddress = Trim(Me.txtTo.Text)
        .CcAddress = Trim(Me.txtCc.Text)
        .AttachmentFilter = Me.cboAttachmentFilter.Text
        .SizeFilter = Me.cboSizeFilter.Text
        .UnreadOnly = Me.chkUnreadOnly.Value
        .ImportantOnly = Me.chkImportantOnly.Value
    End With
    
    ' Save preset (this calls function in ThisOutlookSession)
    SaveSearchPreset presetName, criteria
    Me.lblStatus.Caption = "Preset '" & presetName & "' saved successfully"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving preset: " & Err.Description, vbCritical, "Save Error"
    Me.lblStatus.Caption = "Error saving preset"
End Sub

Private Sub btnLoadPreset_Click()
    On Error GoTo ErrorHandler
    
    Dim presetName As String
    Dim criteria As SearchCriteria
    
    ' Ask which preset to load
    presetName = InputBox("Enter the preset name to load:", "Load Search Preset")
    If Trim(presetName) = "" Then Exit Sub
    
    ' Load preset (this calls function in ThisOutlookSession)
    If LoadSearchPreset(presetName, criteria) Then
        ' Populate form with loaded criteria
        Me.txtFrom.Text = criteria.FromAddress
        Me.txtSearchTerms.Text = criteria.SearchTerms
        Me.chkSearchSubject.Value = criteria.SearchSubject
        Me.chkSearchBody.Value = criteria.SearchBody
        Me.txtDateFrom.Text = criteria.DateFrom
        Me.txtDateTo.Text = criteria.DateTo
        Me.cboFolder.Text = criteria.FolderName
        Me.chkCaseSensitive.Value = criteria.CaseSensitive
        Me.txtTo.Text = criteria.ToAddress
        Me.txtCc.Text = criteria.CcAddress
        Me.cboAttachmentFilter.Text = criteria.AttachmentFilter
        Me.cboSizeFilter.Text = criteria.SizeFilter
        Me.chkUnreadOnly.Value = criteria.UnreadOnly
        Me.chkImportantOnly.Value = criteria.ImportantOnly
        
        Me.lblStatus.Caption = "Preset '" & presetName & "' loaded successfully"
    Else
        MsgBox "Preset '" & presetName & "' not found.", vbExclamation, "Load Preset"
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading preset: " & Err.Description, vbCritical, "Load Error"
    Me.lblStatus.Caption = "Error loading preset"
End Sub

' ============================================
' Bulk Actions
' ============================================

Private Sub btnMarkRead_Click()
    On Error GoTo ErrorHandler
    
    If m_LastSearchResults Is Nothing Then
        MsgBox "Please perform a search first before using bulk actions.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    BulkMarkAsRead m_LastSearchResults
    Me.lblStatus.Caption = "Bulk operation completed - Marked emails as read"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during bulk mark as read: " & Err.Description, vbCritical, "Bulk Action Error"
    Me.lblStatus.Caption = "Error during bulk action"
End Sub

Private Sub btnFlagItems_Click()
    On Error GoTo ErrorHandler
    
    If m_LastSearchResults Is Nothing Then
        MsgBox "Please perform a search first before using bulk actions.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    BulkFlagItems m_LastSearchResults
    Me.lblStatus.Caption = "Bulk operation completed - Flagged emails"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during bulk flag: " & Err.Description, vbCritical, "Bulk Action Error"
    Me.lblStatus.Caption = "Error during bulk action"
End Sub

Private Sub btnMoveItems_Click()
    On Error GoTo ErrorHandler
    
    If m_LastSearchResults Is Nothing Then
        MsgBox "Please perform a search first before using bulk actions.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    BulkMoveItems m_LastSearchResults
    Me.lblStatus.Caption = "Bulk operation completed - Moved emails"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during bulk move: " & Err.Description, vbCritical, "Bulk Action Error"
    Me.lblStatus.Caption = "Error during bulk action"
End Sub

' ============================================
' Export Functions
' ============================================

Private Sub btnExportCSV_Click()
    On Error GoTo ErrorHandler
    
    If m_LastSearchResults Is Nothing Then
        MsgBox "Please perform a search first before exporting.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    ExportSearchResultsToCSV m_LastSearchResults
    Me.lblStatus.Caption = "Export completed - CSV file saved"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during CSV export: " & Err.Description, vbCritical, "Export Error"
    Me.lblStatus.Caption = "Error during export"
End Sub

Private Sub btnExportExcel_Click()
    On Error GoTo ErrorHandler
    
    If m_LastSearchResults Is Nothing Then
        MsgBox "Please perform a search first before exporting.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    ExportSearchResultsToExcel m_LastSearchResults
    Me.lblStatus.Caption = "Export completed - Excel file saved"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during Excel export: " & Err.Description, vbCritical, "Export Error"
    Me.lblStatus.Caption = "Error during export"
End Sub

' ============================================
' Statistics
' ============================================

Private Sub btnShowStats_Click()
    On Error GoTo ErrorHandler
    
    If m_LastSearchResults Is Nothing Then
        MsgBox "Please perform a search first to see statistics.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    ShowSearchStatistics m_LastSearchResults
    Exit Sub
    
ErrorHandler:
    MsgBox "Error displaying statistics: " & Err.Description, vbCritical, "Statistics Error"
End Sub

' ============================================
' Search History Management
' ============================================

Private Sub LoadSearchHistory()
    ' Initialize collection
    Set m_SearchHistory = New Collection
    
    ' Load from registry or file (simplified version - stores in memory only)
    ' In production, this would read from a JSON file or registry
    With Me.cboSearchHistory
        .Clear
        .AddItem "--- Select from recent searches ---"
        .ListIndex = 0
    End With
End Sub

Private Sub SaveToSearchHistory(criteria As SearchCriteria)
    Dim historyText As String
    
    ' Build human-readable summary
    historyText = ""
    If criteria.FromAddress <> "" Then historyText = historyText & "From: " & criteria.FromAddress & " | "
    If criteria.SearchTerms <> "" Then historyText = historyText & "Terms: " & criteria.SearchTerms & " | "
    If criteria.DateFrom <> "" Or criteria.DateTo <> "" Then
        historyText = historyText & "Dates: " & criteria.DateFrom & " to " & criteria.DateTo & " | "
    End If
    
    If Len(historyText) > MAX_HISTORY_TEXT_LENGTH Then
        historyText = Left(historyText, MAX_HISTORY_TEXT_LENGTH - 3) & "..."
    End If
    
    ' Add to history dropdown (keep last MAX_HISTORY_ITEMS)
    If Me.cboSearchHistory.ListCount >= (MAX_HISTORY_ITEMS + 1) Then
        Me.cboSearchHistory.RemoveItem Me.cboSearchHistory.ListCount - 1
    End If
    
    Me.cboSearchHistory.AddItem historyText, 1
End Sub

Private Sub cboSearchHistory_Click()
    ' When user selects from history, this would reload those criteria
    ' Simplified version - just shows the selection
    If Me.cboSearchHistory.ListIndex > 0 Then
        Me.lblStatus.Caption = "Search history: " & Me.cboSearchHistory.Text
    End If
End Sub

' ============================================
' Keyboard Shortcuts
' ============================================

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Ctrl+F - Focus on search terms
    If Shift = 2 And KeyCode = 70 Then ' Ctrl+F
        Me.txtSearchTerms.SetFocus
        KeyCode = 0
    
    ' Ctrl+Enter - Execute search
    ElseIf Shift = 2 And KeyCode = 13 Then ' Ctrl+Enter
        btnSearch_Click
        KeyCode = 0
    
    ' Escape - Close form
    ElseIf KeyCode = 27 Then ' Esc
        Unload Me
        KeyCode = 0
    
    ' Ctrl+E - Export CSV
    ElseIf Shift = 2 And KeyCode = 69 Then ' Ctrl+E
        btnExportCSV_Click
        KeyCode = 0
    End If
End Sub

' ============================================
' Enter Key Handlers for Text Boxes
' ============================================

Private Sub txtFrom_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtSearchTerms_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtTo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtCc_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtDateFrom_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtDateTo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

' ============================================
' Store Last Search Results (for bulk actions)
' ============================================

Public Sub SetLastSearchResults(searchObj As Outlook.Search)
    Set m_LastSearchResults = searchObj
End Sub

' ============================================
' Form Cleanup
' ============================================

Private Sub UserForm_Terminate()
    ' Release all module-level objects to prevent memory leaks
    On Error Resume Next
    Set m_SearchHistory = Nothing
    Set m_LastSearchResults = Nothing
    On Error GoTo 0
End Sub
