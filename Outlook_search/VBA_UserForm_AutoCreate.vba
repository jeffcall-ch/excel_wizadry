' ============================================
' Module: frmEmailSearch (UserForm Code)
' Purpose: Auto-create all controls programmatically
' Version: 1.2 - Programmatic control creation
' ============================================

Option Explicit

' Constants
Private Const MAX_HISTORY_ITEMS As Long = 10
Private Const MAX_HISTORY_TEXT_LENGTH As Long = 100

' Module-level variables
Private m_SearchHistory As Collection

' Control references
Private txtFrom As MSForms.TextBox
Private txtSearchTerms As MSForms.TextBox
Private txtDateFrom As MSForms.TextBox
Private txtDateTo As MSForms.TextBox
Private txtTo As MSForms.TextBox
Private txtCc As MSForms.TextBox
Private cboFolder As MSForms.ComboBox
Private cboAttachmentFilter As MSForms.ComboBox
Private cboSizeFilter As MSForms.ComboBox
Private cboSearchHistory As MSForms.ComboBox
Private chkSearchSubject As MSForms.CheckBox
Private chkSearchBody As MSForms.CheckBox
Private chkCaseSensitive As MSForms.CheckBox
Private chkUnreadOnly As MSForms.CheckBox
Private chkImportantOnly As MSForms.CheckBox
Private btnSearch As MSForms.CommandButton
Private btnClear As MSForms.CommandButton
Private btnSavePreset As MSForms.CommandButton
Private btnLoadPreset As MSForms.CommandButton
Private btnMarkRead As MSForms.CommandButton
Private btnFlagItems As MSForms.CommandButton
Private btnMoveItems As MSForms.CommandButton
Private btnExportCSV As MSForms.CommandButton
Private btnExportExcel As MSForms.CommandButton
Private btnShowStats As MSForms.CommandButton
Private lblStatus As MSForms.Label

' ============================================
' Form Initialization - Create All Controls
' ============================================

Private Sub UserForm_Initialize()
    On Error Resume Next
    
    ' Set form size
    Me.Width = 600
    Me.Height = 550
    Me.Caption = "Advanced Email Search"
    
    ' Create all controls
    CreateControls
    
    ' Initialize controls
    InitializeControls
    
    On Error GoTo 0
End Sub

Private Sub CreateControls()
    Dim leftCol As Long: leftCol = 12
    Dim rightCol As Long: rightCol = 320
    Dim topPos As Long: topPos = 12
    
    ' === BASIC SEARCH SECTION ===
    
    ' From field
    Set txtFrom = Me.Controls.Add("Forms.TextBox.1", "txtFrom")
    With txtFrom
        .Left = leftCol + 80
        .Top = topPos
        .Width = 200
        .Height = 20
    End With
    CreateLabel "From:", leftCol, topPos + 2
    topPos = topPos + 28
    
    ' Search Terms field
    Set txtSearchTerms = Me.Controls.Add("Forms.TextBox.1", "txtSearchTerms")
    With txtSearchTerms
        .Left = leftCol + 80
        .Top = topPos
        .Width = 200
        .Height = 20
    End With
    CreateLabel "Search:", leftCol, topPos + 2
    topPos = topPos + 28
    
    ' Search options checkboxes
    Set chkSearchSubject = Me.Controls.Add("Forms.CheckBox.1", "chkSearchSubject")
    With chkSearchSubject
        .Left = leftCol + 80
        .Top = topPos
        .Width = 100
        .Height = 18
        .Caption = "Subject"
        .Value = True
    End With
    
    Set chkSearchBody = Me.Controls.Add("Forms.CheckBox.1", "chkSearchBody")
    With chkSearchBody
        .Left = leftCol + 180
        .Top = topPos
        .Width = 100
        .Height = 18
        .Caption = "Body"
    End With
    topPos = topPos + 24
    
    Set chkCaseSensitive = Me.Controls.Add("Forms.CheckBox.1", "chkCaseSensitive")
    With chkCaseSensitive
        .Left = leftCol + 80
        .Top = topPos
        .Width = 120
        .Height = 18
        .Caption = "Case Sensitive"
    End With
    topPos = topPos + 28
    
    ' Date range
    Set txtDateFrom = Me.Controls.Add("Forms.TextBox.1", "txtDateFrom")
    With txtDateFrom
        .Left = leftCol + 80
        .Top = topPos
        .Width = 80
        .Height = 20
    End With
    CreateLabel "From Date:", leftCol, topPos + 2
    
    Set txtDateTo = Me.Controls.Add("Forms.TextBox.1", "txtDateTo")
    With txtDateTo
        .Left = leftCol + 240
        .Top = topPos
        .Width = 80
        .Height = 20
    End With
    CreateLabel "To:", leftCol + 210, topPos + 2
    topPos = topPos + 28
    
    ' Folder dropdown
    Set cboFolder = Me.Controls.Add("Forms.ComboBox.1", "cboFolder")
    With cboFolder
        .Left = leftCol + 80
        .Top = topPos
        .Width = 240
        .Height = 20
    End With
    CreateLabel "Folder:", leftCol, topPos + 2
    topPos = topPos + 36
    
    ' === ADVANCED FILTERS ===
    
    ' To/Cc fields
    Set txtTo = Me.Controls.Add("Forms.TextBox.1", "txtTo")
    With txtTo
        .Left = leftCol + 80
        .Top = topPos
        .Width = 240
        .Height = 20
    End With
    CreateLabel "To:", leftCol, topPos + 2
    topPos = topPos + 28
    
    Set txtCc = Me.Controls.Add("Forms.TextBox.1", "txtCc")
    With txtCc
        .Left = leftCol + 80
        .Top = topPos
        .Width = 240
        .Height = 20
    End With
    CreateLabel "Cc:", leftCol, topPos + 2
    topPos = topPos + 28
    
    ' Attachment filter
    Set cboAttachmentFilter = Me.Controls.Add("Forms.ComboBox.1", "cboAttachmentFilter")
    With cboAttachmentFilter
        .Left = leftCol + 80
        .Top = topPos
        .Width = 150
        .Height = 20
    End With
    CreateLabel "Attachments:", leftCol, topPos + 2
    topPos = topPos + 28
    
    ' Size filter
    Set cboSizeFilter = Me.Controls.Add("Forms.ComboBox.1", "cboSizeFilter")
    With cboSizeFilter
        .Left = leftCol + 80
        .Top = topPos
        .Width = 150
        .Height = 20
    End With
    CreateLabel "Size:", leftCol, topPos + 2
    topPos = topPos + 28
    
    ' Filter checkboxes
    Set chkUnreadOnly = Me.Controls.Add("Forms.CheckBox.1", "chkUnreadOnly")
    With chkUnreadOnly
        .Left = leftCol + 80
        .Top = topPos
        .Width = 100
        .Height = 18
        .Caption = "Unread Only"
    End With
    
    Set chkImportantOnly = Me.Controls.Add("Forms.CheckBox.1", "chkImportantOnly")
    With chkImportantOnly
        .Left = leftCol + 180
        .Top = topPos
        .Width = 140
        .Height = 18
        .Caption = "Important/Flagged"
    End With
    topPos = topPos + 28
    
    ' Search history
    Set cboSearchHistory = Me.Controls.Add("Forms.ComboBox.1", "cboSearchHistory")
    With cboSearchHistory
        .Left = leftCol + 100
        .Top = topPos
        .Width = 220
        .Height = 20
    End With
    CreateLabel "Recent:", leftCol, topPos + 2
    topPos = topPos + 36
    
    ' === BUTTON ROW 1 ===
    Dim btnLeft As Long: btnLeft = leftCol
    
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch")
    With btnSearch
        .Left = btnLeft
        .Top = topPos
        .Width = 70
        .Height = 24
        .Caption = "Search"
    End With
    btnLeft = btnLeft + 75
    
    Set btnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClear")
    With btnClear
        .Left = btnLeft
        .Top = topPos
        .Width = 60
        .Height = 24
        .Caption = "Clear"
    End With
    btnLeft = btnLeft + 65
    
    Set btnSavePreset = Me.Controls.Add("Forms.CommandButton.1", "btnSavePreset")
    With btnSavePreset
        .Left = btnLeft
        .Top = topPos
        .Width = 80
        .Height = 24
        .Caption = "Save Preset"
    End With
    btnLeft = btnLeft + 85
    
    Set btnLoadPreset = Me.Controls.Add("Forms.CommandButton.1", "btnLoadPreset")
    With btnLoadPreset
        .Left = btnLeft
        .Top = topPos
        .Width = 80
        .Height = 24
        .Caption = "Load Preset"
    End With
    topPos = topPos + 30
    
    ' === BUTTON ROW 2 - Bulk Actions ===
    btnLeft = leftCol
    
    Set btnMarkRead = Me.Controls.Add("Forms.CommandButton.1", "btnMarkRead")
    With btnMarkRead
        .Left = btnLeft
        .Top = topPos
        .Width = 75
        .Height = 24
        .Caption = "Mark Read"
    End With
    btnLeft = btnLeft + 80
    
    Set btnFlagItems = Me.Controls.Add("Forms.CommandButton.1", "btnFlagItems")
    With btnFlagItems
        .Left = btnLeft
        .Top = topPos
        .Width = 60
        .Height = 24
        .Caption = "Flag"
    End With
    btnLeft = btnLeft + 65
    
    Set btnMoveItems = Me.Controls.Add("Forms.CommandButton.1", "btnMoveItems")
    With btnMoveItems
        .Left = btnLeft
        .Top = topPos
        .Width = 60
        .Height = 24
        .Caption = "Move"
    End With
    btnLeft = btnLeft + 65
    
    Set btnExportCSV = Me.Controls.Add("Forms.CommandButton.1", "btnExportCSV")
    With btnExportCSV
        .Left = btnLeft
        .Top = topPos
        .Width = 70
        .Height = 24
        .Caption = "CSV"
    End With
    btnLeft = btnLeft + 75
    
    Set btnExportExcel = Me.Controls.Add("Forms.CommandButton.1", "btnExportExcel")
    With btnExportExcel
        .Left = btnLeft
        .Top = topPos
        .Width = 70
        .Height = 24
        .Caption = "Excel"
    End With
    topPos = topPos + 30
    
    ' === BUTTON ROW 3 ===
    Set btnShowStats = Me.Controls.Add("Forms.CommandButton.1", "btnShowStats")
    With btnShowStats
        .Left = leftCol
        .Top = topPos
        .Width = 80
        .Height = 24
        .Caption = "Statistics"
    End With
    topPos = topPos + 32
    
    ' === STATUS LABEL ===
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Left = leftCol
        .Top = topPos
        .Width = 560
        .Height = 20
        .Caption = "Ready"
        .BackColor = RGB(240, 240, 240)
        .BorderStyle = 1
    End With
End Sub

Private Sub CreateLabel(caption As String, left As Long, top As Long)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1")
    With lbl
        .caption = caption
        .left = left
        .top = top
        .Width = 70
        .Height = 16
        .TextAlign = 1  ' Right align
    End With
End Sub

Private Sub InitializeControls()
    ' Initialize search history
    LoadSearchHistory
    
    ' Populate folder dropdown
    With cboFolder
        .Clear
        .AddItem "All Folders"
        .AddItem "Inbox"
        .AddItem "Sent Items"
        .AddItem "Deleted Items"
        .AddItem "Drafts"
        .ListIndex = 0
    End With
    
    ' Populate attachment filter
    With cboAttachmentFilter
        .Clear
        .AddItem "Any"
        .AddItem "With Attachments"
        .AddItem "Without Attachments"
        .ListIndex = 0
    End With
    
    ' Populate size filter
    With cboSizeFilter
        .Clear
        .AddItem "Any Size"
        .AddItem "Small (< 100 KB)"
        .AddItem "Medium (100 KB - 1 MB)"
        .AddItem "Large (> 1 MB)"
        .ListIndex = 0
    End With
    
    ' Set default search options
    chkSearchSubject.Value = True
    chkSearchBody.Value = False
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
    
    ' Get the latest search results from ThisOutlookSession
    Dim searchResults As Outlook.Search
    Set searchResults = GetLastSearchObject()
    
    If searchResults Is Nothing Then
        MsgBox "Please perform a search first before using bulk actions.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    BulkMarkAsRead searchResults
    Me.lblStatus.Caption = "Bulk operation completed - Marked emails as read"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during bulk mark as read: " & Err.Description, vbCritical, "Bulk Action Error"
    Me.lblStatus.Caption = "Error during bulk action"
End Sub

Private Sub btnFlagItems_Click()
    On Error GoTo ErrorHandler
    
    Dim searchResults As Outlook.Search
    Set searchResults = GetLastSearchObject()
    
    If searchResults Is Nothing Then
        MsgBox "Please perform a search first before using bulk actions.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    BulkFlagItems searchResults
    Me.lblStatus.Caption = "Bulk operation completed - Flagged emails"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during bulk flag: " & Err.Description, vbCritical, "Bulk Action Error"
    Me.lblStatus.Caption = "Error during bulk action"
End Sub

Private Sub btnMoveItems_Click()
    On Error GoTo ErrorHandler
    
    Dim searchResults As Outlook.Search
    Set searchResults = GetLastSearchObject()
    
    If searchResults Is Nothing Then
        MsgBox "Please perform a search first before using bulk actions.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    BulkMoveItems searchResults
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
    
    Dim searchResults As Outlook.Search
    Set searchResults = GetLastSearchObject()
    
    If searchResults Is Nothing Then
        MsgBox "Please perform a search first before exporting.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    ExportSearchResultsToCSV searchResults
    Me.lblStatus.Caption = "Export completed - CSV file saved"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during CSV export: " & Err.Description, vbCritical, "Export Error"
    Me.lblStatus.Caption = "Error during export"
End Sub

Private Sub btnExportExcel_Click()
    On Error GoTo ErrorHandler
    
    Dim searchResults As Outlook.Search
    Set searchResults = GetLastSearchObject()
    
    If searchResults Is Nothing Then
        MsgBox "Please perform a search first before exporting.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    ExportSearchResultsToExcel searchResults
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
    
    Dim searchResults As Outlook.Search
    Set searchResults = GetLastSearchObject()
    
    If searchResults Is Nothing Then
        MsgBox "Please perform a search first to see statistics.", vbInformation, "No Search Results"
        Exit Sub
    End If
    
    ShowSearchStatistics searchResults
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
    With cboSearchHistory
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
    If cboSearchHistory.ListCount >= (MAX_HISTORY_ITEMS + 1) Then
        cboSearchHistory.RemoveItem cboSearchHistory.ListCount - 1
    End If
    
    cboSearchHistory.AddItem historyText, 1
End Sub

Private Sub cboSearchHistory_Click()
    ' When user selects from history, this would reload those criteria
    ' Simplified version - just shows the selection
    If cboSearchHistory.ListIndex > 0 Then
        lblStatus.Caption = "Search history: " & cboSearchHistory.Text
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
' Form Cleanup
' ============================================

Private Sub UserForm_Terminate()
    ' Release all module-level objects to prevent memory leaks
    On Error Resume Next
    Set m_SearchHistory = Nothing
    On Error GoTo 0
End Sub

