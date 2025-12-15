' ============================================
' UserForm: frmEmailSearch (AQS Version - Programmatic UI)
' Purpose: Email search form with ALL controls created at runtime
' Version: 2.0.0 - AQS with programmatic UI
' Date: 2024-12-15
' ============================================
'
' INSTRUCTIONS:
' 1. Create a new blank UserForm named "frmEmailSearch"
' 2. Set form properties:
'    - Width: 620
'    - Height: 550
'    - Caption: "Email Search (AQS)"
' 3. Copy this entire code into the form's code module
' 4. No manual control placement needed - all controls created programmatically!
' ============================================

Option Explicit

' Control references (module-level with WithEvents for event handling)
Private WithEvents txtFrom As MSForms.TextBox
Private WithEvents txtTo As MSForms.TextBox
Private WithEvents txtCC As MSForms.TextBox
Private WithEvents txtSubject As MSForms.TextBox
Private WithEvents txtBody As MSForms.TextBox
Private WithEvents txtSearchTerms As MSForms.TextBox
Private WithEvents txtDateFrom As MSForms.TextBox
Private WithEvents txtDateTo As MSForms.TextBox

Private WithEvents chkUnreadOnly As MSForms.CheckBox
Private WithEvents chkImportantOnly As MSForms.CheckBox
Private WithEvents chkSearchSubFolders As MSForms.CheckBox

Private WithEvents cboAttachmentFilter As MSForms.ComboBox
Private WithEvents cboSizeFilter As MSForms.ComboBox
Private WithEvents cboFolderName As MSForms.ComboBox

Private WithEvents btnSearch As MSForms.CommandButton
Private WithEvents btnCaptureResults As MSForms.CommandButton
Private WithEvents btnClear As MSForms.CommandButton
Private WithEvents btnExport As MSForms.CommandButton
Private WithEvents btnBulkActions As MSForms.CommandButton
Private WithEvents btnStatistics As MSForms.CommandButton

Private lblStatus As MSForms.Label

' Labels for field descriptions
Private lblFrom As MSForms.Label
Private lblTo As MSForms.Label
Private lblCC As MSForms.Label
Private lblSubject As MSForms.Label
Private lblBody As MSForms.Label
Private lblSearchTerms As MSForms.Label
Private lblDateFrom As MSForms.Label
Private lblDateTo As MSForms.Label
Private lblAttachmentFilter As MSForms.Label
Private lblSizeFilter As MSForms.Label
Private lblFolderName As MSForms.Label

' ============================================
' FORM INITIALIZATION - Creates all controls
' ============================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    ' Set form properties
    Me.Caption = "Email Search (AQS)"
    Me.Width = 620
    Me.Height = 550
    
    ' Create all controls programmatically
    CreateLabels
    CreateTextBoxes
    CreateCheckBoxes
    CreateComboBoxes
    CreateButtons
    CreateStatusLabel
    
    ' Initialize control values
    InitializeControlValues
    
    ' Set initial button states
    SetInitialButtonStates
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing form: " & Err.Description, vbCritical, "Form Initialization Error"
End Sub

' ============================================
' CREATE CONTROLS - Programmatic UI
' ============================================

Private Sub CreateLabels()
    ' From label
    Set lblFrom = Me.Controls.Add("Forms.Label.1", "lblFrom", True)
    lblFrom.Left = 10
    lblFrom.Top = 10
    lblFrom.Width = 80
    lblFrom.Height = 18
    lblFrom.Caption = "From:"
    
    ' To label
    Set lblTo = Me.Controls.Add("Forms.Label.1", "lblTo", True)
    lblTo.Left = 320
    lblTo.Top = 10
    lblTo.Width = 80
    lblTo.Height = 18
    lblTo.Caption = "To:"
    
    ' CC label
    Set lblCC = Me.Controls.Add("Forms.Label.1", "lblCC", True)
    lblCC.Left = 10
    lblCC.Top = 50
    lblCC.Width = 80
    lblCC.Height = 18
    lblCC.Caption = "CC:"
    
    ' Subject label
    Set lblSubject = Me.Controls.Add("Forms.Label.1", "lblSubject", True)
    lblSubject.Left = 320
    lblSubject.Top = 50
    lblSubject.Width = 80
    lblSubject.Height = 18
    lblSubject.Caption = "Subject:"
    
    ' Body label
    Set lblBody = Me.Controls.Add("Forms.Label.1", "lblBody", True)
    lblBody.Left = 10
    lblBody.Top = 90
    lblBody.Width = 80
    lblBody.Height = 18
    lblBody.Caption = "Body:"
    
    ' Search Terms label
    Set lblSearchTerms = Me.Controls.Add("Forms.Label.1", "lblSearchTerms", True)
    lblSearchTerms.Left = 320
    lblSearchTerms.Top = 90
    lblSearchTerms.Width = 80
    lblSearchTerms.Height = 18
    lblSearchTerms.Caption = "Search Terms:"
    
    ' Date From label
    Set lblDateFrom = Me.Controls.Add("Forms.Label.1", "lblDateFrom", True)
    lblDateFrom.Left = 10
    lblDateFrom.Top = 130
    lblDateFrom.Width = 80
    lblDateFrom.Height = 18
    lblDateFrom.Caption = "Date From:"
    
    ' Date To label
    Set lblDateTo = Me.Controls.Add("Forms.Label.1", "lblDateTo", True)
    lblDateTo.Left = 320
    lblDateTo.Top = 130
    lblDateTo.Width = 80
    lblDateTo.Height = 18
    lblDateTo.Caption = "Date To:"
    
    ' Attachment Filter label
    Set lblAttachmentFilter = Me.Controls.Add("Forms.Label.1", "lblAttachmentFilter", True)
    lblAttachmentFilter.Left = 10
    lblAttachmentFilter.Top = 210
    lblAttachmentFilter.Width = 120
    lblAttachmentFilter.Height = 18
    lblAttachmentFilter.Caption = "Attachment Filter:"
    
    ' Size Filter label
    Set lblSizeFilter = Me.Controls.Add("Forms.Label.1", "lblSizeFilter", True)
    lblSizeFilter.Left = 320
    lblSizeFilter.Top = 210
    lblSizeFilter.Width = 80
    lblSizeFilter.Height = 18
    lblSizeFilter.Caption = "Size Filter:"
    
    ' Folder Name label
    Set lblFolderName = Me.Controls.Add("Forms.Label.1", "lblFolderName", True)
    lblFolderName.Left = 10
    lblFolderName.Top = 250
    lblFolderName.Width = 80
    lblFolderName.Height = 18
    lblFolderName.Caption = "Folder:"
End Sub

Private Sub CreateTextBoxes()
    ' From textbox
    Set txtFrom = Me.Controls.Add("Forms.TextBox.1", "txtFrom", True)
    txtFrom.Left = 100
    txtFrom.Top = 10
    txtFrom.Width = 190
    txtFrom.Height = 20
    
    ' To textbox
    Set txtTo = Me.Controls.Add("Forms.TextBox.1", "txtTo", True)
    txtTo.Left = 410
    txtTo.Top = 10
    txtTo.Width = 195
    txtTo.Height = 20
    
    ' CC textbox
    Set txtCC = Me.Controls.Add("Forms.TextBox.1", "txtCC", True)
    txtCC.Left = 100
    txtCC.Top = 50
    txtCC.Width = 190
    txtCC.Height = 20
    
    ' Subject textbox
    Set txtSubject = Me.Controls.Add("Forms.TextBox.1", "txtSubject", True)
    txtSubject.Left = 410
    txtSubject.Top = 50
    txtSubject.Width = 195
    txtSubject.Height = 20
    
    ' Body textbox
    Set txtBody = Me.Controls.Add("Forms.TextBox.1", "txtBody", True)
    txtBody.Left = 100
    txtBody.Top = 90
    txtBody.Width = 190
    txtBody.Height = 20
    
    ' Search Terms textbox
    Set txtSearchTerms = Me.Controls.Add("Forms.TextBox.1", "txtSearchTerms", True)
    txtSearchTerms.Left = 410
    txtSearchTerms.Top = 90
    txtSearchTerms.Width = 195
    txtSearchTerms.Height = 20
    
    ' Date From textbox
    Set txtDateFrom = Me.Controls.Add("Forms.TextBox.1", "txtDateFrom", True)
    txtDateFrom.Left = 100
    txtDateFrom.Top = 130
    txtDateFrom.Width = 190
    txtDateFrom.Height = 20
    
    ' Date To textbox
    Set txtDateTo = Me.Controls.Add("Forms.TextBox.1", "txtDateTo", True)
    txtDateTo.Left = 410
    txtDateTo.Top = 130
    txtDateTo.Width = 195
    txtDateTo.Height = 20
End Sub

Private Sub CreateCheckBoxes()
    ' Unread Only checkbox
    Set chkUnreadOnly = Me.Controls.Add("Forms.CheckBox.1", "chkUnreadOnly", True)
    chkUnreadOnly.Left = 10
    chkUnreadOnly.Top = 170
    chkUnreadOnly.Width = 120
    chkUnreadOnly.Height = 18
    chkUnreadOnly.Caption = "Unread Only"
    
    ' Important Only checkbox
    Set chkImportantOnly = Me.Controls.Add("Forms.CheckBox.1", "chkImportantOnly", True)
    chkImportantOnly.Left = 150
    chkImportantOnly.Top = 170
    chkImportantOnly.Width = 120
    chkImportantOnly.Height = 18
    chkImportantOnly.Caption = "Important Only"
    
    ' Search Subfolders checkbox
    Set chkSearchSubFolders = Me.Controls.Add("Forms.CheckBox.1", "chkSearchSubFolders", True)
    chkSearchSubFolders.Left = 290
    chkSearchSubFolders.Top = 170
    chkSearchSubFolders.Width = 150
    chkSearchSubFolders.Height = 18
    chkSearchSubFolders.Caption = "Include Subfolders"
End Sub

Private Sub CreateComboBoxes()
    ' Attachment Filter combo
    Set cboAttachmentFilter = Me.Controls.Add("Forms.ComboBox.1", "cboAttachmentFilter", True)
    cboAttachmentFilter.Left = 135
    cboAttachmentFilter.Top = 210
    cboAttachmentFilter.Width = 155
    cboAttachmentFilter.Height = 20
    
    ' Size Filter combo
    Set cboSizeFilter = Me.Controls.Add("Forms.ComboBox.1", "cboSizeFilter", True)
    cboSizeFilter.Left = 410
    cboSizeFilter.Top = 210
    cboSizeFilter.Width = 195
    cboSizeFilter.Height = 20
    
    ' Folder Name combo
    Set cboFolderName = Me.Controls.Add("Forms.ComboBox.1", "cboFolderName", True)
    cboFolderName.Left = 100
    cboFolderName.Top = 250
    cboFolderName.Width = 505
    cboFolderName.Height = 20
End Sub

Private Sub CreateButtons()
    ' Search button
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch", True)
    btnSearch.Left = 10
    btnSearch.Top = 300
    btnSearch.Width = 90
    btnSearch.Height = 30
    btnSearch.Caption = "Search"
    
    ' Capture Results button
    Set btnCaptureResults = Me.Controls.Add("Forms.CommandButton.1", "btnCaptureResults", True)
    btnCaptureResults.Left = 110
    btnCaptureResults.Top = 300
    btnCaptureResults.Width = 110
    btnCaptureResults.Height = 30
    btnCaptureResults.Caption = "Capture Results"
    
    ' Clear button
    Set btnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClear", True)
    btnClear.Left = 230
    btnClear.Top = 300
    btnClear.Width = 90
    btnClear.Height = 30
    btnClear.Caption = "Clear"
    
    ' Export button
    Set btnExport = Me.Controls.Add("Forms.CommandButton.1", "btnExport", True)
    btnExport.Left = 10
    btnExport.Top = 340
    btnExport.Width = 90
    btnExport.Height = 30
    btnExport.Caption = "Export"
    
    ' Bulk Actions button
    Set btnBulkActions = Me.Controls.Add("Forms.CommandButton.1", "btnBulkActions", True)
    btnBulkActions.Left = 110
    btnBulkActions.Top = 340
    btnBulkActions.Width = 110
    btnBulkActions.Height = 30
    btnBulkActions.Caption = "Bulk Actions"
    
    ' Statistics button
    Set btnStatistics = Me.Controls.Add("Forms.CommandButton.1", "btnStatistics", True)
    btnStatistics.Left = 230
    btnStatistics.Top = 340
    btnStatistics.Width = 90
    btnStatistics.Height = 30
    btnStatistics.Caption = "Statistics"
End Sub

Private Sub CreateStatusLabel()
    ' Status label at bottom of form
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    lblStatus.Left = 10
    lblStatus.Top = 390
    lblStatus.Width = 595
    lblStatus.Height = 40
    lblStatus.Caption = "Ready to search"
    lblStatus.ForeColor = RGB(0, 0, 0)  ' Black
    lblStatus.BackColor = RGB(240, 240, 240)  ' Light gray background
    lblStatus.BorderStyle = 1  ' Single border
    lblStatus.TextAlign = 1  ' Left align
    lblStatus.WordWrap = True
End Sub

' ============================================
' INITIALIZE CONTROL VALUES
' ============================================

Private Sub InitializeControlValues()
    ' Populate Attachment Filter dropdown
    cboAttachmentFilter.Clear
    cboAttachmentFilter.AddItem "Any"
    cboAttachmentFilter.AddItem "With Attachments"
    cboAttachmentFilter.AddItem "Without Attachments"
    cboAttachmentFilter.AddItem "PDF files"
    cboAttachmentFilter.AddItem "Excel files (.xls, .xlsx)"
    cboAttachmentFilter.AddItem "Word files (.doc, .docx)"
    cboAttachmentFilter.AddItem "Images (.jpg, .png, .gif)"
    cboAttachmentFilter.AddItem "PowerPoint files (.ppt, .pptx)"
    cboAttachmentFilter.AddItem "ZIP/Archives"
    cboAttachmentFilter.ListIndex = 0  ' Default to "Any"
    
    ' Populate Size Filter dropdown
    cboSizeFilter.Clear
    cboSizeFilter.AddItem "Any Size"
    cboSizeFilter.AddItem "< 100 KB"
    cboSizeFilter.AddItem "100 KB - 1 MB"
    cboSizeFilter.AddItem "1 MB - 10 MB"
    cboSizeFilter.AddItem "> 10 MB"
    cboSizeFilter.ListIndex = 0  ' Default to "Any Size"
    
    ' Populate Folder Name dropdown
    cboFolderName.Clear
    cboFolderName.AddItem "Inbox"
    cboFolderName.AddItem "Sent Items"
    cboFolderName.AddItem "Deleted Items"
    cboFolderName.AddItem "Drafts"
    cboFolderName.AddItem "All Folders (including subfolders)"
    cboFolderName.ListIndex = 4  ' Default to "All Folders (including subfolders)"
    
    ' Set default checkbox states
    chkUnreadOnly.Value = False
    chkImportantOnly.Value = False
    chkSearchSubFolders.Value = True  ' Default to include subfolders
End Sub

Private Sub SetInitialButtonStates()
    ' Initial button states (before search)
    btnCaptureResults.Enabled = False
    btnExport.Enabled = False
    btnBulkActions.Enabled = False
    btnStatistics.Enabled = False
    
    ' Update status
    lblStatus.Caption = "Ready to search"
    lblStatus.ForeColor = RGB(0, 0, 0)  ' Black
End Sub

' ============================================
' BUTTON EVENT HANDLERS
' ============================================

Private Sub btnSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim criteria As SearchCriteria
    
    ' Validate form inputs
    If Not ValidateForm() Then Exit Sub
    
    ' Populate criteria from form controls
    criteria = GetSearchCriteriaFromForm()
    
    ' Execute search (calls ThisOutlookSession.ExecuteEmailSearch)
    ThisOutlookSession.ExecuteEmailSearch criteria
    
    ' Update UI state for AQS architecture
    lblStatus.Caption = "Search launched in Outlook. Review results, then click 'Capture Results'."
    lblStatus.ForeColor = RGB(0, 0, 255)  ' Blue
    btnCaptureResults.Enabled = True
    btnExport.Enabled = False
    btnBulkActions.Enabled = False
    btnStatistics.Enabled = False
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Search failed: " & Err.Description
    lblStatus.ForeColor = RGB(255, 0, 0)  ' Red
    MsgBox "Error executing search: " & Err.Description, vbCritical, "Search Error"
End Sub

Private Sub btnCaptureResults_Click()
    On Error GoTo ErrorHandler
    
    Dim resultCount As Long
    Dim capturedResults As Collection
    
    ' Show progress
    lblStatus.Caption = "Capturing results..."
    lblStatus.ForeColor = RGB(0, 0, 255)  ' Blue
    DoEvents
    
    ' Capture from search Explorer
    Set capturedResults = ThisOutlookSession.CaptureCurrentSearchResults()
    resultCount = capturedResults.Count
    
    If resultCount = 0 Then
        lblStatus.Caption = "No results captured. Search may still be running."
        lblStatus.ForeColor = RGB(255, 0, 0)  ' Red
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
    
    ' Show confirmation
    MsgBox "Captured " & resultCount & " email(s)." & vbCrLf & vbCrLf & _
           "You can now use Export or Bulk Actions.", vbInformation, "Capture Complete"
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Capture failed: " & Err.Description
    lblStatus.ForeColor = RGB(255, 0, 0)  ' Red
    MsgBox "Error capturing results: " & Err.Description, vbCritical, "Capture Error"
End Sub

Private Sub btnClear_Click()
    ' Clear all form inputs
    txtFrom.Value = ""
    txtTo.Value = ""
    txtCC.Value = ""
    txtSubject.Value = ""
    txtBody.Value = ""
    txtSearchTerms.Value = ""
    txtDateFrom.Value = ""
    txtDateTo.Value = ""
    
    chkUnreadOnly.Value = False
    chkImportantOnly.Value = False
    chkSearchSubFolders.Value = True
    
    cboAttachmentFilter.ListIndex = 0
    cboSizeFilter.ListIndex = 0
    cboFolderName.ListIndex = 0
    
    ' Reset button states
    SetInitialButtonStates
End Sub

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
            If ThisOutlookSession.ExportSearchResultsToCSV() Then
                lblStatus.Caption = "Export to CSV completed successfully."
                lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
            Else
                lblStatus.Caption = "Export to CSV failed."
                lblStatus.ForeColor = RGB(255, 0, 0)  ' Red
            End If
            
        Case vbNo
            ' Export to Excel
            If ThisOutlookSession.ExportSearchResultsToExcel() Then
                lblStatus.Caption = "Export to Excel completed successfully."
                lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
            Else
                lblStatus.Caption = "Export to Excel failed."
                lblStatus.ForeColor = RGB(255, 0, 0)  ' Red
            End If
            
        Case vbCancel
            Exit Sub
    End Select
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Export error: " & Err.Description
    lblStatus.ForeColor = RGB(255, 0, 0)  ' Red
    MsgBox "Error during export: " & Err.Description, vbCritical, "Export Error"
End Sub

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
            lblStatus.Caption = "Bulk action: Mark as Read completed."
            lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
            
        Case "2"
            ThisOutlookSession.BulkFlagItems
            lblStatus.Caption = "Bulk action: Flag Items completed."
            lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
            
        Case "3"
            ThisOutlookSession.BulkMoveItems
            lblStatus.Caption = "Bulk action: Move Items completed."
            lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
            
        Case Else
            MsgBox "Invalid choice or cancelled.", vbInformation
    End Select
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Bulk action error: " & Err.Description
    lblStatus.ForeColor = RGB(255, 0, 0)  ' Red
    MsgBox "Error during bulk action: " & Err.Description, vbCritical, "Bulk Action Error"
End Sub

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
    
    lblStatus.Caption = "Statistics displayed for " & capturedResults.Count & " results."
    lblStatus.ForeColor = RGB(0, 128, 0)  ' Green
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Statistics error: " & Err.Description
    lblStatus.ForeColor = RGB(255, 0, 0)  ' Red
    MsgBox "Error calculating statistics: " & Err.Description, vbCritical, "Statistics Error"
End Sub

' ============================================
' HELPER FUNCTIONS
' ============================================

Private Function ValidateForm() As Boolean
    ' Basic validation - at least one search criterion must be filled
    If Len(Trim(txtFrom.Value)) = 0 And _
       Len(Trim(txtTo.Value)) = 0 And _
       Len(Trim(txtCC.Value)) = 0 And _
       Len(Trim(txtSubject.Value)) = 0 And _
       Len(Trim(txtBody.Value)) = 0 And _
       Len(Trim(txtSearchTerms.Value)) = 0 And _
       Len(Trim(txtDateFrom.Value)) = 0 And _
       Len(Trim(txtDateTo.Value)) = 0 And _
       Not chkUnreadOnly.Value And _
       Not chkImportantOnly.Value And _
       cboAttachmentFilter.ListIndex = 0 And _
       cboSizeFilter.ListIndex = 0 Then
        
        MsgBox "Please enter at least one search criterion.", vbExclamation, "Validation Error"
        ValidateForm = False
        Exit Function
    End If
    
    ' Validate date formats if provided
    If Len(Trim(txtDateFrom.Value)) > 0 Then
        If Not IsDate(txtDateFrom.Value) Then
            MsgBox "Date From is not a valid date. Please use format: MM/DD/YYYY", vbExclamation, "Validation Error"
            txtDateFrom.SetFocus
            ValidateForm = False
            Exit Function
        End If
    End If
    
    If Len(Trim(txtDateTo.Value)) > 0 Then
        If Not IsDate(txtDateTo.Value) Then
            MsgBox "Date To is not a valid date. Please use format: MM/DD/YYYY", vbExclamation, "Validation Error"
            txtDateTo.SetFocus
            ValidateForm = False
            Exit Function
        End If
    End If
    
    ' Validate date range if both dates provided
    If Len(Trim(txtDateFrom.Value)) > 0 And Len(Trim(txtDateTo.Value)) > 0 Then
        If IsDate(txtDateFrom.Value) And IsDate(txtDateTo.Value) Then
            If CDate(txtDateFrom.Value) > CDate(txtDateTo.Value) Then
                MsgBox "Date From must be earlier than Date To.", vbExclamation, "Validation Error"
                txtDateFrom.SetFocus
                ValidateForm = False
                Exit Function
            End If
        End If
    End If
    
    ValidateForm = True
End Function

Private Function GetSearchCriteriaFromForm() As SearchCriteria
    Dim criteria As SearchCriteria
    
    ' Populate criteria from form controls
    criteria.FromAddress = Trim(txtFrom.Value)
    criteria.ToAddress = Trim(txtTo.Value)
    criteria.CcAddress = Trim(txtCC.Value)
    criteria.Subject = Trim(txtSubject.Value)
    criteria.Body = Trim(txtBody.Value)
    criteria.SearchTerms = Trim(txtSearchTerms.Value)
    
    ' Dates
    If Len(Trim(txtDateFrom.Value)) > 0 And IsDate(txtDateFrom.Value) Then
        criteria.DateFrom = CDate(txtDateFrom.Value)
    End If
    
    If Len(Trim(txtDateTo.Value)) > 0 And IsDate(txtDateTo.Value) Then
        criteria.DateTo = CDate(txtDateTo.Value)
    End If
    
    ' Checkboxes
    criteria.UnreadOnly = chkUnreadOnly.Value
    criteria.ImportantOnly = chkImportantOnly.Value
    criteria.SearchSubFolders = chkSearchSubFolders.Value
    
    ' Dropdowns
    If cboAttachmentFilter.ListIndex >= 0 Then
        criteria.AttachmentFilter = cboAttachmentFilter.List(cboAttachmentFilter.ListIndex)
    End If
    
    If cboSizeFilter.ListIndex >= 0 Then
        criteria.SizeFilter = cboSizeFilter.List(cboSizeFilter.ListIndex)
    End If
    
    If cboFolderName.ListIndex >= 0 Then
        criteria.FolderName = cboFolderName.List(cboFolderName.ListIndex)
    Else
        criteria.FolderName = "Inbox"  ' Default
    End If
    
    GetSearchCriteriaFromForm = criteria
End Function
