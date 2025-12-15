' ============================================
' Module: frmEmailSearch (UserForm)
' Purpose: Advanced email search interface with all Phase 1-3 features
' ============================================

Option Explicit

Private Sub UserForm_Initialize()
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
    
    ' Set placeholders
    Me.txtDateFrom.Tag = "YYYY-MM-DD"
    Me.txtDateTo.Tag = "YYYY-MM-DD"
End Sub

Private Sub btnSearch_Click()
    ' Validate and execute search
    Dim searchCriteria As SearchCriteria
    
    ' Get values
    searchCriteria.FromText = Trim(Me.txtFrom.Text)
    searchCriteria.SubjectText = Trim(Me.txtSubject.Text)
    searchCriteria.SearchBody = Me.chkSearchBody.Value
    searchCriteria.DateFrom = Trim(Me.txtDateFrom.Text)
    searchCriteria.DateTo = Trim(Me.txtDateTo.Text)
    searchCriteria.FolderName = Me.cboFolder.Text
    
    ' Check if at least one criterion is specified
    If searchCriteria.FromText = "" And searchCriteria.SubjectText = "" And _
       searchCriteria.DateFrom = "" And searchCriteria.DateTo = "" Then
        MsgBox "Please enter at least one search criterion.", vbExclamation, "Search Required"
        Exit Sub
    End If
    
    ' Validate dates if provided
    If searchCriteria.DateFrom <> "" Then
        If Not IsDate(searchCriteria.DateFrom) Then
            MsgBox "Invalid Date From format. Use YYYY-MM-DD or MM/DD/YYYY", vbExclamation, "Invalid Date"
            Me.txtDateFrom.SetFocus
            Exit Sub
        End If
    End If
    
    If searchCriteria.DateTo <> "" Then
        If Not IsDate(searchCriteria.DateTo) Then
            MsgBox "Invalid Date To format. Use YYYY-MM-DD or MM/DD/YYYY", vbExclamation, "Invalid Date"
            Me.txtDateTo.SetFocus
            Exit Sub
        End If
    End If
    
    ' Hide form and execute search
    Me.Hide
    
    ' Call the search function in ThisOutlookSession
    ExecuteEmailSearch searchCriteria
    
    ' Close the form
    Unload Me
End Sub

Private Sub btnCancel_Click()
    ' Close without searching
    Unload Me
End Sub

' Handle Enter key in text boxes
Private Sub txtFrom_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtSubject_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtDateFrom_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtDateTo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub
