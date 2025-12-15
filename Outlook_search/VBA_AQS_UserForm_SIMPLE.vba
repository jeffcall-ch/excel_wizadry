' ============================================
' UserForm: frmEmailSearch (SIMPLIFIED VERSION)
' Purpose: Programmatic search form with minimal UI
' Features: Single search box + checkboxes for search scope
' Date: 2024-12-15
' ============================================

Option Explicit

' Module-level controls (WithEvents for event handling)
Private WithEvents lblFrom As MSForms.Label
Private WithEvents lblTo As MSForms.Label
Private WithEvents lblCC As MSForms.Label
Private WithEvents lblSearchTerms As MSForms.Label
Private WithEvents lblSearchHelp As MSForms.Label

Private WithEvents txtFrom As MSForms.TextBox
Private WithEvents txtTo As MSForms.TextBox
Private WithEvents txtCC As MSForms.TextBox
Private WithEvents txtSearchTerms As MSForms.TextBox

Private WithEvents lblSearchIn As MSForms.Label
Private WithEvents chkSubject As MSForms.CheckBox
Private WithEvents chkBody As MSForms.CheckBox

Private WithEvents lblDateFrom As MSForms.Label
Private WithEvents lblDateTo As MSForms.Label
Private WithEvents txtDateFrom As MSForms.TextBox
Private WithEvents txtDateTo As MSForms.TextBox

Private WithEvents lblFolder As MSForms.Label
Private WithEvents cboFolderName As MSForms.ComboBox
Private WithEvents chkSubFolders As MSForms.CheckBox

Private WithEvents btnSearch As MSForms.CommandButton
Private WithEvents btnClear As MSForms.CommandButton

Private WithEvents lblStatus As MSForms.Label

' ============================================
' FORM INITIALIZATION
' ============================================

Private Sub UserForm_Initialize()
    ' Set form properties
    Me.Caption = "Email Search (Simplified)"
    Me.Width = 480
    Me.Height = 435
    
    ' Create all controls programmatically
    CreateLabels
    CreateTextBoxes
    CreateCheckBoxes
    CreateComboBoxes
    CreateButtons
    CreateStatusLabel
    
    ' Populate controls with default values
    InitializeControlValues
End Sub

' ============================================
' CONTROL CREATION
' ============================================

Private Sub CreateLabels()
    ' From Label
    Set lblFrom = Me.Controls.Add("Forms.Label.1", "lblFrom", True)
    With lblFrom
        .Left = 10
        .Top = 10
        .Width = 80
        .Height = 18
        .Caption = "From:"
        .Font.Size = 10
    End With
    
    ' To Label
    Set lblTo = Me.Controls.Add("Forms.Label.1", "lblTo", True)
    With lblTo
        .Left = 10
        .Top = 40
        .Width = 80
        .Height = 18
        .Caption = "To:"
        .Font.Size = 10
    End With
    
    ' CC Label
    Set lblCC = Me.Controls.Add("Forms.Label.1", "lblCC", True)
    With lblCC
        .Left = 10
        .Top = 70
        .Width = 80
        .Height = 18
        .Caption = "CC:"
        .Font.Size = 10
    End With
    
    ' Search Terms Label
    Set lblSearchTerms = Me.Controls.Add("Forms.Label.1", "lblSearchTerms", True)
    With lblSearchTerms
        .Left = 10
        .Top = 110
        .Width = 80
        .Height = 18
        .Caption = "Search Terms:"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    ' Search Help Label (operator reminder)
    Set lblSearchHelp = Me.Controls.Add("Forms.Label.1", "lblSearchHelp", True)
    With lblSearchHelp
        .Left = 100
        .Top = 135
        .Width = 360
        .Height = 14
        .Caption = "Operators: , (OR)  + (AND)  * (wildcard)  ENTER (search)"
        .Font.Size = 8
        .ForeColor = RGB(100, 100, 100)  ' Gray text
    End With
    
    ' Search In Label
    Set lblSearchIn = Me.Controls.Add("Forms.Label.1", "lblSearchIn", True)
    With lblSearchIn
        .Left = 10
        .Top = 160
        .Width = 80
        .Height = 18
        .Caption = "Search In:"
        .Font.Size = 10
    End With
    
    ' Date From Label
    Set lblDateFrom = Me.Controls.Add("Forms.Label.1", "lblDateFrom", True)
    With lblDateFrom
        .Left = 10
        .Top = 205
        .Width = 80
        .Height = 18
        .Caption = "Date From:"
        .Font.Size = 10
    End With
    
    ' Date To Label
    Set lblDateTo = Me.Controls.Add("Forms.Label.1", "lblDateTo", True)
    With lblDateTo
        .Left = 10
        .Top = 235
        .Width = 80
        .Height = 18
        .Caption = "Date To:"
        .Font.Size = 10
    End With
    
    ' Folder Label
    Set lblFolder = Me.Controls.Add("Forms.Label.1", "lblFolder", True)
    With lblFolder
        .Left = 10
        .Top = 275
        .Width = 80
        .Height = 18
        .Caption = "Folder:"
        .Font.Size = 10
    End With
End Sub

Private Sub CreateTextBoxes()
    ' From TextBox
    Set txtFrom = Me.Controls.Add("Forms.TextBox.1", "txtFrom", True)
    With txtFrom
        .Left = 100
        .Top = 10
        .Width = 360
        .Height = 24
        .Font.Size = 10
        .EnterKeyBehavior = False  ' Single line - Enter triggers default action
    End With
    
    ' To TextBox
    Set txtTo = Me.Controls.Add("Forms.TextBox.1", "txtTo", True)
    With txtTo
        .Left = 100
        .Top = 40
        .Width = 360
        .Height = 24
        .Font.Size = 10
        .EnterKeyBehavior = False
    End With
    
    ' CC TextBox
    Set txtCC = Me.Controls.Add("Forms.TextBox.1", "txtCC", True)
    With txtCC
        .Left = 100
        .Top = 70
        .Width = 360
        .Height = 24
        .Font.Size = 10
        .EnterKeyBehavior = False
    End With
    
    ' Search Terms TextBox (wide, bold border)
    Set txtSearchTerms = Me.Controls.Add("Forms.TextBox.1", "txtSearchTerms", True)
    With txtSearchTerms
        .Left = 100
        .Top = 110
        .Width = 360
        .Height = 24
        .Font.Size = 10
        .Font.Bold = True
        .EnterKeyBehavior = False  ' Single line - Enter triggers search
    End With
    
    ' Date From TextBox
    Set txtDateFrom = Me.Controls.Add("Forms.TextBox.1", "txtDateFrom", True)
    With txtDateFrom
        .Left = 100
        .Top = 205
        .Width = 100
        .Height = 24
        .Font.Size = 10
    End With
    
    ' Date To TextBox
    Set txtDateTo = Me.Controls.Add("Forms.TextBox.1", "txtDateTo", True)
    With txtDateTo
        .Left = 100
        .Top = 235
        .Width = 100
        .Height = 24
        .Font.Size = 10
    End With
End Sub

Private Sub CreateCheckBoxes()
    ' Subject CheckBox
    Set chkSubject = Me.Controls.Add("Forms.CheckBox.1", "chkSubject", True)
    With chkSubject
        .Left = 100
        .Top = 160
        .Width = 80
        .Height = 18
        .Caption = "Subject"
        .Font.Size = 9
        .Value = True  ' Default checked
    End With
    
    ' Body CheckBox
    Set chkBody = Me.Controls.Add("Forms.CheckBox.1", "chkBody", True)
    With chkBody
        .Left = 190
        .Top = 160
        .Width = 80
        .Height = 18
        .Caption = "Body"
        .Font.Size = 9
        .Value = True  ' Default checked
    End With
    
    ' Include Subfolders CheckBox
    Set chkSubFolders = Me.Controls.Add("Forms.CheckBox.1", "chkSubFolders", True)
    With chkSubFolders
        .Left = 100
        .Top = 300
        .Width = 150
        .Height = 18
        .Caption = "Include Subfolders"
        .Font.Size = 9
        .Value = True  ' Default checked
    End With
End Sub

Private Sub CreateComboBoxes()
    ' Folder ComboBox
    Set cboFolderName = Me.Controls.Add("Forms.ComboBox.1", "cboFolderName", True)
    With cboFolderName
        .Left = 100
        .Top = 275
        .Width = 360
        .Height = 24
        .Font.Size = 10
        .Style = fmStyleDropDownList
    End With
End Sub

Private Sub CreateButtons()
    ' Search Button
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch", True)
    With btnSearch
        .Left = 10
        .Top = 340
        .Width = 100
        .Height = 30
        .Caption = "Search"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    ' Clear Button
    Set btnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClear", True)
    With btnClear
        .Left = 120
        .Top = 340
        .Width = 100
        .Height = 30
        .Caption = "Clear"
        .Font.Size = 10
    End With
End Sub

Private Sub CreateStatusLabel()
    ' Status Label (bottom of form)
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Left = 10
        .Top = 385
        .Width = 450
        .Height = 35
        .Caption = "Ready to search. Enter search criteria and click Search."
        .Font.Size = 9
        .BackColor = &H8000000F  ' Light gray
        .BorderStyle = fmBorderStyleSingle
        .TextAlign = fmTextAlignLeft
        .WordWrap = True
    End With
End Sub

Private Sub InitializeControlValues()
    ' Populate folder dropdown
    With cboFolderName
        .Clear
        .AddItem "Inbox"
        .AddItem "Sent Items"
        .AddItem "Deleted Items"
        .AddItem "Drafts"
        .AddItem "All Folders (including subfolders)"
        .ListIndex = 4  ' Default to "All Folders"
    End With
    
    ' Clear date fields
    txtDateFrom.Value = ""
    txtDateTo.Value = ""
End Sub

' ============================================
' BUTTON EVENT HANDLERS
' ============================================

' ENTER KEY HANDLER - Execute search when Enter pressed in any textbox
Private Sub txtSearchTerms_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then  ' Enter key
        KeyAscii = 0  ' Suppress beep
        btnSearch_Click  ' Execute search
    End If
End Sub

Private Sub txtFrom_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnSearch_Click
    End If
End Sub

Private Sub txtTo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnSearch_Click
    End If
End Sub

Private Sub txtCC_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnSearch_Click
    End If
End Sub

Private Sub btnSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim criteria As SearchCriteria
    Dim rawSearchTerms As String
    
    ' Build criteria from form fields
    criteria.FromAddress = Trim(txtFrom.Value)
    criteria.ToAddress = Trim(txtTo.Value)
    criteria.CcAddress = Trim(txtCC.Value)
    
    ' Parse search terms: convert , to OR, + to AND, keep * as wildcard
    rawSearchTerms = Trim(txtSearchTerms.Value)
    If Len(rawSearchTerms) > 0 Then
        criteria.SearchTerms = ParseSearchTerms(rawSearchTerms)
    Else
        criteria.SearchTerms = ""
    End If
    
    criteria.FolderName = cboFolderName.Value
    criteria.SearchSubFolders = chkSubFolders.Value
    
    ' Validate: at least one checkbox must be checked if search terms entered
    If Len(criteria.SearchTerms) > 0 Then
        If Not (chkSubject.Value Or chkBody.Value) Then
            MsgBox "Please check where to search (Subject, Body, or both).", vbExclamation, "No Search Location"
            Exit Sub
        End If
    End If
    
    ' Validate: at least SOMETHING to search for
    If Len(criteria.FromAddress) = 0 And Len(criteria.ToAddress) = 0 And Len(criteria.CcAddress) = 0 And Len(criteria.SearchTerms) = 0 Then
        MsgBox "Please enter at least one search criterion (From, To, CC, or Search Terms).", vbExclamation, "No Search Criteria"
        Exit Sub
    End If
    
    ' Parse dates if provided
    If Len(Trim(txtDateFrom.Value)) > 0 Then
        If IsDate(txtDateFrom.Value) Then
            criteria.DateFrom = CDate(txtDateFrom.Value)
        Else
            MsgBox "Invalid 'Date From' format. Please use mm/dd/yyyy.", vbExclamation, "Invalid Date"
            txtDateFrom.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(Trim(txtDateTo.Value)) > 0 Then
        If IsDate(txtDateTo.Value) Then
            criteria.DateTo = CDate(txtDateTo.Value)
        Else
            MsgBox "Invalid 'Date To' format. Please use mm/dd/yyyy.", vbExclamation, "Invalid Date"
            txtDateTo.SetFocus
            Exit Sub
        End If
    End If
    
    ' Update status
    lblStatus.Caption = "Searching..."
    
    ' Execute search via ThisOutlookSession - pass checkboxes for WHERE to search
    ThisOutlookSession.ExecuteSimplifiedSearch criteria, chkSubject.Value, chkBody.Value
    
    lblStatus.Caption = "Search launched. Results shown in Outlook window."
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error: " & Err.Description
    MsgBox "Error during search: " & Err.Description, vbCritical, "Search Error"
End Sub

Private Sub btnClear_Click()
    ' Clear all text fields
    txtFrom.Value = ""
    txtTo.Value = ""
    txtCC.Value = ""
    txtSearchTerms.Value = ""
    txtDateFrom.Value = ""
    txtDateTo.Value = ""
    
    ' Reset checkboxes to default (both checked)
    chkSubject.Value = True
    chkBody.Value = True
    chkSubFolders.Value = True
    
    ' Reset folder to All Folders
    cboFolderName.ListIndex = 4
    
    ' Update status
    lblStatus.Caption = "Ready to search. Enter search criteria and click Search."
    
    ' Focus on first field
    txtFrom.SetFocus
End Sub

' ============================================
' SEARCH TERM PARSER (Operator Conversion)
' ============================================

'-----------------------------------------------------------
' Function: ParseSearchTerms
' Purpose: Convert user-friendly operators to AQS syntax
' Parameters:
'   rawTerms (String) - User input with , + * operators
' Returns: Parsed string with OR/AND operators
' Notes: Comma (,) = OR, Plus (+) = AND, Asterisk (*) = wildcard (kept as-is)
'        Examples:
'          "invoice, receipt" → "invoice OR receipt"
'          "urgent+meeting" → "urgent AND meeting"
'          "proj*" → "proj*" (wildcard stays)
'          "invoice, urgent+project*" → "invoice OR urgent AND project*"
'-----------------------------------------------------------
Private Function ParseSearchTerms(rawTerms As String) As String
    Dim result As String
    
    result = rawTerms
    
    ' Replace comma with OR (spaces around for clarity)
    result = Replace(result, ",", " OR ")
    
    ' Replace plus with AND (spaces around for clarity)
    result = Replace(result, "+", " AND ")
    
    ' Clean up multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ' Trim final result
    ParseSearchTerms = Trim(result)
End Function
