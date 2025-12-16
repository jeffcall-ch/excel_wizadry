' ============================================
' UserForm: frmEmailSearch (ULTRA-COMPACT VERSION)
' Purpose: Email search form - ultra-compact with form-level Enter key
' Features: Narrower form, shorter fields, Enter works anywhere
' Date: 2025-12-16
' Version: 2.4 - Ultra-compact with form-level Enter
' ============================================

Option Explicit

' --------------------------------------------
' THEME CONSTANTS
' --------------------------------------------
Private Const UI_FONT As String = "Segoe UI"
Private Const UI_BG As Long = &HFFFFFF          ' White background
Private Const UI_PANEL As Long = &HF0F0F0       ' Light gray panel
Private Const UI_TEXT As Long = &H000000        ' Black text
Private Const UI_MUTED As Long = &H666666       ' Darker muted gray for labels
Private Const UI_LINE As Long = &HD0D0D0        ' Darker gray separator
Private Const UI_ACCENT As Long = &HC35600      ' Stronger accent (Search button)
Private Const UI_BORDER As Long = &HC0C0C0      ' Panel border

Private Const M As Single = 10                  ' Margin
Private Const G As Single = 8                   ' Gap

' --------------------------------------------
' Module-level controls (WithEvents for event handling)
' --------------------------------------------
Private WithEvents lblFrom As MSForms.Label
Private WithEvents lblTo As MSForms.Label
Private WithEvents lblCC As MSForms.Label
Private WithEvents lblSearchTerms As MSForms.Label

Private WithEvents txtFrom As MSForms.TextBox
Private WithEvents txtTo As MSForms.TextBox
Private WithEvents txtCC As MSForms.TextBox
Private WithEvents txtSearchTerms As MSForms.TextBox

Private WithEvents chkSubject As MSForms.CheckBox
Private WithEvents chkBody As MSForms.CheckBox
Private WithEvents chkCurrentFolderOnly As MSForms.CheckBox

Private WithEvents lblDateFrom As MSForms.Label
Private WithEvents lblDateTo As MSForms.Label
Private WithEvents txtDateFrom As MSForms.TextBox
Private WithEvents txtDateTo As MSForms.TextBox

Private WithEvents lblAttachment As MSForms.Label
Private WithEvents cboAttachment As MSForms.ComboBox

Private WithEvents btnSearch As MSForms.CommandButton
Private WithEvents btnClear As MSForms.CommandButton

Private WithEvents lblStatus As MSForms.Label

' Layout elements
Private lblHeader As MSForms.Label
Private pnlHeaderLine As MSForms.Label
Private fraPeople As MSForms.Frame
Private fraKeywords As MSForms.Frame
Private fraFilters As MSForms.Frame

' ============================================
' FORM INITIALIZATION
' ============================================

Private Sub UserForm_Initialize()
    ApplyTheme
    CreateHeader
    CreatePanels
    CreateLabels
    CreateTextBoxes
    CreateCheckBoxes
    CreateComboBoxes
    CreateButtons
    CreateStatusLabel
    InitializeControlValues
End Sub

Private Sub ApplyTheme()
    Me.Caption = "Email Search"
    Me.Width = 480          ' Ultra-narrow (was 580)
    Me.Height = 340
    Me.BackColor = UI_BG
End Sub

' ============================================
' CONTROL CREATION
' ============================================

Private Sub CreateHeader()
    Set lblHeader = Me.Controls.Add("Forms.Label.1", "lblHeader", True)
    With lblHeader
        .Left = M
        .Top = M
        .Width = Me.Width - (2 * M)
        .Height = 18
        .Caption = "Search mail"
        .Font.Name = UI_FONT
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = UI_TEXT
        .BackStyle = fmBackStyleTransparent
    End With

    Set pnlHeaderLine = Me.Controls.Add("Forms.Label.1", "pnlHeaderLine", True)
    With pnlHeaderLine
        .Left = M
        .Top = lblHeader.Top + lblHeader.Height + 4
        .Width = Me.Width - (2 * M)
        .Height = 1
        .BackColor = UI_LINE
        .Caption = ""
    End With
End Sub

Private Sub CreatePanels()
    Dim topY As Single
    Dim leftX As Single, rightX As Single
    Dim colW As Single

    topY = pnlHeaderLine.Top + 6
    leftX = M
    colW = (Me.Width - (2 * M) - G) / 2
    rightX = leftX + colW + G

    Set fraPeople = AddPanel("fraPeople", "People", leftX, topY, colW, 100)
    Set fraKeywords = AddPanel("fraKeywords", "Keywords", rightX, topY, colW, 100)

    topY = topY + fraPeople.Height + G
    Set fraFilters = AddPanel("fraFilters", "Filters", leftX, topY, Me.Width - (2 * M), 85)
End Sub

Private Function AddPanel(ByVal name As String, ByVal title As String, _
                              ByVal x As Single, ByVal y As Single, _
                              ByVal w As Single, ByVal h As Single) As MSForms.Frame
    Dim f As MSForms.Frame
    Set f = Me.Controls.Add("Forms.Frame.1", name, True)
    With f
        .Left = x
        .Top = y
        .Width = w
        .Height = h
        .Caption = title
        .Font.Name = UI_FONT
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = UI_TEXT
        .BackColor = UI_PANEL
        .BorderColor = UI_BORDER
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With
    Set AddPanel = f
End Function

Private Sub CreateLabels()
    ' PEOPLE panel labels
    Set lblFrom = AddLabel(fraPeople, "From", 8, 22, 35)
    Set lblTo = AddLabel(fraPeople, "To", 8, 48, 35)
    Set lblCC = AddLabel(fraPeople, "CC", 8, 74, 35)

    ' KEYWORDS panel labels
    Set lblSearchTerms = AddLabel(fraKeywords, "Terms", 8, 22, 35)

    ' FILTERS panel labels
    Set lblDateFrom = AddLabel(fraFilters, "From", 8, 24, 30)
    Set lblDateTo = AddLabel(fraFilters, "To", 8, 50, 30)
    Set lblAttachment = AddLabel(fraFilters, "Attachment", 220, 24, 60)
End Sub

Private Function AddLabel(ByVal host As Object, ByVal caption As String, _
                              ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.Label
    Dim l As MSForms.Label
    Set l = host.Controls.Add("Forms.Label.1", "lbl_" & Replace(caption, " ", "_") & "_" & CStr(x) & "_" & CStr(y), True)
    With l
        .Left = x
        .Top = y
        .Width = w
        .Height = 14
        .Caption = caption
        .Font.Name = UI_FONT
        .Font.Size = 8
        .ForeColor = UI_MUTED
        .BackStyle = fmBackStyleTransparent
    End With
    Set AddLabel = l
End Function

Private Sub CreateTextBoxes()
    Dim tbW_People As Single
    Dim tbW_Keywords As Single
    Dim tbW_Date As Single
    
    ' Calculate textbox widths - SHORTER as requested
    tbW_People = (fraPeople.Width - 52) * 0.5     ' 50% of available width
    tbW_Keywords = (fraKeywords.Width - 52) * 0.7 ' 70% of available width
    tbW_Date = 55                                  ' 50% of previous length
    
    ' PEOPLE panel textboxes - HALF LENGTH
    Set txtFrom = AddTextBox(fraPeople, 48, 20, tbW_People)
    Set txtTo = AddTextBox(fraPeople, 48, 46, tbW_People)
    Set txtCC = AddTextBox(fraPeople, 48, 72, tbW_People)
    
    ' KEYWORDS panel textbox - 70% LENGTH
    Set txtSearchTerms = AddTextBox(fraKeywords, 48, 20, tbW_Keywords)
    txtSearchTerms.Height = 28
    txtSearchTerms.MultiLine = True
    
    ' FILTERS panel textboxes - HALF LENGTH
    Set txtDateFrom = AddTextBox(fraFilters, 45, 22, tbW_Date)
    Set txtDateTo = AddTextBox(fraFilters, 45, 48, tbW_Date)
End Sub

Private Function AddTextBox(ByVal host As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox
    Set t = host.Controls.Add("Forms.TextBox.1", "txt_" & CStr(x) & "_" & CStr(y), True)
    With t
        .Left = x
        .Top = y
        .Width = w
        .Height = 18
        .Font.Name = UI_FONT
        .Font.Size = 8
        .BackColor = &HFFFFFF
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With
    Set AddTextBox = t
End Function

Private Sub CreateCheckBoxes()
    Dim chkTop As Single
    chkTop = 52
    
    ' KEYWORDS panel - checkboxes in a row
    Set chkSubject = fraKeywords.Controls.Add("Forms.CheckBox.1", "chkSubject", True)
    With chkSubject
        .Left = 8
        .Top = chkTop
        .Width = 50
        .Height = 16
        .Caption = "Subj"
        .Font.Name = UI_FONT
        .Font.Size = 8
        .ForeColor = UI_TEXT
        .BackColor = UI_PANEL
    End With
    
    Set chkBody = fraKeywords.Controls.Add("Forms.CheckBox.1", "chkBody", True)
    With chkBody
        .Left = 62
        .Top = chkTop
        .Width = 42
        .Height = 16
        .Caption = "Body"
        .Font.Name = UI_FONT
        .Font.Size = 8
        .ForeColor = UI_TEXT
        .BackColor = UI_PANEL
    End With
    
    Set chkCurrentFolderOnly = fraKeywords.Controls.Add("Forms.CheckBox.1", "chkCurrentFolderOnly", True)
    With chkCurrentFolderOnly
        .Left = 108
        .Top = chkTop
        .Width = fraKeywords.Width - 116
        .Height = 16
        .Caption = "Cur.Fld"
        .Font.Name = UI_FONT
        .Font.Size = 8
        .ForeColor = UI_TEXT
        .BackColor = UI_PANEL
        .Value = False
    End With
    
    ' Help text below checkboxes
    Dim lblHelp As MSForms.Label
    Set lblHelp = fraKeywords.Controls.Add("Forms.Label.1", "lblHelp", True)
    With lblHelp
        .Left = 8
        .Top = chkTop + 18
        .Width = fraKeywords.Width - 16
        .Height = 16
        .Caption = ", = OR  + = AND  * = wildcard  Enter = Search"
        .Font.Name = UI_FONT
        .Font.Size = 7
        .ForeColor = UI_MUTED
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Private Sub CreateComboBoxes()
    ' FILTERS panel - Attachment dropdown
    Set cboAttachment = fraFilters.Controls.Add("Forms.ComboBox.1", "cboAttachment", True)
    With cboAttachment
        .Left = 285
        .Top = 22
        .Width = fraFilters.Width - 293
        .Height = 18
        .Font.Name = UI_FONT
        .Font.Size = 8
        .BackColor = &HFFFFFF
        .Style = fmStyleDropDownList
    End With
    
    ' Date format help
    Dim lblDateHelp As MSForms.Label
    Set lblDateHelp = fraFilters.Controls.Add("Forms.Label.1", "lblDateHelp", True)
    With lblDateHelp
        .Left = 110
        .Top = 36
        .Width = 100
        .Height = 12
        .Caption = "dd.mm.yyyy"
        .Font.Name = UI_FONT
        .Font.Size = 7
        .ForeColor = UI_MUTED
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Private Sub CreateButtons()
    Dim btnY As Single
    Dim btnH As Single
    Dim btnW As Single
    Dim btnGap As Single
    
    btnY = fraFilters.Top + fraFilters.Height + G
    btnH = 24
    btnW = 70
    btnGap = 6
    
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch", True)
    With btnSearch
        .Left = M
        .Top = btnY
        .Width = btnW
        .Height = btnH
        .Caption = "Search"
        .Font.Name = UI_FONT
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = &HFFFFFF
        .BackColor = UI_ACCENT
    End With
    
    Set btnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClear", True)
    With btnClear
        .Left = btnSearch.Left + btnSearch.Width + btnGap
        .Top = btnY
        .Width = btnW
        .Height = btnH
        .Caption = "Clear"
        .Font.Name = UI_FONT
        .Font.Size = 9
        .ForeColor = UI_TEXT
        .BackColor = UI_PANEL
    End With
End Sub

Private Sub CreateStatusLabel()
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Left = M
        .Top = btnSearch.Top + btnSearch.Height + 4
        .Width = Me.Width - (2 * M)
        .Height = 14
        .Caption = "Ready. Press Enter to search."
        .Font.Name = UI_FONT
        .Font.Size = 8
        .ForeColor = UI_MUTED
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Private Sub InitializeControlValues()
    With cboAttachment
        .Clear
        .AddItem "--- (Any)"
        .AddItem "With Attachments"
        .AddItem "Without Attachments"
        .AddItem "PDF Files"
        .AddItem "Excel Files"
        .AddItem "Word Files"
        .AddItem "PowerPoint"
        .AddItem "Images"
        .AddItem "Text Files"
        .AddItem "ZIP/Archives"
        .ListIndex = 0
    End With
    
    chkSubject.Value = True
    chkBody.Value = True
    chkCurrentFolderOnly.Value = False
    
    txtDateFrom.Value = ""
    txtDateTo.Value = ""
    
    lblStatus.Caption = "Ready. Press Enter to search."
    txtFrom.SetFocus
End Sub

' ============================================
' INDIVIDUAL TEXTBOX ENTER KEY HANDLERS
' (Keep these for redundancy, though form-level handler covers all)
' ============================================

Private Sub txtSearchTerms_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then KeyAscii = 0: btnSearch_Click
End Sub

Private Sub txtFrom_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then KeyAscii = 0: btnSearch_Click
End Sub

Private Sub txtTo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then KeyAscii = 0: btnSearch_Click
End Sub

Private Sub txtCC_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then KeyAscii = 0: btnSearch_Click
End Sub

Private Sub txtDateFrom_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then KeyAscii = 0: btnSearch_Click
End Sub

Private Sub txtDateTo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then KeyAscii = 0: btnSearch_Click
End Sub

' ============================================
' CHECKBOX AND COMBOBOX ENTER KEY HANDLERS
' ============================================

Private Sub chkSubject_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub

Private Sub chkBody_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub

Private Sub chkCurrentFolderOnly_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub

Private Sub cboAttachment_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub

' ============================================
' BUTTON EVENT HANDLERS
' ============================================

Private Sub btnSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim fromAddr As String, toAddr As String, ccAddr As String
    Dim searchTerms As String, dateFrom As String, dateTo As String
    Dim attachmentType As Integer
    Dim searchInSubject As Boolean, searchInBody As Boolean
    Dim currentFolderOnly As Boolean
    
    fromAddr = Trim(txtFrom.Value)
    toAddr = Trim(txtTo.Value)
    ccAddr = Trim(txtCC.Value)
    searchTerms = Trim(txtSearchTerms.Value)
    dateFrom = Trim(txtDateFrom.Value)
    dateTo = Trim(txtDateTo.Value)
    attachmentType = cboAttachment.ListIndex
    searchInSubject = chkSubject.Value
    searchInBody = chkBody.Value
    currentFolderOnly = chkCurrentFolderOnly.Value
    
    If Len(searchTerms) > 0 Then
        If Not (searchInSubject Or searchInBody) Then
            MsgBox "Please check where to search (Subject, Body, or both).", vbExclamation, "No Search Location"
            lblStatus.Caption = "Error: No search location selected."
            Exit Sub
        End If
    End If
    
    If Len(dateFrom) > 0 Then
        If Not IsValidDateFormat(dateFrom) Then
            MsgBox "Invalid 'Date From' format. Please use dd.mm.yyyy.", vbExclamation, "Invalid Date"
            txtDateFrom.SetFocus
            lblStatus.Caption = "Error: Invalid date format."
            Exit Sub
        End If
    End If
    
    If Len(dateTo) > 0 Then
        If Not IsValidDateFormat(dateTo) Then
            MsgBox "Invalid 'Date To' format. Please use dd.mm.yyyy.", vbExclamation, "Invalid Date"
            txtDateTo.SetFocus
            lblStatus.Caption = "Error: Invalid date format."
            Exit Sub
        End If
    End If
    
    lblStatus.Caption = "Searching..."
    DoEvents
    
    ThisOutlookSession.ExecuteAQSSearch _
        fromAddr, toAddr, ccAddr, searchTerms, _
        dateFrom, dateTo, attachmentType, _
        searchInSubject, searchInBody, currentFolderOnly
    
    lblStatus.Caption = "Search executed. Press Enter to search again."
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error: " & Err.Description
    MsgBox "Error during search: " & Err.Description, vbCritical, "Search Error"
End Sub

Private Sub btnClear_Click()
    txtFrom.Value = ""
    txtTo.Value = ""
    txtCC.Value = ""
    txtSearchTerms.Value = ""
    txtDateFrom.Value = ""
    txtDateTo.Value = ""
    
    chkSubject.Value = True
    chkBody.Value = True
    chkCurrentFolderOnly.Value = False
    
    cboAttachment.ListIndex = 0
    
    lblStatus.Caption = "Ready. Press Enter to search."
    txtFrom.SetFocus
End Sub

Private Function IsValidDateFormat(ByVal dateStr As String) As Boolean
    On Error GoTo InvalidDate
    
    Dim parts() As String
    Dim day As Integer, month As Integer, year As Integer
    
    parts = Split(dateStr, ".")
    
    If UBound(parts) - LBound(parts) + 1 <> 3 Then
        IsValidDateFormat = False
        Exit Function
    End If
    
    day = CInt(parts(0))
    month = CInt(parts(1))
    year = CInt(parts(2))
    
    Dim testDate As Date
    testDate = DateSerial(year, month, day)
    
    IsValidDateFormat = True
    Exit Function
    
InvalidDate:
    IsValidDateFormat = False
End Function
