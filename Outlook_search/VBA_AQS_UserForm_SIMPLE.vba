' ============================================
' UserForm: frmEmailSearch (LAYOUT OPTIMIZED)
' Purpose: Email search with modern Outlook-inspired design
' Features: Panel-based layout, refined typography, better visual hierarchy
' Date: 2025-12-15
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

Private Const M As Single = 14                  ' Margin (14pt)
Private Const G As Single = 12                  ' Gap (12pt)

' --------------------------------------------
' Module-level controls (WithEvents for event handling)
' --------------------------------------------
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

Private WithEvents lblAttachment As MSForms.Label
Private WithEvents cboAttachment As MSForms.ComboBox

Private WithEvents btnSearch As MSForms.CommandButton
Private WithEvents btnClear As MSForms.CommandButton

Private WithEvents lblStatus As MSForms.Label

' Layout elements
Private lblHeader As MSForms.Label
Private lblSubHeader As MSForms.Label
Private pnlHeaderLine As MSForms.Label
Private fraPeople As MSForms.Frame
Private fraQuery As MSForms.Frame
Private fraScope As MSForms.Frame
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
    Me.Width = 640
    ' *** OPTIMIZED: Reduced height for tighter fit ***
    Me.Height = 465 
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
        .Height = 22
        .Caption = "Search mail"
        .Font.Name = UI_FONT
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = UI_TEXT
        .BackStyle = fmBackStyleTransparent
    End With

    Set lblSubHeader = Me.Controls.Add("Forms.Label.1", "lblSubHeader", True)
    With lblSubHeader
        .Left = M
        .Top = lblHeader.Top + lblHeader.Height + 2
        .Width = Me.Width - (2 * M)
        .Height = 16
        .Caption = "Enter criteria and press Enter or click Search."
        .Font.Name = UI_FONT
        .Font.Size = 9
        .ForeColor = UI_MUTED
        .BackStyle = fmBackStyleTransparent
    End With

    Set pnlHeaderLine = Me.Controls.Add("Forms.Label.1", "pnlHeaderLine", True)
    With pnlHeaderLine
        .Left = M
        .Top = lblSubHeader.Top + lblSubHeader.Height + 10
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

    topY = pnlHeaderLine.Top + 10
    leftX = M
    ' Calculate column width based on total form width
    colW = (Me.Width - (2 * M) - G) / 2
    rightX = leftX + colW + G

    Set fraPeople = AddPanel("fraPeople", "People", leftX, topY, colW, 135)
    Set fraQuery = AddPanel("fraQuery", "Keywords", rightX, topY, colW, 135)

    topY = topY + fraPeople.Height + G
    Set fraScope = AddPanel("fraScope", "Scope", leftX, topY, colW, 138)
    Set fraFilters = AddPanel("fraFilters", "Filters", rightX, topY, colW, 138)
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
        .Font.Size = 10
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
    Set lblFrom = AddLabel(fraPeople, "From", 10, 26, 50)
    Set lblTo = AddLabel(fraPeople, "To", 10, 60, 50)
    Set lblCC = AddLabel(fraPeople, "CC", 10, 94, 50)

    ' QUERY panel labels
    Set lblSearchTerms = AddLabel(fraQuery, "Search terms", 10, 26, 120)
    lblSearchTerms.Font.Bold = True

    Set lblSearchHelp = fraQuery.Controls.Add("Forms.Label.1", "lblSearchHelp", True)
    With lblSearchHelp
        .Left = 10
        .Top = 96
        .Width = fraQuery.Width - 20
        .Height = 28
        .Caption = "Operators: comma = OR, plus = AND, * = wildcard. Press Enter to search."
        .Font.Name = UI_FONT
        .Font.Size = 8
        .ForeColor = UI_MUTED
        .BackStyle = fmBackStyleTransparent
    End With

    ' FILTERS panel labels
    Set lblDateFrom = AddLabel(fraFilters, "Date from", 10, 26, 50)
    Set lblDateTo = AddLabel(fraFilters, "Date to", 10, 60, 50)
    
    ' Date format help text
    Dim lblDateHelp As MSForms.Label
    Set lblDateHelp = fraFilters.Controls.Add("Forms.Label.1", "lblDateHelp", True)
    With lblDateHelp
        .Left = 65
        .Top = 80
        .Width = fraFilters.Width - 75
        .Height = 12
        .Caption = "Format: dd.mm.yyyy"
        .Font.Name = UI_FONT
        .Font.Size = 8
        .ForeColor = UI_MUTED
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Private Function AddLabel(ByVal host As Object, ByVal caption As String, _
                              ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.Label
    Dim l As MSForms.Label
    Set l = host.Controls.Add("Forms.Label.1", "lbl_" & Replace(caption, " ", "_") & "_" & CStr(x) & "_" & CStr(y), True)
    With l
        .Left = x
        .Top = y
        .Width = w
        .Height = 16
        .Caption = caption
        .Font.Name = UI_FONT
        .Font.Size = 9
        .ForeColor = UI_MUTED
        .BackStyle = fmBackStyleTransparent
    End With
    Set AddLabel = l
End Function

Private Sub CreateTextBoxes()
    Dim tbW_People As Single 
    Dim tbW_Query As Single  

    ' *** OPTIMIZED: Shorter width for People textboxes ***
    tbW_People = 200 

    Set txtFrom = AddTextBox(fraPeople, "txtFrom", 65, 22, tbW_People)
    Set txtTo = AddTextBox(fraPeople, "txtTo", 65, 56, tbW_People)
    Set txtCC = AddTextBox(fraPeople, "txtCC", 65, 90, tbW_People)

    ' *** OPTIMIZED: Shorter width for Search Terms textbox ***
    tbW_Query = 250 
    Set txtSearchTerms = AddTextBox(fraQuery, "txtSearchTerms", 10, 46, tbW_Query)
    txtSearchTerms.Font.Bold = True

    ' Add Subject and Body checkboxes in Keywords panel (grouped with search terms)
    Set chkSubject = fraQuery.Controls.Add("Forms.CheckBox.1", "chkSubject", True)
    With chkSubject
        .Left = 10
        .Top = 72
        .Width = 100
        .Height = 18
        .Caption = "Subject"
        .Font.Name = UI_FONT
        .Font.Size = 9
        .Value = True
    End With

    Set chkBody = fraQuery.Controls.Add("Forms.CheckBox.1", "chkBody", True)
    With chkBody
        .Left = 120
        .Top = 72
        .Width = 100
        .Height = 18
        .Caption = "Body"
        .Font.Name = UI_FONT
        .Font.Size = 9
        .Value = True
    End With

    Set txtDateFrom = AddTextBox(fraFilters, "txtDateFrom", 65, 22, fraFilters.Width - 75)
    Set txtDateTo = AddTextBox(fraFilters, "txtDateTo", 65, 56, fraFilters.Width - 75)
End Sub

Private Function AddTextBox(ByVal host As Object, ByVal name As String, _
                                ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.TextBox
    Dim t As MSForms.TextBox
    Set t = host.Controls.Add("Forms.TextBox.1", name, True)
    With t
        .Left = x
        .Top = y
        .Width = w
        .Height = 22
        .Font.Name = UI_FONT
        .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = UI_BORDER
        .SpecialEffect = fmSpecialEffectFlat
        .EnterKeyBehavior = False
    End With
    Set AddTextBox = t
End Function

Private Sub CreateCheckBoxes()
    Set chkSubFolders = fraScope.Controls.Add("Forms.CheckBox.1", "chkSubFolders", True)
    With chkSubFolders
        .Left = 10
        .Top = 26
        .Width = fraScope.Width - 20
        .Height = 18
        .Caption = "Include subfolders"
        .Font.Name = UI_FONT
        .Font.Size = 9
        .Value = True
    End With
End Sub

Private Sub CreateComboBoxes()
    ' Folder label + combobox (Scope panel)
    Set lblFolder = AddLabel(fraScope, "Folder", 10, 46, 80)
    Set cboFolderName = fraScope.Controls.Add("Forms.ComboBox.1", "cboFolderName", True)
    With cboFolderName
        .Left = 10
        .Top = 60
        .Width = fraScope.Width - 20
        .Height = 22
        .Font.Name = UI_FONT
        .Font.Size = 10
        .Style = fmStyleDropDownList
    End With

    ' Attachment label + combobox (Filters panel)
    Set lblAttachment = AddLabel(fraFilters, "Attachments", 10, 84, 90)
    Set cboAttachment = fraFilters.Controls.Add("Forms.ComboBox.1", "cboAttachment", True)
    With cboAttachment
        .Left = 10
        .Top = 100
        .Width = fraFilters.Width - 20
        .Height = 22
        .Font.Name = UI_FONT
        .Font.Size = 10
        .Style = fmStyleDropDownList
    End With
End Sub

Private Sub CreateButtons()
    ' Positioned below the main panels
    Set btnSearch = Me.Controls.Add("Forms.CommandButton.1", "btnSearch", True)
    With btnSearch
        .Left = M
        .Top = fraScope.Top + fraScope.Height + G
        .Width = 130
        .Height = 30
        .Caption = "Search"
        .Font.Name = UI_FONT
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = UI_ACCENT
        .ForeColor = vbWhite
        .TakeFocusOnClick = False
        .Default = True  ' Makes Enter key trigger this button from anywhere in the form
    End With

    Set btnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClear", True)
    With btnClear
        .Left = btnSearch.Left + btnSearch.Width + 10
        .Top = btnSearch.Top
        .Width = 110
        .Height = 30
        .Caption = "Clear"
        .Font.Name = UI_FONT
        .Font.Size = 10
        .BackColor = &HEFEFEF
        .ForeColor = UI_TEXT
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub CreateStatusLabel()
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Left = M
        .Top = btnSearch.Top + btnSearch.Height + 10
        .Width = Me.Width - (2 * M)
        .Height = 34
        .Caption = "Ready to search."
        .Font.Name = UI_FONT
        .Font.Size = 9
        .BackColor = &HFAFAFA
        .BorderStyle = fmBorderStyleSingle
        .ForeColor = UI_TEXT
        .TextAlign = fmTextAlignLeft
        .WordWrap = True
    End With
End Sub

Private Sub InitializeControlValues()
    ' Populate folder dropdown
    With cboFolderName
        .Clear
        .AddItem "Current Folder"
        .AddItem "Inbox"
        .AddItem "Sent Items"
        .AddItem "Deleted Items"
        .AddItem "Drafts"
        .AddItem "All Folders (including subfolders)"
        .ListIndex = 5  ' Default to "All Folders"
    End With
    
    ' Populate attachment filter dropdown
    With cboAttachment
        .Clear
        .AddItem "--- (Any)"
        .AddItem "With Attachments"
        .AddItem "Without Attachments"
        .AddItem "PDF Files"
        .AddItem "Excel Files (.xls, .xlsx)"
        .AddItem "Word Files (.doc, .docx)"
        .AddItem "Images (.jpg, .png, .gif)"
        .AddItem "PowerPoint Files (.ppt, .pptx)"
        .AddItem "ZIP/Archives"
        .ListIndex = 0  ' Default to "Any"
    End With
    
    ' Set search scope defaults
    chkSubject.Value = True
    chkBody.Value = True
    chkSubFolders.Value = True
    
    ' Clear date fields
    txtDateFrom.Value = ""
    txtDateTo.Value = ""
    
    lblStatus.Caption = "Ready to search."
    txtFrom.SetFocus
End Sub

' ============================================
' BUTTON EVENT HANDLERS (Functionality)
' ============================================

' ENTER KEY HANDLER - Execute search when Enter pressed in any relevant textbox
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

Private Sub btnSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim criteria As SearchCriteria
    Dim rawSearchTerms As String
    
    ' 1. Collect and parse data
    criteria.FromAddress = Trim(txtFrom.Value)
    criteria.ToAddress = Trim(txtTo.Value)
    criteria.CcAddress = Trim(txtCC.Value)
    
    rawSearchTerms = Trim(txtSearchTerms.Value)
    If Len(rawSearchTerms) > 0 Then
        criteria.SearchTerms = ParseSearchTerms(rawSearchTerms)
    Else
        criteria.SearchTerms = ""
    End If
    
    criteria.FolderName = cboFolderName.Value
    criteria.SearchSubFolders = chkSubFolders.Value
    criteria.AttachmentFilter = BuildAttachmentFilterToken()
    
    ' 2. Validation Checks
    ' Check if search location is selected if terms exist
    If Len(criteria.SearchTerms) > 0 Then
        If Not (chkSubject.Value Or chkBody.Value) Then
            MsgBox "Please check where to search (Subject, Body, or both).", vbExclamation, "No Search Location"
            Exit Sub
        End If
    End If
    
    ' Check for any criteria
    If Len(criteria.FromAddress) = 0 And Len(criteria.ToAddress) = 0 And Len(criteria.CcAddress) = 0 And Len(criteria.SearchTerms) = 0 Then
        MsgBox "Please enter at least one search criterion (From, To, CC, or Search Terms).", vbExclamation, "No Search Criteria"
        Exit Sub
    End If
    
    ' Parse Date From
    If Len(Trim(txtDateFrom.Value)) > 0 Then
        Dim parsedDateFrom As Date
        If ParseDDMMYYYY(txtDateFrom.Value, parsedDateFrom) Then
            criteria.DateFrom = parsedDateFrom
        Else
            MsgBox "Invalid 'Date From' format. Please use dd.mm.yyyy.", vbExclamation, "Invalid Date"
            txtDateFrom.SetFocus
            Exit Sub
        End If
    End If
    
    ' Parse Date To
    If Len(Trim(txtDateTo.Value)) > 0 Then
        Dim parsedDateTo As Date
        If ParseDDMMYYYY(txtDateTo.Value, parsedDateTo) Then
            criteria.DateTo = parsedDateTo
        Else
            MsgBox "Invalid 'Date To' format. Please use dd.mm.yyyy.", vbExclamation, "Invalid Date"
            txtDateTo.SetFocus
            Exit Sub
        End If
    End If
    
    ' 3. Execute Search
    lblStatus.Caption = "Searching..."
    
    ' Execute search via ThisOutlookSession - criteria is passed by value (copy)
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
    
    ' Reset checkboxes to default
    chkSubject.Value = True
    chkBody.Value = True
    chkSubFolders.Value = True
    
    ' Reset dropdowns
    cboFolderName.ListIndex = 5
    cboAttachment.ListIndex = 0
    
    lblStatus.Caption = "Ready to search."
    txtFrom.SetFocus
End Sub

' ============================================
' UTILITY FUNCTIONS
' ============================================

Private Function BuildAttachmentFilterToken() As String
    Select Case cboAttachment.ListIndex
        Case 0  ' --- (Any)
            BuildAttachmentFilterToken = ""
        Case 1  ' With Attachments
            BuildAttachmentFilterToken = "HASATT"
        Case 2  ' Without Attachments
            BuildAttachmentFilterToken = "NOATT"
        Case 3  ' PDF Files
            BuildAttachmentFilterToken = "HASATT|PDF"
        Case 4  ' Excel Files (.xls, .xlsx)
            BuildAttachmentFilterToken = "HASATT|XLS"
        Case 5  ' Word Files (.doc, .docx)
            BuildAttachmentFilterToken = "HASATT|DOC"
        Case 6  ' Images (.jpg, .png, .gif)
            BuildAttachmentFilterToken = "HASATT|IMG"
        Case 7  ' PowerPoint Files (.ppt, .pptx)
            BuildAttachmentFilterToken = "HASATT|PPT"
        Case 8  ' ZIP/Archives
            BuildAttachmentFilterToken = "HASATT|ZIP"
        Case Else
            BuildAttachmentFilterToken = ""
    End Select
End Function

Private Function ParseDDMMYYYY(ByVal dateStr As String, ByRef outDate As Date) As Boolean
    On Error GoTo ParseError
    
    Dim parts() As String
    Dim day As Integer, month As Integer, year As Integer
    
    ' Split by dot
    parts = Split(Trim(dateStr), ".")
    
    ' Must have exactly 3 parts
    If UBound(parts) - LBound(parts) + 1 <> 3 Then
        ParseDDMMYYYY = False
        Exit Function
    End If
    
    ' Parse day, month, year
    day = CInt(parts(0))
    month = CInt(parts(1))
    year = CInt(parts(2))
    
    ' Create date using DateSerial (handles date validity automatically)
    outDate = DateSerial(year, month, day)
    ParseDDMMYYYY = True
    Exit Function
    
ParseError:
    ParseDDMMYYYY = False
End Function

Private Function ParseSearchTerms(rawTerms As String) As String
    Dim result As String
    
    result = rawTerms
    
    ' Replace comma with OR
    result = Replace(result, ",", " OR ")
    
    ' Replace plus with AND
    result = Replace(result, "+", " AND ")
    
    ' Clean up multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ParseSearchTerms = Trim(result)
End Function
