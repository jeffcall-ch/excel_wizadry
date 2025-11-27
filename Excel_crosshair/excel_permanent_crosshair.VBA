' ========================================
' PERMANENT EXCEL CROSSHAIR - CONDITIONAL FORMATTING VERSION
' ========================================
' This module creates a red border around the active cell's row and column
' using conditional formatting with formulas.
'
' INSTALLATION:
' 1. Open Excel VBA Editor (Alt+F11)
' 2. Insert this code into a new Module
' 3. Insert the Workbook event code into ThisWorkbook
' 4. Save and restart Excel
' ========================================

Option Explicit

' ========================================
' MODULE-LEVEL VARIABLES
' ========================================

Private m_IsRunning As Boolean

' ========================================
' PUBLIC INTERFACE
' ========================================

Public Sub StartCrosshair()
    ' Initialize and start the crosshair using conditional formatting
    If m_IsRunning Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Set up conditional formatting for the crosshair
    SetupConditionalFormatting ws
    
    m_IsRunning = True
    
    MsgBox "Crosshair enabled! The active cell's row and column will be highlighted with red borders.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error starting crosshair: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
End Sub

Public Sub StopCrosshair()
    ' Stop and remove the crosshair conditional formatting
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    RemoveConditionalFormatting ws
    
    m_IsRunning = False
    
    MsgBox "Crosshair disabled.", vbInformation
End Sub

Public Sub SetupAllSheets()
    ' Apply crosshair to all sheets in workbook
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim count As Long
    
    For Each ws In ThisWorkbook.Worksheets
        SetupConditionalFormatting ws
        count = count + 1
    Next ws
    
    m_IsRunning = True
    
    MsgBox "Crosshair enabled on " & count & " sheets.", vbInformation
End Sub

Public Sub RemoveAllSheets()
    ' Remove crosshair from all sheets in workbook
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim count As Long
    
    For Each ws In ThisWorkbook.Worksheets
        RemoveConditionalFormatting ws
        count = count + 1
    Next ws
    
    m_IsRunning = False
    
    MsgBox "Crosshair removed from " & count & " sheets.", vbInformation
End Sub

' ========================================
' PRIVATE HELPER FUNCTIONS
' ========================================

Private Sub SetupConditionalFormatting(ws As Worksheet)
    ' Set up conditional formatting rules for row and column highlighting
    On Error Resume Next
    
    Dim usedRange As Range
    Dim fc As FormatCondition
    Dim ruleRange As String
    
    ' Remove existing crosshair rules first
    RemoveConditionalFormatting ws
    
    ' Get the used range (or larger if you want it to cover everything)
    Set usedRange = ws.UsedRange
    If usedRange Is Nothing Then
        Set usedRange = ws.Range("A1:Z100") ' Default range if sheet is empty
    Else
        ' Expand to cover more area
        Set usedRange = ws.Range(usedRange.Address).Resize(usedRange.Rows.count + 100, _
                                                             usedRange.Columns.count + 26)
    End If
    
    ruleRange = usedRange.Address
    
    ' Rule 1: Highlight row of active cell (yellow background)
    Set fc = ws.Range(ruleRange).FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=ROW()=CELL(""row"")")
    
    With fc
        .StopIfTrue = False
        .Interior.Color = RGB(255, 255, 0) ' Yellow
    End With
    
    ' Rule 2: Highlight column of active cell (yellow background)
    Set fc = ws.Range(ruleRange).FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=COLUMN()=CELL(""col"")")
    
    With fc
        .StopIfTrue = False
        .Interior.Color = RGB(255, 255, 0) ' Yellow
    End With
    
End Sub

Private Sub RemoveConditionalFormatting(ws As Worksheet)
    ' Remove all conditional formatting rules with CELL formulas (crosshair rules)
    On Error Resume Next
    
    Dim fc As FormatCondition
    Dim i As Long
    
    ' Loop backwards through format conditions to safely delete
    For i = ws.Cells.FormatConditions.count To 1 Step -1
        Set fc = ws.Cells.FormatConditions(i)
        If InStr(1, fc.Formula1, "CELL(", vbTextCompare) > 0 Then
            fc.Delete
        End If
    Next i
    
End Sub

' ========================================
' MANUAL TESTING
' ========================================
' Run these from VBA Immediate Window (Ctrl+G):
'
' StartCrosshair       - Enable crosshair on active sheet
' StopCrosshair        - Disable crosshair on active sheet
' SetupAllSheets       - Enable crosshair on all sheets
' RemoveAllSheets      - Disable crosshair on all sheets
'
' ========================================
' EVENT HANDLERS (Add to ThisWorkbook)
' ========================================
'
' Copy the following code to ThisWorkbook module:
'
' Private Sub Workbook_Open()
'     SetupAllSheets
' End Sub
'
' Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'     Application.ScreenUpdating = True
'     Application.Calculate
' End Sub
