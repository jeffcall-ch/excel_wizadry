' ================================================================
' MODULE SPLIT INSTRUCTIONS:
'
' 1. In VBA editor (Alt+F11), expand PERSONAL.XLSB
' 2. Double-click "ThisWorkbook" and paste ONLY Workbook_Open there
' 3. In a Standard Module (e.g. Module1), paste RegisterShortcut
'    and ClearAllFilters
' ================================================================

' ---------------------------------------------------------------
' >>> ThisWorkbook module of PERSONAL.XLSB <<<
' ---------------------------------------------------------------
Private Sub Workbook_Open()
    Dim targetMacro As String
    ' Dynamically fetch the exact workbook name (correct extension casing)
    targetMacro = "'" & ThisWorkbook.Name & "'!RegisterShortcut"
    ' Delay 2 seconds to allow Excel UI to finish initializing
    Application.OnTime Now + TimeValue("00:00:02"), targetMacro
End Sub

' ---------------------------------------------------------------
' >>> Standard Module of PERSONAL.XLSB (e.g. Module1) <<<
' NOTE: Do NOT name the module RegisterShortcut or ClearAllFilters
' ---------------------------------------------------------------
Sub RegisterShortcut()
    Dim targetMacro As String
    ' Dynamically build the reference so casing always matches
    targetMacro = "'" & ThisWorkbook.Name & "'!ClearAllFilters"
    Application.OnKey "^+q", targetMacro
End Sub

Sub ClearAllFilters()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim pt As PivotTable
    Dim pf As PivotField

    Set ws = ActiveSheet

    ' --- 1. Clear filters on defined Tables (ListObjects) ---
    For Each lo In ws.ListObjects
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    Next lo

    ' --- 2. Clear worksheet-level AutoFilter (normal sheets) ---
    If ws.AutoFilterMode Then
        If ws.FilterMode Then ws.ShowAllData
    End If

    ' --- 3. Clear filters on all PivotTables ---
    For Each pt In ws.PivotTables
        For Each pf In pt.PivotFields
            On Error Resume Next
            pf.ClearAllFilters
            On Error GoTo 0
        Next pf
        On Error Resume Next
        pt.ClearAllFilters
        On Error GoTo 0
    Next pt

End Sub
