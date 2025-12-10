' ================================================================================
' SU-Based Comment Replicator
' ================================================================================
' Purpose: Automatically replicate column T values to all rows with matching
'          column D values (case-insensitive)
'
' How to Install:
' 1. Open your Excel file
' 2. Press ALT + F11 to open VBA Editor
' 3. Find your worksheet in the left panel (e.g., "Sheet1")
' 4. Double-click the worksheet to open its code window
' 5. Copy and paste this code into that window
' 6. Close VBA Editor and save the file as .xlsm (macro-enabled)
'
' How it Works:
' - When you enter/change a value in column T and press ENTER
' - Script checks if the corresponding column D value is valid (not -, empty, N/A, error)
' - Finds ALL rows with the same column D value
' - Copies the T value to all matching rows
' - Also copies column H to column U for all matching rows
' - Shows a quick status message when done
' ================================================================================

Option Explicit

' Module-level variables
Private previousFilterState As String
Private lastStatusMessage As String

' ================================================================================
' Initialize filter tracking when workbook opens
' ================================================================================
Private Sub Worksheet_Activate()
    previousFilterState = GetFilterState()
    UpdateUniqueCountStatus
End Sub

' ================================================================================
' Update status bar when selection changes (show cached stats)
' ================================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Check if filter state changed
    Dim currentFilterState As String
    currentFilterState = GetFilterState()
    
    If currentFilterState <> previousFilterState Then
        ' Filter changed - recalculate
        previousFilterState = currentFilterState
        UpdateUniqueCountStatus
    ElseIf lastStatusMessage <> "" Then
        ' Restore cached message if it was overwritten
        Application.StatusBar = lastStatusMessage
    End If
End Sub

' ================================================================================
' Update when Calculate event fires (backup for filter changes)
' ================================================================================
Private Sub Worksheet_Calculate()
    Dim currentFilterState As String
    currentFilterState = GetFilterState()
    
    ' Only update if filter state actually changed
    If currentFilterState <> previousFilterState Then
        previousFilterState = currentFilterState
        UpdateUniqueCountStatus
    End If
End Sub

' ================================================================================
' Get current filter state as a string
' ================================================================================
Private Function GetFilterState() As String
    Dim filterRange As Range
    Dim result As String
    Dim i As Long
    
    On Error Resume Next
    If Me.AutoFilterMode And Not Me.AutoFilter Is Nothing Then
        Set filterRange = Me.AutoFilter.Range
        If Not filterRange Is Nothing Then
            result = "FILTERED:"
            ' Track which columns are filtered
            For i = 1 To Me.AutoFilter.Filters.Count
                If Me.AutoFilter.Filters(i).On Then
                    result = result & i & ","
                End If
            Next i
        Else
            result = "NOFILTER"
        End If
    Else
        result = "NOFILTER"
    End If
    On Error GoTo 0
    
    GetFilterState = result
End Function

' ================================================================================
' Helper function to update status bar with unique count from column D
' ================================================================================
Private Sub UpdateUniqueCountStatus()
    On Error Resume Next
    Dim lastRow As Long
    Dim visibleRange As Range
    Dim cell As Range
    Dim uniqueDict As Object
    Dim dValue As Variant
    Dim totalVisible As Long
    Dim uniqueCount As Long
    
    ' Create dictionary to track unique values
    Set uniqueDict = CreateObject("Scripting.Dictionary")
    If uniqueDict Is Nothing Then Exit Sub
    uniqueDict.CompareMode = 1 ' Case-insensitive
    
    ' Find last used row in column D
    lastRow = Me.Cells(Me.Rows.Count, 4).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' Get the range (assuming row 1 is header)
    Set visibleRange = Me.Range(Me.Cells(2, 4), Me.Cells(lastRow, 4))
    
    ' Check if autofilter is active
    If Me.AutoFilterMode And Not Me.AutoFilter Is Nothing Then
        ' Get only visible cells
        On Error Resume Next
        Set visibleRange = visibleRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If visibleRange Is Nothing Then
            Application.StatusBar = "Column D: 0 visible rows | 0 unique values"
            Exit Sub
        End If
    End If
    
    ' Count unique values in visible cells
    totalVisible = 0
    For Each cell In visibleRange
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            dValue = Trim(CStr(cell.Value))
            ' Exclude invalid values
            If UCase(dValue) <> "-" And UCase(dValue) <> "N/A" And Len(dValue) > 0 Then
                If Not uniqueDict.Exists(dValue) Then
                    uniqueDict.Add dValue, 1
                End If
                totalVisible = totalVisible + 1
            End If
        End If
    Next cell
    
    uniqueCount = uniqueDict.Count
    
    ' Build status message
    If Me.AutoFilterMode And Not Me.AutoFilter Is Nothing Then
        lastStatusMessage = "Column D: " & totalVisible & " visible rows | " & uniqueCount & " unique values (filtered)"
    Else
        lastStatusMessage = "Column D: " & totalVisible & " rows | " & uniqueCount & " unique values"
    End If
    
    ' Update status bar
    Application.StatusBar = lastStatusMessage
    
    Set uniqueDict = Nothing
    On Error GoTo 0
End Sub

' ================================================================================
' Main replication macro
' ================================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim changedCell As Range
    Dim dValue As Variant
    Dim pValue As Variant
    Dim lastRow As Long
    Dim searchRange As Range
    Dim cell As Range
    Dim matchCount As Long
    Dim startTime As Double
    Dim elapsedTime As Double
    
    ' Exit if multiple cells changed or not in column T
    If Target.Column <> 20 Then Exit Sub ' Column T = 20
    If Target.Cells.Count > 1 Then Exit Sub
    
    ' Start timer for performance tracking
    startTime = Timer
    
    ' Disable events to prevent cascading triggers
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Set changedCell = Target
    
    ' Get the value from column D in the same row
    dValue = Me.Cells(changedCell.Row, 4).Value ' Column D = 4
    
    ' Validate column D value
    If IsEmpty(dValue) Then GoTo Cleanup
    If IsError(dValue) Then GoTo Cleanup
    If VarType(dValue) = vbString Then
        If Trim(UCase(dValue)) = "-" Or Trim(UCase(dValue)) = "N/A" Or Len(Trim(dValue)) = 0 Then
            GoTo Cleanup
        End If
    End If
    
    ' Get the T column value to replicate (can be empty to clear)
    pValue = changedCell.Value
    
    ' Find the last used row in column D
    lastRow = Me.Cells(Me.Rows.Count, 4).End(xlUp).Row
    
    ' Only proceed if there are rows to search
    If lastRow < 2 Then GoTo Cleanup
    
    ' Set search range in column D (skip header row if exists)
    Set searchRange = Me.Range(Me.Cells(1, 4), Me.Cells(lastRow, 4))
    
    ' Initialize counter
    matchCount = 0
    
    ' Loop through column D and find matches (case-insensitive)
    For Each cell In searchRange
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            ' Case-insensitive comparison
            If UCase(Trim(CStr(cell.Value))) = UCase(Trim(CStr(dValue))) Then
                ' Copy T value to matching row (overwrite even if empty)
                Me.Cells(cell.Row, 20).Value = pValue
                ' Copy H value to column U for this row
                Me.Cells(cell.Row, 21).Value = Me.Cells(cell.Row, 8).Value ' Column H = 8, Column U = 21
                matchCount = matchCount + 1
            End If
        End If
    Next cell
    
    ' Calculate elapsed time
    elapsedTime = Timer - startTime
    
Cleanup:
    ' Re-enable events and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' Show brief success message (only if we actually replicated something)
    If matchCount > 0 Then
        Application.StatusBar = "✓ Replicated to " & matchCount & " row(s) in " & Format(elapsedTime, "0.00") & "s"
        ' Brief pause so user can see the message (non-blocking)
        Dim pauseTime As Double
        pauseTime = Timer
        Do While Timer < pauseTime + 10
            DoEvents
        Loop
    End If
    
    ' Restore the column D status after showing replication message
    UpdateUniqueCountStatus
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error in SU Comment Replicator: " & Err.Description, vbExclamation, "Macro Error"
End Sub
