' ================================================================================
' SU-Based Comment Replicator
' ================================================================================
' Purpose: Automatically replicate column P values to all rows with matching
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
' - When you enter/change a value in column P and press ENTER
' - Script checks if the corresponding column D value is valid (not -, empty, N/A, error)
' - Finds ALL rows with the same column D value
' - Copies the P value to all matching rows
' - Shows a quick status message when done
' ================================================================================

Option Explicit

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
    
    ' Exit if multiple cells changed or not in column P
    If Target.Column <> 16 Then Exit Sub ' Column P = 16
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
    
    ' Get the P column value to replicate (can be empty to clear)
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
                ' Copy P value to matching row (overwrite even if empty)
                Me.Cells(cell.Row, 16).Value = pValue
                matchCount = matchCount + 1
            End If
        End If
    Next cell
    
    ' Calculate elapsed time
    elapsedTime = Timer - startTime
    
    ' Show non-intrusive status message (will clear on next Excel action)
    Application.StatusBar = "✓ Replicated to " & matchCount & " row(s) in " & Format(elapsedTime, "0.00") & "s"
    
Cleanup:
    ' Re-enable events and screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error in SU Comment Replicator: " & Err.Description, vbExclamation, "Macro Error"
End Sub
