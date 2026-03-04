Sub PasteAndRestoreFilter()
    Dim targetSelection As Range
    Dim visibleTarget As Range
    Dim ws As Worksheet
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim infoMessage As String
    Dim targetAddresses As Collection
    Dim copiedValues As Collection
    Dim targetCount As Long
    Dim copiedCount As Long
    Dim mismatch As Long
    Dim i As Long
    Dim hadAutoFilter As Boolean
    Dim hadFilteredRows As Boolean
    Dim filterOn() As Boolean
    Dim criteria1() As Variant
    Dim criteria2() As Variant
    Dim filterOp() As Variant

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents

    On Error GoTo HandleFailure

    If TypeName(Selection) <> "Range" Then
        infoMessage = "Please select target cells first."
        GoTo HandleMissing
    End If

    If Application.CutCopyMode <> xlCopy Then
        infoMessage = "No copied data found. Copy source cells first (Ctrl+C)."
        GoTo HandleMissing
    End If

    Set targetSelection = Selection
    Set ws = targetSelection.Worksheet

    On Error Resume Next
    Set visibleTarget = targetSelection.SpecialCells(xlCellTypeVisible)
    On Error GoTo HandleFailure
    If visibleTarget Is Nothing Then
        infoMessage = "No visible target cells found in your current selection."
        GoTo HandleMissing
    End If

    Set copiedValues = GetClipboardValues(ws.Parent)
    If copiedValues Is Nothing Then
        infoMessage = "Could not read copied values from clipboard."
        GoTo HandleMissing
    End If

    Set targetAddresses = GetVisibleCellAddresses(visibleTarget)
    targetCount = targetAddresses.Count
    copiedCount = copiedValues.Count

    If copiedCount <> targetCount Then
        mismatch = copiedCount - targetCount
        infoMessage = "Source/target cell count mismatch." & vbCrLf & _
                      "Copied cells: " & copiedCount & vbCrLf & _
                      "Selected visible target cells: " & targetCount & vbCrLf & _
                      "Mismatch: " & mismatch & " (copied - selected)"
        GoTo HandleMissing
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    hadAutoFilter = ws.AutoFilterMode
    hadFilteredRows = ws.FilterMode

    If hadAutoFilter Then
        CaptureFilterState ws, filterOn, criteria1, criteria2, filterOp
        If hadFilteredRows Then
            On Error Resume Next
            ws.ShowAllData
            On Error GoTo HandleFailure
        End If
    End If

    For i = 1 To targetCount
        With ws.Range(CStr(targetAddresses(i)))
            .Value = copiedValues(i)
            .Font.Color = vbRed
        End With
    Next i

    If hadAutoFilter Then
        RestoreFilterState ws, filterOn, criteria1, criteria2, filterOp
    End If

    Application.CutCopyMode = False
    GoTo Cleanup

HandleMissing:
    MsgBox infoMessage, vbExclamation, "Paste To Filtered Rows"
    GoTo Cleanup

HandleFailure:
    MsgBox "Paste failed: " & Err.Description, vbCritical, "Paste To Filtered Rows"

Cleanup:
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
End Sub

Private Function GetVisibleCellAddresses(ByVal visibleTarget As Range) As Collection
    Dim result As Collection
    Dim c As Range

    Set result = New Collection
    For Each c In visibleTarget.Cells
        result.Add c.Address(False, False)
    Next c

    Set GetVisibleCellAddresses = result
End Function

Private Function GetClipboardValues(ByVal wb As Workbook) As Collection
    Dim htmlDoc As Object
    Dim clipText As String
    Dim lineParts As Variant
    Dim rowParts As Variant
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim result As Collection

    On Error GoTo ClipboardTextFailed

    Set result = New Collection

    Set htmlDoc = CreateObject("htmlfile")
    clipText = CStr(htmlDoc.ParentWindow.ClipboardData.GetData("Text"))
    If Len(clipText) = 0 Then GoTo ClipboardFailed

    clipText = Replace(clipText, vbCrLf, vbLf)
    clipText = Replace(clipText, vbCr, vbLf)
    If Right$(clipText, 1) = vbLf Then clipText = Left$(clipText, Len(clipText) - 1)
    If Len(clipText) = 0 Then GoTo ClipboardFailed

    lineParts = Split(clipText, vbLf)
    For rowIdx = LBound(lineParts) To UBound(lineParts)
        rowParts = Split(CStr(lineParts(rowIdx)), vbTab)
        For colIdx = LBound(rowParts) To UBound(rowParts)
            result.Add rowParts(colIdx)
        Next colIdx
    Next rowIdx

    If result.Count = 0 Then GoTo ClipboardFailed

    Set GetClipboardValues = result
    Exit Function

ClipboardTextFailed:
    On Error GoTo ClipboardFailed
    Set result = GetClipboardValuesFromTempPaste(wb)
    If Not result Is Nothing Then
        Set GetClipboardValues = result
        Exit Function
    End If

ClipboardFailed:
    Set GetClipboardValues = Nothing
End Function

Private Function GetClipboardValuesFromTempPaste(ByVal wb As Workbook) As Collection
    Dim tempWs As Worksheet
    Dim usedRng As Range
    Dim r As Long
    Dim c As Long
    Dim result As Collection
    Dim firstCellRow As Long
    Dim firstCellCol As Long
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim originalWs As Worksheet
    Dim originalSel As Range

    On Error GoTo FallbackFailed

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevDisplayAlerts = Application.DisplayAlerts

    On Error Resume Next
    Set originalWs = ActiveSheet
    Set originalSel = Selection
    On Error GoTo FallbackFailed

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set tempWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    tempWs.Range("A1").PasteSpecial Paste:=xlPasteValues

    Set usedRng = tempWs.UsedRange
    If usedRng Is Nothing Then GoTo FallbackFailed

    firstCellRow = usedRng.Row
    firstCellCol = usedRng.Column
    Set result = New Collection

    For r = 1 To usedRng.Rows.Count
        For c = 1 To usedRng.Columns.Count
            result.Add tempWs.Cells(firstCellRow + r - 1, firstCellCol + c - 1).Value
        Next c
    Next r

    If result.Count = 0 Then GoTo FallbackFailed

    tempWs.Delete
    If Not originalWs Is Nothing Then originalWs.Activate
    If Not originalSel Is Nothing Then originalSel.Select

    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating

    Set GetClipboardValuesFromTempPaste = result
    Exit Function

FallbackFailed:
    On Error Resume Next
    If Not tempWs Is Nothing Then tempWs.Delete
    If Not originalWs Is Nothing Then originalWs.Activate
    If Not originalSel Is Nothing Then originalSel.Select
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Set GetClipboardValuesFromTempPaste = Nothing
End Function

Private Sub CaptureFilterState(ByVal ws As Worksheet, _
                               ByRef filterOn() As Boolean, _
                               ByRef criteria1() As Variant, _
                               ByRef criteria2() As Variant, _
                               ByRef filterOp() As Variant)
    Dim af As AutoFilter
    Dim i As Long

    Set af = ws.AutoFilter
    ReDim filterOn(1 To af.Filters.Count)
    ReDim criteria1(1 To af.Filters.Count)
    ReDim criteria2(1 To af.Filters.Count)
    ReDim filterOp(1 To af.Filters.Count)

    For i = 1 To af.Filters.Count
        If af.Filters(i).On Then
            filterOn(i) = True
            criteria1(i) = af.Filters(i).Criteria1
            filterOp(i) = af.Filters(i).Operator
            On Error Resume Next
            criteria2(i) = af.Filters(i).Criteria2
            On Error GoTo 0
        End If
    Next i
End Sub

Private Sub RestoreFilterState(ByVal ws As Worksheet, _
                               ByRef filterOn() As Boolean, _
                               ByRef criteria1() As Variant, _
                               ByRef criteria2() As Variant, _
                               ByRef filterOp() As Variant)
    Dim afRange As Range
    Dim i As Long

    If Not ws.AutoFilterMode Then Exit Sub

    Set afRange = ws.AutoFilter.Range
    For i = LBound(filterOn) To UBound(filterOn)
        If filterOn(i) Then
            If IsEmpty(criteria2(i)) Then
                afRange.AutoFilter Field:=i, Criteria1:=criteria1(i), Operator:=filterOp(i)
            Else
                afRange.AutoFilter Field:=i, Criteria1:=criteria1(i), Operator:=filterOp(i), Criteria2:=criteria2(i)
            End If
        End If
    Next i
End Sub
