Option Explicit

Private Const MAX_TARGET_CELLS As Long = 100000

Public Sub PasteAndRestoreFilter()
    On Error GoTo Failed

    If TypeName(Selection) <> "Range" Then
        Err.Raise vbObjectError + 2100, , "Please select target cells first."
    End If

    PasteVisibleValues Selection
    Exit Sub

Failed:
    MsgBox Err.Description, vbExclamation, "Paste To Filtered Rows"
End Sub

Public Function PasteVisibleValues(ByVal targetSelection As Range) As Long
    Dim visibleTarget As Range
    Dim targetCells() As Range
    Dim cellKeys() As Double
    Dim originalValues() As Variant
    Dim originalFormulas() As Variant
    Dim originalFormats() As Variant
    Dim originalColors() As Variant
    Dim originalColorIndexes() As Variant
    Dim originalThemes() As Variant
    Dim originalTints() As Variant
    Dim sourceValues As Variant
    Dim sourceValue As Variant
    Dim scratchBook As Workbook
    Dim scratchSheet As Worksheet
    Dim pastedRange As Range
    Dim targetCell As Range
    Dim targetSheet As Worksheet
    Dim targetCount As Long
    Dim sourceCount As Double
    Dim sourceColumns As Long
    Dim cellIndex As Long
    Dim writtenCount As Long
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim rollbackFailed As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    On Error GoTo Failed

    If targetSelection Is Nothing Then
        Err.Raise vbObjectError + 2101, , "Please select target cells first."
    End If
    If Application.CutCopyMode <> xlCopy Then
        Err.Raise vbObjectError + 2102, , "Copy source cells in Excel first (Ctrl+C). Cut and external clipboard data are not supported."
    End If
    Set targetSheet = targetSelection.Worksheet
    If targetSheet.ProtectContents Then
        Err.Raise vbObjectError + 2103, , "The target sheet is protected. Unprotect it before pasting."
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set scratchBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set scratchSheet = scratchBook.Worksheets(1)
    scratchSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    Set pastedRange = Application.Selection
    If Not (pastedRange.Parent Is scratchSheet) Then
        Err.Raise vbObjectError + 2108, , "Excel could not read the copied cell range."
    End If
    If pastedRange.Areas.Count <> 1 Then
        Err.Raise vbObjectError + 2109, , "The copied data must paste as one rectangular range."
    End If
    sourceCount = pastedRange.CountLarge
    If sourceCount > MAX_TARGET_CELLS Then
        Err.Raise vbObjectError + 2105, , "Select a smaller source range (maximum 100,000 copied cells)."
    End If
    sourceColumns = pastedRange.Columns.Count
    sourceValues = pastedRange.Value2

    If targetSelection.CountLarge = 1 Then
        Set visibleTarget = GetAnchorTargets(targetSelection, pastedRange.Rows.Count, sourceColumns)
    Else
        On Error Resume Next
        Set visibleTarget = targetSelection.SpecialCells(xlCellTypeVisible)
        Set visibleTarget = Application.Intersect(targetSelection, visibleTarget)
        On Error GoTo Failed
    End If
    scratchBook.Close SaveChanges:=False
    Set scratchBook = Nothing

    If visibleTarget Is Nothing Then
        Err.Raise vbObjectError + 2104, , "No visible target cells found."
    End If
    If visibleTarget.CountLarge > MAX_TARGET_CELLS Then
        Err.Raise vbObjectError + 2105, , "Select a smaller target range (maximum 100,000 visible cells)."
    End If

    targetCount = CLng(visibleTarget.CountLarge)
    ReDim targetCells(1 To targetCount)
    ReDim cellKeys(1 To targetCount)
    cellIndex = 0
    For Each targetCell In visibleTarget.Cells
        If targetCell.MergeCells Or targetCell.HasArray Or HasSpillCells(targetCell) Then
            Err.Raise vbObjectError + 2106, , "Merged cells, array formulas and spill ranges are not supported: " & targetCell.Address(False, False)
        End If
        cellIndex = cellIndex + 1
        Set targetCells(cellIndex) = targetCell
        cellKeys(cellIndex) = CDbl(targetCell.Row) * targetSheet.Columns.Count + targetCell.Column
    Next targetCell
    SortTargetCells targetCells, cellKeys, 1, targetCount
    For cellIndex = 2 To targetCount
        If cellKeys(cellIndex) = cellKeys(cellIndex - 1) Then
            Err.Raise vbObjectError + 2107, , "The target selection contains overlapping areas."
        End If
    Next cellIndex

    If sourceCount <> targetCount Then
        Err.Raise vbObjectError + 2110, , "Source/target cell count mismatch." & vbCrLf & _
                  "Copied cells: " & sourceCount & vbCrLf & _
                  "Selected visible target cells: " & targetCount & vbCrLf & _
                  "Mismatch: " & (sourceCount - targetCount) & " (copied - selected)"
    End If
    ReDim originalValues(1 To targetCount)
    ReDim originalFormulas(1 To targetCount)
    ReDim originalFormats(1 To targetCount)
    ReDim originalColors(1 To targetCount)
    ReDim originalColorIndexes(1 To targetCount)
    ReDim originalThemes(1 To targetCount)
    ReDim originalTints(1 To targetCount)
    For cellIndex = 1 To targetCount
        With targetCells(cellIndex)
            originalValues(cellIndex) = .Value2
            If .HasFormula Then originalFormulas(cellIndex) = .Formula
            originalFormats(cellIndex) = .NumberFormat
            originalColors(cellIndex) = .Font.Color
            originalColorIndexes(cellIndex) = .Font.ColorIndex
            originalTints(cellIndex) = .Font.TintAndShade
            On Error Resume Next
            originalThemes(cellIndex) = .Font.ThemeColor
            On Error GoTo Failed
        End With
    Next cellIndex

    For cellIndex = 1 To targetCount
        If targetCount = 1 Then
            sourceValue = sourceValues
        Else
            sourceValue = sourceValues(((cellIndex - 1) \ sourceColumns) + 1, ((cellIndex - 1) Mod sourceColumns) + 1)
        End If
        writtenCount = cellIndex
        SetCellValue targetCells(cellIndex), sourceValue, originalFormats(cellIndex)
        targetCells(cellIndex).Font.Color = vbRed
    Next cellIndex
    PasteVisibleValues = targetCount
    Application.CutCopyMode = False
    GoTo Cleanup

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    For cellIndex = writtenCount To 1 Step -1
        Err.Clear
        SetCellValue targetCells(cellIndex), originalValues(cellIndex), originalFormats(cellIndex)
        With targetCells(cellIndex)
            If Not IsEmpty(originalFormulas(cellIndex)) Then .Formula = originalFormulas(cellIndex)
            If Not IsEmpty(originalThemes(cellIndex)) Then
                .Font.ThemeColor = originalThemes(cellIndex)
            ElseIf originalColorIndexes(cellIndex) = xlColorIndexAutomatic Then
                .Font.ColorIndex = xlColorIndexAutomatic
            Else
                .Font.Color = originalColors(cellIndex)
            End If
            .Font.TintAndShade = originalTints(cellIndex)
        End With
        If Err.Number <> 0 Then rollbackFailed = True
    Next cellIndex
    If rollbackFailed Then failureDescription = failureDescription & vbCrLf & "Some original cells could not be restored. Check the target range."

Cleanup:
    On Error Resume Next
    If Not scratchBook Is Nothing Then scratchBook.Close SaveChanges:=False
    If Not targetSheet Is Nothing Then
        targetSheet.Parent.Activate
        targetSheet.Activate
        targetSelection.Select
    End If
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "PasteVisibleValues", failureDescription
End Function

Private Function GetAnchorTargets(ByVal anchor As Range, ByVal sourceRows As Long, ByVal sourceColumns As Long) As Range
    Dim targetSheet As Worksheet
    Dim table As ListObject
    Dim boundary As Range
    Dim rowScan As Range
    Dim visibleRows As Range
    Dim visibleArea As Range
    Dim rowLimit As Long
    Dim columnLimit As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim remainingRows As Long
    Dim remainingColumns As Long

    Set targetSheet = anchor.Worksheet
    If anchor.EntireRow.Hidden Or anchor.EntireColumn.Hidden Then
        Err.Raise vbObjectError + 2104, , "No visible target cells found at the starting cell."
    End If
    rowLimit = targetSheet.Rows.Count
    columnLimit = targetSheet.Columns.Count
    For Each table In targetSheet.ListObjects
        If Not Application.Intersect(anchor, table.Range) Is Nothing Then
            Set boundary = table.DataBodyRange
            If Not boundary Is Nothing Then
                If Application.Intersect(anchor, boundary) Is Nothing Then Set boundary = Nothing
            End If
            If boundary Is Nothing Then Err.Raise vbObjectError + 2111, , "Select a starting cell in the table's data rows, not its header or totals."
            Exit For
        End If
    Next table
    If boundary Is Nothing And targetSheet.AutoFilterMode Then
        If Not targetSheet.AutoFilter Is Nothing Then
            If Not Application.Intersect(anchor, targetSheet.AutoFilter.Range) Is Nothing Then
                Set boundary = targetSheet.AutoFilter.Range
                If anchor.Row = boundary.Row Then Err.Raise vbObjectError + 2111, , "Select a starting cell below the filter header."
            End If
        End If
    End If
    If Not boundary Is Nothing Then
        rowLimit = boundary.Row + boundary.Rows.Count - 1
        columnLimit = boundary.Column + boundary.Columns.Count - 1
    End If

    remainingColumns = sourceColumns
    For lastColumn = anchor.Column To columnLimit
        If Not targetSheet.Columns(lastColumn).Hidden Then remainingColumns = remainingColumns - 1
        If remainingColumns = 0 Then Exit For
    Next lastColumn
    If remainingColumns <> 0 Then Err.Raise vbObjectError + 2112, , "Not enough visible target columns from the starting cell."

    Set rowScan = targetSheet.Range(anchor, targetSheet.Cells(rowLimit, anchor.Column))
    Set visibleRows = Application.Intersect(rowScan, rowScan.SpecialCells(xlCellTypeVisible))
    remainingRows = sourceRows
    For Each visibleArea In visibleRows.Areas
        If remainingRows <= visibleArea.Rows.Count Then
            lastRow = visibleArea.Row + remainingRows - 1
            remainingRows = 0
            Exit For
        End If
        remainingRows = remainingRows - visibleArea.Rows.Count
    Next visibleArea
    If remainingRows <> 0 Then Err.Raise vbObjectError + 2113, , "Not enough visible target rows from the starting cell."

    Set rowScan = targetSheet.Range(anchor, targetSheet.Cells(lastRow, lastColumn))
    Set GetAnchorTargets = Application.Intersect(rowScan, rowScan.SpecialCells(xlCellTypeVisible))
End Function

Private Sub SetCellValue(ByVal targetCell As Range, ByVal cellValue As Variant, ByVal originalFormat As Variant)
    If VarType(cellValue) = vbString Then targetCell.NumberFormat = "@"
    targetCell.Value2 = cellValue
    targetCell.NumberFormat = originalFormat
End Sub

Private Function HasSpillCells(ByVal targetCell As Range) As Boolean
    On Error Resume Next
    HasSpillCells = CallByName(targetCell, "HasSpill", VbGet)
    On Error GoTo 0
End Function

Private Sub SortTargetCells(ByRef targetCells() As Range, ByRef cellKeys() As Double, ByVal firstIndex As Long, ByVal lastIndex As Long)
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim pivotKey As Double
    Dim swapKey As Double
    Dim swapCell As Range

    leftIndex = firstIndex
    rightIndex = lastIndex
    pivotKey = cellKeys(firstIndex + (lastIndex - firstIndex) \ 2)
    Do While leftIndex <= rightIndex
        Do While cellKeys(leftIndex) < pivotKey
            leftIndex = leftIndex + 1
        Loop
        Do While cellKeys(rightIndex) > pivotKey
            rightIndex = rightIndex - 1
        Loop
        If leftIndex <= rightIndex Then
            swapKey = cellKeys(leftIndex)
            cellKeys(leftIndex) = cellKeys(rightIndex)
            cellKeys(rightIndex) = swapKey
            Set swapCell = targetCells(leftIndex)
            Set targetCells(leftIndex) = targetCells(rightIndex)
            Set targetCells(rightIndex) = swapCell
            leftIndex = leftIndex + 1
            rightIndex = rightIndex - 1
        End If
    Loop
    If firstIndex < rightIndex Then SortTargetCells targetCells, cellKeys, firstIndex, rightIndex
    If leftIndex < lastIndex Then SortTargetCells targetCells, cellKeys, leftIndex, lastIndex
End Sub
