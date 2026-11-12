Attribute VB_Name = "ActiveFilters"
Option Explicit

Private Const POPUP_NAME As String = "ActiveFiltersPopup"
Private Const SOURCE_RANGE As String = "sheet range"
Private Const SOURCE_TABLE As String = "table"
Private Const STALE_MESSAGE As String = "The filter list was out of date - please open the menu again."
Private Const PROTECTED_MESSAGE As String = "Sheet is protected - filter could not be cleared."

Private mTargets As Collection
Private mMenuSheet As Worksheet
Private mPopupOpen As Boolean
Private mTestMode As Boolean

Public Sub ShowActiveFiltersMenu()
    Dim ws As Worksheet
    Dim popup As CommandBar
    Dim button As CommandBarButton
    Dim target As CFilterTarget
    Dim targetIndex As Long
    Dim caption As String
    Dim errorText As String

    On Error GoTo Failed
    If mPopupOpen Then Exit Sub
    If Not TypeOf Application.ActiveSheet Is Worksheet Then
        Notify "Open a worksheet to view active filters."
        Exit Sub
    End If
    If Application.ActiveWindow.SelectedSheets.Count > 1 Then
        Notify "Ungroup the sheets before clearing filters."
        Exit Sub
    End If

    Set ws = Application.ActiveSheet
    Set mTargets = GetActiveFilterTargets(ws)
    If mTargets Is Nothing Then Exit Sub
    Set mMenuSheet = ws
    DeletePopup
    Set popup = Application.CommandBars.Add(Name:=POPUP_NAME, _
        Position:=msoBarPopup, Temporary:=True)
    mPopupOpen = True

    For targetIndex = 1 To mTargets.Count
        Set target = mTargets(targetIndex)
        caption = "Col " & target.ColumnLetter & " - " & CleanHeader(target.HeaderText)
        If target.SourceKind = SOURCE_TABLE Then
            caption = caption & " [" & target.TableName & "]"
        End If
        Set button = popup.Controls.Add(Type:=msoControlButton)
        button.Caption = Replace(caption, "&", "&&")
        button.OnAction = "'" & ThisWorkbook.Name & "'!ActiveFilters_OnClick"
        button.Parameter = CStr(targetIndex)
    Next targetIndex

    If mTargets.Count >= 2 Then
        Set button = popup.Controls.Add(Type:=msoControlButton)
        button.Caption = "Clear all filters on this sheet"
        button.BeginGroup = True
        button.OnAction = "'" & ThisWorkbook.Name & "'!ActiveFilters_OnClick"
        button.Parameter = "ALL"
    ElseIf mTargets.Count = 0 Then
        Set button = popup.Controls.Add(Type:=msoControlButton)
        If ws.FilterMode Then
            button.Caption = "Advanced filter active - clear all"
            button.OnAction = "'" & ThisWorkbook.Name & "'!ActiveFilters_OnClick"
            button.Parameter = "ALL"
        Else
            button.Caption = "No active filters on this sheet"
            button.Enabled = False
        End If
    End If

    popup.ShowPopup
CleanUp:
    DeletePopup
    Set mTargets = Nothing
    Set mMenuSheet = Nothing
    mPopupOpen = False
    If Len(errorText) > 0 Then Notify errorText
    Exit Sub
Failed:
    errorText = "Could not show active filters: " & Err.Description
    Resume CleanUp
End Sub

Public Sub ActiveFilters_OnClick()
    Dim parameter As String
    Dim targetIndex As Long
    Dim target As CFilterTarget
    Dim cleared As Boolean

    On Error GoTo Failed
    If mTargets Is Nothing Then GoTo Stale
    If mMenuSheet Is Nothing Then GoTo Stale
    If Not IsActiveWorksheet(mMenuSheet) Then GoTo Stale
    parameter = Application.CommandBars.ActionControl.Parameter
    If parameter = "ALL" Then
        ClearAllFilters mMenuSheet
    Else
        targetIndex = CLng(parameter)
        If targetIndex < 1 Or targetIndex > mTargets.Count Then GoTo Stale
        Set target = mTargets(targetIndex)
        cleared = ClearFilterTarget(target)
    End If
    Exit Sub
Stale:
    Notify STALE_MESSAGE
    Exit Sub
Failed:
    Notify "Could not process the filter menu: " & Err.Description
End Sub

Public Function GetActiveFilterTargets(ws As Worksheet) As Collection
    Dim targets As Collection
    Dim autoFilter As AutoFilter
    Dim table As ListObject

    On Error GoTo Failed
    Set targets = New Collection
    If ws Is Nothing Then Err.Raise vbObjectError + 2000, , "No worksheet supplied."
    If ws.AutoFilterMode Then
        Set autoFilter = ws.AutoFilter
        If Not autoFilter Is Nothing Then
            If Not IsTableRange(ws, autoFilter.Range) Then
                AppendTargets targets, ws, autoFilter, SOURCE_RANGE, vbNullString
            End If
        End If
    End If
    For Each table In ws.ListObjects
        If table.ShowHeaders And table.ShowAutoFilter Then
            Set autoFilter = table.AutoFilter
            If Not autoFilter Is Nothing Then
                AppendTargets targets, ws, autoFilter, SOURCE_TABLE, table.Name
            End If
        End If
    Next table
    Set GetActiveFilterTargets = targets
    Exit Function
Failed:
    Notify "Could not read active filters: " & Err.Description
End Function

Public Function ClearFilterTarget(t As CFilterTarget) As Boolean
    Dim ws As Worksheet
    Dim table As ListObject
    Dim autoFilter As AutoFilter
    Dim filterRange As Range
    Dim window As Window
    Dim scrollRow As Long
    Dim scrollColumn As Long
    Dim scrollSaved As Boolean
    Dim message As String

    On Error GoTo Stale
    If t Is Nothing Then GoTo Stale
    Set ws = Application.Workbooks(t.WorkbookName).Worksheets(t.SheetName)
    If Not IsActiveWorksheet(ws) Then GoTo Stale
    Select Case t.SourceKind
        Case SOURCE_RANGE
            If Len(t.TableName) <> 0 Then GoTo Stale
            If Not ws.AutoFilterMode Then GoTo Stale
            Set autoFilter = ws.AutoFilter
            If autoFilter Is Nothing Then GoTo Stale
            If IsTableRange(ws, autoFilter.Range) Then GoTo Stale
            Set filterRange = autoFilter.Range
        Case SOURCE_TABLE
            Set table = ws.ListObjects(t.TableName)
            If Not table.ShowHeaders Or Not table.ShowAutoFilter Then GoTo Stale
            Set autoFilter = table.AutoFilter
            If autoFilter Is Nothing Then GoTo Stale
            Set filterRange = autoFilter.Range
        Case Else
            GoTo Stale
    End Select
    If filterRange.Address <> t.FilterRangeAddress Then GoTo Stale
    If t.FieldIndex < 1 Or t.FieldIndex > autoFilter.Filters.Count Then GoTo Stale
    If t.FieldIndex > filterRange.Columns.Count Then GoTo Stale
    If filterRange.Columns(t.FieldIndex).Column <> t.SheetColumn Then GoTo Stale
    If ColumnLetters(t.SheetColumn) <> t.ColumnLetter Then GoTo Stale
    If CStr(filterRange.Cells(1, t.FieldIndex).Text) <> t.HeaderText Then GoTo Stale
    If Not autoFilter.Filters(t.FieldIndex).On Then GoTo Stale

    On Error GoTo Failed
    Set window = Application.ActiveWindow
    scrollRow = window.ScrollRow
    scrollColumn = window.ScrollColumn
    scrollSaved = True
    If t.SourceKind = SOURCE_TABLE Then
        table.Range.AutoFilter Field:=t.FieldIndex
    Else
        ws.AutoFilter.Range.AutoFilter Field:=t.FieldIndex
    End If
    ClearFilterTarget = True
CleanUp:
    If scrollSaved Then RestoreScroll window, scrollRow, scrollColumn
    If Len(message) > 0 Then Notify message
    Exit Function
Stale:
    message = STALE_MESSAGE
    GoTo CleanUp
Failed:
    If ws.ProtectContents Then
        message = PROTECTED_MESSAGE
    Else
        message = "Filter could not be cleared: " & Err.Description
    End If
    Resume CleanUp
End Function

Public Sub ClearAllFilters(ws As Worksheet)
    Dim table As ListObject
    Dim window As Window
    Dim scrollRow As Long
    Dim scrollColumn As Long
    Dim scrollSaved As Boolean
    Dim message As String

    On Error GoTo Stale
    If ws Is Nothing Then GoTo Stale
    If Not IsActiveWorksheet(ws) Then GoTo Stale
    On Error GoTo Failed
    Set window = Application.ActiveWindow
    scrollRow = window.ScrollRow
    scrollColumn = window.ScrollColumn
    scrollSaved = True
    If ws.FilterMode Then ws.ShowAllData
    For Each table In ws.ListObjects
        If HasActiveFilters(table.AutoFilter) Then table.AutoFilter.ShowAllData
    Next table
CleanUp:
    If scrollSaved Then RestoreScroll window, scrollRow, scrollColumn
    If Len(message) > 0 Then Notify message
    Exit Sub
Stale:
    message = STALE_MESSAGE
    GoTo CleanUp
Failed:
    If ws.ProtectContents Then
        message = PROTECTED_MESSAGE
    Else
        message = "Filters could not all be cleared: " & Err.Description
    End If
    Resume CleanUp
End Sub

Private Sub AppendTargets(targets As Collection, ws As Worksheet, _
    autoFilter As AutoFilter, sourceKind As String, tableName As String)

    Dim fieldIndex As Long
    Dim target As CFilterTarget
    Dim filterRange As Range

    Set filterRange = autoFilter.Range
    For fieldIndex = 1 To autoFilter.Filters.Count
        If autoFilter.Filters(fieldIndex).On Then
            Set target = New CFilterTarget
            target.WorkbookName = ws.Parent.Name
            target.SheetName = ws.Name
            target.SourceKind = sourceKind
            target.TableName = tableName
            target.FieldIndex = fieldIndex
            target.SheetColumn = filterRange.Columns(fieldIndex).Column
            target.ColumnLetter = ColumnLetters(target.SheetColumn)
            target.HeaderText = CStr(filterRange.Cells(1, fieldIndex).Text)
            target.FilterRangeAddress = filterRange.Address
            InsertSorted targets, target
        End If
    Next fieldIndex
End Sub

Private Sub InsertSorted(targets As Collection, target As CFilterTarget)
    Dim targetIndex As Long
    Dim existing As CFilterTarget

    For targetIndex = 1 To targets.Count
        Set existing = targets(targetIndex)
        If target.SheetColumn < existing.SheetColumn Then
            targets.Add Item:=target, Before:=targetIndex
            Exit Sub
        ElseIf target.SheetColumn = existing.SheetColumn Then
            If StrComp(target.TableName, existing.TableName, vbTextCompare) < 0 Then
                targets.Add Item:=target, Before:=targetIndex
                Exit Sub
            End If
        End If
    Next targetIndex
    targets.Add target
End Sub

Private Function IsTableRange(ws As Worksheet, filterRange As Range) As Boolean
    Dim table As ListObject
    Dim tableFilter As AutoFilter

    For Each table In ws.ListObjects
        If table.Range.Address = filterRange.Address Then
            IsTableRange = True
            Exit Function
        End If
        If table.ShowHeaders Then
            Set tableFilter = table.AutoFilter
            If Not tableFilter Is Nothing Then
                If tableFilter.Range.Address = filterRange.Address Then
                    IsTableRange = True
                    Exit Function
                End If
            End If
        End If
    Next table
End Function

Private Function IsActiveWorksheet(ws As Worksheet) As Boolean
    Dim active As Worksheet

    If Not TypeOf Application.ActiveSheet Is Worksheet Then Exit Function
    If Application.ActiveWindow Is Nothing Then Exit Function
    If Application.ActiveWindow.SelectedSheets.Count > 1 Then Exit Function
    Set active = Application.ActiveSheet
    IsActiveWorksheet = (active Is ws)
End Function

Private Function HasActiveFilters(autoFilter As AutoFilter) As Boolean
    Dim fieldIndex As Long

    If autoFilter Is Nothing Then Exit Function
    For fieldIndex = 1 To autoFilter.Filters.Count
        If autoFilter.Filters(fieldIndex).On Then
            HasActiveFilters = True
            Exit Function
        End If
    Next fieldIndex
End Function

Private Function ColumnLetters(columnNumber As Long) As String
    Dim remainder As Long

    Do While columnNumber > 0
        remainder = (columnNumber - 1) Mod 26
        ColumnLetters = Chr$(65 + remainder) & ColumnLetters
        columnNumber = (columnNumber - 1) \ 26
    Loop
End Function

Private Function CleanHeader(headerText As String) As String
    headerText = Replace(headerText, vbCrLf, " ")
    headerText = Replace(headerText, vbCr, " ")
    headerText = Replace(headerText, vbLf, " ")
    CleanHeader = Left$(headerText, 40)
End Function

Private Sub DeletePopup()
    On Error Resume Next
    Application.CommandBars(POPUP_NAME).Delete
    On Error GoTo 0
End Sub

Private Sub RestoreScroll(window As Window, scrollRow As Long, scrollColumn As Long)
    On Error Resume Next
    window.ScrollRow = scrollRow
    window.ScrollColumn = scrollColumn
    On Error GoTo 0
End Sub

Private Sub Notify(message As String)
    On Error Resume Next
    If Not mTestMode Then MsgBox message, vbExclamation, "Active Filters"
    On Error GoTo 0
End Sub

Public Sub ActiveFilters_SelfTest()
    Dim testWorkbook As Workbook
    Dim ws As Worksheet
    Dim shiftedSheet As Worksheet
    Dim tableA As ListObject
    Dim tableB As ListObject
    Dim targets As Collection
    Dim target As CFilterTarget
    Dim staleTarget As CFilterTarget
    Dim data(1 To 200, 1 To 25) As Variant
    Dim tableData(1 To 21, 1 To 3) As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim passed As Long
    Dim failed As Long
    Dim cleared As Boolean
    Dim expectedCount As Long
    Dim originalTestMode As Boolean
    Dim scrollRow As Long
    Dim scrollColumn As Long
    Dim failureText As String

    On Error GoTo FailedTest
    originalTestMode = mTestMode
    mTestMode = True
    Debug.Print "ActiveFilters self-test"
    Set testWorkbook = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = testWorkbook.Worksheets(1)
    For columnIndex = 1 To 25
        data(1, columnIndex) = "Header " & ColumnLetters(columnIndex)
    Next columnIndex
    data(1, 3) = "Category"
    data(1, 20) = "State"
    data(1, 24) = "Shade"
    For rowIndex = 2 To 200
        For columnIndex = 1 To 25
            Select Case columnIndex Mod 3
                Case 0: data(rowIndex, columnIndex) = "Text " & CStr(rowIndex Mod 5)
                Case 1: data(rowIndex, columnIndex) = rowIndex * columnIndex
                Case 2: data(rowIndex, columnIndex) = DateSerial(2026, 1, 1) + rowIndex
            End Select
        Next columnIndex
        If rowIndex Mod 2 = 0 Then
            data(rowIndex, 3) = "Keep"
        Else
            data(rowIndex, 3) = "Drop"
        End If
        Select Case rowIndex Mod 3
            Case 0: data(rowIndex, 20) = "Open"
            Case 1: data(rowIndex, 20) = "Hold"
            Case 2: data(rowIndex, 20) = "Closed"
        End Select
        data(rowIndex, 24) = "Paint " & CStr(rowIndex Mod 2)
    Next rowIndex
    ws.Range("A1:Y200").Value = data
    ws.Range("B2:B200").NumberFormat = "yyyy-mm-dd"
    ws.Range("X2:X200").Interior.Color = RGB(255, 255, 0)
    For rowIndex = 2 To 200 Step 2
        ws.Cells(rowIndex, 24).Interior.Color = RGB(0, 176, 80)
    Next rowIndex
    ws.Range("A1:Y200").AutoFilter
    ws.AutoFilter.Range.AutoFilter Field:=3, Criteria1:="Keep"
    ws.AutoFilter.Range.AutoFilter Field:=20, _
        Criteria1:=Array("Open", "Hold"), Operator:=xlFilterValues
    ws.AutoFilter.Range.AutoFilter Field:=24, _
        Criteria1:=RGB(0, 176, 80), Operator:=xlFilterCellColor

    tableData(1, 1) = "ID"
    tableData(1, 2) = "Status"
    tableData(1, 3) = "Quantity"
    For rowIndex = 2 To 21
        tableData(rowIndex, 1) = rowIndex - 1
        If rowIndex Mod 2 = 0 Then
            tableData(rowIndex, 2) = "Keep"
        Else
            tableData(rowIndex, 2) = "Drop"
        End If
        tableData(rowIndex, 3) = rowIndex * 10
    Next rowIndex
    ws.Range("AA1:AC21").Value = tableData
    ws.Range("AE1:AG21").Value = tableData
    Set tableA = ws.ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=ws.Range("AA1:AC21"), XlListObjectHasHeaders:=xlYes)
    tableA.Name = "TblA"
    Set tableB = ws.ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=ws.Range("AE1:AG21"), XlListObjectHasHeaders:=xlYes)
    tableB.Name = "TblB"
    tableA.Range.AutoFilter Field:=2, Criteria1:="Keep"
    tableB.Range.AutoFilter Field:=2, Criteria1:="Keep"

    Set targets = GetActiveFilterTargets(ws)
    If targets Is Nothing Then Err.Raise vbObjectError + 2001, , "Enumeration failed."
    CheckTest targets.Count = 5, "Exactly five active targets", passed, failed
    CheckTest TargetMatches(targets, 1, "C", "Category", "", 3, 3, ws), _
        "C: sheet value filter", passed, failed
    CheckTest TargetMatches(targets, 2, "T", "State", "", 20, 20, ws), _
        "T: off-screen array filter", passed, failed
    CheckTest TargetMatches(targets, 3, "X", "Shade", "", 24, 24, ws), _
        "X: off-screen cell-colour filter", passed, failed
    CheckTest TargetMatches(targets, 4, "AB", "Status", "TblA", 2, 28, ws), _
        "AB: TblA duplicate Status header", passed, failed
    CheckTest TargetMatches(targets, 5, "AF", "Status", "TblB", 2, 32, ws), _
        "AF: TblB duplicate Status header", passed, failed

    Set target = FindTarget(targets, "", 20)
    scrollRow = Application.ActiveWindow.ScrollRow
    scrollColumn = Application.ActiveWindow.ScrollColumn
    cleared = ClearFilterTarget(target)
    CheckTest cleared And Not ws.AutoFilter.Filters(20).On, _
        "Clear T only", passed, failed
    Set targets = GetActiveFilterTargets(ws)
    CheckTest targets.Count = 4 And ws.AutoFilter.Filters(3).On _
        And ws.AutoFilter.Filters(24).On And tableA.AutoFilter.Filters(2).On _
        And tableB.AutoFilter.Filters(2).On, "Other four filters remain", passed, failed
    CheckTest Application.ActiveWindow.ScrollRow = scrollRow _
        And Application.ActiveWindow.ScrollColumn = scrollColumn, _
        "Scroll position preserved", passed, failed

    Set target = FindTarget(targets, "TblB", 2)
    cleared = ClearFilterTarget(target)
    CheckTest cleared And Not tableB.AutoFilter.Filters(2).On _
        And tableA.AutoFilter.Filters(2).On, "Clear TblB without clearing TblA", passed, failed
    Set targets = GetActiveFilterTargets(ws)
    expectedCount = targets.Count
    Set staleTarget = FindTarget(targets, "", 3)
    ws.Cells(1, 3).Value = "Changed category"
    cleared = ClearFilterTarget(staleTarget)
    Set targets = GetActiveFilterTargets(ws)
    CheckTest Not cleared And targets.Count = expectedCount _
        And ws.AutoFilter.Filters(3).On And ws.AutoFilter.Filters(24).On _
        And tableA.AutoFilter.Filters(2).On And Not tableB.AutoFilter.Filters(2).On _
        And Not ws.AutoFilter.Filters(20).On, "Stale header rejected without clearing", passed, failed

    ClearAllFilters ws
    Set targets = GetActiveFilterTargets(ws)
    CheckTest targets.Count = 0, "Clear all leaves no active filters", passed, failed
    CheckTest ws.AutoFilterMode And tableA.ShowAutoFilter And tableB.ShowAutoFilter, _
        "Filter arrows retained", passed, failed

    Set shiftedSheet = testWorkbook.Worksheets.Add(After:=ws)
    shiftedSheet.Range("D1").Value = "ID"
    shiftedSheet.Range("E1").Value = "Offset status"
    shiftedSheet.Range("F1").Value = "Quantity"
    For rowIndex = 2 To 8
        shiftedSheet.Cells(rowIndex, 4).Value = rowIndex
        shiftedSheet.Cells(rowIndex, 5).Value = rowIndex Mod 2
        shiftedSheet.Cells(rowIndex, 6).Value = rowIndex * 10
    Next rowIndex
    shiftedSheet.Range("D1:F8").AutoFilter Field:=2, Criteria1:="0"
    Set targets = GetActiveFilterTargets(shiftedSheet)
    CheckTest TargetMatches(targets, 1, "E", "Offset status", "", 2, 5, shiftedSheet), _
        "Offset range: field 2 is sheet column E", passed, failed
    Set target = FindTarget(targets, "", 2)
    cleared = ClearFilterTarget(target)
    CheckTest cleared And Not shiftedSheet.AutoFilter.Filters(2).On _
        And shiftedSheet.AutoFilterMode, "Clear offset range field", passed, failed

Finish:
    On Error Resume Next
    If Not testWorkbook Is Nothing Then
        Err.Clear
        testWorkbook.Close SaveChanges:=False
        If Err.Number <> 0 Then
            failed = failed + 1
            Debug.Print "FAIL: Temporary workbook cleanup - " & Err.Description
        End If
    End If
    mTestMode = originalTestMode
    Debug.Print "Summary: " & CStr(passed) & " PASS, " & CStr(failed) & " FAIL"
    On Error GoTo 0
    Exit Sub
FailedTest:
    failureText = Err.Description
    failed = failed + 1
    Debug.Print "FAIL: Self-test runtime error - " & failureText
    Resume Finish
End Sub

Private Sub CheckTest(condition As Boolean, label As String, passed As Long, failed As Long)
    If condition Then
        passed = passed + 1
        Debug.Print "PASS: " & label
    Else
        failed = failed + 1
        Debug.Print "FAIL: " & label
    End If
End Sub

Private Function FindTarget(targets As Collection, tableName As String, _
    fieldIndex As Long) As CFilterTarget

    Dim target As CFilterTarget

    For Each target In targets
        If target.TableName = tableName And target.FieldIndex = fieldIndex Then
            Set FindTarget = target
            Exit Function
        End If
    Next target
    Err.Raise vbObjectError + 2002, , "Expected target not found."
End Function

Private Function TargetMatches(targets As Collection, targetIndex As Long, _
    letter As String, header As String, tableName As String, fieldIndex As Long, _
    sheetColumn As Long, ws As Worksheet) As Boolean

    Dim target As CFilterTarget
    Dim expectedAddress As String
    Dim expectedKind As String

    If targetIndex > targets.Count Then Exit Function
    Set target = targets(targetIndex)
    If Len(tableName) = 0 Then
        expectedKind = SOURCE_RANGE
        expectedAddress = ws.AutoFilter.Range.Address
    Else
        expectedKind = SOURCE_TABLE
        expectedAddress = ws.ListObjects(tableName).AutoFilter.Range.Address
    End If
    TargetMatches = target.ColumnLetter = letter And target.HeaderText = header _
        And target.TableName = tableName And target.FieldIndex = fieldIndex _
        And target.SheetColumn = sheetColumn And target.WorkbookName = ws.Parent.Name _
        And target.SheetName = ws.Name And target.SourceKind = expectedKind _
        And target.FilterRangeAddress = expectedAddress
End Function