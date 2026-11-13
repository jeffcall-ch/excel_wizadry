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
Private mCurrentStep As String
Private mReport As String
Private mRemembered As Collection
Private mTestingOptional As Boolean
Private mOptionalSelfTestCalled As Boolean
Private mForceSnapshotFailure As Boolean
Private mTracePath As String

Public Sub ShowActiveFiltersMenu()
    Dim ws As Worksheet
    Dim popup As CommandBar
    Dim errorText As String

    On Error GoTo Failed
    If mPopupOpen Then Exit Sub
    If Not TypeOf Application.ActiveSheet Is Worksheet Then
        AF_Notify "Open a worksheet to view active filters."
        Exit Sub
    End If
    If Application.ActiveWindow.SelectedSheets.Count > 1 Then
        AF_Notify "Ungroup the sheets before clearing filters."
        Exit Sub
    End If

    Set ws = Application.ActiveSheet
    Set popup = AF_BuildPopup(ws)
    If popup Is Nothing Then Exit Sub
    mPopupOpen = True
    popup.ShowPopup
CleanUp:
    AF_DeletePopup
    If Len(errorText) > 0 Then AF_Notify errorText
    Exit Sub
Failed:
    errorText = "Could not show active filters: " & Err.Description
    Resume CleanUp
End Sub

Private Function AF_BuildPopup(ws As Worksheet) As CommandBar
    Dim popup As CommandBar
    Dim button As CommandBarButton
    Dim target As AF_FilterTarget
    Dim targetIndex As Long
    Dim caption As String
    Dim activeCount As Long
    Dim reapplyCount As Long
    Dim rememberedCount As Long
    Dim bottomAdded As Boolean

    On Error GoTo Failed
    AF_DeletePopup
    If Not AF_IsActiveWorksheet(ws) Then Err.Raise vbObjectError + 2003, , STALE_MESSAGE
    Set mTargets = AF_GetMenuTargets(ws)
    If mTargets Is Nothing Then Err.Raise vbObjectError + 2004, , "Could not read filter targets."
    Set mMenuSheet = ws
    Set popup = Application.CommandBars.Add(Name:=POPUP_NAME, _
        Position:=msoBarPopup, Temporary:=True)

    For targetIndex = 1 To mTargets.Count
        Set target = mTargets(targetIndex)
        caption = "Col " & target.ColumnLetter & " - " & AF_CleanHeader(target.HeaderText)
        If target.SourceKind = SOURCE_TABLE Then
            caption = caption & " [" & target.TableName & "]"
        End If
        If target.Remembered Then
            rememberedCount = rememberedCount + 1
            If target.CanReapply Then
                caption = caption & " (off)"
                reapplyCount = reapplyCount + 1
            Else
                caption = caption & " (cannot re-apply)"
            End If
        Else
            activeCount = activeCount + 1
        End If
        Set button = popup.Controls.Add(Type:=msoControlButton)
        button.Caption = Replace(caption, "&", "&&")
        button.Style = msoButtonCaption
        button.FaceId = 0
        If Not target.Remembered Then button.State = msoButtonDown
        button.Enabled = Not target.Remembered Or target.CanReapply
        button.OnAction = AF_Callback()
        button.Parameter = CStr(targetIndex)
    Next targetIndex

    If activeCount >= 2 Then
        AF_AddBottom popup, "Clear all filters on this sheet", "ALL", bottomAdded
    ElseIf activeCount = 0 And ws.FilterMode Then
        AF_AddBottom popup, "Advanced filter active - clear all", "ALL", bottomAdded
    End If
    If reapplyCount >= 2 Then
        AF_AddBottom popup, "Re-apply all cleared filters", "REAPPLY", bottomAdded
    End If
    If rememberedCount > 0 Then
        AF_AddBottom popup, "Forget cleared filters", "FORGET", bottomAdded
    ElseIf mTargets.Count = 0 Then
        If bottomAdded Then
            Set AF_BuildPopup = popup
            Exit Function
        End If
        Set button = popup.Controls.Add(Type:=msoControlButton)
        button.Caption = "No active filters on this sheet"
        button.Enabled = False
    End If

    Set AF_BuildPopup = popup
    Exit Function
Failed:
    AF_Notify "Could not build the filter menu: " & Err.Description
    AF_DeletePopup
End Function

Public Sub ActiveFilters_OnClick(Optional ByVal unused As Variant)
    Dim parameter As String

    On Error GoTo Failed
    parameter = Application.CommandBars.ActionControl.Parameter
    AF_HandleClick parameter
    Exit Sub
Failed:
    AF_Notify "Could not process the filter menu: " & Err.Description
End Sub

Private Sub AF_HandleClick(ByVal parameter As String)
    Dim targetIndex As Long
    Dim target As AF_FilterTarget
    Dim cleared As Boolean

    On Error GoTo Stale
    If mTargets Is Nothing Then GoTo Stale
    If mMenuSheet Is Nothing Then GoTo Stale
    If Not AF_IsActiveWorksheet(mMenuSheet) Then GoTo Stale
    If parameter = "ALL" Then
        AF_ClearAllFilters mMenuSheet
    ElseIf parameter = "REAPPLY" Then
        ' Tables share row visibility with sheet ranges; restore the sheet range last.
        For Each target In mTargets
            If target.SourceKind = SOURCE_TABLE Then
                If target.Remembered And target.CanReapply Then AF_ReapplyTarget target
            End If
        Next target
        For Each target In mTargets
            If target.SourceKind = SOURCE_RANGE Then
                If target.Remembered And target.CanReapply Then AF_ReapplyTarget target
            End If
        Next target
    ElseIf parameter = "FORGET" Then
        AF_ForgetSheet mMenuSheet
    Else
        targetIndex = CLng(parameter)
        If targetIndex < 1 Or targetIndex > mTargets.Count Then GoTo Stale
        Set target = mTargets(targetIndex)
        If target.Remembered Then
            AF_ReapplyTarget target
        Else
            cleared = AF_ClearFilterTarget(target)
        End If
    End If
    Exit Sub
Stale:
    AF_Notify STALE_MESSAGE
End Sub

Private Function AF_GetActiveFilterTargets(ws As Worksheet, _
    Optional includeHidden As Boolean = False) As Collection
    Dim targets As Collection
    Dim autoFilter As AutoFilter
    Dim table As ListObject

    On Error GoTo Failed
    Set targets = New Collection
    If ws Is Nothing Then Err.Raise vbObjectError + 2000, , "No worksheet supplied."
    If ws.AutoFilterMode Then
        Set autoFilter = ws.AutoFilter
        If Not autoFilter Is Nothing Then
            If Not AF_IsTableRange(ws, autoFilter.Range) Then
                AF_AppendTargets targets, ws, autoFilter, SOURCE_RANGE, vbNullString
            End If
        End If
    End If
    For Each table In ws.ListObjects
        If includeHidden Or (table.ShowHeaders And table.ShowAutoFilter) Then
            Set autoFilter = table.AutoFilter
            If Not autoFilter Is Nothing Then
                AF_AppendTargets targets, ws, autoFilter, SOURCE_TABLE, table.Name
            End If
        End If
    Next table
    Set AF_GetActiveFilterTargets = targets
    Exit Function
Failed:
    AF_Notify "Could not read active filters: " & Err.Description
End Function

Private Function AF_ClearFilterTarget(t As AF_FilterTarget) As Boolean
    Dim ws As Worksheet
    Dim table As ListObject
    Dim autoFilter As AutoFilter
    Dim filterRange As Range
    Dim window As Window
    Dim scrollRow As Long
    Dim scrollColumn As Long
    Dim scrollSaved As Boolean
    Dim message As String
    Dim snapshot As AF_FilterTarget
    Dim snapshotError As String

    On Error GoTo Stale
    If t Is Nothing Then GoTo Stale
    Set filterRange = AF_ResolveTarget(t, autoFilter)
    If filterRange Is Nothing Then GoTo Stale
    Set ws = filterRange.Worksheet
    If Not AF_IsActiveWorksheet(ws) Then GoTo Stale
    If Not autoFilter.Filters(t.FieldIndex).On Then GoTo Stale

    On Error GoTo Failed
    Set snapshot = AF_CaptureForClear(t, autoFilter, snapshotError)
    Set window = Application.ActiveWindow
    scrollRow = window.ScrollRow
    scrollColumn = window.ScrollColumn
    scrollSaved = True
    filterRange.AutoFilter Field:=t.FieldIndex
    AF_Remember snapshot
    AF_ClearFilterTarget = True
    If Len(snapshotError) > 0 Then
        message = "Filter cleared, but it cannot be re-applied: " & snapshotError
    End If
CleanUp:
    If scrollSaved Then AF_RestoreScroll window, scrollRow, scrollColumn
    If Len(message) > 0 Then AF_Notify message
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

Private Sub AF_ClearAllFilters(ws As Worksheet)
    Dim table As ListObject
    Dim window As Window
    Dim scrollRow As Long
    Dim scrollColumn As Long
    Dim scrollSaved As Boolean
    Dim message As String
    Dim targets As Collection
    Dim target As AF_FilterTarget

    On Error GoTo Stale
    If ws Is Nothing Then GoTo Stale
    If Not AF_IsActiveWorksheet(ws) Then GoTo Stale
    On Error GoTo Failed
    Set window = Application.ActiveWindow
    scrollRow = window.ScrollRow
    scrollColumn = window.ScrollColumn
    scrollSaved = True
    Set targets = AF_GetActiveFilterTargets(ws, True)
    If targets Is Nothing Then Err.Raise vbObjectError + 2004, , "Could not read filter targets."
    For Each target In targets
        If Not AF_ClearFilterTarget(target) Then GoTo CleanUp
    Next target
    If ws.FilterMode Then ws.ShowAllData
CleanUp:
    If scrollSaved Then AF_RestoreScroll window, scrollRow, scrollColumn
    If Len(message) > 0 Then AF_Notify message
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

Private Function AF_Callback() As String
    AF_Callback = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!ActiveFilters_OnClick"
End Function

Private Function AF_BindingMatches(ByVal binding As String) As Boolean
    Dim fullBinding As String

    fullBinding = "'" & Replace(ThisWorkbook.FullName, "'", "''") & "'!ActiveFilters_OnClick"
    AF_BindingMatches = (binding = AF_Callback() Or binding = fullBinding)
End Function

Private Sub AF_AddBottom(popup As CommandBar, caption As String, parameter As String, _
    bottomAdded As Boolean)
    Dim button As CommandBarButton

    Set button = popup.Controls.Add(Type:=msoControlButton)
    button.Caption = caption
    button.BeginGroup = Not bottomAdded
    button.OnAction = AF_Callback()
    button.Parameter = parameter
    bottomAdded = True
End Sub

Private Function AF_ResolveTarget(t As AF_FilterTarget, autoFilter As AutoFilter) As Range
    Dim ws As Worksheet
    Dim filterRange As Range

    On Error GoTo Stale
    Set autoFilter = Nothing
    Set ws = Application.Workbooks(t.WorkbookName).Worksheets(t.SheetName)
    If Not ws Is t.TargetSheet Then Exit Function
    Select Case t.SourceKind
        Case SOURCE_RANGE
            If Len(t.TableName) <> 0 Or Not ws.AutoFilterMode Then Exit Function
            Set autoFilter = ws.AutoFilter
            If autoFilter Is Nothing Then Exit Function
            If AF_IsTableRange(ws, autoFilter.Range) Then Exit Function
        Case SOURCE_TABLE
            Set autoFilter = ws.ListObjects(t.TableName).AutoFilter
            If autoFilter Is Nothing Then Exit Function
        Case Else
            Exit Function
    End Select
    Set filterRange = autoFilter.Range
    If t.SourceKind = SOURCE_RANGE Then
        If filterRange.Address <> t.FilterRangeAddress Then Exit Function
    End If
    If t.FieldIndex < 1 Or t.FieldIndex > autoFilter.Filters.Count Then Exit Function
    If t.FieldIndex > filterRange.Columns.Count Then Exit Function
    If filterRange.Columns(t.FieldIndex).Column <> t.SheetColumn Then Exit Function
    If AF_ColumnLetters(t.SheetColumn) <> t.ColumnLetter Then Exit Function
    If CStr(filterRange.Cells(1, t.FieldIndex).Text) <> t.HeaderText Then Exit Function
    Set AF_ResolveTarget = filterRange
Stale:
    ' Missing objects are stale identities, not a reason to open or repair workbooks.
End Function

Private Function AF_ReadCriterion(filter As Filter, ByVal second As Boolean, _
    value As Variant) As Boolean
    Dim number As Long
    Dim description As String

    On Error GoTo Missing
    If second Then
        value = filter.Criteria2
    Else
        value = filter.Criteria1
    End If
    AF_ReadCriterion = True
    Exit Function
Missing:
    number = Err.Number
    description = Err.Description
    ' Excel raises 1004 for an absent criterion (notably Criteria1 on date trees).
    If number <> 1004 Then Err.Raise number, "AF_ReadCriterion", description
End Function

Private Function AF_Snapshot(t As AF_FilterTarget, filter As Filter) As AF_FilterTarget
    Dim criteria1 As Variant
    Dim criteria2 As Variant

    If mTestMode And mForceSnapshotFailure Then
        Err.Raise vbObjectError + 2007, "AF_Snapshot", "Forced self-test snapshot failure."
    End If
    t.FilterOperator = filter.Operator
    t.CanReapply = False
    t.HasCriteria1 = False
    t.HasCriteria2 = False
    Select Case t.FilterOperator
        Case 0, xlAnd, xlOr, xlFilterValues, xlFilterDynamic
            t.HasCriteria1 = AF_ReadCriterion(filter, False, criteria1)
            If t.FilterOperator = xlAnd Or t.FilterOperator = xlOr Then
                t.HasCriteria2 = AF_ReadCriterion(filter, True, criteria2)
            ElseIf t.FilterOperator = xlFilterValues And Not t.HasCriteria1 Then
                t.HasCriteria2 = AF_ReadCriterion(filter, True, criteria2)
            ElseIf t.FilterOperator = xlFilterDynamic Then
                t.HasCriteria2 = AF_ReadCriterion(filter, True, criteria2)
            End If
            If Not t.HasCriteria1 And Not t.HasCriteria2 Then
                ' Excel can expose neither criterion for date-tree selections.
                Set AF_Snapshot = t
                Exit Function
            End If
            t.Criteria1 = criteria1
            t.Criteria2 = criteria2
            t.CanReapply = True
    End Select
    Set AF_Snapshot = t
End Function

Private Function AF_CaptureForClear(t As AF_FilterTarget, autoFilter As AutoFilter, _
    snapshotError As String) As AF_FilterTarget
    On Error GoTo Unsupported
    Set AF_CaptureForClear = AF_Snapshot(t, autoFilter.Filters(t.FieldIndex))
    Exit Function
Unsupported:
    snapshotError = CStr(Err.Number) & " (" & Err.Description & ")"
    t.CanReapply = False
    t.HasCriteria1 = False
    t.HasCriteria2 = False
    t.Criteria1 = Empty
    t.Criteria2 = Empty
    Set AF_CaptureForClear = t
End Function

Private Sub AF_Remember(t As AF_FilterTarget)
    Dim index As Long
    Dim existing As AF_FilterTarget

    If mRemembered Is Nothing Then Set mRemembered = New Collection
    For index = mRemembered.Count To 1 Step -1
        Set existing = mRemembered(index)
        If existing.TargetSheet Is t.TargetSheet Then
            If existing.SheetColumn = t.SheetColumn Then mRemembered.Remove index
        End If
    Next index
    t.Remembered = True
    mRemembered.Add t
End Sub

Private Function AF_RemoveRemembered(t As AF_FilterTarget) As Boolean
    Dim index As Long
    Dim existing As AF_FilterTarget

    If mRemembered Is Nothing Then Exit Function
    For index = mRemembered.Count To 1 Step -1
        Set existing = mRemembered(index)
        If existing Is t Then
            mRemembered.Remove index
            AF_RemoveRemembered = True
            Exit Function
        End If
    Next index
End Function

Private Sub AF_ForgetSheet(ws As Worksheet)
    Dim index As Long
    Dim target As AF_FilterTarget

    If mRemembered Is Nothing Then Exit Sub
    For index = mRemembered.Count To 1 Step -1
        Set target = mRemembered(index)
        If target.TargetSheet Is ws Then mRemembered.Remove index
    Next index
End Sub

Private Function AF_GetMenuTargets(ws As Worksheet) As Collection
    Dim targets As Collection
    Dim target As AF_FilterTarget
    Dim active As AF_FilterTarget
    Dim autoFilter As AutoFilter
    Dim filterRange As Range
    Dim index As Long
    Dim drop As Boolean
    Dim visible As Boolean

    Set targets = AF_GetActiveFilterTargets(ws)
    If targets Is Nothing Then Exit Function
    If Not mRemembered Is Nothing Then
        For index = mRemembered.Count To 1 Step -1
            Set target = mRemembered(index)
            Set filterRange = AF_ResolveTarget(target, autoFilter)
            drop = (filterRange Is Nothing)
            If Not drop Then drop = autoFilter.Filters(target.FieldIndex).On
            If Not drop And target.TargetSheet Is ws Then
                For Each active In targets
                    If active.SheetColumn = target.SheetColumn Then drop = True
                Next active
            End If
            If drop Then
                mRemembered.Remove index
            ElseIf target.TargetSheet Is ws Then
                visible = True
                If target.SourceKind = SOURCE_TABLE Then
                    With ws.ListObjects(target.TableName)
                        visible = .ShowHeaders And .ShowAutoFilter
                    End With
                End If
                If visible Then AF_InsertSorted targets, target
            End If
        Next index
    End If
    Set AF_GetMenuTargets = targets
End Function

Private Sub AF_ReapplyTarget(t As AF_FilterTarget)
    Dim filterRange As Range
    Dim autoFilter As AutoFilter
    Dim existing As AF_FilterTarget
    Dim found As Boolean
    Dim window As Window
    Dim scrollRow As Long
    Dim scrollColumn As Long
    Dim scrollSaved As Boolean
    Dim message As String

    On Error GoTo Failed
    If Not t.CanReapply Or Not t.Remembered Then GoTo Stale
    If mRemembered Is Nothing Then GoTo Stale
    For Each existing In mRemembered
        If existing Is t Then found = True
    Next existing
    If Not found Then GoTo Stale
    Set filterRange = AF_ResolveTarget(t, autoFilter)
    If filterRange Is Nothing Then GoTo Stale
    If Not AF_IsActiveWorksheet(filterRange.Worksheet) Then GoTo Stale
    If autoFilter.Filters(t.FieldIndex).On Then GoTo Stale
    Set window = Application.ActiveWindow
    scrollRow = window.ScrollRow
    scrollColumn = window.ScrollColumn
    scrollSaved = True
    Select Case t.FilterOperator
        Case 0
            filterRange.AutoFilter Field:=t.FieldIndex, Criteria1:=t.Criteria1
        Case xlAnd, xlOr
            filterRange.AutoFilter Field:=t.FieldIndex, Criteria1:=t.Criteria1, _
                Operator:=t.FilterOperator, Criteria2:=t.Criteria2
        Case xlFilterValues
            If t.HasCriteria1 Then
                filterRange.AutoFilter Field:=t.FieldIndex, Criteria1:=t.Criteria1, Operator:=xlFilterValues
            Else
                filterRange.AutoFilter Field:=t.FieldIndex, Operator:=xlFilterValues, Criteria2:=t.Criteria2
            End If
        Case xlFilterDynamic
            If t.HasCriteria2 Then
                filterRange.AutoFilter Field:=t.FieldIndex, Criteria1:=t.Criteria1, _
                    Operator:=xlFilterDynamic, Criteria2:=t.Criteria2
            Else
                filterRange.AutoFilter Field:=t.FieldIndex, Criteria1:=t.Criteria1, _
                    Operator:=xlFilterDynamic
            End If
    End Select
    If Not autoFilter.Filters(t.FieldIndex).On Then
        Err.Raise vbObjectError + 2006, , "Excel did not re-apply the filter."
    End If
    AF_RemoveRemembered t
CleanUp:
    If scrollSaved Then AF_RestoreScroll window, scrollRow, scrollColumn
    If Len(message) > 0 Then AF_Notify message
    Exit Sub
Stale:
    message = STALE_MESSAGE
    GoTo CleanUp
Failed:
    message = "Filter could not be re-applied: " & Err.Description
    Resume CleanUp
End Sub

Private Sub AF_AppendTargets(targets As Collection, ws As Worksheet, _
    autoFilter As AutoFilter, sourceKind As String, tableName As String)

    Dim fieldIndex As Long
    Dim target As AF_FilterTarget
    Dim filterRange As Range

    Set filterRange = autoFilter.Range
    For fieldIndex = 1 To autoFilter.Filters.Count
        If autoFilter.Filters(fieldIndex).On Then
            Set target = New AF_FilterTarget
            target.WorkbookName = ws.Parent.Name
            Set target.TargetSheet = ws
            target.SheetName = ws.Name
            target.SourceKind = sourceKind
            target.TableName = tableName
            target.FieldIndex = fieldIndex
            target.SheetColumn = filterRange.Columns(fieldIndex).Column
            target.ColumnLetter = AF_ColumnLetters(target.SheetColumn)
            target.HeaderText = CStr(filterRange.Cells(1, fieldIndex).Text)
            target.FilterRangeAddress = filterRange.Address
            AF_InsertSorted targets, target
        End If
    Next fieldIndex
End Sub

Private Sub AF_InsertSorted(targets As Collection, target As AF_FilterTarget)
    Dim targetIndex As Long
    Dim existing As AF_FilterTarget

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

Private Function AF_IsTableRange(ws As Worksheet, filterRange As Range) As Boolean
    Dim table As ListObject
    Dim tableFilter As AutoFilter

    For Each table In ws.ListObjects
        If table.Range.Address = filterRange.Address Then
            AF_IsTableRange = True
            Exit Function
        End If
        If table.ShowHeaders Then
            Set tableFilter = table.AutoFilter
            If Not tableFilter Is Nothing Then
                If tableFilter.Range.Address = filterRange.Address Then
                    AF_IsTableRange = True
                    Exit Function
                End If
            End If
        End If
    Next table
End Function

Private Function AF_IsActiveWorksheet(ws As Worksheet) As Boolean
    Dim active As Worksheet

    If Not TypeOf Application.ActiveSheet Is Worksheet Then Exit Function
    If Application.ActiveWindow Is Nothing Then Exit Function
    If Application.ActiveWindow.SelectedSheets.Count > 1 Then Exit Function
    Set active = Application.ActiveSheet
    AF_IsActiveWorksheet = (active Is ws)
End Function

Private Function AF_HasActiveFilters(autoFilter As AutoFilter) As Boolean
    Dim fieldIndex As Long

    If autoFilter Is Nothing Then Exit Function
    For fieldIndex = 1 To autoFilter.Filters.Count
        If autoFilter.Filters(fieldIndex).On Then
            AF_HasActiveFilters = True
            Exit Function
        End If
    Next fieldIndex
End Function

Private Function AF_ColumnLetters(ByVal columnNumber As Long) As String
    Dim remainder As Long

    Do While columnNumber > 0
        remainder = (columnNumber - 1) Mod 26
        AF_ColumnLetters = Chr$(65 + remainder) & AF_ColumnLetters
        columnNumber = (columnNumber - 1) \ 26
    Loop
End Function

Private Function AF_CleanHeader(ByVal headerText As String) As String
    headerText = Replace(headerText, vbCrLf, " ")
    headerText = Replace(headerText, vbCr, " ")
    headerText = Replace(headerText, vbLf, " ")
    AF_CleanHeader = Left$(headerText, 40)
End Function

Private Sub AF_DeletePopup()
    Dim popup As CommandBar
    On Error Resume Next
    For Each popup In Application.CommandBars
        If popup.Name = POPUP_NAME Then
            popup.Delete
            Exit For
        End If
    Next popup
    Set mTargets = Nothing
    Set mMenuSheet = Nothing
    mPopupOpen = False
    On Error GoTo 0
End Sub

Private Function AF_PopupExists() As Boolean
    Dim popup As CommandBar
    For Each popup In Application.CommandBars
        If popup.Name = POPUP_NAME Then
            AF_PopupExists = True
            Exit Function
        End If
    Next popup
End Function

Private Sub AF_RestoreScroll(window As Window, scrollRow As Long, scrollColumn As Long)
    On Error Resume Next
    window.ScrollRow = scrollRow
    window.ScrollColumn = scrollColumn
    On Error GoTo 0
End Sub

Private Sub AF_Notify(message As String)
    On Error Resume Next
    If Not mTestMode Then MsgBox message, vbExclamation, "Active Filters"
    On Error GoTo 0
End Sub

Public Sub ActiveFilters_SelfTest(Optional ByVal unused As Variant)
    Dim testWorkbook As Workbook
    Dim ws As Worksheet
    Dim shiftedSheet As Worksheet
    Dim tableA As ListObject
    Dim tableB As ListObject
    Dim targets As Collection
    Dim target As AF_FilterTarget
    Dim staleTarget As AF_FilterTarget
    Dim data() As Variant
    Dim tableData() As Variant
    Dim valueCriteria() As Variant
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
    Dim failureNumber As Long
    Dim popup As CommandBar
    Dim bindingsMatch As Boolean
    Dim controlIndex As Long
    Dim bindingDetails As String
    Dim originalRemembered As Collection

    If mTestingOptional Then
        mOptionalSelfTestCalled = True
        Exit Sub
    End If
    On Error GoTo FailedTest
    mReport = vbNullString
    mCurrentStep = "initialize self-test"
    originalTestMode = mTestMode
    mTestMode = True
    Set originalRemembered = mRemembered
    Set mRemembered = New Collection
    AF_ReportLine "ActiveFilters self-test"
    mCurrentStep = "create temporary workbook"
    Set testWorkbook = Application.Workbooks.Add(xlWBATWorksheet)
    mCurrentStep = "get test worksheet"
    Set ws = testWorkbook.Worksheets(1)
    mCurrentStep = "dimension test arrays"
    ReDim data(1 To 200, 1 To 25)
    ReDim tableData(1 To 21, 1 To 3)
    ReDim valueCriteria(1 To 2)
    mCurrentStep = "build test data"
    For columnIndex = LBound(data, 2) To UBound(data, 2)
        data(1, columnIndex) = "Header " & AF_ColumnLetters(columnIndex)
    Next columnIndex
    data(1, 3) = "Category & Type"
    data(1, 20) = "State"
    data(1, 24) = "Shade"
    For rowIndex = LBound(data, 1) + 1 To UBound(data, 1)
        For columnIndex = LBound(data, 2) To UBound(data, 2)
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
    mCurrentStep = "write test data"
    ws.Range("A1").Resize(UBound(data, 1), UBound(data, 2)).Value = data
    mCurrentStep = "format test dates and colours"
    ws.Range("B2:B200").NumberFormat = "yyyy-mm-dd"
    ws.Range("X2:X200").Interior.Color = RGB(255, 255, 0)
    For rowIndex = LBound(data, 1) + 1 To UBound(data, 1) Step 2
        ws.Cells(rowIndex, 24).Interior.Color = RGB(0, 176, 80)
    Next rowIndex
    mCurrentStep = "enable sheet AutoFilter"
    ws.Range("A1").Resize(UBound(data, 1), UBound(data, 2)).AutoFilter
    mCurrentStep = "apply C value filter"
    ws.AutoFilter.Range.AutoFilter Field:=3, Criteria1:="Keep"
    mCurrentStep = "build array filter criteria"
    valueCriteria(1) = "Open"
    valueCriteria(2) = "Hold"
    mCurrentStep = "apply T array filter"
    ws.AutoFilter.Range.AutoFilter Field:=20, _
        Criteria1:=valueCriteria, Operator:=xlFilterValues
    mCurrentStep = "apply X cell-colour filter"
    ws.AutoFilter.Range.AutoFilter Field:=24, _
        Criteria1:=RGB(0, 176, 80), Operator:=xlFilterCellColor

    mCurrentStep = "build table data"
    tableData(1, 1) = "ID"
    tableData(1, 2) = "Status"
    tableData(1, 3) = "Quantity"
    For rowIndex = LBound(tableData, 1) + 1 To UBound(tableData, 1)
        tableData(rowIndex, 1) = rowIndex - 1
        If rowIndex Mod 2 = 0 Then
            tableData(rowIndex, 2) = "Keep"
        Else
            tableData(rowIndex, 2) = "Drop"
        End If
        tableData(rowIndex, 3) = rowIndex * 10
    Next rowIndex
    mCurrentStep = "write TblA data"
    ws.Range("AA1").Resize(UBound(tableData, 1), UBound(tableData, 2)).Value = tableData
    mCurrentStep = "write TblB data"
    ws.Range("AE1").Resize(UBound(tableData, 1), UBound(tableData, 2)).Value = tableData
    mCurrentStep = "create TblA"
    Set tableA = ws.ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=ws.Range("AA1").Resize(UBound(tableData, 1), UBound(tableData, 2)), _
        XlListObjectHasHeaders:=xlYes)
    tableA.Name = "TblA"
    mCurrentStep = "create TblB"
    Set tableB = ws.ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=ws.Range("AE1").Resize(UBound(tableData, 1), UBound(tableData, 2)), _
        XlListObjectHasHeaders:=xlYes)
    tableB.Name = "TblB"
    mCurrentStep = "apply TblA filter"
    tableA.Range.AutoFilter Field:=2, Criteria1:="Keep"
    mCurrentStep = "apply TblB filter"
    tableB.Range.AutoFilter Field:=2, Criteria1:="Keep"

    mCurrentStep = "enumerate five active targets"
    Set targets = AF_GetActiveFilterTargets(ws)
    If targets Is Nothing Then Err.Raise vbObjectError + 2001, , "Enumeration failed."
    mCurrentStep = "assert exactly five active targets"
    AF_CheckTest targets.Count = 5, "Exactly five active targets", passed, failed
    mCurrentStep = "assert C sheet value filter"
    AF_CheckTest AF_TargetMatches(targets, 1, "C", "Category & Type", "", 3, 3, ws), _
        "C: sheet value filter", passed, failed
    mCurrentStep = "assert T off-screen array filter"
    AF_CheckTest AF_TargetMatches(targets, 2, "T", "State", "", 20, 20, ws), _
        "T: off-screen array filter", passed, failed
    mCurrentStep = "assert X off-screen cell-colour filter"
    AF_CheckTest AF_TargetMatches(targets, 3, "X", "Shade", "", 24, 24, ws), _
        "X: off-screen cell-colour filter", passed, failed
    mCurrentStep = "assert TblA target metadata"
    AF_CheckTest AF_TargetMatches(targets, 4, "AB", "Status", "TblA", 2, 28, ws), _
        "AB: TblA duplicate Status header", passed, failed
    mCurrentStep = "assert TblB target metadata"
    AF_CheckTest AF_TargetMatches(targets, 5, "AF", "Status", "TblB", 2, 32, ws), _
        "AF: TblB duplicate Status header", passed, failed

    mCurrentStep = "build five-filter popup without showing UI"
    Set popup = AF_BuildPopup(ws)
    mCurrentStep = "assert five targets plus clear-all control"
    AF_CheckTest popup.Controls.Count = 6, "Popup: five targets plus clear all", passed, failed
    mCurrentStep = "assert popup captions and escaped ampersand"
    AF_CheckTest popup.Controls(1).Caption = "Col C - Category && Type" _
        And popup.Controls(2).Caption = "Col T - State" _
        And popup.Controls(3).Caption = "Col X - Shade" _
        And popup.Controls(4).Caption = "Col AB - Status [TblA]" _
        And popup.Controls(5).Caption = "Col AF - Status [TblB]", _
        "Popup: column/header/table captions and escaped ampersand", passed, failed
    mCurrentStep = "assert multi-filter clear-all control"
    AF_CheckTest popup.Controls(6).Caption = "Clear all filters on this sheet" _
        And popup.Controls(6).BeginGroup And popup.Controls(6).Parameter = "ALL", _
        "Popup: separated clear-all item for multiple filters", passed, failed
    mCurrentStep = "inspect popup callback bindings"
    bindingsMatch = True
    For controlIndex = 1 To 5
        If popup.Controls(controlIndex).Parameter <> CStr(controlIndex) Then bindingsMatch = False
        If Not AF_BindingMatches(popup.Controls(controlIndex).OnAction) Then bindingsMatch = False
        bindingDetails = bindingDetails & " [" & CStr(controlIndex) & ": " _
            & popup.Controls(controlIndex).Parameter & ", " & popup.Controls(controlIndex).OnAction & "]"
    Next controlIndex
    mCurrentStep = "assert popup callback bindings"
    AF_CheckTest bindingsMatch, "Popup: target parameters and callback bindings", passed, failed, _
        "parameters 1..5; callback qualified by host Name or FullName", bindingDetails
    AF_ReportLine "Popup OnAction: " & popup.Controls(1).OnAction
    AF_DeletePopup
    Set popup = Nothing

    mCurrentStep = "find T filter target"
    Set target = AF_FindTarget(targets, "", 20)
    mCurrentStep = "save scroll position"
    scrollRow = Application.ActiveWindow.ScrollRow
    scrollColumn = Application.ActiveWindow.ScrollColumn
    mCurrentStep = "clear T filter"
    cleared = AF_ClearFilterTarget(target)
    mCurrentStep = "assert T filter cleared"
    AF_CheckTest cleared And Not ws.AutoFilter.Filters(20).On, _
        "Clear T only", passed, failed
    mCurrentStep = "enumerate remaining four targets"
    Set targets = AF_GetActiveFilterTargets(ws)
    mCurrentStep = "assert other four filters remain"
    AF_CheckTest targets.Count = 4 And ws.AutoFilter.Filters(3).On _
        And ws.AutoFilter.Filters(24).On And tableA.AutoFilter.Filters(2).On _
        And tableB.AutoFilter.Filters(2).On, "Other four filters remain", passed, failed
    mCurrentStep = "assert scroll position preserved"
    AF_CheckTest Application.ActiveWindow.ScrollRow = scrollRow _
        And Application.ActiveWindow.ScrollColumn = scrollColumn, _
        "Scroll position preserved", passed, failed

    mCurrentStep = "reapply T filter and build dispatcher snapshot"
    ws.AutoFilter.Range.AutoFilter Field:=20, Criteria1:=valueCriteria, Operator:=xlFilterValues
    Set popup = AF_BuildPopup(ws)
    mCurrentStep = "dispatch valid T parameter"
    AF_HandleClick "2"
    Set targets = AF_GetActiveFilterTargets(ws)
    mCurrentStep = "assert dispatcher clears exactly T"
    AF_CheckTest targets.Count = 4 And Not ws.AutoFilter.Filters(20).On _
        And ws.AutoFilter.Filters(3).On And ws.AutoFilter.Filters(24).On _
        And tableA.AutoFilter.Filters(2).On And tableB.AutoFilter.Filters(2).On, _
        "Dispatcher: valid parameter clears only T", passed, failed
    mCurrentStep = "dispatch out-of-date parameter"
    AF_HandleClick "9999"
    Set targets = AF_GetActiveFilterTargets(ws)
    mCurrentStep = "assert invalid parameter changes nothing"
    AF_CheckTest targets.Count = 4 And ws.AutoFilter.Filters(3).On _
        And ws.AutoFilter.Filters(24).On And tableA.AutoFilter.Filters(2).On _
        And tableB.AutoFilter.Filters(2).On And Not ws.AutoFilter.Filters(20).On, _
        "Dispatcher: invalid parameter changes nothing", passed, failed
    mCurrentStep = "change popup target header and dispatch stale target"
    ws.Cells(1, 3).Value = "Temporarily changed"
    AF_HandleClick "1"
    Set targets = AF_GetActiveFilterTargets(ws)
    mCurrentStep = "assert stale dispatcher target changes nothing"
    AF_CheckTest targets.Count = 4 And ws.AutoFilter.Filters(3).On _
        And ws.AutoFilter.Filters(24).On And tableA.AutoFilter.Filters(2).On _
        And tableB.AutoFilter.Filters(2).On And Not ws.AutoFilter.Filters(20).On, _
        "Dispatcher: stale header changes nothing", passed, failed
    ws.Cells(1, 3).Value = "Category & Type"
    AF_DeletePopup
    Set popup = Nothing
    mCurrentStep = "assert popup and callback state deleted"
    AF_CheckTest Not AF_PopupExists() And mTargets Is Nothing And mMenuSheet Is Nothing, _
        "Popup: deleted with callback state after use", passed, failed

    mCurrentStep = "find TblB filter target"
    Set target = AF_FindTarget(targets, "TblB", 2)
    mCurrentStep = "clear TblB filter"
    cleared = AF_ClearFilterTarget(target)
    mCurrentStep = "assert TblB cleared and TblA retained"
    AF_CheckTest cleared And Not tableB.AutoFilter.Filters(2).On _
        And tableA.AutoFilter.Filters(2).On, "Clear TblB without clearing TblA", passed, failed
    mCurrentStep = "enumerate targets before stale-header test"
    Set targets = AF_GetActiveFilterTargets(ws)
    expectedCount = targets.Count
    mCurrentStep = "find C filter target for stale-header test"
    Set staleTarget = AF_FindTarget(targets, "", 3)
    mCurrentStep = "change C header"
    ws.Cells(1, 3).Value = "Changed category"
    mCurrentStep = "attempt stale-header clear"
    cleared = AF_ClearFilterTarget(staleTarget)
    mCurrentStep = "enumerate targets after stale-header attempt"
    Set targets = AF_GetActiveFilterTargets(ws)
    mCurrentStep = "assert stale header rejected without clearing"
    AF_CheckTest Not cleared And targets.Count = expectedCount _
        And ws.AutoFilter.Filters(3).On And ws.AutoFilter.Filters(24).On _
        And tableA.AutoFilter.Filters(2).On And Not tableB.AutoFilter.Filters(2).On _
        And Not ws.AutoFilter.Filters(20).On, "Stale header rejected without clearing", passed, failed

    mCurrentStep = "clear all sheet and table filters"
    AF_ClearAllFilters ws
    mCurrentStep = "enumerate targets after clear all"
    Set targets = AF_GetActiveFilterTargets(ws)
    mCurrentStep = "assert no active filters"
    AF_CheckTest targets.Count = 0, "Clear all leaves no active filters", passed, failed
    mCurrentStep = "assert filter arrows retained"
    AF_CheckTest ws.AutoFilterMode And tableA.ShowAutoFilter And tableB.ShowAutoFilter, _
        "Filter arrows retained", passed, failed

    mCurrentStep = "build popup on unfiltered sheet"
    AF_ForgetSheet ws
    Set popup = AF_BuildPopup(ws)
    mCurrentStep = "assert disabled empty-state item without clear all"
    AF_CheckTest popup.Controls.Count = 1 And Not popup.Controls(1).Enabled _
        And popup.Controls(1).Caption = "No active filters on this sheet", _
        "Popup: disabled empty-state item and no clear all", passed, failed
    AF_DeletePopup
    Set popup = Nothing

    mCurrentStep = "create offset-range worksheet"
    Set shiftedSheet = testWorkbook.Worksheets.Add(After:=ws)
    mCurrentStep = "build offset-range data"
    shiftedSheet.Range("D1").Value = "ID"
    shiftedSheet.Range("E1").Value = "Offset status"
    shiftedSheet.Range("F1").Value = "Quantity"
    For rowIndex = 2 To 8
        shiftedSheet.Cells(rowIndex, 4).Value = rowIndex
        shiftedSheet.Cells(rowIndex, 5).Value = rowIndex Mod 2
        shiftedSheet.Cells(rowIndex, 6).Value = rowIndex * 10
    Next rowIndex
    mCurrentStep = "apply offset-range filter"
    shiftedSheet.Range("D1:F8").AutoFilter Field:=2, Criteria1:="0"
    mCurrentStep = "enumerate offset-range target"
    Set targets = AF_GetActiveFilterTargets(shiftedSheet)
    mCurrentStep = "assert offset-range field and column"
    AF_CheckTest AF_TargetMatches(targets, 1, "E", "Offset status", "", 2, 5, shiftedSheet), _
        "Offset range: field 2 is sheet column E", passed, failed
    mCurrentStep = "build single-filter popup"
    Set popup = AF_BuildPopup(shiftedSheet)
    mCurrentStep = "assert single filter has no clear-all item"
    AF_CheckTest popup.Controls.Count = 1 And popup.Controls(1).Caption = "Col E - Offset status" _
        And popup.Controls(1).Enabled And popup.Controls(1).Parameter = "1", _
        "Popup: single-filter item and no clear all", passed, failed
    AF_DeletePopup
    Set popup = Nothing
    mCurrentStep = "assert single-filter popup deleted"
    AF_CheckTest Not AF_PopupExists(), "Popup: single-filter popup deleted", passed, failed
    mCurrentStep = "find offset-range target"
    Set target = AF_FindTarget(targets, "", 2)
    mCurrentStep = "clear offset-range filter"
    cleared = AF_ClearFilterTarget(target)
    mCurrentStep = "assert offset-range filter cleared"
    AF_CheckTest cleared And Not shiftedSheet.AutoFilter.Filters(2).On _
        And shiftedSheet.AutoFilterMode, "Clear offset range field", passed, failed
    AF_ForgetSheet shiftedSheet
    AF_TestToggle testWorkbook, passed, failed

Finish:
    On Error Resume Next
    AF_DeletePopup
    Set popup = Nothing
    mCurrentStep = "close temporary workbook without saving"
    If Not testWorkbook Is Nothing Then
        Err.Clear
        testWorkbook.Close SaveChanges:=False
        If Err.Number <> 0 Then
            failureNumber = Err.Number
            failureText = Err.Description
            failed = failed + 1
            AF_ReportLine "FAIL: runtime error " & CStr(failureNumber) & " (" & failureText _
                & ") at step: " & mCurrentStep
        End If
    End If
    mTestMode = originalTestMode
    Set mRemembered = originalRemembered
    mTestingOptional = False
    mForceSnapshotFailure = False
    AF_ReportLine "Summary: " & CStr(passed) & " PASS, " & CStr(failed) & " FAIL"
    On Error GoTo 0
    Exit Sub
FailedTest:
    failureNumber = Err.Number
    failureText = Err.Description
    failed = failed + 1
    AF_ReportLine "FAIL: runtime error " & CStr(failureNumber) & " (" & failureText _
        & ") at step: " & mCurrentStep
    Resume Finish
End Sub

Public Function ActiveFilters_SelfTestReport(Optional ByVal tracePath As String = vbNullString) As String
    On Error GoTo Failed
    mTracePath = tracePath
    ActiveFilters_SelfTest
    ActiveFilters_SelfTestReport = mReport
    mTracePath = vbNullString
    Exit Function
Failed:
    mTracePath = vbNullString
    AF_ReportLine "FAIL: report error " & CStr(Err.Number) & " (" & Err.Description & ")"
    AF_ReportLine "Summary: 0 PASS, 1 FAIL"
    ActiveFilters_SelfTestReport = mReport
End Function

Private Sub AF_TestToggle(testWorkbook As Workbook, passed As Long, failed As Long)
    Dim ws As Worksheet
    Dim otherSheet As Worksheet
    Dim table As ListObject
    Dim popup As CommandBar
    Dim targets As Collection
    Dim target As AF_FilterTarget
    Dim snapshot As AF_FilterTarget
    Dim data() As Variant
    Dim names As Variant
    Dim testIndex As Long
    Dim rowIndex As Long
    Dim field As Long
    Dim beforeRows As String
    Dim afterRows As String
    Dim expected As String
    Dim actual As String
    Dim scrollRow As Long
    Dim scrollColumn As Long
    Dim blockIndex As Long

    On Error GoTo Failed
    mCurrentStep = "create toggle test worksheet"
    Set ws = testWorkbook.Worksheets.Add
    ReDim data(1 To 121, 1 To 12)
    For field = 1 To 12
        data(1, field) = "Toggle " & CStr(field)
    Next field
    For rowIndex = 2 To 121
        data(rowIndex, 1) = "Text " & CStr(rowIndex Mod 4)
        data(rowIndex, 2) = rowIndex - 1
        data(rowIndex, 3) = DateSerial(2026, 1, 1) + (rowIndex - 2) * 2
        data(rowIndex, 4) = Choose((rowIndex Mod 4) + 1, "Alpha", "Beta", "Gamma", "Delta")
        data(rowIndex, 12) = Choose((rowIndex Mod 5 = 0) + 2, "Drop", "Keep")
    Next rowIndex
    ws.Range("A1").Resize(121, 12).Value = data
    ws.Range("C2:C121").NumberFormat = "yyyy-mm-dd"
    ws.Range("A1:L121").AutoFilter
    names = Array("values list", "begins-with text", "number between (xlAnd)", _
        "number xlOr", "top 10 items", "bottom 10 items", "top 10 percent", _
        "bottom 10 percent", "date tree Criteria2", "above-average dynamic", _
        "below-average dynamic", "date-period dynamic", "single number condition", "values array")
    On Error GoTo CaseFailed
    For testIndex = LBound(names) To UBound(names)
        mCurrentStep = "toggle: " & names(testIndex)
        If ws.FilterMode Then ws.ShowAllData
        AF_ForgetSheet ws
        ws.AutoFilter.Range.AutoFilter Field:=12, Criteria1:="Keep"
        field = 2
        Select Case testIndex
            Case 0
                field = 4
                ws.AutoFilter.Range.AutoFilter Field:=field, _
                    Criteria1:=Array("Alpha", "Beta"), Operator:=xlFilterValues
            Case 1
                field = 1
                ws.AutoFilter.Range.AutoFilter Field:=field, Criteria1:="Text 1*"
            Case 2
                ws.AutoFilter.Range.AutoFilter Field:=field, Criteria1:=">=25", _
                    Operator:=xlAnd, Criteria2:="<=85"
            Case 3
                ws.AutoFilter.Range.AutoFilter Field:=field, Criteria1:="<=20", _
                    Operator:=xlOr, Criteria2:=">=100"
            Case 4
                ws.AutoFilter.Range.AutoFilter Field:=field, Criteria1:=10, Operator:=xlTop10Items
            Case 5
                ws.AutoFilter.Range.AutoFilter Field:=field, Criteria1:=10, Operator:=xlBottom10Items
            Case 6
                ws.AutoFilter.Range.AutoFilter Field:=field, Criteria1:=10, Operator:=xlTop10Percent
            Case 7
                ws.AutoFilter.Range.AutoFilter Field:=field, Criteria1:=10, Operator:=xlBottom10Percent
            Case 8
                field = 3
                ws.AutoFilter.Range.AutoFilter Field:=field, Operator:=xlFilterValues, _
                    Criteria2:=Array(1, "1/31/2026", 1, "3/31/2026")
            Case 9
                ws.AutoFilter.Range.AutoFilter Field:=field, _
                    Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic
            Case 10
                ws.AutoFilter.Range.AutoFilter Field:=field, _
                    Criteria1:=xlFilterBelowAverage, Operator:=xlFilterDynamic
            Case 11
                field = 3
                ws.AutoFilter.Range.AutoFilter Field:=field, _
                    Criteria1:=xlFilterThisYear, Operator:=xlFilterDynamic
            Case 12
                ws.AutoFilter.Range.AutoFilter Field:=field, Criteria1:=">=45"
            Case 13
                field = 4
                ws.AutoFilter.Range.AutoFilter Field:=field, _
                    Criteria1:=Array("Alpha", "Beta", "Gamma"), Operator:=xlFilterValues
        End Select
        beforeRows = AF_HiddenRows(ws, 121)
        Set targets = AF_GetActiveFilterTargets(ws)
        Set target = AF_FindTarget(targets, "", field)
        Set snapshot = AF_Snapshot(target, ws.AutoFilter.Filters(field))
        expected = AF_FilterSignature(snapshot)
        Set popup = AF_BuildPopup(ws)
        scrollRow = Application.ActiveWindow.ScrollRow
        scrollColumn = Application.ActiveWindow.ScrollColumn
        popup.Controls(1).Execute
        AF_CheckTest Not ws.AutoFilter.Filters(field).On And ws.AutoFilter.Filters(12).On, _
            "Toggle " & names(testIndex) & ": callback clears only target", passed, failed
        Set popup = AF_BuildPopup(ws)
        If testIndex >= 4 And testIndex <= 8 Then
            AF_CheckTest Not snapshot.CanReapply And popup.Controls.Count = 3 _
                And popup.Controls(1).Caption = "Col " & target.ColumnLetter & " - " _
                    & target.HeaderText & " (cannot re-apply)" _
                And Not popup.Controls(1).Enabled And popup.Controls(1).State = msoButtonUp _
                And popup.Controls(2).State = msoButtonDown And popup.Controls(3).Parameter = "FORGET", _
                "Toggle " & names(testIndex) & ": disabled unsupported entry, no re-apply all", passed, failed
            AF_HandleClick "1"
            AF_CheckTest Not ws.AutoFilter.Filters(field).On And ws.AutoFilter.Filters(12).On _
                And mRemembered.Count = 1, _
                "Toggle " & names(testIndex) & ": unsupported dispatch keeps cleared snapshot", passed, failed
        Else
            AF_CheckTest popup.Controls.Count = 3 And popup.Controls(1).Caption = _
                "Col " & target.ColumnLetter & " - " & target.HeaderText & " (off)" _
                And popup.Controls(1).State = msoButtonUp And popup.Controls(1).FaceId = 0 _
                And popup.Controls(1).Enabled And popup.Controls(2).State = msoButtonDown _
                And popup.Controls(2).FaceId = 0 And popup.Controls(3).Parameter = "FORGET" _
                And popup.Controls(3).BeginGroup, _
                "Toggle " & names(testIndex) & ": off/active states and conditional bottom items", passed, failed
            popup.Controls(1).Execute
            afterRows = AF_HiddenRows(ws, 121)
            Set targets = AF_GetActiveFilterTargets(ws)
            Set target = AF_FindTarget(targets, "", field)
            Set target = AF_Snapshot(target, ws.AutoFilter.Filters(field))
            actual = AF_FilterSignature(target)
            AF_CheckTest beforeRows = afterRows And expected = actual And mRemembered.Count = 0, _
                "Toggle " & names(testIndex) & ": exact rows/operator/criteria restored", passed, failed, _
                expected & "; hidden=" & beforeRows & "; memory=0", _
                actual & "; hidden=" & afterRows & "; memory=" & CStr(mRemembered.Count)
        End If
        AF_CheckTest Application.ActiveWindow.ScrollRow = scrollRow _
            And Application.ActiveWindow.ScrollColumn = scrollColumn, _
            "Toggle " & names(testIndex) & ": scroll preserved", passed, failed
CaseCleanup:
        AF_TestToggleCleanup ws, failed
    Next testIndex

    On Error GoTo BlockFailed
    For blockIndex = 0 To 9
    mCurrentStep = "initialize toggle block " & CStr(blockIndex)
    Set otherSheet = Nothing
    Set ws = testWorkbook.Worksheets.Add
    ws.Range("A1").Resize(121, 12).Value = data
    ws.Range("C2:C121").NumberFormat = "yyyy-mm-dd"
    ws.Range("A1:L121").AutoFilter
    If blockIndex <= 2 Then
        ws.Range("N1:O1").Value = Array("Table text", "Table number")
        For rowIndex = 2 To 121
            ws.Cells(rowIndex, 14).Value = data(rowIndex, 4)
            ws.Cells(rowIndex, 15).Value = rowIndex - 1
        Next rowIndex
        Set table = ws.ListObjects.Add(xlSrcRange, ws.Range("N1:O121"), , xlYes)
        table.Name = "ToggleTable" & CStr(blockIndex)
        table.Range.AutoFilter Field:=1, Criteria1:=Array("Alpha", "Beta"), Operator:=xlFilterValues
    ElseIf blockIndex >= 6 And blockIndex <= 8 Then
        Set otherSheet = ws
        otherSheet.AutoFilterMode = False
        otherSheet.Range("D1").Value = "ID"
        otherSheet.Range("E1").Value = "Value"
        otherSheet.Range("D2").Value = 1
        otherSheet.Range("E2").Value = "Keep"
        otherSheet.Range("D3").Value = 2
        otherSheet.Range("E3").Value = "Drop"
    End If
    Select Case blockIndex
    Case 0
    mCurrentStep = "toggle table column"
    beforeRows = AF_HiddenRows(ws, 121)
    Set targets = AF_GetActiveFilterTargets(ws)
    Set target = AF_FindTarget(targets, table.Name, 1)
    Set target = AF_Snapshot(target, table.AutoFilter.Filters(1))
    expected = AF_FilterSignature(target)
    Set popup = AF_BuildPopup(ws)
    popup.Controls(1).Execute
    Set popup = AF_BuildPopup(ws)
    AF_CheckTest popup.Controls.Count = 2 And popup.Controls(1).Caption = _
        "Col N - Table text [" & table.Name & "] (off)" And Not table.AutoFilter.Filters(1).On _
        And popup.Controls(2).Parameter = "FORGET", "Toggle Table: off entry and no re-apply all", passed, failed
    popup.Controls(1).Execute
    Set targets = AF_GetActiveFilterTargets(ws)
    Set target = AF_FindTarget(targets, table.Name, 1)
    Set target = AF_Snapshot(target, table.AutoFilter.Filters(1))
    actual = AF_FilterSignature(target)
    AF_CheckTest beforeRows = AF_HiddenRows(ws, 121) And expected = actual And mRemembered.Count = 0, _
        "Toggle Table: exact rows/operator/criteria restored", passed, failed, expected, actual

    Case 1
    mCurrentStep = "toggle table after appending a row"
    beforeRows = AF_HiddenRows(ws, 121)
    Set targets = AF_GetActiveFilterTargets(ws)
    Set target = AF_FindTarget(targets, table.Name, 1)
    Set target = AF_Snapshot(target, table.AutoFilter.Filters(1))
    expected = AF_FilterSignature(target)
    Set popup = AF_BuildPopup(ws)
    popup.Controls(1).Execute
    table.ListRows.Add.Range.Value = Array("Gamma", 121)
    Set popup = AF_BuildPopup(ws)
    AF_CheckTest table.Range.Address <> target.FilterRangeAddress _
        And popup.Controls.Count = 2 And popup.Controls(1).Enabled _
        And popup.Controls(1).Caption = "Col N - Table text [" & table.Name & "] (off)", _
        "Toggle Table growth: off entry survives changed range address", passed, failed
    popup.Controls(1).Execute
    Set targets = AF_GetActiveFilterTargets(ws)
    Set target = AF_FindTarget(targets, table.Name, 1)
    Set target = AF_Snapshot(target, table.AutoFilter.Filters(1))
    actual = AF_FilterSignature(target)
    AF_CheckTest beforeRows = AF_HiddenRows(ws, 121) And ws.Rows(122).Hidden _
        And expected = actual And mRemembered.Count = 0, _
        "Toggle Table growth: criteria restored and appended row filtered", passed, failed, expected, actual

    Case 2
    mCurrentStep = "toggle clear all and re-apply all"
    ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:="Text*"
    ws.AutoFilter.Range.AutoFilter Field:=2, Criteria1:=">=25", Operator:=xlAnd, Criteria2:="<=100"
    ws.AutoFilter.Range.AutoFilter Field:=4, Criteria1:=Array("Alpha", "Beta"), Operator:=xlFilterValues
    beforeRows = AF_HiddenRows(ws, 121)
    Set popup = AF_BuildPopup(ws)
    popup.Controls(5).Execute
    Set popup = AF_BuildPopup(ws)
    actual = vbNullString
    For field = 1 To 4
        actual = actual & popup.Controls(field).Caption & "|"
    Next field
    AF_CheckTest popup.Controls.Count = 6 And mRemembered.Count = 4 _
        And InStr(actual, " (off)|") > 0 And Not ws.FilterMode _
        And popup.Controls(5).Parameter = "REAPPLY" And popup.Controls(5).BeginGroup _
        And popup.Controls(6).Parameter = "FORGET" And Not popup.Controls(6).BeginGroup, _
        "Toggle all: four off entries, re-apply all and forget only", passed, failed, _
        "4 off entries; 6 controls; memory=4", actual & "; controls=" & CStr(popup.Controls.Count)
    popup.Controls(5).Execute
    AF_CheckTest beforeRows = AF_HiddenRows(ws, 121) And mRemembered.Count = 0, _
        "Toggle all: exact hidden rows restored and memory consumed", passed, failed, _
        beforeRows & "; memory=0", AF_HiddenRows(ws, 121) & "; memory=" & CStr(mRemembered.Count)

    Case 3
    mCurrentStep = "toggle unsupported colour"
    AF_ClearAllFilters ws
    AF_ForgetSheet ws
    ws.Range("A2:A121").Interior.Color = RGB(255, 255, 0)
    ws.Range("A2:A20").Interior.Color = RGB(0, 176, 80)
    ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:=RGB(0, 176, 80), Operator:=xlFilterCellColor
    Set popup = AF_BuildPopup(ws)
    popup.Controls(1).Execute
    Set popup = AF_BuildPopup(ws)
    AF_CheckTest Not ws.AutoFilter.Filters(1).On And popup.Controls.Count = 2 _
        And Not popup.Controls(1).Enabled And popup.Controls(1).State = msoButtonUp _
        And popup.Controls(1).Caption = "Col A - Toggle 1 (cannot re-apply)" _
        And popup.Controls(2).Parameter = "FORGET", _
        "Toggle colour: cleared, disabled unsupported entry, no re-apply all", passed, failed

    Case 4
    mCurrentStep = "toggle re-filter invalidation"
    ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:="Text 1*"
    Set popup = AF_BuildPopup(ws)
    popup.Controls(1).Execute
    ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:="Text 2*"
    Set popup = AF_BuildPopup(ws)
    AF_CheckTest mRemembered.Count = 0 And popup.Controls.Count = 1 _
        And popup.Controls(1).State = msoButtonDown, _
        "Toggle memory: re-filtered by code drops old entry", passed, failed
    popup.Controls(1).Execute
    ws.Range("A1").Value = "Changed toggle"
    Set popup = AF_BuildPopup(ws)
    AF_CheckTest mRemembered.Count = 0 And popup.Controls.Count = 1 _
        And Not popup.Controls(1).Enabled, "Toggle memory: header change drops entry", passed, failed

    Case 5
    mCurrentStep = "toggle sheet isolation"
    ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:="Text 1*"
    Set popup = AF_BuildPopup(ws)
    popup.Controls(1).Execute
    Set otherSheet = testWorkbook.Worksheets.Add
    Set popup = AF_BuildPopup(otherSheet)
    AF_CheckTest mRemembered.Count = 1 And popup.Controls.Count = 1 _
        And popup.Controls(1).Caption = "No active filters on this sheet", _
        "Toggle memory: other-sheet entries never appear", passed, failed
    ' Adding a disposable sheet makes it active without Select/Activate.
    AF_ForgetSheet ws
    AF_CheckTest mRemembered.Count = 0, "Toggle memory: forget is sheet scoped", passed, failed
    otherSheet.Range("D1").Value = "ID"
    otherSheet.Range("E1").Value = "Value"
    otherSheet.Range("D2").Value = 1
    otherSheet.Range("E2").Value = "Keep"
    otherSheet.Range("D3").Value = 2
    otherSheet.Range("E3").Value = "Drop"
    otherSheet.Range("D1:E3").AutoFilter Field:=2, Criteria1:="Keep"
    Set popup = AF_BuildPopup(otherSheet)
    popup.Controls(1).Execute
    Set popup = AF_BuildPopup(otherSheet)
    popup.Controls(2).Execute
    Set popup = AF_BuildPopup(otherSheet)
    AF_CheckTest mRemembered.Count = 0 And popup.Controls.Count = 1 _
        And popup.Controls(1).Caption = "No active filters on this sheet", _
        "Toggle memory: Forget callback empties list", passed, failed

    Case 6
    mCurrentStep = "toggle missing source"
    otherSheet.Range("D1:E3").AutoFilter Field:=2, Criteria1:="Keep"
    Set popup = AF_BuildPopup(otherSheet)
    popup.Controls(1).Execute
    otherSheet.AutoFilterMode = False
    Set popup = AF_BuildPopup(otherSheet)
    AF_CheckTest mRemembered.Count = 0, "Toggle memory: removed range drops entry", passed, failed
    Set table = otherSheet.ListObjects.Add(xlSrcRange, otherSheet.Range("D1:E3"), , xlYes)
    table.Range.AutoFilter Field:=2, Criteria1:="Keep"
    Set popup = AF_BuildPopup(otherSheet)
    popup.Controls(1).Execute
    table.Unlist
    Set popup = AF_BuildPopup(otherSheet)
    AF_CheckTest mRemembered.Count = 0, "Toggle memory: removed table drops entry", passed, failed

    Case 7
    mCurrentStep = "toggle failed re-apply retains snapshot"
    otherSheet.AutoFilterMode = False
    otherSheet.Range("D1:E3").AutoFilter Field:=2, Criteria1:="Keep"
    Set popup = AF_BuildPopup(otherSheet)
    popup.Controls(1).Execute
    otherSheet.Protect
    Set popup = AF_BuildPopup(otherSheet)
    popup.Controls(1).Execute
    AF_CheckTest mRemembered.Count = 1 And Not otherSheet.AutoFilter.Filters(2).On, _
        "Toggle failure: protected re-apply keeps entry", passed, failed
    otherSheet.Unprotect
    Set popup = AF_BuildPopup(otherSheet)
    popup.Controls(1).Execute
    AF_CheckTest mRemembered.Count = 0 And otherSheet.AutoFilter.Filters(2).On, _
        "Toggle failure: retry restores filter", passed, failed

    Case 8
    mCurrentStep = "optional callback via Application.Run without arguments"
    otherSheet.Range("D1:E3").AutoFilter Field:=2, Criteria1:="Keep"
    Set popup = AF_BuildPopup(otherSheet)
    Application.Run AF_Callback()
    AF_CheckTest otherSheet.AutoFilter.Filters(2).On, _
        "Macros: Optional OnClick runs without arguments", passed, failed
    mCurrentStep = "optional self-test via Application.Run without arguments"
    mTestingOptional = True
    mOptionalSelfTestCalled = False
    Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!ActiveFilters_SelfTest"
    mTestingOptional = False
    AF_CheckTest mOptionalSelfTestCalled, _
        "Macros: Optional SelfTest runs without arguments", passed, failed
    Case 9
    mCurrentStep = "forced snapshot failure still clears filter"
    ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:="Text 1*"
    ws.AutoFilter.Range.AutoFilter Field:=12, Criteria1:="Keep"
    Set popup = AF_BuildPopup(ws)
    mForceSnapshotFailure = True
    popup.Controls(1).Execute
    mForceSnapshotFailure = False
    Set popup = AF_BuildPopup(ws)
    AF_CheckTest Not ws.AutoFilter.Filters(1).On And ws.AutoFilter.Filters(12).On _
        And mRemembered.Count = 1 And popup.Controls.Count = 3 _
        And popup.Controls(1).Caption = "Col A - Toggle 1 (cannot re-apply)" _
        And Not popup.Controls(1).Enabled And popup.Controls(1).State = msoButtonUp _
        And Not mRemembered(1).CanReapply, _
        "Toggle snapshot failure: target cleared and disabled entry remembered", passed, failed
    End Select
BlockCleanup:
    mForceSnapshotFailure = False
    mTestingOptional = False
    AF_TestToggleCleanup ws, failed
    If Not otherSheet Is Nothing Then AF_TestToggleCleanup otherSheet, failed
    Next blockIndex
    Exit Sub
CaseFailed:
    failed = failed + 1
    AF_ReportLine "FAIL: toggle case runtime error " & CStr(Err.Number) & " (" & Err.Description _
        & ") at step: " & mCurrentStep
    Resume CaseCleanup
BlockFailed:
    failed = failed + 1
    AF_ReportLine "FAIL: toggle block runtime error " & CStr(Err.Number) & " (" & Err.Description _
        & ") at step: " & mCurrentStep
    Resume BlockCleanup
Failed:
    failed = failed + 1
    AF_ReportLine "FAIL: toggle runtime error " & CStr(Err.Number) & " (" & Err.Description _
        & ") at step: " & mCurrentStep & "; expected: test completes; actual: runtime error"
End Sub

Private Sub AF_TestToggleCleanup(ws As Worksheet, failed As Long)
    Dim table As ListObject

    On Error GoTo Failed
    AF_DeletePopup
    If ws Is Nothing Then Exit Sub
    If ws.ProtectContents Then ws.Unprotect
    For Each table In ws.ListObjects
        If AF_HasActiveFilters(table.AutoFilter) Then table.AutoFilter.ShowAllData
    Next table
    If ws.FilterMode Then ws.ShowAllData
    AF_ForgetSheet ws
    Exit Sub
Failed:
    failed = failed + 1
    AF_ReportLine "FAIL: toggle cleanup runtime error " & CStr(Err.Number) & " (" & Err.Description _
        & ") at step: " & mCurrentStep
End Sub

Private Function AF_HiddenRows(ws As Worksheet, lastRow As Long) As String
    Dim rowIndex As Long

    For rowIndex = 2 To lastRow
        If ws.Rows(rowIndex).Hidden Then
            AF_HiddenRows = AF_HiddenRows & "1"
        Else
            AF_HiddenRows = AF_HiddenRows & "0"
        End If
    Next rowIndex
End Function

Private Function AF_FilterSignature(t As AF_FilterTarget) As String
    AF_FilterSignature = "Operator=" & CStr(t.FilterOperator) _
        & "; C1 present=" & CStr(t.HasCriteria1) & "; C2 present=" & CStr(t.HasCriteria2)
    If t.HasCriteria1 Then AF_FilterSignature = AF_FilterSignature & "; C1=" & AF_CriterionSignature(t.Criteria1)
    If t.HasCriteria2 Then AF_FilterSignature = AF_FilterSignature & "; C2=" & AF_CriterionSignature(t.Criteria2)
End Function

Private Function AF_CriterionSignature(value As Variant) As String
    Dim index As Long
    Dim item As String

    If IsArray(value) Then
        AF_CriterionSignature = "array(" & CStr(LBound(value)) & "," & CStr(UBound(value)) & ")"
        For index = LBound(value) To UBound(value)
            item = AF_CriterionSignature(value(index))
            AF_CriterionSignature = AF_CriterionSignature & "[" & CStr(Len(item)) & ":" & item & "]"
        Next index
    Else
        AF_CriterionSignature = CStr(VarType(value)) & ":" & CStr(value)
    End If
End Function

Private Sub AF_ReportLine(text As String)
    Dim fileNumber As Integer
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo TraceFailed
    mReport = mReport & text & vbCrLf
    Debug.Print text
    If Len(mTracePath) > 0 Then
        fileNumber = FreeFile
        Open mTracePath For Append As #fileNumber
        Print #fileNumber, mCurrentStep
        Close #fileNumber
    End If
    Exit Sub
TraceFailed:
    errorNumber = Err.Number
    errorText = Err.Description
    mTracePath = vbNullString
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
    Err.Raise errorNumber, "AF_ReportLine", "Could not write self-test trace: " & errorText
End Sub

Private Sub AF_CheckTest(condition As Boolean, label As String, passed As Long, failed As Long, _
    Optional expected As String = "True", Optional actual As String = "False")
    If condition Then
        passed = passed + 1
        AF_ReportLine "PASS: " & label
    Else
        failed = failed + 1
        AF_ReportLine "FAIL: " & label & "; expected: " & expected & "; actual: " & actual
    End If
End Sub

Private Function AF_FindTarget(targets As Collection, tableName As String, _
    fieldIndex As Long) As AF_FilterTarget

    Dim target As AF_FilterTarget

    For Each target In targets
        If target.TableName = tableName And target.FieldIndex = fieldIndex Then
            Set AF_FindTarget = target
            Exit Function
        End If
    Next target
    Err.Raise vbObjectError + 2002, , "Expected target not found."
End Function

Private Function AF_TargetMatches(targets As Collection, targetIndex As Long, _
    letter As String, header As String, tableName As String, fieldIndex As Long, _
    sheetColumn As Long, ws As Worksheet) As Boolean

    Dim target As AF_FilterTarget
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
    AF_TargetMatches = target.ColumnLetter = letter And target.HeaderText = header _
        And target.TableName = tableName And target.FieldIndex = fieldIndex _
        And target.SheetColumn = sheetColumn And target.WorkbookName = ws.Parent.Name _
        And target.SheetName = ws.Name And target.SourceKind = expectedKind _
        And target.FilterRangeAddress = expectedAddress
End Function
