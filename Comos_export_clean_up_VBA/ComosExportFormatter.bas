Attribute VB_Name = "ComosExportFormatter"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function UrlUnescapeW Lib "shlwapi.dll" ( _
    ByVal source As LongPtr, ByVal destination As LongPtr, _
    ByRef characterCount As Long, ByVal flags As Long) As Long
#Else
Private Declare Function UrlUnescapeW Lib "shlwapi.dll" ( _
    ByVal source As Long, ByVal destination As Long, _
    ByRef characterCount As Long, ByVal flags As Long) As Long
#End If

Public Sub FormatComosExport()
    Dim summary As String

    On Error GoTo Failed
    summary = FormatComosWorkbook(Application.ActiveWorkbook)
    MsgBox summary, vbInformation, "COMOS export"
    Exit Sub

Failed:
    MsgBox Err.Description, vbExclamation, "COMOS export"
End Sub

Public Function FormatComosWorkbook(ByVal target As Workbook) As String
    Dim query As Worksheet
    Dim sheet As Worksheet
    Dim header As Range
    Dim data As Range
    Dim column As Range
    Dim dataRow As Range
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim outputPath As String
    Dim converted As Long
    Dim skipped As Long
    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim settingsCaptured As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim currentStep As String
    Dim outputCreated As Boolean

    On Error GoTo Failed
    currentStep = "Checking the source workbook"
    If target Is Nothing Then
        Err.Raise vbObjectError + 2100, , "Open a COMOS export workbook first."
    End If
    If target Is ThisWorkbook Or target.IsAddin Then
        Err.Raise vbObjectError + 2101, , "Select the export, not the macro workbook."
    End If
    If Len(target.Path) = 0 Then
        Err.Raise vbObjectError + 2102, , "Save the source workbook before formatting it."
    End If
    currentStep = "Checking the output path"
    outputPath = FormattedPath(target)
    If CreateObject("Scripting.FileSystemObject").FileExists(outputPath) Then
        Err.Raise vbObjectError + 2103, , "Output already exists; rename or move it first:" & vbCrLf & outputPath
    End If

    currentStep = "Checking the Query worksheet"
    On Error Resume Next
    Set query = target.Worksheets("Query")
    On Error GoTo Failed
    If query Is Nothing Then
        Err.Raise vbObjectError + 2104, , "The workbook has no Query worksheet."
    End If

    headerRow = 1
    If query.Rows(1).Hidden Then headerRow = 2
    Set header = query.Rows(headerRow).Find(What:="Line No.", _
        After:=query.Cells(headerRow, query.Columns.Count), LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=True, MatchByte:=False, SearchFormat:=False)
    If header Is Nothing Then
        Err.Raise vbObjectError + 2105, , "Line No. was not found in the expected header row. No columns were deleted."
    End If

    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    currentStep = "Creating the formatted copy in the source folder"
    target.SaveAs Filename:=outputPath, FileFormat:=target.FileFormat, _
        CreateBackup:=False, AddToMru:=False
    outputCreated = True

    currentStep = "Removing worksheet protection"
    For Each sheet In target.Worksheets
        sheet.Unprotect Password:=vbNullString
        If sheet.ProtectContents Or sheet.ProtectDrawingObjects Or sheet.ProtectScenarios Then
            Err.Raise vbObjectError + 2106, , "Cannot remove protection from sheet: " & sheet.Name
        End If
    Next sheet

    currentStep = "Formatting the Query worksheet"
    target.Activate
    query.Activate
    If query.AutoFilterMode Then query.AutoFilterMode = False
    If header.Column > 1 Then
        query.Range(query.Columns(1), query.Columns(header.Column - 1)).EntireColumn.Delete
    End If

    Set data = ContentRange(query)
    lastColumn = data.Columns.Count
    Do While lastColumn > 0
        If Not query.Columns(lastColumn).Hidden Then Exit Do
        query.Columns(lastColumn).Delete
        lastColumn = lastColumn - 1
    Loop
    If headerRow = 2 Then query.Rows(1).Delete
    If query.Cells(1, 1).Value2 <> "Line No." Then
        Err.Raise vbObjectError + 2112, , "The header is not Line No. in A1 after removing the technical row."
    End If
    query.Rows(1).Hidden = False

    With query.Cells
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With

    Set data = ContentRange(query)
    ConvertNumericText data, converted, skipped
    query.Rows(1).Font.Bold = True
    data.AutoFilter

    With data.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    data.Borders(xlDiagonalDown).LineStyle = xlNone
    data.Borders(xlDiagonalUp).LineStyle = xlNone

    For Each column In data.Columns
        If Not column.Hidden Then column.AutoFit
    Next column
    For Each dataRow In data.Rows
        If Not dataRow.Hidden Then dataRow.AutoFit
    Next dataRow
    currentStep = "Freezing the top row"
    With target.Windows(1)
        .Activate
        query.Activate
        .View = xlNormalView
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 0
        query.Rows(1).Hidden = False
        query.Rows(1).AutoFit
        .ScrollRow = 1
        .ScrollColumn = 1
        query.Range("A1").Select
        .SplitRow = 1
        .FreezePanes = True
    End With

    currentStep = "Saving the formatted workbook"
    Application.Calculation = oldCalculation
    target.Save

    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    FormatComosWorkbook = "Saved: " & outputPath & vbCrLf & _
        "Converted text numbers: " & CStr(converted) & vbCrLf & _
        "Numeric texts retained for precision/range limits: " & CStr(skipped)
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorDescription = "Step: " & currentStep & vbCrLf & _
        "Error " & CStr(errorNumber) & ": " & errorDescription
    If Len(outputPath) > 0 Then
        errorDescription = errorDescription & vbCrLf & "Output: " & outputPath
    End If
    On Error Resume Next
    If settingsCaptured Then
        Application.Calculation = oldCalculation
        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating
    End If
    On Error GoTo 0
    If outputCreated Then
        errorDescription = errorDescription & vbCrLf & _
            "The original was not formatted. The _formatted copy may be incomplete; close it and delete it before retrying."
    Else
        errorDescription = errorDescription & vbCrLf & _
            "Formatting did not start. No source data was changed by the macro."
    End If
    Err.Raise errorNumber, "ComosExportFormatter", errorDescription
End Function

Private Function FormattedPath(ByVal target As Workbook) As String
    Dim dotPosition As Long
    Dim extension As String

    dotPosition = InStrRev(target.Name, ".")
    If dotPosition = 0 Then
        Err.Raise vbObjectError + 2107, , "The source needs an Excel file extension."
    End If
    extension = Mid$(target.Name, dotPosition)
    Select Case LCase$(extension)
        Case ".xlsx", ".xlsm", ".xlsb", ".xls"
        Case Else
            Err.Raise vbObjectError + 2108, , "Only Excel workbook formats are supported."
    End Select
    FormattedPath = LocalWorkbookFolder(target.Path) & Application.PathSeparator & _
        Left$(target.Name, dotPosition - 1) & "_formatted_" & Format$(Now, "yymmdd_hhnn") & extension
End Function

Private Function LocalWorkbookFolder(ByVal sourcePath As String) As String
    Const currentUser As Long = &H80000001
    Const providerKey As String = "Software\SyncEngines\Providers\OneDrive"
    Dim registry As Object
    Dim fileSystem As Object
    Dim subkeys As Variant
    Dim subkey As Variant
    Dim namespace As Variant
    Dim mountPoint As Variant
    Dim sourceUrl As String
    Dim rootUrl As String
    Dim relativePath As String
    Dim candidate As String
    Dim bestMatchLength As Long

    If InStr(1, sourcePath, "://", vbTextCompare) = 0 Then
        LocalWorkbookFolder = sourcePath
        Exit Function
    End If

    sourceUrl = DecodeUrlPath(sourcePath)
    If Right$(sourceUrl, 1) <> "/" Then sourceUrl = sourceUrl & "/"
    Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If registry.EnumKey(currentUser, providerKey, subkeys) = 0 Then
        If IsArray(subkeys) Then
            For Each subkey In subkeys
                namespace = Empty
                mountPoint = Empty
                registry.GetStringValue currentUser, providerKey & "\" & CStr(subkey), "UrlNamespace", namespace
                registry.GetStringValue currentUser, providerKey & "\" & CStr(subkey), "MountPoint", mountPoint
                If VarType(namespace) = vbString And VarType(mountPoint) = vbString Then
                    If Len(namespace) > 0 And Len(mountPoint) > 0 Then
                        rootUrl = DecodeUrlPath(CStr(namespace))
                        If Right$(rootUrl, 1) <> "/" Then rootUrl = rootUrl & "/"
                        If Len(rootUrl) > bestMatchLength Then
                            If StrComp(Left$(sourceUrl, Len(rootUrl)), rootUrl, vbTextCompare) = 0 Then
                                relativePath = Replace(Mid$(sourceUrl, Len(rootUrl) + 1), "/", "\")
                                candidate = CStr(mountPoint)
                                If Right$(candidate, 1) <> "\" Then candidate = candidate & "\"
                                candidate = candidate & relativePath
                                If fileSystem.FolderExists(candidate) Then
                                    LocalWorkbookFolder = candidate
                                    bestMatchLength = Len(rootUrl)
                                End If
                            End If
                        End If
                    End If
                End If
            Next subkey
        End If
    End If

    If Len(LocalWorkbookFolder) = 0 Then
        Err.Raise vbObjectError + 2110, , _
            "No matching OneDrive sync folder was found for:" & vbCrLf & sourcePath & vbCrLf & _
            "Enable OneDrive sync for this folder, then run the macro again. No manual file copy is needed."
    End If
    Do While Right$(LocalWorkbookFolder, 1) = "\"
        LocalWorkbookFolder = Left$(LocalWorkbookFolder, Len(LocalWorkbookFolder) - 1)
    Loop
End Function

Private Function DecodeUrlPath(ByVal encodedPath As String) As String
    Dim buffer As String
    Dim characterCount As Long

    characterCount = Len(encodedPath) + 1
    buffer = String$(characterCount, vbNullChar)
    If UrlUnescapeW(StrPtr(encodedPath), StrPtr(buffer), characterCount, &H40000) <> 0 Then
        Err.Raise vbObjectError + 2111, , "Cannot decode the OneDrive folder address: " & encodedPath
    End If
    DecodeUrlPath = Left$(buffer, characterCount)
End Function

Private Function ContentRange(ByVal sheet As Worksheet) As Range
    Dim lastRowCell As Range
    Dim lastColumnCell As Range

    Set lastRowCell = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    Set lastColumnCell = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    If lastRowCell Is Nothing Or lastColumnCell Is Nothing Then
        Err.Raise vbObjectError + 2109, , "The Query worksheet is empty."
    End If
    Set ContentRange = sheet.Range(sheet.Cells(1, 1), _
        sheet.Cells(lastRowCell.Row, lastColumnCell.Column))
End Function

Private Sub ConvertNumericText(ByVal data As Range, ByRef converted As Long, ByRef skipped As Long)
    Dim texts As Range
    Dim cell As Range
    Dim pattern As Object
    Dim value As Double
    Dim decimals As Long
    Dim numericText As String
    Dim numberFormat As String

    On Error Resume Next
    Set texts = data.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If texts Is Nothing Then Exit Sub

    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Pattern = "^[+-]?[0-9]+(\.[0-9]+)?$"
    pattern.Global = False

    For Each cell In texts.Cells
        numericText = TrimNumericWhitespace(CStr(cell.Value2))
        If pattern.Test(numericText) Then
            If TryNumericValue(numericText, value, decimals) Then
                numberFormat = "0"
                If decimals > 0 Then numberFormat = numberFormat & "." & String$(decimals, "0")
                cell.NumberFormat = numberFormat
                cell.Value2 = value
                converted = converted + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next cell
End Sub

Private Function TryNumericValue(ByVal text As String, ByRef value As Double, _
    ByRef decimals As Long) As Boolean
    Dim dotPosition As Long
    Dim digits As String

    On Error GoTo NotRepresentable
    decimals = 0
    dotPosition = InStr(1, text, ".", vbBinaryCompare)
    If dotPosition > 0 Then decimals = Len(text) - dotPosition
    If decimals > 253 Then Exit Function

    digits = Replace(Replace(Replace(text, ".", ""), "+", ""), "-", "")
    Do While Len(digits) > 0
        If Left$(digits, 1) <> "0" Then Exit Do
        digits = Mid$(digits, 2)
    Loop
    If Len(digits) > 15 Then Exit Function

    value = Val(text)
    If value = 0 And Len(digits) > 0 Then Exit Function
    TryNumericValue = True
    Exit Function

NotRepresentable:
    TryNumericValue = False
End Function

Private Function TrimNumericWhitespace(ByVal text As String) As String
    Dim first As Long
    Dim last As Long

    first = 1
    last = Len(text)
    Do While first <= last
        If Not IsNumericWhitespace(Mid$(text, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not IsNumericWhitespace(Mid$(text, last, 1)) Then Exit Do
        last = last - 1
    Loop
    TrimNumericWhitespace = Mid$(text, first, last - first + 1)
End Function

Private Function IsNumericWhitespace(ByVal character As String) As Boolean
    Select Case AscW(character)
        Case 9 To 13, 28 To 32, 133, 160, 5760, 8192 To 8202, 8232, 8233, 8239, 8287, 12288
            IsNumericWhitespace = True
    End Select
End Function