Option Explicit

' =============================
' Configuration
' =============================
Private Const OUTPUT_FOLDER As String = "C:\temp\"
Private Const FILE_PREFIX As String = "CommonMeetingSlots_KVI"
Private Const CSV_DELIMITER As String = ","   ' Comma delimiter as requested.

Private Const SLOT_GRANULARITY_MIN As Long = 30
Private Const EARLIEST_START_HOUR As Long = 8
Private Const LUNCH_START_HOUR As Long = 12
Private Const LUNCH_END_HOUR As Long = 14
Private Const LATEST_ALLOWED_START_HOUR As Long = 16

Private mIndicatorInitialized As Boolean
Private mOriginalStatusBar As Variant

' =============================
' Entry Point
' =============================
Public Sub ExportCommonMeetingSlotsToCsv()
    On Error GoTo EH

    StartRunningIndicator "Initializing"

    Dim queryTimestamp As Date
    queryTimestamp = Now

    Dim startDate As Date
    Dim endDate As Date
    startDate = Date
    endDate = DateSerial(2026, 6, 23)

    If DateValue(startDate) > DateValue(endDate) Then
        StopRunningIndicator
        MsgBox "Start date is later than end date. Nothing to process.", vbExclamation, "Common Meeting Slots"
        Exit Sub
    End If

    Dim earliestStart As Date
    Dim lunchStart As Date
    Dim lunchEnd As Date
    Dim latestAllowedStart As Date

    earliestStart = TimeSerial(EARLIEST_START_HOUR, 0, 0)
    lunchStart = TimeSerial(LUNCH_START_HOUR, 0, 0)
    lunchEnd = TimeSerial(LUNCH_END_HOUR, 0, 0)
    latestAllowedStart = TimeSerial(LATEST_ALLOWED_START_HOUR, 0, 0)

    Dim requiredAttendees As Variant
    requiredAttendees = BuildRequiredAttendees()

    Dim optionalAttendee As Variant
    optionalAttendee = BuildOptionalAttendee()

    Dim ns As Outlook.NameSpace
    Set ns = Application.Session

    Dim requiredMaps As Collection
    Set requiredMaps = New Collection

    Dim requiredCount As Long
    requiredCount = UBound(requiredAttendees, 1) - LBound(requiredAttendees, 1) + 1

    Dim i As Long
    For i = LBound(requiredAttendees, 1) To UBound(requiredAttendees, 1)
        UpdateRunningIndicator "Reading required free/busy", i - LBound(requiredAttendees, 1) + 1, requiredCount

        Dim reqRecipient As Outlook.Recipient
        Set reqRecipient = ResolveRecipient(ns, CStr(requiredAttendees(i, 2)))

        If reqRecipient Is Nothing Then
            StopRunningIndicator
            MsgBox "Required attendee could not be resolved: " & CStr(requiredAttendees(i, 1)) & " <" & CStr(requiredAttendees(i, 2)) & ">" & vbCrLf & _
                   "No slots were exported to avoid false positives.", vbCritical, "Common Meeting Slots"
            Exit Sub
        End If

        Dim reqMap As Object
        Dim reqErr As String
        If Not TryBuildFreeBusyMap(reqRecipient, startDate, endDate, SLOT_GRANULARITY_MIN, reqMap, reqErr) Then
            StopRunningIndicator
            MsgBox "Free/busy could not be read for required attendee: " & CStr(requiredAttendees(i, 1)) & " <" & CStr(requiredAttendees(i, 2)) & ">" & vbCrLf & _
                   reqErr & vbCrLf & _
                   "No slots were exported to avoid false positives.", vbCritical, "Common Meeting Slots"
            Exit Sub
        End If

        requiredMaps.Add reqMap
    Next i

    Dim optionalName As String
    Dim optionalEmail As String
    optionalName = CStr(optionalAttendee(1))
    optionalEmail = CStr(optionalAttendee(2))

    Dim optionalMap As Object
    Dim optionalAvailable As Boolean
    optionalAvailable = False

    Dim optionalWarning As String
    optionalWarning = ""

    Dim optRecipient As Outlook.Recipient
    Set optRecipient = ResolveRecipient(ns, optionalEmail)

    UpdateRunningIndicator "Reading optional free/busy", 1, 1

    If optRecipient Is Nothing Then
        optionalWarning = "Optional attendee could not be resolved. OptionalAttendeeFree is FALSE for all exported rows."
    Else
        Dim optErr As String
        If TryBuildFreeBusyMap(optRecipient, startDate, endDate, SLOT_GRANULARITY_MIN, optionalMap, optErr) Then
            optionalAvailable = True
        Else
            optionalWarning = "Optional attendee free/busy could not be read. OptionalAttendeeFree is FALSE for all exported rows. " & optErr
        End If
    End If

    Dim durations As Variant
    durations = Array(60, 30)  ' Preferred first, fallback second.

    Dim queryStartDateTime As Date
    queryStartDateTime = Now

    UpdateRunningIndicator "Finding common slots", 0, DateDiff("d", DateValue(startDate), DateValue(endDate)) + 1

    Dim slots As Collection
    Set slots = FindCommonSlots( _
        requiredMaps:=requiredMaps, _
        optionalMap:=optionalMap, _
        optionalAvailable:=optionalAvailable, _
        optionalName:=optionalName, _
        queryTimestamp:=queryTimestamp, _
        queryStartDateTime:=queryStartDateTime, _
        startDate:=startDate, _
        endDate:=endDate, _
        earliestStart:=earliestStart, _
        lunchStart:=lunchStart, _
        lunchEnd:=lunchEnd, _
        latestAllowedStart:=latestAllowedStart, _
        durations:=durations _
    )

    Dim csvPath As String
    Dim writeErr As String

    UpdateRunningIndicator "Writing CSV", 1, 1

    If Not WriteSlotsToCsv(slots, queryTimestamp, csvPath, writeErr) Then
        StopRunningIndicator
        MsgBox "Failed to write CSV file: " & writeErr, vbCritical, "Common Meeting Slots"
        Exit Sub
    End If

    Dim msg As String
    msg = "CSV export completed." & vbCrLf & _
          "File: " & csvPath & vbCrLf & _
          "Candidate slots exported: " & CStr(slots.Count)

    If Len(optionalWarning) > 0 Then
        msg = msg & vbCrLf & vbCrLf & optionalWarning
    End If

    StopRunningIndicator
    MsgBox msg, vbInformation, "Common Meeting Slots"
    Exit Sub

EH:
    StopRunningIndicator
    MsgBox "Unexpected error " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "Common Meeting Slots"
End Sub

' =============================
' Attendee Definitions
' =============================
Private Function BuildRequiredAttendees() As Variant
    Dim attendees(1 To 5, 1 To 2) As String

    attendees(1, 1) = "Sziva, Laszlo"
    attendees(1, 2) = "Laszlo.Sziva@kanadevia-inova.com"

    attendees(2, 1) = "Maties, Oana"
    attendees(2, 2) = "Oana.Maties@kanadevia-inova.com"

    attendees(3, 1) = "Chavlisek, Georgios"
    attendees(3, 2) = "Georgios.Chavlisek@kanadevia-inova.com"

    attendees(4, 1) = "Zelem, David"
    attendees(4, 2) = "David.Zelem@kanadevia-inova.com"

    attendees(5, 1) = "Balla, Andrej"
    attendees(5, 2) = "Andrej.Balla@kanadevia-inova.com"

    BuildRequiredAttendees = attendees
End Function

Private Function BuildOptionalAttendee() As Variant
    Dim attendee(1 To 2) As String
    attendee(1) = "Lorenzoni, Enrico"
    attendee(2) = "Enrico.Lorenzoni@kanadevia-inova.com"
    BuildOptionalAttendee = attendee
End Function

' =============================
' Free/Busy Acquisition
' =============================
Private Function ResolveRecipient(ByVal ns As Outlook.NameSpace, ByVal emailAddress As String) As Outlook.Recipient
    On Error GoTo EH

    Dim r As Outlook.Recipient
    Set r = ns.CreateRecipient(emailAddress)

    If r Is Nothing Then
        Set ResolveRecipient = Nothing
        Exit Function
    End If

    r.Resolve
    If r.Resolved Then
        Set ResolveRecipient = r
    Else
        Set ResolveRecipient = Nothing
    End If
    Exit Function

EH:
    Set ResolveRecipient = Nothing
End Function

Private Function TryBuildFreeBusyMap( _
    ByVal recipient As Outlook.Recipient, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByVal intervalMinutes As Long, _
    ByRef dayMap As Object, _
    ByRef errorText As String _
) As Boolean

    On Error GoTo EH

    Dim intervalsPerDay As Long
    intervalsPerDay = (24 * 60) \ intervalMinutes

    Set dayMap = CreateObject("Scripting.Dictionary")

    Dim d As Date
    For d = DateValue(startDate) To DateValue(endDate)
        Dim dayStart As Date
        dayStart = DateSerial(Year(d), Month(d), Day(d))

        Dim freeBusyRaw As String
        freeBusyRaw = recipient.FreeBusy(dayStart, intervalMinutes, True)

        Dim normalized As String
        normalized = NormalizeDailyFreeBusy(freeBusyRaw, intervalsPerDay)

        If Len(normalized) <> intervalsPerDay Then
            errorText = "Returned free/busy data is incomplete for " & Format$(dayStart, "yyyy-mm-dd") & "."
            Set dayMap = Nothing
            TryBuildFreeBusyMap = False
            Exit Function
        End If

        dayMap(Format$(dayStart, "yyyymmdd")) = normalized
    Next d

    TryBuildFreeBusyMap = True
    Exit Function

EH:
    errorText = "Error " & CStr(Err.Number) & ": " & Err.Description
    Set dayMap = Nothing
    TryBuildFreeBusyMap = False
End Function

Private Function NormalizeDailyFreeBusy(ByVal freeBusyRaw As String, ByVal intervalsPerDay As Long) As String
    Dim cleaned As String
    cleaned = Trim$(freeBusyRaw)

    If Len(cleaned) = 0 Then
        NormalizeDailyFreeBusy = ""
        Exit Function
    End If

    If Len(cleaned) >= intervalsPerDay Then
        NormalizeDailyFreeBusy = Left$(cleaned, intervalsPerDay)
    Else
        ' If Outlook returns less data than expected, pad with a non-zero marker to stay conservative.
        NormalizeDailyFreeBusy = cleaned & String$(intervalsPerDay - Len(cleaned), "9")
    End If
End Function

' =============================
' Slot Search
' =============================
Private Function FindCommonSlots( _
    ByVal requiredMaps As Collection, _
    ByVal optionalMap As Object, _
    ByVal optionalAvailable As Boolean, _
    ByVal optionalName As String, _
    ByVal queryTimestamp As Date, _
    ByVal queryStartDateTime As Date, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByVal earliestStart As Date, _
    ByVal lunchStart As Date, _
    ByVal lunchEnd As Date, _
    ByVal latestAllowedStart As Date, _
    ByVal durations As Variant _
) As Collection

    Dim results As New Collection

    Dim totalDays As Long
    totalDays = DateDiff("d", DateValue(startDate), DateValue(endDate)) + 1

    Dim dayCounter As Long
    dayCounter = 0

    Dim d As Date
    For d = DateValue(startDate) To DateValue(endDate)
        dayCounter = dayCounter + 1
        UpdateRunningIndicator "Finding common slots", dayCounter, totalDays

        Dim dayStart As Date
        dayStart = DateSerial(Year(d), Month(d), Day(d))

        Dim minuteOffset As Long
        For minuteOffset = 0 To (24 * 60 - SLOT_GRANULARITY_MIN) Step SLOT_GRANULARITY_MIN
            Dim slotStart As Date
            slotStart = DateAdd("n", minuteOffset, dayStart)

            If slotStart >= queryStartDateTime Then
                Dim durationItem As Variant
                For Each durationItem In durations
                    Dim durationMinutes As Long
                    durationMinutes = CLng(durationItem)

                    If IsSlotAllowedByTimeRules(slotStart, durationMinutes, earliestStart, lunchStart, lunchEnd, latestAllowedStart) Then
                        If AreAllRequiredAttendeesFree(requiredMaps, slotStart, durationMinutes, SLOT_GRANULARITY_MIN) Then
                            Dim optionalFree As Boolean
                            optionalFree = False

                            If optionalAvailable Then
                                optionalFree = IsAttendeeFreeForSlot(optionalMap, slotStart, durationMinutes, SLOT_GRANULARITY_MIN)
                            End If

                            Dim noteText As String
                            noteText = BuildNote(durationMinutes, optionalFree)

                            Dim rowData As Variant
                            rowData = BuildRow(queryTimestamp, slotStart, durationMinutes, optionalFree, optionalName, noteText)
                            results.Add rowData
                        End If
                    End If
                Next durationItem
            End If
        Next minuteOffset
    Next d

    Set FindCommonSlots = results
End Function

Private Function IsSlotAllowedByTimeRules( _
    ByVal slotStart As Date, _
    ByVal durationMinutes As Long, _
    ByVal earliestStart As Date, _
    ByVal lunchStart As Date, _
    ByVal lunchEnd As Date, _
    ByVal latestAllowedStart As Date _
) As Boolean

    IsSlotAllowedByTimeRules = False

    ' Weekend is not allowed (Saturday/Sunday).
    If Weekday(DateValue(slotStart), vbMonday) > 5 Then Exit Function

    Dim t As Date
    t = TimeValue(slotStart)

    If t < earliestStart Then Exit Function
    If t >= latestAllowedStart Then Exit Function

    Dim slotEnd As Date
    slotEnd = DateAdd("n", durationMinutes, slotStart)

    Dim dayBase As Date
    dayBase = DateSerial(Year(slotStart), Month(slotStart), Day(slotStart))

    Dim lunchWindowStart As Date
    Dim lunchWindowEnd As Date
    lunchWindowStart = dayBase + lunchStart
    lunchWindowEnd = dayBase + lunchEnd

    If TimeRangesOverlap(slotStart, slotEnd, lunchWindowStart, lunchWindowEnd) Then Exit Function

    IsSlotAllowedByTimeRules = True
End Function

Private Function AreAllRequiredAttendeesFree( _
    ByVal requiredMaps As Collection, _
    ByVal slotStart As Date, _
    ByVal durationMinutes As Long, _
    ByVal intervalMinutes As Long _
) As Boolean

    AreAllRequiredAttendeesFree = False

    Dim mapItem As Variant
    For Each mapItem In requiredMaps
        If Not IsAttendeeFreeForSlot(mapItem, slotStart, durationMinutes, intervalMinutes) Then
            Exit Function
        End If
    Next mapItem

    AreAllRequiredAttendeesFree = True
End Function

Private Function IsAttendeeFreeForSlot( _
    ByVal dayMap As Object, _
    ByVal slotStart As Date, _
    ByVal durationMinutes As Long, _
    ByVal intervalMinutes As Long _
) As Boolean

    IsAttendeeFreeForSlot = False

    If dayMap Is Nothing Then Exit Function

    Dim dayKey As String
    dayKey = Format$(DateValue(slotStart), "yyyymmdd")

    If Not dayMap.Exists(dayKey) Then Exit Function

    Dim dailyBusy As String
    dailyBusy = CStr(dayMap(dayKey))

    Dim startIndex As Long
    startIndex = (DateDiff("n", DateValue(slotStart), slotStart) \ intervalMinutes) + 1

    Dim requiredIntervals As Long
    requiredIntervals = durationMinutes \ intervalMinutes

    If startIndex < 1 Then Exit Function
    If Len(dailyBusy) < (startIndex + requiredIntervals - 1) Then Exit Function

    Dim i As Long
    For i = 0 To requiredIntervals - 1
        Dim statusChar As String
        statusChar = Mid$(dailyBusy, startIndex + i, 1)

        ' Mapping: 0 (free) and 1 (tentative) are considered usable.
        ' Busy/OOF/unknown values remain blocked.
        If statusChar <> "0" And statusChar <> "1" Then Exit Function
    Next i

    IsAttendeeFreeForSlot = True
End Function

Private Function TimeRangesOverlap( _
    ByVal startA As Date, _
    ByVal endA As Date, _
    ByVal startB As Date, _
    ByVal endB As Date _
) As Boolean
    TimeRangesOverlap = (startA < endB) And (endA > startB)
End Function

' =============================
' CSV Output
' =============================
Private Function WriteSlotsToCsv( _
    ByVal slots As Collection, _
    ByVal queryTimestamp As Date, _
    ByRef csvPath As String, _
    ByRef errorText As String _
) As Boolean

    On Error GoTo EH

    WriteSlotsToCsv = False

    If Not EnsureFolderExists(OUTPUT_FOLDER, errorText) Then Exit Function

    csvPath = BuildOutputCsvPath(OUTPUT_FOLDER, FILE_PREFIX, queryTimestamp)

    Dim fileNo As Integer
    fileNo = FreeFile

    Open csvPath For Output As #fileNo

    Print #fileNo, BuildHeaderLine()

    Dim rowData As Variant
    For Each rowData In slots
        Print #fileNo, RowToCsvLine(rowData)
    Next rowData

    Close #fileNo
    WriteSlotsToCsv = True
    Exit Function

EH:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0

    errorText = "Error " & CStr(Err.Number) & ": " & Err.Description
    WriteSlotsToCsv = False
End Function

Private Function EnsureFolderExists(ByVal folderPath As String, ByRef errorText As String) As Boolean
    On Error GoTo EH

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    EnsureFolderExists = fso.FolderExists(folderPath)

    If Not EnsureFolderExists Then
        errorText = "Could not create output folder: " & folderPath
    End If
    Exit Function

EH:
    errorText = "Unable to access or create output folder '" & folderPath & "'. Error " & CStr(Err.Number) & ": " & Err.Description
    EnsureFolderExists = False
End Function

Private Function BuildOutputCsvPath(ByVal folderPath As String, ByVal filePrefix As String, ByVal ts As Date) As String
    Dim normalizedFolder As String
    normalizedFolder = folderPath

    If Right$(normalizedFolder, 1) <> "\" Then
        normalizedFolder = normalizedFolder & "\"
    End If

    BuildOutputCsvPath = normalizedFolder & filePrefix & "_" & Format$(ts, "yyyymmdd_Hhnnss") & ".csv"
End Function

Private Function BuildHeaderLine() As String
    BuildHeaderLine = JoinCsvFields(Array( _
        "QueryTimestamp", _
        "SlotStart", _
        "SlotEnd", _
        "DurationMinutes", _
        "RequiredAttendeesFree", _
        "OptionalAttendeeFree", _
        "OptionalAttendeeName", _
        "Notes" _
    ))
End Function

Private Function BuildRow( _
    ByVal queryTimestamp As Date, _
    ByVal slotStart As Date, _
    ByVal durationMinutes As Long, _
    ByVal optionalFree As Boolean, _
    ByVal optionalName As String, _
    ByVal notes As String _
) As Variant

    Dim rowValues(0 To 7) As String
    rowValues(0) = FormatDateTimeSortable(queryTimestamp)
    rowValues(1) = FormatDateTimeSortable(slotStart)
    rowValues(2) = FormatDateTimeSortable(DateAdd("n", durationMinutes, slotStart))
    rowValues(3) = CStr(durationMinutes)
    rowValues(4) = "TRUE"

    If optionalFree Then
        rowValues(5) = "TRUE"
    Else
        rowValues(5) = "FALSE"
    End If

    rowValues(6) = optionalName
    rowValues(7) = notes

    BuildRow = rowValues
End Function

Private Function RowToCsvLine(ByVal rowData As Variant) As String
    RowToCsvLine = JoinCsvFields(rowData)
End Function

Private Function JoinCsvFields(ByVal values As Variant) As String
    Dim i As Long
    Dim s As String

    For i = LBound(values) To UBound(values)
        If i > LBound(values) Then s = s & CSV_DELIMITER
        s = s & CsvEscape(CStr(values(i)))
    Next i

    JoinCsvFields = s
End Function

Private Function CsvEscape(ByVal rawValue As String) As String
    Dim v As String
    v = rawValue

    If InStr(v, Chr$(34)) > 0 Then
        v = Replace(v, Chr$(34), Chr$(34) & Chr$(34))
    End If

    If InStr(v, CSV_DELIMITER) > 0 Or InStr(v, Chr$(34)) > 0 Or InStr(v, vbCr) > 0 Or InStr(v, vbLf) > 0 Then
        CsvEscape = Chr$(34) & v & Chr$(34)
    Else
        CsvEscape = v
    End If
End Function

Private Function FormatDateTimeSortable(ByVal dt As Date) As String
    FormatDateTimeSortable = Format$(dt, "yyyy-mm-dd HH:nn")
End Function

Private Function BuildNote(ByVal durationMinutes As Long, ByVal optionalFree As Boolean) As String
    Select Case durationMinutes
        Case 60
            If optionalFree Then
                BuildNote = "Preferred 60 min slot; optional attendee free"
            Else
                BuildNote = "Preferred 60 min slot; optional attendee blocked"
            End If
        Case 30
            If optionalFree Then
                BuildNote = "Fallback 30 min slot; optional attendee free"
            Else
                BuildNote = "Fallback 30 min slot; optional attendee blocked"
            End If
        Case Else
            If optionalFree Then
                BuildNote = "Slot; optional attendee free"
            Else
                BuildNote = "Slot; optional attendee blocked"
            End If
    End Select
End Function

Private Sub StartRunningIndicator(ByVal phase As String)
    On Error Resume Next

    Dim exp As Outlook.Explorer
    Set exp = Application.ActiveExplorer

    If exp Is Nothing Then Exit Sub

    mOriginalStatusBar = exp.StatusBar
    mIndicatorInitialized = True

    exp.StatusBar = "RUNNING: " & phase
    DoEvents
End Sub

Private Sub UpdateRunningIndicator(ByVal phase As String, ByVal currentValue As Long, ByVal totalValue As Long)
    On Error Resume Next

    Dim exp As Outlook.Explorer
    Set exp = Application.ActiveExplorer

    If exp Is Nothing Then Exit Sub

    If Not mIndicatorInitialized Then
        mOriginalStatusBar = exp.StatusBar
        mIndicatorInitialized = True
    End If

    Dim percentText As String
    If totalValue > 0 Then
        percentText = CStr(Int((CDbl(currentValue) / CDbl(totalValue)) * 100)) & "%"
        exp.StatusBar = "RUNNING: " & phase & " " & CStr(currentValue) & "/" & CStr(totalValue) & " (" & percentText & ")"
    Else
        exp.StatusBar = "RUNNING: " & phase
    End If

    DoEvents
End Sub

Private Sub StopRunningIndicator()
    On Error Resume Next

    Dim exp As Outlook.Explorer
    Set exp = Application.ActiveExplorer

    If exp Is Nothing Then Exit Sub

    If mIndicatorInitialized Then
        If IsEmpty(mOriginalStatusBar) Then
            exp.StatusBar = False
        Else
            exp.StatusBar = mOriginalStatusBar
        End If
    End If

    mIndicatorInitialized = False
    mOriginalStatusBar = Empty
    DoEvents
End Sub