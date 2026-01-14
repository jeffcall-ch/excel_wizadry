Option Explicit  ' Force variable declaration - prevents typos

Sub CalculateFastenerSpares()
    '============================================================
    ' CONFIGURABLE PARAMETERS
    '============================================================
    Const HEADER_ROW As Long = 6
    
    Const COL_DESCRIPTION As String = "C"
    Const COL_QTY As String = "E"
    Const COL_SPARE As String = "F"
    
    ' Nut/Washer/Bolt keywords
    Const NUT_WASHER_BOLT_KEYWORDS As String = "nut,washer,bolt"
    Const NWB_THRESHOLD_1 As Long = 30
    Const NWB_THRESHOLD_2 As Long = 100
    Const NWB_PERCENT_LOW As Double = 1#
    Const NWB_PERCENT_MID As Double = 0.5
    Const NWB_PERCENT_HIGH As Double = 0.2
    
    ' Gasket keywords
    Const GASKET_KEYWORD As String = "gasket"
    Const GASKET_THRESHOLD_1 As Long = 30
    Const GASKET_THRESHOLD_2 As Long = 500
    Const GASKET_PERCENT_LOW As Double = 1#
    Const GASKET_PERCENT_MID As Double = 0.5
    Const GASKET_PERCENT_HIGH As Double = 0.3
    
    ' Standard thresholds
    Const STANDARD_THRESHOLD As Long = 400
    Const STANDARD_PERCENT_LOW As Double = 0.35
    Const STANDARD_PERCENT_HIGH As Double = 0.2
    
    '============================================================
    ' MAIN LOGIC
    '============================================================
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim qty As Double
    Dim spareValue As Long
    Dim noteText As String
    Dim ruleNum As String
    Dim baseSpare As Long
    Dim totalQty As Long
    Dim processedCount As Long
    Dim nutWasherBoltCount As Long
    Dim gasketCount As Long
    Dim standardCount As Long
    Dim isNutWasherBolt As Boolean
    Dim isGasket As Boolean
    Dim description As String
    
    Dim descArr As Variant
    Dim qtyArr As Variant
    Dim spareArr As Variant
    Dim notesArr() As String
    
    Dim calcMode As XlCalculation
    Dim startTime As Double
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    
    startTime = Timer
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ws.Range(COL_SPARE & HEADER_ROW).value = "Spare"
    
    lastRow = FindLastDataRow(ws, HEADER_ROW)
    
    If lastRow < HEADER_ROW + 1 Then
        MsgBox "No data found to process.", vbInformation
        GoTo CleanUp
    End If
    
    ' Clear existing spare column (values and comments only, preserve formatting)
    ws.Range(COL_SPARE & (HEADER_ROW + 1) & ":" & COL_SPARE & lastRow).ClearContents
    ws.Range(COL_SPARE & (HEADER_ROW + 1) & ":" & COL_SPARE & lastRow).ClearComments
    
    descArr = ws.Range(COL_DESCRIPTION & (HEADER_ROW + 1) & ":" & COL_DESCRIPTION & lastRow).value
    qtyArr = ws.Range(COL_QTY & (HEADER_ROW + 1) & ":" & COL_QTY & lastRow).value
    
    ReDim spareArr(1 To UBound(descArr, 1), 1 To 1)
    ReDim notesArr(1 To UBound(descArr, 1))
    
    processedCount = 0
    nutWasherBoltCount = 0
    gasketCount = 0
    standardCount = 0
    
    For i = 1 To UBound(descArr, 1)
        spareValue = 0
        noteText = ""
        ruleNum = ""
        
        description = GetStringValue(descArr(i, 1))
        qty = GetNumericValue(qtyArr(i, 1))
        
        If qty < 0 Then qty = 0
        
        isNutWasherBolt = HasAnyKeyword(description, NUT_WASHER_BOLT_KEYWORDS)
        isGasket = HasKeyword(description, GASKET_KEYWORD)
        
        If isNutWasherBolt Then
            nutWasherBoltCount = nutWasherBoltCount + 1
        ElseIf isGasket Then
            gasketCount = gasketCount + 1
        Else
            standardCount = standardCount + 1
        End If
        
        If qty > 0 Then
            ' NUT/WASHER/BOLT
            If isNutWasherBolt Then
                If qty <= NWB_THRESHOLD_1 Then
                    baseSpare = RoundUpToInteger(qty * NWB_PERCENT_LOW)
                    totalQty = RoundUpToNext10(qty + baseSpare)
                    spareValue = totalQty - qty
                    ruleNum = "#1"
                    noteText = "Rule #1: Nut/Washer/Bolt qty <= " & NWB_THRESHOLD_1 & _
                               " -> base spare = " & PercentText(NWB_PERCENT_LOW) & _
                               " of qty = " & baseSpare & "; total rounded to next 10 = " & totalQty & "; spare = " & spareValue & "."
                ElseIf qty <= NWB_THRESHOLD_2 Then
                    baseSpare = RoundUpToInteger(qty * NWB_PERCENT_MID)
                    totalQty = RoundUpToNext10(qty + baseSpare)
                    spareValue = totalQty - qty
                    ruleNum = "#2"
                    noteText = "Rule #2: Nut/Washer/Bolt qty between " & (NWB_THRESHOLD_1 + 1) & _
                               " and " & NWB_THRESHOLD_2 & " -> base spare = " & PercentText(NWB_PERCENT_MID) & _
                               " of qty = " & baseSpare & "; total rounded to next 10 = " & totalQty & "; spare = " & spareValue & "."
                Else
                    baseSpare = RoundUpToInteger(qty * NWB_PERCENT_HIGH)
                    totalQty = RoundUpToNext10(qty + baseSpare)
                    spareValue = totalQty - qty
                    ruleNum = "#3"
                    noteText = "Rule #3: Nut/Washer/Bolt qty > " & NWB_THRESHOLD_2 & _
                               " -> base spare = " & PercentText(NWB_PERCENT_HIGH) & _
                               " of qty = " & baseSpare & "; total rounded to next 10 = " & totalQty & "; spare = " & spareValue & "."
                End If
            
            ' GASKET
            ElseIf isGasket Then
                If qty <= GASKET_THRESHOLD_1 Then
                    spareValue = RoundUpToInteger(qty * GASKET_PERCENT_LOW)
                    ruleNum = "#4"
                    noteText = "Rule #4: Gasket qty <= " & GASKET_THRESHOLD_1 & _
                               " -> spare = " & PercentText(GASKET_PERCENT_LOW) & _
                               " of qty rounded up to the next integer."
                ElseIf qty <= GASKET_THRESHOLD_2 Then
                    baseSpare = RoundUpToInteger(qty * GASKET_PERCENT_MID)
                    totalQty = RoundUpToNext10(qty + baseSpare)
                    spareValue = totalQty - qty
                    ruleNum = "#5"
                    noteText = "Rule #5: Gasket qty between " & (GASKET_THRESHOLD_1 + 1) & _
                               " and " & GASKET_THRESHOLD_2 & " -> base spare = " & PercentText(GASKET_PERCENT_MID) & _
                               " of qty = " & baseSpare & "; total rounded to next 10 = " & totalQty & "; spare = " & spareValue & "."
                Else
                    baseSpare = RoundUpToInteger(qty * GASKET_PERCENT_HIGH)
                    totalQty = RoundUpToNext10(qty + baseSpare)
                    spareValue = totalQty - qty
                    ruleNum = "#6"
                    noteText = "Rule #6: Gasket qty > " & GASKET_THRESHOLD_2 & _
                               " -> base spare = " & PercentText(GASKET_PERCENT_HIGH) & _
                               " of qty = " & baseSpare & "; total rounded to next 10 = " & totalQty & "; spare = " & spareValue & "."
                End If
            
            ' STANDARD
            Else
                If qty < STANDARD_THRESHOLD Then
                    spareValue = RoundUpToInteger(qty * STANDARD_PERCENT_LOW)
                    ruleNum = "#7"
                    noteText = "Rule #7: Standard components qty < " & STANDARD_THRESHOLD & _
                               " -> spare = " & PercentText(STANDARD_PERCENT_LOW) & _
                               " of qty rounded up to the next integer."
                Else
                    spareValue = RoundUpToInteger(qty * STANDARD_PERCENT_HIGH)
                    ruleNum = "#8"
                    noteText = "Rule #8: Standard components qty >= " & STANDARD_THRESHOLD & _
                               " -> spare = " & PercentText(STANDARD_PERCENT_HIGH) & _
                               " of qty rounded up to the next integer."
                End If
            End If
            
            processedCount = processedCount + 1
        End If
        
        spareArr(i, 1) = spareValue
        notesArr(i) = noteText
    Next i
    
    ' Write values to sheet
    ws.Range(COL_SPARE & (HEADER_ROW + 1) & ":" & COL_SPARE & lastRow).value = spareArr
    
    ' Add notes/comments to cells
    Dim currentRow As Long
    For i = 1 To UBound(notesArr)
        If Len(notesArr(i)) > 0 Then
            currentRow = HEADER_ROW + i
            AddCellNote ws.Range(COL_SPARE & currentRow), notesArr(i)
        End If
    Next i
    
    ' Add rules documentation at bottom
    Call AddFastenerRulesDocumentation(ws, lastRow, COL_DESCRIPTION)
    
    MsgBox "Fastener spare calculation completed in " & _
           Format(Timer - startTime, "0.00") & " seconds." & vbCrLf & vbCrLf & _
           "Processed: " & processedCount & " rows" & vbCrLf & _
           "Nut/Washer/Bolt: " & nutWasherBoltCount & " rows" & vbCrLf & _
           "Gasket: " & gasketCount & " rows" & vbCrLf & _
           "Standard: " & standardCount & " rows" & vbCrLf & _
           "Skipped: " & (lastRow - HEADER_ROW - processedCount) & " rows" & vbCrLf & vbCrLf & _
           "Rules documentation added below data.", _
           vbInformation, "Calculation Complete"

CleanUp:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "An error occurred during calculation:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description & vbCrLf & _
           "Processing was stopped.", _
           vbCritical, "Error"
End Sub

Sub AddFastenerRulesDocumentation(ws As Worksheet, lastDataRow As Long, ruleColumn As String)
    ' Add comprehensive rules documentation below the data
    
    Dim startRow As Long
    Dim currentRow As Long
    
    ' Start 5 rows below last data
    startRow = lastDataRow + 5
    currentRow = startRow
    
    ' Header
    ws.Range(ruleColumn & currentRow).value = "SPARE CALCULATION RULES"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    ws.Range(ruleColumn & currentRow).Font.Size = 12
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "Priority: NUT/WASHER/BOLT > GASKET > STANDARD"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 2
    
    ' NUT/WASHER/BOLT RULES
    ws.Range(ruleColumn & currentRow).value = "NUT/WASHER/BOLT COMPONENTS (keywords: nut, washer, bolt)"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #1: If QTY <= 30 -> Base spare = 100% of QTY; Total (QTY + base spare) rounded up to next 10"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "           Examples: QTY=5->base=5,total=10,spare=5; QTY=15->base=15,total=30,spare=15; QTY=25->base=25,total=50,spare=25"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #2: If 30 < QTY <= 100 -> Base spare = 50% of QTY; Total (QTY + base spare) rounded up to next 10"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "           Examples: QTY=40->base=20,total=60,spare=20; QTY=50->base=25,total=80,spare=30; QTY=60->base=30,total=90,spare=30"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #3: If QTY > 100 -> Base spare = 20% of QTY; Total (QTY + base spare) rounded up to next 10"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "           Examples: QTY=150->base=30,total=180,spare=30; QTY=200->base=40,total=240,spare=40; QTY=250->base=50,total=300,spare=50"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 2
    
    ' GASKET RULES
    ws.Range(ruleColumn & currentRow).value = "GASKET COMPONENTS (keyword: gasket)"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #4: If QTY <= 30 -> Spare = 100% of QTY, rounded up to next integer"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "           Examples: QTY=5->5, QTY=15->15, QTY=30->30"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #5: If 30 < QTY <= 500 -> Base spare = 50% of QTY; Total (QTY + base spare) rounded up to next 10"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "           Examples: QTY=40->base=20,total=60,spare=20; QTY=100->base=50,total=150,spare=50; QTY=200->base=100,total=300,spare=100"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #6: If QTY > 500 -> Base spare = 30% of QTY; Total (QTY + base spare) rounded up to next 10"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "           Examples: QTY=600->base=180,total=780,spare=180; QTY=700->base=210,total=910,spare=210; QTY=1000->base=300,total=1300,spare=300"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 2
    
    ' STANDARD RULES
    ws.Range(ruleColumn & currentRow).value = "STANDARD COMPONENTS (all others)"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #7: If QTY < 400 -> Spare = 35% of QTY, rounded up to next integer"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "           Examples: QTY=100->35, QTY=200->70, QTY=399->140"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #8: If QTY >= 400 -> Spare = 20% of QTY, rounded up to next integer"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "           Examples: QTY=400->80, QTY=500->100, QTY=1000->200"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 2
    
    ' Notes
    ws.Range(ruleColumn & currentRow).value = "NOTES:"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  - Component classification is case-insensitive"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  - 'Round up to next 10' means: 3->10, 15->20, 21->30, 95->100"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  - 'Round up to next integer' uses ceiling function: 1.1->2, 5.2->6"
    
    ' Auto-fit column width
    ws.Columns(ruleColumn).AutoFit
End Sub

'============================================================
' HELPER FUNCTIONS
'============================================================

Sub AddCellNote(cell As Range, noteText As String)
    ' Add or update a note/comment to a cell
    On Error Resume Next
    cell.Comment.Delete  ' Remove existing comment if any
    On Error GoTo 0
    
    If Len(noteText) > 0 Then
        cell.AddComment noteText
        cell.Comment.Shape.TextFrame.AutoSize = True
    End If
End Sub

Function HasKeyword(description As String, keyword As String) As Boolean
    Dim descLower As String
    HasKeyword = False
    If Len(Trim(description)) = 0 Then Exit Function
    If Len(Trim(keyword)) = 0 Then Exit Function
    descLower = LCase(Trim(description))
    If InStr(1, descLower, LCase(Trim(keyword)), vbTextCompare) > 0 Then
        HasKeyword = True
    End If
End Function

Function HasAnyKeyword(description As String, keywordList As String) As Boolean
    Dim keywords() As String
    Dim keyword As Variant
    Dim descLower As String
    
    HasAnyKeyword = False
    If Len(Trim(description)) = 0 Then Exit Function
    If Len(Trim(keywordList)) = 0 Then Exit Function
    
    descLower = LCase(Trim(description))
    keywords = Split(keywordList, ",")
    
    For Each keyword In keywords
        If Len(Trim(CStr(keyword))) > 0 Then
            If InStr(1, descLower, LCase(Trim(CStr(keyword))), vbTextCompare) > 0 Then
                HasAnyKeyword = True
                Exit Function
            End If
        End If
    Next keyword
End Function

Function GetStringValue(cellValue As Variant) As String
    If IsEmpty(cellValue) Then
        GetStringValue = ""
    ElseIf IsError(cellValue) Then
        GetStringValue = ""
    Else
        GetStringValue = Trim(CStr(cellValue))
    End If
End Function

Function FindLastDataRow(ws As Worksheet, headerRow As Long) As Long
    ' Find the last row where columns A, B, C, and D all have non-empty values
    Dim maxRow As Long
    Dim testRow As Long
    Dim valA As Variant, valB As Variant, valC As Variant, valD As Variant
    
    ' Find the maximum row that has data in any of the columns A-D
    maxRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "B").End(xlUp).Row > maxRow Then
        maxRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    End If
    If ws.Cells(ws.Rows.Count, "C").End(xlUp).Row > maxRow Then
        maxRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    End If
    If ws.Cells(ws.Rows.Count, "D").End(xlUp).Row > maxRow Then
        maxRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    End If
    
    ' If no data found beyond header, return header row
    If maxRow <= headerRow Then
        FindLastDataRow = headerRow
        Exit Function
    End If
    
    ' Search backwards from maxRow to find the last row where A, B, C, D all have values
    For testRow = maxRow To headerRow + 1 Step -1
        valA = ws.Cells(testRow, "A").Value
        valB = ws.Cells(testRow, "B").Value
        valC = ws.Cells(testRow, "C").Value
        valD = ws.Cells(testRow, "D").Value
        
        ' Check if all four cells have non-empty values
        If Not IsEmpty(valA) And Not IsEmpty(valB) And _
           Not IsEmpty(valC) And Not IsEmpty(valD) And _
           Len(Trim(CStr(valA))) > 0 And Len(Trim(CStr(valB))) > 0 And _
           Len(Trim(CStr(valC))) > 0 And Len(Trim(CStr(valD))) > 0 Then
            FindLastDataRow = testRow
            Exit Function
        End If
    Next testRow
    
    ' If no valid row found, return header row
    FindLastDataRow = headerRow
End Function

Function GetNumericValue(cellValue As Variant) As Double
    If IsEmpty(cellValue) Then
        GetNumericValue = 0
    ElseIf IsError(cellValue) Then
        GetNumericValue = 0
    ElseIf IsNumeric(cellValue) Then
        GetNumericValue = CDbl(cellValue)
    Else
        GetNumericValue = 0
    End If
End Function

Function PercentText(value As Double) As String
    PercentText = Format$(value * 100, "0.##") & "%"
End Function

Function RoundUpToInteger(value As Double) As Long
    If value = Int(value) Then
        RoundUpToInteger = CLng(value)
    ElseIf value > 0 Then
        RoundUpToInteger = Int(value) + 1
    Else
        RoundUpToInteger = Int(value)
    End If
End Function

Function RoundUpToNext10(value As Double) As Long
    If value <= 0 Then
        RoundUpToNext10 = 0
        Exit Function
    End If
    RoundUpToNext10 = CLng(Application.WorksheetFunction.Ceiling_Math(value, 10))
End Function


