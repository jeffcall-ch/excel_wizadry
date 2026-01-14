Option Explicit  ' Force variable declaration - prevents typos

Sub CalculateSpares()
    '============================================================
    ' CONFIGURABLE PARAMETERS
    '============================================================
    Const HEADER_ROW As Long = 7
    
    Const COL_MATERIAL As String = "A"
    Const COL_TYPE_LABEL As String = "B"
    Const COL_TYPE As String = "C"
    Const COL_DESCRIPTION As String = "D"
    Const COL_QTY_PCS As String = "E"
    Const COL_QTY_M As String = "F"
    Const COL_SPARE As String = "H"
    
    ' Seal ring keywords/filters
    Const SEAL_REQUIRED_KEYWORDS As String = "seal,ring,fkm"
    Const SEAL_EXCLUDED_KEYWORDS As String = "elbow,flange,tee,reducer"
    Const SEAL_MIN_SPARE_LOW_QTY As Long = 5
    Const SEAL_MIN_SPARE_HIGH_QTY As Long = 5
    Const SEAL_PERCENT As Double = 0.1
    Const SEAL_QTY_THRESHOLD As Long = 10
    
    ' Critical component keywords
    Const CRITICAL_KEYWORDS As String = "nipple,nipples,union,connection,connections,welding,adaptor"
    Const CRITICAL_MIN_SPARE As Long = 3
    Const CRITICAL_PERCENT_LOW_QTY As Double = 0.5
    Const CRITICAL_PERCENT_HIGH_QTY As Double = 0.1
    Const CRITICAL_QTY_THRESHOLD As Long = 10
    
    ' MAPRESS fitting keywords
    Const MAPRESS_KEYWORD As String = "mapress"
    Const MAPRESS_FITTING_KEYWORDS As String = "cap,coupling,elbow,flange,tee"
    Const MAPRESS_PERCENT As Double = 0.1
    Const MAPRESS_QTY_THRESHOLD As Long = 20
    
    ' Standard thresholds
    Const THRESHOLD_1 As Long = 5
    Const THRESHOLD_2 As Long = 20
    Const TYPE_THRESHOLD As Double = 50
    
    ' Standard percentages
    Const PERCENT_E_SMALL As Double = 0.1
    Const PERCENT_E_LARGE As Double = 0.05
    Const PERCENT_F_SMALL As Double = 0.15
    Const PERCENT_F_LARGE As Double = 0.1
    
    '============================================================
    ' MAIN LOGIC
    '============================================================
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim typeLabel As String
    Dim typeValue As Double
    Dim qtyPcs As Double
    Dim qtyM As Double
    Dim spareValue As Double
    Dim noteText As String
    Dim ruleNum As String
    Dim unitLabel As String
    Dim tubiPercent As Double
    Dim tubiBaseSpare As Long
    Dim tubiTotalQty As Long
    Dim processedCount As Long
    Dim criticalCount As Long
    Dim sealCount As Long
    Dim mapressCount As Long
    Dim tubiCount As Long
    Dim isCritical As Boolean
    Dim isSeal As Boolean
    Dim isMapress As Boolean
    Dim isTubi As Boolean
    Dim description As String
    
    Dim typeLabelArr As Variant
    Dim typeArr As Variant
    Dim descArr As Variant
    Dim qtyPcsArr As Variant
    Dim qtyMArr As Variant
    Dim spareArr As Variant
    Dim notesArr() As String
    
    Dim calcMode As XlCalculation
    Dim startTime As Double
    
    On Error GoTo ErrorHandler
    
    If Not ValidateParameters(HEADER_ROW, THRESHOLD_1, THRESHOLD_2) Then
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    startTime = Timer
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ws.Range(COL_SPARE & HEADER_ROW).value = "Spare"
    
    lastRow = FindLastDataRow(ws, HEADER_ROW, COL_MATERIAL, COL_TYPE_LABEL, COL_TYPE, COL_DESCRIPTION)
    
    If lastRow < HEADER_ROW + 1 Then
        MsgBox "No data found to process.", vbInformation
        GoTo CleanUp
    End If
    
    ' Clear existing spare column (values and comments only, preserve formatting)
    ws.Range(COL_SPARE & (HEADER_ROW + 1) & ":" & COL_SPARE & lastRow).ClearContents
    ws.Range(COL_SPARE & (HEADER_ROW + 1) & ":" & COL_SPARE & lastRow).ClearComments
    
    typeLabelArr = ws.Range(COL_TYPE_LABEL & (HEADER_ROW + 1) & ":" & COL_TYPE_LABEL & lastRow).value
    typeArr = ws.Range(COL_TYPE & (HEADER_ROW + 1) & ":" & COL_TYPE & lastRow).value
    descArr = ws.Range(COL_DESCRIPTION & (HEADER_ROW + 1) & ":" & COL_DESCRIPTION & lastRow).value
    qtyPcsArr = ws.Range(COL_QTY_PCS & (HEADER_ROW + 1) & ":" & COL_QTY_PCS & lastRow).value
    qtyMArr = ws.Range(COL_QTY_M & (HEADER_ROW + 1) & ":" & COL_QTY_M & lastRow).value
    
    ReDim spareArr(1 To UBound(typeArr, 1), 1 To 1)
    ReDim notesArr(1 To UBound(typeArr, 1))
    
    processedCount = 0
    criticalCount = 0
    sealCount = 0
    mapressCount = 0
    tubiCount = 0
    
    For i = 1 To UBound(typeArr, 1)
        spareValue = 0
        noteText = ""
        ruleNum = ""
        
        description = GetStringValue(descArr(i, 1))
        qtyPcs = GetNumericValue(qtyPcsArr(i, 1))
        qtyM = GetNumericValue(qtyMArr(i, 1))
        typeLabel = GetStringValue(typeLabelArr(i, 1))
        typeValue = ExtractNumericFromType(typeArr(i, 1))
        
        If qtyPcs < 0 Then qtyPcs = 0
        If qtyM < 0 Then qtyM = 0
        
        isSeal = IsSealComponent(description, SEAL_REQUIRED_KEYWORDS, SEAL_EXCLUDED_KEYWORDS)
        isCritical = HasAnyKeyword(description, CRITICAL_KEYWORDS)
        isTubi = IsTubiComponent(typeLabel, CStr(typeArr(i, 1)), description)
        isMapress = HasMapressWithFitting(description, MAPRESS_KEYWORD, MAPRESS_FITTING_KEYWORDS)
        If isTubi Then isMapress = False
        
        If isSeal Then sealCount = sealCount + 1
        If isCritical And Not isSeal Then criticalCount = criticalCount + 1
        If isTubi And Not isSeal And Not isCritical Then tubiCount = tubiCount + 1
        If isMapress And Not isSeal And Not isCritical And Not isTubi Then mapressCount = mapressCount + 1
        
        If qtyPcs > 0 Or qtyM > 0 Then
            If qtyPcs > 0 Then
                unitLabel = "pcs"
                ' SEAL RING
                If isSeal Then
                    If qtyPcs <= SEAL_QTY_THRESHOLD Then
                        spareValue = SEAL_MIN_SPARE_LOW_QTY
                        ruleNum = "#1"
                        noteText = "Rule #1: Seal components qty <= " & SEAL_QTY_THRESHOLD & _
                                   " -> spare = fixed minimum of " & SEAL_MIN_SPARE_LOW_QTY & _
                                   " " & unitLabel & "."
                    Else
                        spareValue = Application.WorksheetFunction.Max(SEAL_MIN_SPARE_HIGH_QTY, RoundUpToInteger(qtyPcs * SEAL_PERCENT))
                        ruleNum = "#2"
                        noteText = "Rule #2: Seal components qty > " & SEAL_QTY_THRESHOLD & _
                                   " -> spare = max(" & SEAL_MIN_SPARE_HIGH_QTY & ", ceiling(qty * " & _
                                   PercentText(SEAL_PERCENT) & ")) " & unitLabel & "."
                    End If
                
                ' CRITICAL
                ElseIf isCritical Then
                    If qtyPcs <= CRITICAL_QTY_THRESHOLD Then
                        spareValue = Application.WorksheetFunction.Max(CRITICAL_MIN_SPARE, RoundUpToInteger(qtyPcs * CRITICAL_PERCENT_LOW_QTY))
                        ruleNum = "#3"
                        noteText = "Rule #3: Critical components qty <= " & CRITICAL_QTY_THRESHOLD & _
                                   " -> spare = max(" & CRITICAL_MIN_SPARE & ", ceiling(qty * " & _
                                   PercentText(CRITICAL_PERCENT_LOW_QTY) & ")) " & unitLabel & "."
                    Else
                        spareValue = Application.WorksheetFunction.Max(CRITICAL_MIN_SPARE, RoundUpToInteger(qtyPcs * CRITICAL_PERCENT_HIGH_QTY))
                        ruleNum = "#4"
                        noteText = "Rule #4: Critical components qty > " & CRITICAL_QTY_THRESHOLD & _
                                   " -> spare = max(" & CRITICAL_MIN_SPARE & ", ceiling(qty * " & _
                                   PercentText(CRITICAL_PERCENT_HIGH_QTY) & ")) " & unitLabel & "."
                    End If
                
                ' TUBI
                ElseIf isTubi Then
                    If typeValue > 0 And typeValue < TYPE_THRESHOLD Then
                        tubiPercent = PERCENT_F_SMALL
                    Else
                        tubiPercent = PERCENT_F_LARGE
                    End If
                    tubiBaseSpare = RoundUpToInteger(qtyPcs * tubiPercent)
                    tubiTotalQty = RoundUpToMultiple(qtyPcs + tubiBaseSpare, 6)
                    spareValue = Application.WorksheetFunction.Max(0, tubiTotalQty - qtyPcs)
                    ruleNum = "#T1"
                    noteText = "Rule #T1: TUBI components -> base spare = ceiling(qty * " & _
                               PercentText(tubiPercent) & ") = " & tubiBaseSpare & " " & unitLabel & _
                               "; adjusted total rounded to next multiple of 6 = " & tubiTotalQty & _
                               "; spare = " & spareValue & " " & unitLabel & "."
                
                ' MAPRESS
                ElseIf isMapress Then
                    If qtyPcs > MAPRESS_QTY_THRESHOLD Then
                        spareValue = RoundUpToInteger(qtyPcs * MAPRESS_PERCENT)
                        ruleNum = "#7"
                        noteText = "Rule #7: MAPRESS fittings qty > " & MAPRESS_QTY_THRESHOLD & _
                                   " -> spare = ceiling(qty * " & PercentText(MAPRESS_PERCENT) & ") " & unitLabel & "."
                    ElseIf qtyPcs <= THRESHOLD_1 Then
                        spareValue = 0
                        ruleNum = "#5"
                        noteText = "Rule #5: MAPRESS fittings qty <= " & THRESHOLD_1 & _
                                   " -> spare = 0 " & unitLabel & "."
                    Else
                        spareValue = 1
                        ruleNum = "#6"
                        noteText = "Rule #6: MAPRESS fittings qty between " & (THRESHOLD_1 + 1) & _
                                   " and " & MAPRESS_QTY_THRESHOLD & " -> spare = 1 " & unitLabel & "."
                    End If
                
                ' STANDARD
                Else
                    If qtyPcs <= THRESHOLD_1 Then
                        spareValue = 0
                        ruleNum = "#8"
                        noteText = "Rule #8: Standard components qty <= " & THRESHOLD_1 & _
                                   " -> spare = 0 " & unitLabel & "."
                    ElseIf qtyPcs <= THRESHOLD_2 Then
                        spareValue = 1
                        ruleNum = "#9"
                        noteText = "Rule #9: Standard components qty between " & (THRESHOLD_1 + 1) & _
                                   " and " & THRESHOLD_2 & " -> spare = 1 " & unitLabel & "."
                    ElseIf typeValue < TYPE_THRESHOLD Then
                        spareValue = RoundUpToInteger(qtyPcs * PERCENT_E_SMALL)
                        ruleNum = "#10a"
                        noteText = "Rule #10a: Standard components qty > " & THRESHOLD_2 & _
                                   " and DN < " & TYPE_THRESHOLD & " -> spare = ceiling(qty * " & _
                                   PercentText(PERCENT_E_SMALL) & ") " & unitLabel & "."
                    Else
                        spareValue = RoundUpToInteger(qtyPcs * PERCENT_E_LARGE)
                        ruleNum = "#10b"
                        noteText = "Rule #10b: Standard components qty > " & THRESHOLD_2 & _
                                   " and DN >= " & TYPE_THRESHOLD & " -> spare = ceiling(qty * " & _
                                   PercentText(PERCENT_E_LARGE) & ") " & unitLabel & "."
                    End If
                End If
            
            ' QTY M
            ElseIf qtyM > 0 Then
                unitLabel = "m"
                If isSeal Then
                    If qtyM <= SEAL_QTY_THRESHOLD Then
                        spareValue = SEAL_MIN_SPARE_LOW_QTY
                        ruleNum = "#1"
                        noteText = "Rule #1: Seal components qty <= " & SEAL_QTY_THRESHOLD & _
                                   " -> spare = fixed minimum of " & SEAL_MIN_SPARE_LOW_QTY & _
                                   " " & unitLabel & "."
                    Else
                        spareValue = Application.WorksheetFunction.Max(SEAL_MIN_SPARE_HIGH_QTY, RoundUpToInteger(qtyM * SEAL_PERCENT))
                        ruleNum = "#2"
                        noteText = "Rule #2: Seal components qty > " & SEAL_QTY_THRESHOLD & _
                                   " -> spare = max(" & SEAL_MIN_SPARE_HIGH_QTY & ", ceiling(qty * " & _
                                   PercentText(SEAL_PERCENT) & ")) " & unitLabel & "."
                    End If
                
                ElseIf isCritical Then
                    If qtyM <= CRITICAL_QTY_THRESHOLD Then
                        spareValue = Application.WorksheetFunction.Max(CRITICAL_MIN_SPARE, RoundUpToInteger(qtyM * CRITICAL_PERCENT_LOW_QTY))
                        ruleNum = "#3"
                        noteText = "Rule #3: Critical components qty <= " & CRITICAL_QTY_THRESHOLD & _
                                   " -> spare = max(" & CRITICAL_MIN_SPARE & ", ceiling(qty * " & _
                                   PercentText(CRITICAL_PERCENT_LOW_QTY) & ")) " & unitLabel & "."
                    Else
                        spareValue = Application.WorksheetFunction.Max(CRITICAL_MIN_SPARE, RoundUpToInteger(qtyM * CRITICAL_PERCENT_HIGH_QTY))
                        ruleNum = "#4"
                        noteText = "Rule #4: Critical components qty > " & CRITICAL_QTY_THRESHOLD & _
                                   " -> spare = max(" & CRITICAL_MIN_SPARE & ", ceiling(qty * " & _
                                   PercentText(CRITICAL_PERCENT_HIGH_QTY) & ")) " & unitLabel & "."
                    End If
                
                ' TUBI
                ElseIf isTubi Then
                    If typeValue > 0 And typeValue < TYPE_THRESHOLD Then
                        tubiPercent = PERCENT_F_SMALL
                    Else
                        tubiPercent = PERCENT_F_LARGE
                    End If
                    tubiBaseSpare = RoundUpToInteger(qtyM * tubiPercent)
                    tubiTotalQty = RoundUpToMultiple(qtyM + tubiBaseSpare, 6)
                    spareValue = Application.WorksheetFunction.Max(0, tubiTotalQty - qtyM)
                    ruleNum = "#T1"
                    noteText = "Rule #T1: TUBI components -> base spare = ceiling(qty * " & _
                               PercentText(tubiPercent) & ") = " & tubiBaseSpare & " " & unitLabel & _
                               "; adjusted total rounded to next multiple of 6 = " & tubiTotalQty & _
                               "; spare = " & spareValue & " " & unitLabel & "."
                
                ElseIf isMapress Then
                    If qtyM > MAPRESS_QTY_THRESHOLD Then
                        spareValue = RoundUpToInteger(qtyM * MAPRESS_PERCENT)
                        ruleNum = "#7"
                        noteText = "Rule #7: MAPRESS fittings qty > " & MAPRESS_QTY_THRESHOLD & _
                                   " -> spare = ceiling(qty * " & PercentText(MAPRESS_PERCENT) & ") " & unitLabel & "."
                    Else
                        spareValue = 0
                        ruleNum = "#5"
                        noteText = "Rule #5: MAPRESS fittings qty <= " & MAPRESS_QTY_THRESHOLD & _
                                   " -> spare = 0 " & unitLabel & "."
                    End If
                
                Else
                    If typeValue < TYPE_THRESHOLD Then
                        spareValue = RoundUpToInteger(qtyM * PERCENT_F_SMALL)
                        ruleNum = "#11a"
                        noteText = "Rule #11a: Standard meter qty with DN < " & TYPE_THRESHOLD & _
                                   " -> spare = ceiling(qty * " & PercentText(PERCENT_F_SMALL) & ") " & unitLabel & "."
                    Else
                        spareValue = RoundUpToInteger(qtyM * PERCENT_F_LARGE)
                        ruleNum = "#11b"
                        noteText = "Rule #11b: Standard meter qty with DN >= " & TYPE_THRESHOLD & _
                                   " -> spare = ceiling(qty * " & PercentText(PERCENT_F_LARGE) & ") " & unitLabel & "."
                    End If
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
    Call AddRulesDocumentation(ws, lastRow, COL_DESCRIPTION)
    
    MsgBox "Spare calculation completed successfully in " & _
           Format(Timer - startTime, "0.00") & " seconds." & vbCrLf & vbCrLf & _
           "Processed: " & processedCount & " rows" & vbCrLf & _
           "Seal rings: " & sealCount & " rows" & vbCrLf & _
           "Critical components: " & criticalCount & " rows" & vbCrLf & _
            "TUBI components: " & tubiCount & " rows" & vbCrLf & _
           "MAPRESS fittings: " & mapressCount & " rows" & vbCrLf & _
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

Sub AddRulesDocumentation(ws As Worksheet, lastDataRow As Long, ruleColumn As String)
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
    
    ws.Range(ruleColumn & currentRow).value = "Priority: SEAL > CRITICAL > TUBI > MAPRESS > STANDARD"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 2
    
    ' SEAL RING RULES
    ws.Range(ruleColumn & currentRow).value = "SEAL RING COMPONENTS (require: seal, ring, FKM; exclude: elbow, flange, tee, reducer)"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #1: If QTY <= 10 -> Spare = 5 (absolute minimum)"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #2: If QTY > 10 -> Spare = max(5, ceiling(QTY x 10%))"
    currentRow = currentRow + 2
    
    ' CRITICAL RULES
    ws.Range(ruleColumn & currentRow).value = "CRITICAL COMPONENTS (keywords: nipple, union, connection, welding, adaptor)"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #3: If QTY <= 10 -> Spare = max(3, ceiling(QTY x 50%))"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #4: If QTY > 10 -> Spare = max(3, ceiling(QTY x 10%))"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Note: Minimum 3 spares always enforced for critical components"
    ws.Range(ruleColumn & currentRow).Font.Italic = True
    currentRow = currentRow + 2

    ' TUBI RULES
    ws.Range(ruleColumn & currentRow).value = "TUBI COMPONENTS (Type column = 'TUBI'; qty rounded to multiples of 6)"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #T1: Base spare = ceiling(QTY x 15%) if DN < 50, else ceiling(QTY x 10%);"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "             add to original QTY, round total up to the next multiple of 6,"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "             spare = rounded total - original QTY"
    currentRow = currentRow + 2
    
    ' MAPRESS RULES
    ws.Range(ruleColumn & currentRow).value = "MAPRESS FITTINGS (keywords: mapress + cap/coupling/elbow/flange/tee)"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #5: If QTY <= 5 -> Spare = 0"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #6: If 5 < QTY <= 20 -> Spare = 1"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #7: If QTY > 20 -> Spare = ceiling(QTY x 10%)"
    currentRow = currentRow + 2
    
    ' STANDARD RULES
    ws.Range(ruleColumn & currentRow).value = "STANDARD COMPONENTS (all others)"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    
    ws.Range(ruleColumn & currentRow).value = "  Rule #8: If QTY <= 5 -> Spare = 0"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #9: If 5 < QTY <= 20 -> Spare = 1"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #10a: If QTY pcs > 20 and DN < 50 -> Spare = ceiling(QTY x 10%)"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #10b: If QTY pcs > 20 and DN >= 50 -> Spare = ceiling(QTY x 5%)"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #11a: If QTY m > 0 and DN < 50 -> Spare = ceiling(QTY x 15%)"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  Rule #11b: If QTY m > 0 and DN >= 50 -> Spare = ceiling(QTY x 10%)"
    currentRow = currentRow + 2
    
    ' Notes
    ws.Range(ruleColumn & currentRow).value = "NOTES:"
    ws.Range(ruleColumn & currentRow).Font.Bold = True
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  - QTY pcs (pieces) takes priority over QTY m (meters)"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  - All percentages are rounded UP to next integer (ceiling function)"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  - DN value is extracted from column C (e.g., '50mm' -> 50)"
    currentRow = currentRow + 1
    ws.Range(ruleColumn & currentRow).value = "  - Component classification is case-insensitive"
    
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

Function IsSealComponent(description As String, requiredKeywords As String, excludedKeywords As String) As Boolean
    If Len(Trim(description)) = 0 Then Exit Function
    If Not HasAllKeywords(description, requiredKeywords) Then Exit Function
    If HasForbiddenKeyword(description, excludedKeywords) Then Exit Function
    IsSealComponent = True
End Function

Function IsTubiComponent(typeLabel As String, typeCellValue As String, description As String) As Boolean
    If ContainsKeyword(typeLabel, "tubi") Then
        IsTubiComponent = True
        Exit Function
    End If
    If ContainsKeyword(typeCellValue, "tubi") Then
        IsTubiComponent = True
        Exit Function
    End If
    If ContainsKeyword(description, "tubi") Then
        IsTubiComponent = True
    End If
End Function

Function HasMapressWithFitting(description As String, mapressWord As String, fittingWords As String) As Boolean
    Dim descLower As String
    HasMapressWithFitting = False
    If Len(Trim(description)) = 0 Then Exit Function
    descLower = LCase(Trim(description))
    If InStr(1, descLower, LCase(Trim(mapressWord)), vbTextCompare) = 0 Then Exit Function
    If Not HasAnyKeyword(description, fittingWords) Then Exit Function
    HasMapressWithFitting = True
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

Function HasAllKeywords(description As String, keywordList As String) As Boolean
    Dim keywords() As String
    Dim keyword As Variant
    Dim descLower As String
    Dim token As String
    HasAllKeywords = False
    If Len(Trim(description)) = 0 Then Exit Function
    If Len(Trim(keywordList)) = 0 Then Exit Function
    descLower = LCase(Trim(description))
    keywords = Split(keywordList, ",")
    For Each keyword In keywords
        token = LCase(Trim(CStr(keyword)))
        If Len(token) = 0 Then GoTo ContinueLoop
        If InStr(1, descLower, token, vbTextCompare) = 0 Then Exit Function
ContinueLoop:
    Next keyword
    HasAllKeywords = True
End Function

Function HasForbiddenKeyword(description As String, keywordList As String) As Boolean
    Dim keywords() As String
    Dim keyword As Variant
    Dim descLower As String
    HasForbiddenKeyword = False
    If Len(Trim(description)) = 0 Then Exit Function
    If Len(Trim(keywordList)) = 0 Then Exit Function
    descLower = LCase(Trim(description))
    keywords = Split(keywordList, ",")
    For Each keyword In keywords
        If Len(Trim(CStr(keyword))) = 0 Then GoTo NextKeyword
        If InStr(1, descLower, LCase(Trim(CStr(keyword))), vbTextCompare) > 0 Then
            HasForbiddenKeyword = True
            Exit Function
        End If
NextKeyword:
    Next keyword
End Function

Function ContainsKeyword(textValue As String, keyword As String) As Boolean
    ContainsKeyword = False
    If Len(Trim(textValue)) = 0 Then Exit Function
    If Len(Trim(keyword)) = 0 Then Exit Function
    If InStr(1, textValue, keyword, vbTextCompare) > 0 Then
        ContainsKeyword = True
    End If
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

Function ValidateParameters(headerRow As Long, threshold1 As Long, threshold2 As Long) As Boolean
    ValidateParameters = True
    If headerRow < 1 Then
        MsgBox "Error: HEADER_ROW must be at least 1." & vbCrLf & "Current value: " & headerRow, vbCritical, "Invalid Parameter"
        ValidateParameters = False
        Exit Function
    End If
    If threshold1 >= threshold2 Then
        MsgBox "Error: THRESHOLD_1 must be less than THRESHOLD_2." & vbCrLf & "Current values: THRESHOLD_1=" & threshold1 & ", THRESHOLD_2=" & threshold2, vbCritical, "Invalid Parameter"
        ValidateParameters = False
        Exit Function
    End If
    If threshold1 < 0 Or threshold2 < 0 Then
        MsgBox "Error: Thresholds must be non-negative." & vbCrLf & "Current values: THRESHOLD_1=" & threshold1 & ", THRESHOLD_2=" & threshold2, vbCritical, "Invalid Parameter"
        ValidateParameters = False
        Exit Function
    End If
End Function

Function FindLastDataRow(ws As Worksheet, headerRow As Long, colA As String, colB As String, colC As String, colD As String) As Long
    Dim i As Long
    Dim maxRow As Long
    
    ' Find the maximum row with any data in the checked columns
    maxRow = ws.Cells(ws.Rows.Count, colA).End(xlUp).Row
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, colB).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, colC).End(xlUp).Row)
    maxRow = Application.WorksheetFunction.Max(maxRow, ws.Cells(ws.Rows.Count, colD).End(xlUp).Row)
    
    If maxRow <= headerRow Then
        FindLastDataRow = headerRow
        Exit Function
    End If
    
    ' Search backwards from maxRow to find the last row where ALL columns A, B, C, D have values
    For i = maxRow To headerRow + 1 Step -1
        If Len(Trim(CStr(ws.Cells(i, colA).Value))) > 0 And _
           Len(Trim(CStr(ws.Cells(i, colB).Value))) > 0 And _
           Len(Trim(CStr(ws.Cells(i, colC).Value))) > 0 And _
           Len(Trim(CStr(ws.Cells(i, colD).Value))) > 0 Then
            FindLastDataRow = i
            Exit Function
        End If
    Next i
    
    ' If no row found with all columns filled, return header row
    FindLastDataRow = headerRow
End Function

Function ExtractNumericFromType(typeCell As Variant) As Double
    Dim cleanedString As String
    Dim i As Integer
    Dim char As String
    Dim result As String
    
    If IsEmpty(typeCell) Then
        ExtractNumericFromType = 0
        Exit Function
    End If
    
    If IsError(typeCell) Then
        ExtractNumericFromType = 0
        Exit Function
    End If
    
    cleanedString = CStr(typeCell)
    result = ""
    
    For i = 1 To Len(cleanedString)
        char = Mid(cleanedString, i, 1)
        If IsNumeric(char) Or char = "." Or (char = "-" And i = 1) Then
            result = result & char
        ElseIf result <> "" Then
            Exit For
        End If
    Next i
    
    If result = "" Or result = "." Or result = "-" Then
        ExtractNumericFromType = 0
    Else
        On Error Resume Next
        ExtractNumericFromType = CDbl(result)
        If Err.Number <> 0 Then
            ExtractNumericFromType = 0
            Err.Clear
        End If
        On Error GoTo 0
    End If
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

Function RoundUpToMultiple(value As Double, multiple As Long) As Long
    If multiple <= 0 Then
        RoundUpToMultiple = RoundUpToInteger(value)
        Exit Function
    End If
    If value <= 0 Then
        RoundUpToMultiple = 0
        Exit Function
    End If
    RoundUpToMultiple = CLng(Application.WorksheetFunction.Ceiling_Math(value, multiple))
End Function


