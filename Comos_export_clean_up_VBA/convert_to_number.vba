Option Explicit

Public Sub ConvertTextNumbersInActiveSheet()
	Dim ws As Worksheet
	Dim rng As Range
	Dim cell As Range
	Dim convertedCount As Long
	Dim skippedCount As Long
	Dim parsedValue As Double
	Dim decimalPlaces As Long
	Dim usePercentFormat As Boolean
	Dim fmt As String
	Dim oldCalc As XlCalculation
	Dim oldScreenUpdating As Boolean

	Set ws = ActiveSheet

	On Error Resume Next
	Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues)
	On Error GoTo 0

	If rng Is Nothing Then
		MsgBox "No text constants found on the active sheet.", vbInformation
		Exit Sub
	End If

	oldCalc = Application.Calculation
	oldScreenUpdating = Application.ScreenUpdating
	Application.ScreenUpdating = False
	Application.Calculation = xlCalculationManual
	On Error GoTo SafeExit

	For Each cell In rng.Cells
		If TryParseNumericText(CStr(cell.Value2), parsedValue, decimalPlaces, usePercentFormat) Then
			cell.Value2 = parsedValue

			fmt = "0"
			If decimalPlaces > 0 Then
				fmt = fmt & "." & String$(decimalPlaces, "0")
			End If
			If usePercentFormat Then
				fmt = fmt & "%"
			End If
			cell.NumberFormat = fmt

			convertedCount = convertedCount + 1
		Else
			skippedCount = skippedCount + 1
		End If
	Next cell

	Application.Calculation = oldCalc
	Application.ScreenUpdating = oldScreenUpdating

	MsgBox "Conversion complete on sheet '" & ws.Name & "'." & vbCrLf & _
		   "Converted: " & convertedCount & vbCrLf & _
		   "Skipped: " & skippedCount, vbInformation
	Exit Sub

SafeExit:
	Application.Calculation = oldCalc
	Application.ScreenUpdating = oldScreenUpdating
	MsgBox "Error during conversion: " & Err.Description, vbExclamation
End Sub

Private Function TryParseNumericText(ByVal inputText As String, _
									 ByRef outValue As Double, _
									 ByRef outDecimalPlaces As Long, _
									 ByRef outUsePercentFormat As Boolean) As Boolean
	Dim s As String
	Dim isNegative As Boolean
	Dim hasPlus As Boolean
	Dim dotPos As Long
	Dim commaPos As Long
	Dim decSepInText As String
	Dim groupSeps As String
	Dim lastSepPos As Long
	Dim ePos As Long
	Dim mantissaPart As String
	Dim exponentPart As String
	Dim hasExponent As Boolean
	Dim exponentPower As Long
	Dim hasPercent As Boolean
	Dim i As Long
	Dim ch As String
	Dim normalized As String
	Dim localDecSep As String
	Dim digitsRight As Long
	Dim digitsLeft As Long
	Dim sepCount As Long
	Dim normalizedFractionDigits As Long
	Dim sigDigits As Long

	TryParseNumericText = False
	outDecimalPlaces = 0
	outUsePercentFormat = False

	s = NormalizeUnicodeNumericChars(inputText)
	s = Replace(s, ChrW$(160), "")
	s = Replace(s, ChrW$(8239), "")
	s = Replace(s, ChrW$(8201), "")
	s = Replace(s, ChrW$(8200), "")
	s = Replace(s, ChrW$(9), "")
	s = Replace(s, " ", "")
	s = Trim$(s)

	If Len(s) = 0 Then Exit Function
	s = StripOuterCurrencySymbols(s)
	If Len(s) = 0 Then Exit Function

	If Left$(s, 1) = "'" Then
		s = Mid$(s, 2)
	End If

	If Len(s) = 0 Then Exit Function

	If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
		isNegative = True
		s = Mid$(s, 2, Len(s) - 2)
	End If

	If Len(s) = 0 Then Exit Function

	If Left$(s, 1) = "+" Then
		hasPlus = True
		s = Mid$(s, 2)
	ElseIf Left$(s, 1) = "-" Then
		isNegative = True
		s = Mid$(s, 2)
	ElseIf Right$(s, 1) = "-" Then
		isNegative = True
		s = Left$(s, Len(s) - 1)
	End If

	If Len(s) = 0 Then Exit Function
	If hasPlus And isNegative Then Exit Function

	s = StripOuterCurrencySymbols(s)
	If Len(s) = 0 Then Exit Function

	If Right$(s, 1) = "%" Then
		hasPercent = True
		s = Left$(s, Len(s) - 1)
	End If

	If Len(s) = 0 Then Exit Function
	If HasInvalidAlphaNumericMix(s) Then Exit Function

	ePos = InStrRev(s, "E")
	If ePos = 0 Then ePos = InStrRev(s, "e")
	If ePos > 0 Then
		mantissaPart = Left$(s, ePos - 1)
		exponentPart = Mid$(s, ePos + 1)
		If Len(mantissaPart) = 0 Or Not IsValidExponentPart(exponentPart) Then Exit Function
		hasExponent = True
		exponentPower = CLng(exponentPart)
		s = mantissaPart
	End If

	For i = 1 To Len(s)
		ch = Mid$(s, i, 1)
		If Not (IsAsciiDigitChar(ch) Or ch = "." Or ch = "," Or ch = "'" Or ch = "_") Then
			Exit Function
		End If
	Next i

	dotPos = InStrRev(s, ".")
	commaPos = InStrRev(s, ",")

	decSepInText = ""
	groupSeps = "'_"
	lastSepPos = GetLastPosOfAny(s, ".,")

	If dotPos > 0 And commaPos > 0 Then
		If dotPos > commaPos Then
			decSepInText = "."
			groupSeps = groupSeps & ","
		Else
			decSepInText = ","
			groupSeps = groupSeps & "."
		End If
	ElseIf dotPos > 0 Or commaPos > 0 Then
		If dotPos > 0 Then
			decSepInText = "."
		Else
			decSepInText = ","
		End If

		sepCount = CountChar(s, decSepInText)
		digitsRight = CountDigitsInString(Mid$(s, lastSepPos + 1))
		digitsLeft = CountDigitsInString(Left$(s, lastSepPos - 1))

		If sepCount = 1 Then
			If digitsRight = 3 And digitsLeft >= 1 Then
				Exit Function
			End If
		ElseIf sepCount > 1 Then
			If digitsRight = 3 Then
				groupSeps = groupSeps & decSepInText
				decSepInText = ""
			End If
		End If
	End If

	normalized = s
	normalized = RemoveSeparators(normalized, groupSeps)

	If decSepInText <> "" Then
		normalizedFractionDigits = CountDigitsInString(Mid$(s, lastSepPos + 1))
		outDecimalPlaces = normalizedFractionDigits
	Else
		outDecimalPlaces = 0
	End If

	If decSepInText <> "" Then
		If CountChar(normalized, decSepInText) > 1 Then Exit Function
	End If

	If Not HasAnyDigit(normalized) Then Exit Function
	sigDigits = SignificantDigitCount(normalized)
	If sigDigits > 15 Then Exit Function

	localDecSep = Application.International(xlDecimalSeparator)

	If decSepInText <> "" Then
		If decSepInText <> localDecSep Then
			normalized = Replace(normalized, decSepInText, localDecSep)
		End If
	End If

	On Error GoTo ParseFail
	outValue = CDbl(normalized)
	If hasExponent Then
		outValue = outValue * (10# ^ exponentPower)
	End If
	If hasPercent Then
		outValue = outValue / 100#
		outUsePercentFormat = True
	End If
	If isNegative Then outValue = -Abs(outValue)

	TryParseNumericText = True
	Exit Function

ParseFail:
	TryParseNumericText = False
End Function

Private Function NormalizeUnicodeNumericChars(ByVal s As String) As String
	Dim i As Long
	Dim ch As String
	Dim cp As Long
	Dim resultText As String

	resultText = ""

	For i = 1 To Len(s)
		ch = Mid$(s, i, 1)
		cp = AscW(ch)
		If cp < 0 Then cp = cp + 65536

		Select Case cp
			Case &H660 To &H669
				resultText = resultText & Chr$(&H30 + (cp - &H660))
			Case &H6F0 To &H6F9
				resultText = resultText & Chr$(&H30 + (cp - &H6F0))
			Case &H66B
				resultText = resultText & "."
			Case &H66C
				resultText = resultText & ","
			Case Else
				resultText = resultText & ch
		End Select
	Next i

	NormalizeUnicodeNumericChars = resultText
End Function

Private Function StripOuterCurrencySymbols(ByVal s As String) As String
	Dim symbols As String
	Dim changed As Boolean

	symbols = "$€£¥₹₽₩₪₫₴₦฿₺₱₨₲₡₵¢"
	StripOuterCurrencySymbols = s

	Do
		changed = False
		If Len(StripOuterCurrencySymbols) > 0 Then
			If InStr(1, symbols, Left$(StripOuterCurrencySymbols, 1), vbBinaryCompare) > 0 Then
				StripOuterCurrencySymbols = Mid$(StripOuterCurrencySymbols, 2)
				changed = True
			End If
		End If

		If Len(StripOuterCurrencySymbols) > 0 Then
			If InStr(1, symbols, Right$(StripOuterCurrencySymbols, 1), vbBinaryCompare) > 0 Then
				StripOuterCurrencySymbols = Left$(StripOuterCurrencySymbols, Len(StripOuterCurrencySymbols) - 1)
				changed = True
			End If
		End If
	Loop While changed
End Function

Private Function HasInvalidAlphaNumericMix(ByVal s As String) As Boolean
	Dim i As Long
	Dim ch As String
	Dim hasDigit As Boolean
	Dim hasOtherLetter As Boolean

	For i = 1 To Len(s)
		ch = Mid$(s, i, 1)
		If ch Like "[0-9]" Then
			hasDigit = True
		ElseIf ch Like "[A-Za-z]" Then
			If ch <> "E" And ch <> "e" Then
				hasOtherLetter = True
			End If
		End If

		If hasDigit And hasOtherLetter Then
			HasInvalidAlphaNumericMix = True
			Exit Function
		End If
	Next i

	HasInvalidAlphaNumericMix = False
End Function

Private Function IsValidExponentPart(ByVal s As String) As Boolean
	Dim i As Long
	Dim ch As String

	IsValidExponentPart = False
	If Len(s) = 0 Then Exit Function

	If Left$(s, 1) = "+" Or Left$(s, 1) = "-" Then
		s = Mid$(s, 2)
	End If

	If Len(s) = 0 Then Exit Function

	For i = 1 To Len(s)
		ch = Mid$(s, i, 1)
		If Not (ch Like "[0-9]") Then Exit Function
	Next i

	IsValidExponentPart = True
End Function

Private Function IsAsciiDigitChar(ByVal ch As String) As Boolean
	IsAsciiDigitChar = (ch Like "[0-9]")
End Function

Private Function CountChar(ByVal s As String, ByVal target As String) As Long
	Dim i As Long

	CountChar = 0
	If Len(target) = 0 Then Exit Function

	For i = 1 To Len(s)
		If Mid$(s, i, 1) = target Then
			CountChar = CountChar + 1
		End If
	Next i
End Function

Private Function CountDigitsInString(ByVal s As String) As Long
	Dim i As Long
	Dim ch As String

	CountDigitsInString = 0
	For i = 1 To Len(s)
		ch = Mid$(s, i, 1)
		If IsAsciiDigitChar(ch) Then
			CountDigitsInString = CountDigitsInString + 1
		End If
	Next i
End Function

Private Function SignificantDigitCount(ByVal s As String) As Long
	Dim i As Long
	Dim ch As String
	Dim digitsOnly As String

	digitsOnly = ""
	For i = 1 To Len(s)
		ch = Mid$(s, i, 1)
		If ch Like "[0-9]" Then
			digitsOnly = digitsOnly & ch
		End If
	Next i

	Do While Len(digitsOnly) > 0 And Left$(digitsOnly, 1) = "0"
		digitsOnly = Mid$(digitsOnly, 2)
	Loop

	If Len(digitsOnly) = 0 Then
		SignificantDigitCount = 1
	Else
		SignificantDigitCount = Len(digitsOnly)
	End If
End Function

Private Function HasAnyDigit(ByVal s As String) As Boolean
	HasAnyDigit = (CountDigitsInString(s) > 0)
End Function

Private Function GetLastPosOfAny(ByVal s As String, ByVal chars As String) As Long
	Dim i As Long
	Dim pos As Long

	GetLastPosOfAny = 0
	For i = 1 To Len(chars)
		pos = InStrRev(s, Mid$(chars, i, 1))
		If pos > GetLastPosOfAny Then
			GetLastPosOfAny = pos
		End If
	Next i
End Function

Private Function RemoveSeparators(ByVal s As String, ByVal chars As String) As String
	Dim i As Long

	RemoveSeparators = s
	For i = 1 To Len(chars)
		RemoveSeparators = Replace(RemoveSeparators, Mid$(chars, i, 1), "")
	Next i
End Function

