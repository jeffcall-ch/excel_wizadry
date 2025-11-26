Option Explicit

' Module for cascading dropdowns to lookup data from Spans sheet
' NOTE: The Worksheet_Change event is in the Calcs sheet code, not here

Sub InitializeDropdowns()
    ' This sub sets up the initial dropdowns
    ' Run this once to set up the system
    
    Dim ws As Worksheet
    Dim spansSheet As Worksheet
    Dim materialRange As Range
    
    Set ws = ThisWorkbook.Sheets("Calcs")
    Set spansSheet = ThisWorkbook.Sheets("Spans_with_Mapress_from TV046")
    
    Application.ScreenUpdating = False
    
    ' Add labels
    ws.Range("E2").Value = "Material:"
    ws.Range("E3").Value = "DN [mm]:"
    ws.Range("E4").Value = "Mass Type:"
    ws.Range("E5").Value = "Result [kg]:"
    
    ' Clear dropdown cells
    ws.Range("F2:F5").ClearContents
    ws.Range("F2:F4").Validation.Delete
    
    ' Set up Material dropdown in F2 (CS, SS, Mapress)
    With ws.Range("F2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="CS,SS,Mapress"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' Set up Mass Type dropdown in F4 (4 options)
    With ws.Range("F4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Gas w/o insulation,Gas w insulation,Liquid w/o insulation,Liquid w insulation"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "Dropdowns initialized successfully!", vbInformation
End Sub

Sub UpdateDNDropdown()
    ' Updates the DN dropdown based on selected Material
    
    Dim ws As Worksheet
    Dim spansSheet As Worksheet
    Dim material As String
    Dim dnList As String
    Dim i As Long
    Dim lastRow As Long
    Dim dnValues As Collection
    Dim dnVal As Variant
    
    Set ws = ThisWorkbook.Sheets("Calcs")
    Set spansSheet = ThisWorkbook.Sheets("Spans_with_Mapress_from TV046")
    
    material = ws.Range("F2").Value
    
    If material = "" Then
        ws.Range("F3").Validation.Delete
        ws.Range("F3").ClearContents
        ws.Range("F5").ClearContents
        Exit Sub
    End If
    
    ' Find last row in Spans sheet
    lastRow = spansSheet.Cells(spansSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Collect unique DN values for selected material
    Set dnValues = New Collection
    On Error Resume Next
    For i = 2 To lastRow ' Start from row 2 (skip header)
        If spansSheet.Cells(i, 1).Value = material Then
            dnValues.Add spansSheet.Cells(i, 2).Value, CStr(spansSheet.Cells(i, 2).Value)
        End If
    Next i
    On Error GoTo 0
    
    ' Build comma-separated list
    dnList = ""
    For Each dnVal In dnValues
        If dnList = "" Then
            dnList = CStr(dnVal)
        Else
            dnList = dnList & "," & CStr(dnVal)
        End If
    Next dnVal
    
    ' Clear F3 and F5
    ws.Range("F3").ClearContents
    ws.Range("F5").ClearContents
    
    ' Set up DN dropdown in F3
    If dnList <> "" Then
        With ws.Range("F3").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:=dnList
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End If
End Sub

Sub UpdateResult()
    ' Updates the result in F5 based on all three selections
    
    Dim ws As Worksheet
    Dim spansSheet As Worksheet
    Dim material As String
    Dim dn As Variant
    Dim massType As String
    Dim i As Long
    Dim lastRow As Long
    Dim colIndex As Integer
    Dim result As Variant
    
    Set ws = ThisWorkbook.Sheets("Calcs")
    Set spansSheet = ThisWorkbook.Sheets("Spans_with_Mapress_from TV046")
    
    material = ws.Range("F2").Value
    dn = ws.Range("F3").Value
    massType = ws.Range("F4").Value
    
    ' Clear result if any selection is missing
    If material = "" Or dn = "" Or massType = "" Then
        ws.Range("F5").ClearContents
        Exit Sub
    End If
    
    ' Determine column based on mass type
    Select Case massType
        Case "Gas w/o insulation"
            colIndex = 6 ' Column F: Mass per support [kg] Gas w/o insulation
        Case "Gas w insulation"
            colIndex = 9 ' Column I: Mass per support [kg] Gas w insulation
        Case "Liquid w/o insulation"
            colIndex = 12 ' Column L: Mass per support [kg] Liquid w/o insulation
        Case "Liquid w insulation"
            colIndex = 15 ' Column O: Mass per support [kg] Liquid w insulation
        Case Else
            ws.Range("F5").Value = "Invalid mass type"
            Exit Sub
    End Select
    
    ' Find the matching row
    lastRow = spansSheet.Cells(spansSheet.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If spansSheet.Cells(i, 1).Value = material And _
           spansSheet.Cells(i, 2).Value = dn Then
            result = spansSheet.Cells(i, colIndex).Value
            ws.Range("F5").Value = result
            Exit Sub
        End If
    Next i
    
    ' If no match found
    ws.Range("F5").Value = "Not found"
End Sub
