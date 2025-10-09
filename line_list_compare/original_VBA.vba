
Sub vergleich()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Call inhalt_prüfen
Call erste_spalte_bearb 'wenn werte in Spalte A zweimal vorkommen, kann es eine Fehlermeldung geben, weil dann zweimal in der gleichen zeile gearbeitet wird => Zelle hat bereits kommentar
Call neue_tabellen


sprache = Application.LanguageSettings.LanguageID(msoLanguageIDUI)

letztespalte = Cells(4, 16000).End(xlToLeft).Column

If letztespalte < 7 Then letztespalte = 8
For anz = 1 To 2
    x = ""
    neu_sp = 0
        If anz = 1 Then
            Set ws1 = ThisWorkbook.Sheets("Old Table")
            Set ws2 = ThisWorkbook.Sheets("New Table")
            Set ws3 = ThisWorkbook.Sheets("Comparison, Base is OLD Table")
            Set ws4 = ThisWorkbook.Sheets("Comparison, Base is NEW Table")
            ws3.UsedRange.Interior.Color = &HC0C0FF
        Else
            Set ws2 = ThisWorkbook.Sheets("Old Table")
            Set ws1 = ThisWorkbook.Sheets("New Table")
            Set ws4 = ThisWorkbook.Sheets("Comparison, Base is OLD Table")
            Set ws3 = ThisWorkbook.Sheets("Comparison, Base is NEW Table")
            ws3.UsedRange.Interior.Color = vbGreen
        End If
    
    
    ws3.Rows(1).Interior.Color = vbBlack
    If sprache = 1031 Then
        ws3.Range("B1").Value = "Vergleich: Grundlage ist NEUE Tabelle"
        ws3.Range("E1").Value = "Geändert"
        ws3.Range("F1").Value = "Hinzugef."
        ws3.Range("G1").Value = "Gelöscht"
    Else
        ws3.Range("B1").Value = "Comparison, base is the NEW table"
        ws3.Range("E1").Value = "Changed"
        ws3.Range("F1").Value = "Added"
        ws3.Range("G1").Value = "Deleted"
    End If
    ws3.Range("E1").Interior.Color = vbYellow
    ws3.Range("F1").Interior.Color = vbGreen
    ws3.Range("G1").Interior.Color = &HC0C0FF
    ws3.Range("E1:G1").Font.ColorIndex = xlAutomatic
    
    Letze_z_alt = ws1.Cells(Rows.Count, 1).End(xlUp).Row
    Letze_z_neu = ws2.Cells(Rows.Count, 1).End(xlUp).Row
    letze_sp_alt = ws1.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    Letze_sp_neu = ws2.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
    For i = 2 To Letze_z_neu
    
    Set Found = ws1.Columns(1).Find(ws2.Cells(i, 1), LookIn:=xlValues, LookAt:=xlWhole)
    
    If Found Is Nothing Then
        'MsgBox Cells(i, 1) & " nicht gefunden!", vbInformation, "Meldung"
        If i = 1 Then GoTo 33
        neu_sp = neu_sp + 1
        If ws1.Name = "Old Table" Then
            ws3.Cells(Letze_z_alt + 2, 1).Value = "Not found in OLD Table => Row was added in New"
            ws3.Cells(Letze_z_alt + 2, 5).Value = "Nicht in ALTER Tabelle gefunden => Zeile wurde in Neu hinzugefügt"
            ws2.Rows(i).Copy Destination:=ws3.Rows(Letze_z_alt + neu_sp + 2)
            ws3.Rows(Letze_z_alt + neu_sp + 2).Interior.Color = vbGreen
            ws3.Cells(Letze_z_alt + neu_sp + 2, letztespalte + 2).Value = "N"
        Else
            ws3.Cells(Letze_z_alt + 2, 1).Value = "Not found in NEW Table => Row was deleted in New"
            ws3.Cells(Letze_z_alt + 2, 5).Value = "Nicht in NEUER Tabelle gefunden => Zeile wurde in Neu gelöscht"
            ws2.Rows(i).Copy Destination:=ws3.Rows(Letze_z_alt + neu_sp + 2)
            ws3.Rows(Letze_z_alt + neu_sp + 2).Interior.Color = &HC0C0FF
            ws3.Cells(Letze_z_alt + neu_sp + 2, letztespalte + 2).Value = "D"
        End If
        
        ws3.Rows(Letze_z_alt + 2).Font.Bold = True
        ws3.Rows(Letze_z_alt + 2).Font.ThemeColor = xlThemeColorDark1
        ws3.Rows(Letze_z_alt + 2).Interior.Color = 10498160
        
    Else
        ws3.Rows(Found.Row).Interior.Pattern = xlNone
        For x = 2 To letze_sp_alt
            
            If ws3.Cells(Found.Row, x).Text = ws2.Cells(i, x).Text Then
                GoTo 55
            ElseIf ws3.Cells(Found.Row, x).Text = "" And ws2.Cells(i, x).Text <> "" Then
                If ws1.Name = "Old Table" Then
                    ws3.Cells(Found.Row, x).Interior.Color = vbGreen
                    ws3.Cells(Found.Row, letztespalte + 3).Value = "N"
                Else
                    ws3.Cells(Found.Row, x).Interior.Color = &HC0C0FF
                    ws3.Cells(Found.Row, letztespalte + 5).Value = "D"
                End If
                Cells(Found.Row, x).Select
                 With ws3.Cells(Found.Row, x)
                    .AddComment
                    .Comment.Visible = False
                    .Comment.Text Text:=ws2.Name & " value: " & vbCrLf & ws2.Cells(i, x).Text
                End With
            ElseIf ws3.Cells(Found.Row, x).Text <> "" And ws2.Cells(i, x).Text = "" Then
                If ws1.Name = "Old Table" Then
                    ws3.Cells(Found.Row, x).Interior.Color = &HC0C0FF
                    ws3.Cells(Found.Row, letztespalte + 5).Value = "D"
                Else
                    ws3.Cells(Found.Row, x).Interior.Color = vbGreen
                    ws3.Cells(Found.Row, letztespalte + 3).Value = "N"
                End If
            Else
                ws3.Cells(Found.Row, x).Interior.Color = vbYellow
                ws3.Cells(Found.Row, letztespalte + 4).Value = "Ch"
                Cells(Found.Row, x).Select
                If Found.Row = 447 And x = 3 Then
                    d = ii
                End If
                With ws3.Cells(Found.Row, x)
                    .AddComment
                    .Comment.Visible = False
                    .Comment.Text Text:=ws2.Name & " value: " & vbCrLf & ws2.Cells(i, x).Text
                End With
            End If
55
        Next x
    End If
33
    If Found Is Nothing Then
    Else
    ws3.Cells(Found.Row, letztespalte + 2).Value = ws3.Cells(Found.Row, letztespalte + 5).Text & ws3.Cells(Found.Row, letztespalte + 4).Text & ws3.Cells(Found.Row, letztespalte + 3).Text
    ws3.Cells(Found.Row, letztespalte + 5).Value = ""
    ws3.Cells(Found.Row, letztespalte + 4).Value = ""
    ws3.Cells(Found.Row, letztespalte + 3).Value = ""
    End If
    Next
    

Next anz

    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.AutoFilter
    
    x = 3
    Do While Cells(x, 1).Value <> ""
        x = x + 1
        If ActiveSheet.Cells(x, 1).Interior.Color = vbGreen Then
            If InStr(1, ActiveSheet.Cells(x, letztespalte + 2).Value, "N") = 0 Then
                ActiveSheet.Cells(x, letztespalte + 2).Value = ActiveSheet.Cells(x, letztespalte + 2).Text & "N"
            End If
        ElseIf Cells(x, 1).Interior.Color = &HC0C0FF Then
            If InStr(1, ActiveSheet.Cells(x, letztespalte + 2).Value, "D") = 0 Then
                ActiveSheet.Cells(x, letztespalte + 2).Value = ActiveSheet.Cells(x, letztespalte + 2).Text & "D"
            End If
        End If
    Loop
    ActiveSheet.Cells(1, letztespalte + 2).Value = "Geändert"
    ActiveSheet.Cells(1, letztespalte + 2).Font.Color = vbBlack
    ActiveSheet.Columns(letztespalte + 2).Interior.Color = vbYellow
    
Sheets("Comparison, Base is NEW Table").Select

    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.AutoFilter
    Range("A3").Select
    
    x = 3
    Do While Cells(x, 1).Value <> ""
        x = x + 1
        If ActiveSheet.Cells(x, 1).Interior.Color = vbGreen Then
            If InStr(1, ActiveSheet.Cells(x, letztespalte + 2).Value, "N") = 0 Then
                ActiveSheet.Cells(x, letztespalte + 2).Value = ActiveSheet.Cells(x, letztespalte + 2).Text & "N"
            End If
        ElseIf Cells(x, 1).Interior.Color = &HC0C0FF Then
            If InStr(1, ActiveSheet.Cells(x, letztespalte + 2).Value, "D") = 0 Then
                ActiveSheet.Cells(x, letztespalte + 2).Value = ActiveSheet.Cells(x, letztespalte + 2).Text & "D"
            End If
        End If
    Loop
    ActiveSheet.Cells(1, letztespalte + 2).Value = "Geändert"
    ActiveSheet.Cells(1, letztespalte + 2).Font.Color = vbBlack
    ActiveSheet.Columns(letztespalte + 2).Interior.Color = vbYellow
    
ws1.Columns("A:A").Delete Shift:=xlToLeft
ws2.Columns("A:A").Delete Shift:=xlToLeft
ws3.Columns("A:A").Delete Shift:=xlToLeft
ws4.Columns("A:A").Delete Shift:=xlToLeft


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub neue_tabellen()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Application.DisplayAlerts = False
    For Each Sheet In ThisWorkbook.Sheets
        If Sheet.Name = "Comparison, Base is OLD Table" Then
            Sheets("Comparison, Base is OLD Table").Delete
        ElseIf Sheet.Name = "Comparison, Base is NEW Table" Then
            Sheets("Comparison, Base is NEW Table").Delete
        End If
    Next
    Application.DisplayAlerts = True

    
    Sheets("New Table").Copy after:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Comparison, Base is NEW Table"
    Range("L1").Value = ""
    Range("o1").Value = ""
    
    Sheets("Old Table").Copy after:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Comparison, Base is OLD Table"
        Range("L1").Value = ""
    Range("o1").Value = ""
'    Range("A1").value = "Comparison, base is the OLD table"
'    Range("D1").value = "Vergleich mit ALTER Tabelle als Grundlage"
'        Range("I1:L1").Interior.Color = vbYellow
'    Range("I1").value = "Changed"
    Range("I1:L1").Font.ColorIndex = xlAutomatic
    'Range("k1").value = "Geändert"

    Range("n1:q1").Interior.Color = vbGreen
    'Range("n1").value = "Added"
    Range("n1:q1").Font.ColorIndex = xlAutomatic
   ' Range("p1").value = "Hinzugef."
    
    Range("s1:v1").Interior.Color = &HC0C0FF
  '  Range("s1").value = "Deleted"
    Range("s1:v1").Font.ColorIndex = xlAutomatic
   ' Range("u1").value = "Gelöscht"
    Range("a1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub inhalt_prüfen()
    If Sheets("Old Table").Range("A2").Value = "" Then
        MsgBox "Old Table is empty => please fill in (First Value of Header Cell in A2)" & vbCrLf & vbCrLf & _
            "Old Table ist leer => bitte mit Werten füllen (Erster Wert der Titelzeile in A2)"
            End
    ElseIf Sheets("New Table").Range("A2").Value = "" Then
        MsgBox "New Table is empty => please fill in (First Value of Header Cell in A2)" & vbCrLf & vbCrLf & _
            "New Table ist leer => bitte mit Werten füllen (Erster Wert der Titelzeile in Zelle A2)"
            End
    End If
    
    Set ws1 = ThisWorkbook.Sheets("Old Table")
    Set ws2 = ThisWorkbook.Sheets("New Table")
    letze_sp_alt = ws1.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    Letze_sp_neu = ws2.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
    For x = 1 To letze_sp_alt
        If ws1.Cells(2, x).Value <> ws2.Cells(2, x).Value Then
            MsgBox "Header (Row2) in Old and New table is not identical => please adapt" & vbCrLf & vbCrLf & _
            "Titelzeile (Zeile2) in Old und New ist nicht identisch => bitte entsprechend anpassen"
            End
        End If
    Next x
    
    'falls noch kommentare vorhanden sind, so werden die gelöscht, damit
    Sheets("New Table").Cells.ClearComments
    Sheets("Old Table").Cells.ClearComments
End Sub

Sub erste_spalte_bearb()
    Set ws1 = ThisWorkbook.Sheets("Old Table")
    Set ws2 = ThisWorkbook.Sheets("New Table")
    
    Letze_z_alt = ws1.Cells(Rows.Count, 1).End(xlUp).Row
    Letze_z_neu = ws2.Cells(Rows.Count, 1).End(xlUp).Row
'    letze_sp_alt = ws1.UsedRange.SpecialCells(xlCellTypeLastCell).Column
'    Letze_sp_neu = ws2.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
    ws1.Columns("A:A").Insert Shift:=xlToRight
    ws2.Columns("A:A").Insert Shift:=xlToRight
    anzeigen = "ja"
    
    For zz = 2 To Letze_z_alt
        Set Found = ws1.Columns(1).Find(ws1.Cells(zz, 2), LookIn:=xlValues, LookAt:=xlWhole)
        If Found Is Nothing Then
            ws1.Cells(zz, 1).Value = ws1.Cells(zz, 2).Value
        Else
            ws1.Cells(zz, 1).Value = ws1.Cells(zz, 2).Value & "_" & zz
            If anzeigen = "ja" Then
                MsgBox "Caution: there are double values in OLD Table, Column A. " & vbCrLf & vbCrLf _
                 & " The correct output for these values is not guarantied" & vbCrLf & vbCrLf _
                 & " Text of double values is marked RED"
                 anzeigen = "nein"
             End If
             ws1.Cells(zz, 2).Font.Color = vbRed
        End If
    Next zz
    
    anzeigen = "ja"
    For xx = 2 To Letze_z_neu
        Set Found = ws2.Columns(1).Find(ws2.Cells(xx, 2), LookIn:=xlValues, LookAt:=xlWhole)
        If Found Is Nothing Then
            ws2.Cells(xx, 1) = ws2.Cells(xx, 2).Value
        Else
            ws2.Cells(xx, 1).Value = ws2.Cells(xx, 2).Value & "_" & xx
            If anzeigen = "ja" Then
                MsgBox "Caution: there are double values in NEW Table, Column A. " & vbCrLf & vbCrLf _
                 & " The correct output for these values is not guarantied" & vbCrLf & vbCrLf & vbCrLf & vbCrLf _
                 & " Text of double values is marked RED"
                 anzeigen = "nein"
             End If
             ws2.Cells(xx, 2).Font.Color = vbRed
        End If

    Next xx
End Sub
