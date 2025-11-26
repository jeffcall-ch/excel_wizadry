Private Sub Worksheet_Change(ByVal Target As Range)

   'Do nothing if more than one cell is changed or content deleted
   If Target.Cells.Count > 1 Or IsEmpty(Target) Then Exit Sub

   ' Handle dropdown changes for F2, F3, F4
   If Target.Address = "$F$2" Then
       Call UpdateDNDropdown
       Exit Sub
   ElseIf Target.Address = "$F$3" Then
       Call UpdateResult
       Exit Sub
   ElseIf Target.Address = "$F$4" Then
       Call UpdateResult
       Exit Sub
   End If
      
   ' Handle original B2, B3, B4 changes
   If (Target.Address = "$B$2") Or (Target.Address = "$B$3") Or (Target.Address = "$B$4") Then
       'Ensure target is a number before multiplying by 2
       If IsNumeric(Target) Then
           'Stop any possible runtime errors and halting code
           On Error Resume Next
               'Turn off ALL events so the Target * 2 does not put the code into a loop.
               Application.EnableEvents = False
               Call szam
               'Turn events back on
               Application.EnableEvents = True
           'Allow run time errors again
           On Error GoTo 0
       End If
   End If
      
End Sub
