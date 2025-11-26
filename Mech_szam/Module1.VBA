Public Count As Integer


Sub szam()
With Application
        .ScreenUpdating = False
    End With


For i = 8 To 100
  Cells(i, "A") = ""
  Cells(i, "B") = ""
  Cells(i, "C") = ""
Next


Count = 0
For i = 1 To 50
  If (Cells(i + 7, "H") = "X") And (Cells(i + 7, "R") <> "") And (Cells(i + 7, "R") <= 5) Then
    Count = Count + 1
  End If
Next

Dim CArray()
ReDim CArray(1 To Count, 1 To 3)


j = 1
For i = 1 To 50

   If (Cells(i + 7, "H") = "X") And (Cells(i + 7, "R") <> "") And (Cells(i + 7, "R") <= 5) Then
      CArray(j, 1) = Cells(7 + i, "I")
      CArray(j, 2) = Cells(7 + i, "Q")
      CArray(j, 3) = Cells(7 + i, "R")
      j = j + 1
   End If
   
Next



Dim SArray()
ReDim SArray(1 To Count, 1 To 3)


'SArray feltöltése 0-val
For i = 1 To Count
   SArray(i, 1) = 0
   SArray(i, 2) = 0
   SArray(i, 3) = 0
Next


SArray(1, 1) = CArray(1, 1)
SArray(1, 2) = CArray(1, 2)
SArray(1, 3) = CArray(1, 3)


'For i = 1 To Count
'MsgBox (CArray(i, 1))
'Next

'MsgBox (Count)




Call Sort(SArray, CArray)


' Show status bar notification with stats
Application.StatusBar = "Calculation complete - " & Count & " items processed and sorted"

' Clear status bar after 3 seconds
Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"


End Sub

Sub ClearStatusBar()
    Application.StatusBar = False
End Sub


