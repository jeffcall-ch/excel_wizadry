Sub shift(SArray(), CArray())

With Application
        .ScreenUpdating = False
    End With


j = 2 'A CArray 1.-eleme

For w = 1 To Count - 1
For i = 1 To Count - 1
  'Ha a beillesztendõ elem nagyobb, és mögötte nincs semmi
  If (CArray(j, 3) >= SArray(i, 3)) And (SArray(i + 1, 3) = 0) Then
     SArray(i + 1, 1) = CArray(j, 1)
     SArray(i + 1, 2) = CArray(j, 2)
     SArray(i + 1, 3) = CArray(j, 3)
     
     
     MsgBox ("i= " & i & "  SArray= " & SArray(i, 1) & "  CArray= " & CArray(i, 1))
     Exit For
    End If
  
  
  'Ha a beillesztendo elem kisebb az elsõ elemnél
  If (CArray(j, 3) < SArray(i, 3)) And i = 1 Then
       k = i
       d = Count
       Do While d > k
       SArray(d, 1) = SArray(d - 1, 1)
       SArray(d, 2) = SArray(d - 1, 2)
       SArray(d, 3) = SArray(d - 1, 3)
       d = d - 1
       Loop
   MsgBox ("i= " & i & "  beillesztendo kisebb az elsõ elemnél")
   Exit For
  End If
  
  
  'Ha a beillesztendo elem köztes, és nem az elsõ elemrõl van szó!
  If i > 1 Then
  If (CArray(j, 3) < SArray(i, 3)) And (CArray(j, 3) >= SArray(i - 1, 3)) Then
     k = i
     d = Count
       Do While d >= k
       SArray(d, 1) = SArray(d - 1, 1)
       SArray(d, 2) = SArray(d - 1, 2)
       SArray(d, 3) = SArray(d - 1, 3)
       d = d - 1
       Loop
        MsgBox ("i= " & i & "  beillesztendo köztes")
  End If
  Exit For
  End If
 Next
 j = j + 1
 Next
  
For i = 1 To Count
 MsgBox (SArray(i, 1))
Next
  
End Sub


