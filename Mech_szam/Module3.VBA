Sub Sort(SArray(), CArray())

With Application
        .ScreenUpdating = False
    End With



For nr = 2 To Count

   SC = 1 'SArray - Step counter
  
  'Ha a beillesztenõ elem kisebb az SArray elsõ eleménél. Akkor minden elem egyel jobbra.
  'És az elsõ helyre beilleszt!
  If CArray(nr, 3) < SArray(SC, 3) Then
     k = Count
     Do While k > 1
        SArray(k, 1) = SArray(k - 1, 1)
        SArray(k, 2) = SArray(k - 1, 2)
        SArray(k, 3) = SArray(k - 1, 3)
        k = k - 1
     Loop
        
     SArray(1, 1) = CArray(nr, 1)
     SArray(1, 2) = CArray(nr, 2)
     SArray(1, 3) = CArray(nr, 3)
     
     'MsgBox ("nr= " & nr & " SC= " & SC & " " & SArray(1, 1))
     
     
     
     SC = Count + 1 'hogy ne induljon el a While, és ugorjunk a köv nr-re!
  End If
  

'Beillesztendõ nagyobb, és nincs utána elem
  Do While (SC < Count)

   If (CArray(nr, 3) >= SArray(SC, 3)) And (SArray(SC + 1, 3) = 0) Then
      SArray(SC + 1, 1) = CArray(nr, 1)
      SArray(SC + 1, 2) = CArray(nr, 2)
      SArray(SC + 1, 3) = CArray(nr, 3)
      SC = Count + 1 'hogy kilépjen a While-ból
   End If
   
   'Beillesztendõ köztes
   
   If SC < Count Then
   If (CArray(nr, 3) >= SArray(SC, 3)) And (CArray(nr, 3) < SArray(SC + 1, 3)) Then
      
        k = Count
     Do While k > (SC + 1)
        SArray(k, 1) = SArray(k - 1, 1)
        SArray(k, 2) = SArray(k - 1, 2)
        SArray(k, 3) = SArray(k - 1, 3)
        k = k - 1
     Loop
        
     SArray(SC + 1, 1) = CArray(nr, 1)
     SArray(SC + 1, 2) = CArray(nr, 2)
     SArray(SC + 1, 3) = CArray(nr, 3)
     SC = Count + 1 'hogy kilépjen a While-ból
   End If
   End If
   
   SC = SC + 1
  Loop
   
   
Next 'nr

'For i = 1 To Count
' MsgBox (SArray(i, 1))
' Next
 

'For i = 1 To Count

' Cells(i + 7, "A") = SArray(i, 1)
' Cells(i + 7, "B") = SArray(i, 3)
' Cells(i + 7, "C") = SArray(i, 2)
 
'Next

i = 8
k = Count
Do While k >= 1
  Cells(i, "A") = SArray(k, 1)
 Cells(i, "B") = SArray(k, 3)
 Cells(i, "C") = SArray(k, 2)
 k = k - 1
 i = i + 1
Loop


End Sub
