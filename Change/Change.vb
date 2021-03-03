' ByVal : 값에 의한 전달
' ByRef : 참조에 의한 전달(디폴트)

Private Sub Worksheet_Change(ByVal Target As Range)
 With Target
  If .Column = 3 And .Row > 5 Then
   If .Count > 1 Then Exit Sub
     If IsNumeric(.Value) Then
        .Offset(0, 1) = .Offset(0, 1) + .Value
      End If
    End If
   End With
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim rng As Range
Dim i As Integer
Dim T As String
 If Target.Address = "$B$5" Then
        T = Target.Value
  Range("B8:F" & Rows.Count).ClearContents
   With Sheet3
   For i = 3 To .Cells(Rows.Count, 1).End(3).Row
     If .Cells(i, 1) = T Then
       Set rng = Cells(Rows.Count, 2).End(3).Offset(1)
       
        rng = .Cells(i, 2)
        rng.Offset(, 1) = .Cells(i, 3)
        rng.Offset(, 2) = .Cells(i, 4)
        rng.Offset(, 3) = .Cells(i, 5)
        rng.Offset(, 4) = .Cells(i, 6)
         
'        rng.Resize(1, 5) = .Cells(i, 2).Resize(1, 5).Value
      End If
    Next i
   End With
   Target.Select
  End If
End Sub