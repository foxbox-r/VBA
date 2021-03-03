
' let : 일반변수(기본값이여서 빼도됨)
' set : 개체변수

MsgBox Rows.Count

Sub UnionMethod_1()
 Dim rngUnion As Range
  Set rngUnion = Union(Range("B7"), Range("C8"), Range("D9"), Range("E10"))
  rngUnion.Interior.Color = vbYellow
End Sub


Sub UnionMethod_2()
  Dim i As Integer, r As Integer
  Dim rngRow As Range
    
   r = Cells(Rows.Count, 2).End(xlUp).Row
    Set rngRow = Rows(40)
   For i = 40 To r Step 4
      Set rngRow = Union(rngRow, Rows(i))
     Next i
     rngRow.Select
End Sub

Sub UnionMethod_3()
  Dim i As Integer, r As Integer
  Dim rngRow As Range
    
    r = Cells(Rows.Count, 2).End(xlUp).Row
    Set rngRow = Rows(40)
   For i = 40 To r Step 4
       Set rngRow = Union(rngRow, Rows(i))
     Next i
     rngRow.Copy Sheets(2).Rows(1)
End Sub
