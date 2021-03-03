
Sub Macro1()
 Dim lngSum As Long
 Dim i As Integer, r As Integer
    r = Cells(3, 2).End(4).Row
  For i = 3 To r
    lngSum = lngSum + Cells(i, 2)
   Next i
  MsgBox "합계는 " & lngSum & "입니다."
End Sub


Sub Macro2()
 Dim lngSum As Long
 Dim i As Integer, r As Integer
    r = Cells(3, 2).End(4).Row
  For i = 3 To r Step 2 '1칸씩 건너 뜀
    lngSum = lngSum + Cells(i, 2)
   Next i
  MsgBox "합계는 " & lngSum & "입니다."
End Sub


Sub Macro3()
 Dim lngSum As Long
 Dim n As Integer
     n = 3
  Do While Not IsEmpty(Cells(n, 2)) 
   lngSum = lngSum + Cells(n, 2)
     n = n + 1
  Loop
  MsgBox "합계는 " & lngSum & "입니다."
End Sub


Sub Macro4()
 Dim lngSum As Long
 Dim n As Integer
     n = 3
  Do
   lngSum = lngSum + Cells(n, 2)
     n = n + 1
  Loop While Not IsEmpty(Cells(n, 2))
   MsgBox "합계는 " & lngSum & "입니다."
End Sub

'while : 값이 False가 될때까지
'until : 값이 True가 될때까지

Sub Macro5()
 Dim lngSum As Long
 Dim n As Integer
     n = 3
  Do Until IsEmpty(Cells(n, 2)) 
    lngSum = lngSum + Cells(n, 2)
     n = n + 1
  Loop
  MsgBox "합계는 " & lngSum & "입니다."
End Sub


