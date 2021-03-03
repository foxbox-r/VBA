
Sub OffsetTest_1()
 Range("B2").Offset(1, 0).Select
End Sub

Sub OffsetTest_2()
  Range("B2").Offset(0, 1).Select
End Sub

Sub OffsetTest_3()
  Range("B2").Offset(2, 5).Select
End Sub

Sub OffsetTest_4()
  Range("B2").Offset(0, -1).Select
End Sub

Sub OffsetTest_5()
  Range("B2").Offset(0, -2).Select
End Sub

Sub OffsetTest_6()
  Range("A1").Offset(-1, 0).Select
End Sub


Sub OffsetTest_7()
  Columns("B").Offset(0, 2).Select
End Sub

Sub OffsetTest_8()
 Rows(2).Offset(3).Select
End Sub

Sub OffsetTest_9()
 Range("B2:D7").Offset(7, 4).Select
End Sub

Sub OffsetTest_10()
 Range("B2:D7").Offset(0, 4).Resize(6, 1).Select
End Sub

Sub OffsetTest_11()
 Range("B2:D7").Offset(0, 4).Resize(3, 2).Select
End Sub

Sub ResizeTest()
 Range("G2").Resize(10, 5).Interior.Color = vbRed
End Sub



