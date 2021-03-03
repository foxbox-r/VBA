' Byte    1byte 0~255
' Boolean 2Byte true/false
' Integer 2Byte +2^15-1 ~ 0 ~ -2^15
' Long    4Byte +2^31-1 ~ 0 ~ -2^31
' Double  8Byte (소수)
' Date    14Byte (날짜)
' String(가변) 10+길이
' Variant or (데이터형 생략) : 모든데이터형
	
Sub Variable_Test()
 Dim c As Integer
 Dim r As Integer
  With Range("B3").CurrentRegion
   c = .Columns.Count
   r = .Rows.Count
  End With
  MsgBox c
  MsgBox r
End Sub

Sub 곱하기()
 Dim i As Integer
 Dim j As Integer
 Dim k As Integer
       
    i = Range("B21").Value
    j = Range("B22").Value
       k = i * j
    Range("B23") = k
End Sub
