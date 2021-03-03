' 전역변수 : 어디에서나 호출이 가능
' 지역변수 : 함수가 끝나면 사라짐

Public strText As String '전역변수

Sub ChangeT()
 Dim 문자 As String '지역변수
   문자 = Range("B7")
  Range("B7") = Range("C7")
  Range("C7") = 문자
End Sub

Sub My_String()
  MsgBox strText
End Sub

Sub PublicTest()
 Range("H45") = strText
End Sub

