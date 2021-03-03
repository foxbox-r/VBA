Function ConcatText(범위 As Range)
 Dim Rev() As Variant
 Dim rng As Range
 Dim n As Integer
 
  For Each rng In 범위
   If Len(rng) Then
     ReDim Preserve Rev(n)
       Rev(n) = rng.Value
       n = n + 1
     End If
    Next rng

 ConcatText = Join(Rev, ",")
End Function
Function Add(ByVal a As Integer, ByVal b As Integer)
    Add = a + b
End Function

' 값	함수 범주			
' 0	    모두			
' 1	    재무			
' 2	    날짜/시간			
' 3	    수학/삼각			
' 4	    통계			
' 5	    찾기/참조 영역			
' 6	    데이터베이스			
' 7	    텍스트			
' 8	    논리			
' 9	    정보			

Sub ChangeCategory()
  Application.MacroOptions macro:="ConcatText", Category:=7
  Application.MacroOptions macro:="Add", Category:=3
End Sub

' 함수          설명
' IsNumeric()   데이터 형이 숫자인지 확인
' IsDate()      데이터 형이 날짜인지 확인
' IsObject()    데이터 형이 개체인지 확인
' IsNull()      Null 값인지 확인
' IsError()     데이터 형이 Error 인지 확인
' IsEmpty()     변수가 초기화된 상태인지 확인
' IsMissing()   Optional 로 정의된 인수가 전달되지 않았는지 확인
' IsArray()     데이터 형이 배열인지 확인

' Val : 문자 형태의 숫자를 숫자로 변경하는 VBA 내장 함수

Function 숫자만(문자열)
    Dim strT As String
    Dim i As Integer
    For i = 1 To Len(문자열)
      If IsNumeric(Mid(문자열, i, 1)) Then
          strT = strT & Mid(문자열, i, 1)
      End If
    Next i
    숫자만 = Val(strT)
End Function

' hasArray : 수식이 배열인지 리턴
' hasFormula : 수식인지 리턴
' Formula : 수식리턴

Function 수식보기(대상셀)
  Dim strT As String
   If 대상셀.HasArray Then
       strT = "{" & 대상셀.Formula & "}"
    ElseIf 대상셀.HasFormula Then
       strT = 대상셀.Formula
    Else
       strT = "수식아님!"
    End If

 수식보기 = strT
End Function