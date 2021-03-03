' ""	        length 가 0인 값 (메모리 개념에서는 6 byte 할당됨. 값 존재)
' Empty	        객체는 존재하되, 초기화 하지 않은 상태. 변수는 존재하지만 아무것도 대입하지 않음.
' Nothing	    객제 참조를 삭제
' Null	        알 수 없는 값. 아무것도 참조 하지 않는 값.
' vbNullChar	값이 0인 문자
' vbNullString	메모리가 할당되지 않은 값이 0인 문자열
' Missing	    누락

' ByVal : 값에 의한 전달
' ByRef : 참조에 의한 전달(디폴트)

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
 Dim rngSelect As Range
 
  Set rngSelect = Intersect(Target, Range("B5:G12"))
  If Not rngSelect Is Nothing Then Target.Next.Select
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  If Target.Row > 5 And Target.Column = 3 Then
    If Target.Count > 1 Then Exit Sub
     If Target.Offset(0, -1) = "" Then Exit Sub
      If Target.Value = "" Then
         Target.Value = "√"
         Target.Offset(0, 1) = Time
       Else
         Target.ClearContents
         Target.Offset(0, 1).ClearContents
       End If
   End If
 End Sub

 Private Sub Worksheet_SelectionChange(ByVal Target As Range)
With Target
 If Not Intersect(Target, Range("F6:H31")) Is Nothing Then
  If .Count > 1 Then Exit Sub
     Me.Cells(.Row, 4) = .Value
   End If
 End With
End Sub