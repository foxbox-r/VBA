
Sub getGoTo()
  Dim 이름 As String
     이름 = InputBox("귀하의 이름을 입력하세요", "이름입력")
  If Len(이름) = 0 Then GoTo ET
    MsgBox 이름 & "님 안녕하세요?"
  Exit Sub
ET:
    MsgBox "아무 것도 입력하지 않으셨군요."
End Sub

Sub IfThen_1()
   If Range("J28") >= 80 Then
      MsgBox "합격!!!"
    Else
      MsgBox "불합격~!"
    End If
End Sub

Sub IfThen_2()
    If Range("J39") >= 95 Then
        MsgBox "우수!!!"
    ElseIf Range("J39") >= 80 Then
       MsgBox "합격!!!"
     
    Else
      MsgBox "불합격~!"
    End If
End Sub

Sub SelectCase1()
  Dim Msg As String
   Select Case Range("J39")
     Case Is < 80
       Msg = "불합격~!"
     Case Is < 95
       Msg = "합격!!!"
     Case Else
       Msg = "우수!!!"
    End Select
    
   MsgBox Msg
End Sub

Sub SelectCase2()
 Select Case Range("J77")
   Case Is > 90
     MsgBox "최고급"
   Case Is > 80
    MsgBox "고급"
   Case Is > 70
    MsgBox "중급"
   Case Is > 60
    MsgBox "일반"
   Case Else
    MsgBox "초급"
 End Select
End Sub



