'Controls로 폼에 있는 모든 요소들을 관리할수있음

For i = 1 To 3
    Me.Controls("Text" & i).Value = i * 10 ' 이름이 Text1,Text2,Text3 이 텍스트박스 요소를 선택하는 코드
Next i