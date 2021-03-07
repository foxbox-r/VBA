'리스트박스에 리스트 넣는 방법
.RowSource = 범위.Address ' Range.Address (O) Variant.Address (X)
.List = 1차월배열 || 2차원배열
'리스트박스의 리스트는 0부터시작한다 선택을 안했을때는 -1 이다

Sub userform_initialize()
    Dim db As Range
    
    Set db = Range("c3").CurrentRegion
    With db
        Set db = .Offset(1).Resize(.Rows.Count - 1)
    End With
    
    With ListBox
        .ColumnHeads = True '지정한 범위 바로 위에 있는 데이터를 머릿글로 설정함
        .ColumnCount = 3
        .RowSource = db.Address
        .ColumnWidths = "60 pt;55 pt;80 pt" '각 리스트의 크기 정의 1 Column:60pt | 2 Column:55px | 3 Column:80pt
    End With
    
End Sub

with ListBox
    .MultiSelect = fmMultiSelectMulti ' 리스트박스를 여러개 선택하는 모드
    .Selected i ' 리스트박스의 리스트에서 선택되었는지 True/False 반환
    .List ' 리스트박스의 인덱스
    .ListIndex ' 선택한 인덱스번호 반환
    .RemoveItem i ' 리스트박스의 리스트에서 i번째 값을 삭제
    .AddItem 값 ' 리스트박스의 리스트를 뒤에 추가
End with