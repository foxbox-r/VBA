'============= module ====================
Sub dropdown_change()
Dim drop As DropDown
Dim selection As String

With Sheets("Sheet1")
    Set drop = .DropDowns("ListBox")
    selection = drop.List(drop.ListIndex) ' 드롭다운에 선택된 문자 가져오기
    .Range("g2") = WorksheetFunction.SumIf(Columns(1), selection, Columns(2))
End With

End Sub

'=============== Sheet1 ===================
Sub worksheet_change(ByVal Target As Range)
    Dim Col As New Collection
    Dim db, arr() As Variant
    Dim T As String
    
    If Target.Column <> 1 Then Exit Sub
    
    db = Range("a1").CurrentRegion
    
    On Error Resume Next
    For i = 1 To UBound(db)
        Col.Add db(i, 1), CStr(db(i, 1))
    Next i
    
    ReDim arr(1 To Col.Count)
    
    For i = 1 To Col.Count
        arr(i) = Col(i)
    Next i
    
    Me.DropDowns("ListBox").List = arr ' 드롭다운에 배열 넣기
    
    
End Sub