'============== module ==============
' 범위.find(찾을 값) range 반환

Sub dropBox_change()
    Dim drop As DropDown
    Dim arr As Variant
    Dim numString As String
    Dim rng As Range
    
    Set drop = Sheets("Main").DropDowns("등록번호")
    numString = drop.List(drop.ListIndex)
    
    With Sheets("Data")
        Set rng = .Columns(6).Find(numString, , , 1)
        If Not rng Is Nothing Then
            arr = Array("d4", "f4", "h4", "d5")
        End If
    For i = 1 To UBound(arr)
        Range(arr(i)) = .Cells(rng.Row, 10 + i - 1)
    Next i
    MsgBox .Cells(rng.Row, 14)
    End With
    
    
End Sub

'============== Sheet(Data) ==============

Private Sub worksheet_change(ByVal Target As Range)
    Dim n As Long
    Dim rngList As Range
    Dim address As String
    
    With Target
        If .Row = 1 Then Exit Sub
        If .Column <> 6 Then Exit Sub
    End With
    
    Set rngList = Range("f2:f" & Range("f2").End(4).Row)
    address = "Data!" & rngList.address
    Sheets("Main").DropDowns("등록번호").ListFillRange = address
End Sub
