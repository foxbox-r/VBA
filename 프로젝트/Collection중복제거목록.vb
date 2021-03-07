'통합문서 5회차
Dim db, arr() As Variant
Dim Col As New Collection
Dim n, i As Integer
Dim T As String

Sub clear()
    n = 0
    T = ""
    Erase arr
    Set Col = Nothing
End Sub

Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    If Target.Column <> 5 Then Exit Sub
    If Target.Row = 1 Then Exit Sub
    If Target.Count > 1 Then Exit Sub
        db = Range("a1").CurrentRegion
        
        On Error Resume Next
        For i = 2 To UBound(db)
            Col.Add db(i, 1), CStr(db(i, 1))
        Next i
        
        Err.clear
            
    ReDim arr(1 To Col.Count)
            
    For i = 1 To Col.Count
        arr(i) = Col(i)
    Next i
    
    T = Join(arr, ",")
    'MsgBox T
    
    With Target.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, T
    End With
    clear
End Sub

Sub Worksheet_change(ByVal Target As Range)
    If Target.Column <> 5 Then Exit Sub
    If Target.Row = 1 Then Exit Sub
    If Target.Count <> 1 Then Exit Sub
    
    If IsEmpty(Target) Then '이름 지울때
        With Target.Offset(, 1)
            .Validation.Delete
            .clear
        End With
        Target.Offset(, 2).clear
    Else
        db = Range("a1").CurrentRegion
        For i = 2 To UBound(db)
            On Error Resume Next
            If db(i, 1) = Target Then
                Col.Add db(i, 2), CStr(db(i, 2))
            End If
        Next i
        ReDim arr(1 To Col.Count)
        For i = 1 To Col.Count
            arr(i) = Col(i)
        Next i
        T = Join(arr, ",")
        'MsgBox T
        
        With Target.Offset(, 1).Validation
            .Delete
            .Add xlValidateList, xlValidAlertStop, xlBetween, T
        End With
        
        With Target.Offset(, 1)
            .Value = arr(1)
            .Select
         SendKeys "%{down}"
        End With
        
        Target.Offset(, 2).Value = Now()
        
    End If
    
    clear
    
End Sub
