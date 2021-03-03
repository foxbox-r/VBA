' range(머릿글,값).AdvancedFilter xlFilterCopy,비교range(머릿글,값),결과출력위치range


Sub worksheet_change(ByVal Target As Range)
    Dim dataRange As Range, compRange As Range
    Set dataRange = Sheets("Data").Range("b5").CurrentRegion
    Set compRange = Me.Range("D3:k4")
    
    If Not Intersect(Target, Range("d4:k4")) Is Nothing Then
        If IsEmpty(Target) Then Exit Sub
        
        Me.Range(Cells(8, 4), Cells(Rows.Count, 11)).ClearContents
        dataRange.AdvancedFilter xlFilterCopy, compRange, Me.Range("d7")
        Target.Select
    End If
End Sub