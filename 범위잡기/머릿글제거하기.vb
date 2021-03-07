    dim rngArea as Range
    
    With Range("c4").CurrentRegion
        Set rngArea = .Offset(1).Resize(.Rows.Count - 1)
    End With

    Set rngArea = Range(Cells(4, 2), Cells(Rows.Count, 2).End(3))