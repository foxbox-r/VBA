Range("A7").Value = "안녕하세요?"
Worksheets("Sheet1").Range("A1").Value = "안녕하세요?"
Worksheets("Sheet1").Activate
Range("A1:A10") = "안녕하세요?"

Worksheets("Sheet1").Range("A1:A10").ClearContents
' clear 내용,서식 모두
' clearContents 내용만 
' clearFormats 서식만
' delete shift:=xlup '전체셀을 위로 끌어 올리며 행 삭제

Range("B49").CurrentRegion.Select
Range("C5").CurrentRegion.ClearContents
Range("C20").CurrentRegion.ClearContents

Cells(5, 3) = "안녕하세요?"
Range(Cells(3, 2), Cells(9, 5)).Select
Range("B3:E9").Cells(3, 2) = "안녕!"
Worksheets("Sheet3").Cells(9, 7) = "안녕하세요?"
Cells(Rows.Count, 2).End(3)(2).Resize(1, 3) = Range("H4:J4").Value
' .End(1) (= xlToLeft)은 좌로 이동하라는 명령이고
' .End(2) (= xlToRight)은 우로 이동하라는 명령이고
' .End(3) (= xlUp)은 위로 이동하라는 명령이고
' .End(4) (= xlDown)은 아래로 이동하라는 명령이다
' = Cells(Rows.Count,1).End(3).Cells(2,1)
' = Cells(Rows.Count, 1).End(3)(2,1)
' = Cells(Rows.Count,1).End(3)(2)



