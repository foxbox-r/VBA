'           Name	          Value	        Description	                        설명
' xlCellTypeAllFormatConditions	14	Cells of any format.	            조건부서식으로 지정되어 있는 셀
' xlCellTypeAllValidation	    15	Cells having validation criteria.	유효성 조건이 설정된 셀
' xlCellTypeBlanks	            4	Empty cells.	                    빈 문자열인 셀
' xlCellTypeComments	        1	Cells containing notes.	            메모가 포함된 셀
' xlCellTypeConstants	        2	Cells containing constants.	        상수가 입력되어 있는 셀
' xlCellTypeFormulas	        3	Cells containing formulas.	        수식이 들어있는 셀
' xlCellTypeLastCell	        11	The last cell in the used range.	사용된 범위 내의 마지막 셀
' xlCellTypeSameFormatConditions	Cells having the same format.	    같은 서식을 가진 셀
' xlCellTypeSameValidation		    Cells having the same validation criteria.	같은 유효성 조건을 가진 셀
' xlCellTypeVisible	            12	All visible cells.	                화면에 보이는 모든 셀

Sub SpecialCells_1()
  Columns("B").SpecialCells(xlCellTypeConstants).Select
End Sub

Sub SpecialCells_2()
  Columns("B").SpecialCells(xlCellTypeFormulas).Select
End Sub


Sub SpecialCells_3()
  Columns("B").SpecialCells(xlCellTypeBlanks).Select
End Sub

Sub End_Up()
 Range("E8").End(xlUp).Select
End Sub

Sub End_Down()
 Range("E8").End(xlDown).Select
End Sub

Sub End_ToLeft()
 Range("E8").End(xlToLeft).Select
End Sub

Sub End_ToRight()
 Range("E8").End(xlToRight).Select
End Sub


Sub EndTest()
 Range("B3").End(xlDown).Offset(1, 0) = "합계"
End Sub

