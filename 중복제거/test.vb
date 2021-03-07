' 계산할범위.RemoveDuplicates Colums, Header 로 되어 있습니다.
' Header:=xlNo (열 머리글이 존재하지 않음)
' Header:=xlYes (열 머리글이 존재)

ActiveSheet.Range("A1:C100").RemoveDuplicates Array(1,2), xlYes

Range("A1:C" & Cells(Rows.Count, "A").End(3).Row).RemoveDuplicates Columns:=Array(2, 3), Header:=xlYes

'// A1:C의 데이타가 있는 마지막열 구간범위내에서 2열, 3열 기준으로 중복된 것을 제거하라. 헤더는 포함

