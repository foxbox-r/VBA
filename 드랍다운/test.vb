Sheets("Main").DropDowns("name1").ListFillRange = "a1:a5" ' 범위로 드롭다운 리스트 만들기
Sheets("Main")..DropDowns("name2").List = Array(1,2,3) ' 배열로 드롭다운 리스트 만들기

Sheets("Main").DropDowns("name2").ListIndex ' 현재 리스트 인덱스 위치 반환
Sheets("Main").DropDowns("name2").ListIndex n ' 현재 리스트 인덱스 위치 변환

with Sheets("Main").DropDowns
    msgbox .List(.ListIndex) '현재 리스트에서 선택된 데이터 추출
end with