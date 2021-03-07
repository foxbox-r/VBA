' VLOOKUP
' INDEX,MATCH
' CHOOSE

VLOOKUP(선택셀,테이블,열 번호,FALSE)

' INDEX,MATCH

'INDEX 2차워범위에서 Y,X로 값을 리턴
INDEX(2차원테이블,Y,x)

'MATCH 범위에서 찾는값의 위치 반환
MATCH(찾는값,비교범위,0)

EX) INDEX(크로스테이블,MATCH(찾는값1),MATCH(찾는값2))

'CHOOSE
CHOOSE(2,"월","화","수","목")
EX) SUM(CHOOSE(2,A3,A4,A5))