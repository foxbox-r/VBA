' LEFT,RIGHT,MID,CONCATENATE
' REPLACE,SUBSTITUE
' TEXT
' EXACT,FIND,SEARCH

LEFT(string,number) '문자열의 왼쪽에서 숫자의 크기대로 잘라 리턴
RIGHT(string,number)'문자열의 왼쪽에서 숫자의 크기대로 잘라 리턴
MID(string,start,number) ' 문자열을 start숫자에서 number의 크기대로 잘라 리턴
EX) LEFT("NICE",2) '"NI"
    RIGHT("NICE",2) '"CE"
    MID("ABCDE",2,2) '"BC"

' REPLACE,SUBSTITUE
REPLACE(OLD_STRING,START,CHAR_NUM,NEW_STRING)
EX) REPLACE("010-1234-5678",5,4,"****") '"010-****-5678"

SUBSTITUTE( 문자열, 찾을문자, 새로운문자, [바꿀지점] )
EX) SUBSTITUTE("사과나무","사과","포도") '"포도나무"

' TEXT 문자서식을 바꾼다
EX)
' 1000 => TEXT(CELL,"#,##0") => 1,000
' 1 => TEXT(CELL,"000") => 001
' 06월 12일 => TEXT(CELL,"YYYY-MM-DD") => 2020-06-12
' 오전 9:05:07 => TEXT(CELL,"HH:MM AM/PM") => 09:05 AM
' 0.5 => TEXT(CELL,"# ?/?") => 1/2

' EXACT,FIND,SEARCH
' EXACT(text1, text2) 두 텍스트가 같은지 비교한다(대소문자 구분함)
EX) EXACT("Smith","Smith") => TRUE
'FIND
FIND("사과", "사과나무 사과열렸네", 4) 6 '4번째 문자부터 검색을 시작하여 두번째 '사과'가 위치한 6을 반환합니다.
'SEARCH
=SEARCH("특별시", "서울특별시") '// =3 을 반환합니다.