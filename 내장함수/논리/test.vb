' AND,OR 불리언(TRUE,FALSE) 반환
AND(TRUE,TRUE) 'TURE
AND(TRUE,FALSE) 'FALSE

OR(TRUE,FALSE) 'TRUE
OR(FALSE,FALSE) 'FALSE

'IF(조건식,참일때 값,거짓일때 값)
IF(1<2,"2가 큽니다","2가 작습니다.")

'IFERROR(수식,에러일때 값) 수식(첫번째인자)에 에러가 생기면 두번째인자 리턴
IFERROR("1A2"+1,"수식이 잘못되었습니다.")
