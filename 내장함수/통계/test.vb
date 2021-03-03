' COUNT,COUNTA,COUNTBLANK,COUNTIF,COUNTIFS
COUNT(범위) 범위내의 수치데이터(숫자,날짜,시간)의 개수를 반환 (빈칸제외)

COUNTA(범위) 범위내의 데이터의 개수를 반환 (빈칸제외)

COUNTBLANK(범위) 범위내의 빈칸의 개수를 반환

COUNTIF(범위,"조건식") 범위내에서 조건식에 맞는 셀의 개수를 반환 (조건식은 ""로 묶는다.)
EX) COUNTIF(범위,"수아") 값이 "수아"일때 | COUNTIF(범위,">=150") 값이 150이상일때

COUNTIFS(범위,"조건식",범위,"조건식".....) COUNTIF함수가 여러개일때
EX) COUNTIFS("A1:A5","대리","B1:B5",">=100") ' 직위가 대리이고 월급이 100 이상일때

' AVERAGE,AVERAGEIF,AVERAGEIFS
AVERAGE(계산범위,[계산범위]) 범위에 있는셀의 평균을 구한다. 

AVERAGEIF(조건범위,"조건식",[계산범위]) '처번째인자:조건식에 비교할 범위 | 두번째인자:조건식 | 세번째인자:계산할 범위(없으면 첫번째 인자를 계산)

AVERAGEIFS(계산범위,조건범위,"조건식",조건범위,"조건식"...) '첫번째인자로 계산할 범위를 정하고 뒤에 조건범위와 조건식을 넣는다

' SUMIF,SUMIFS
SUMIF(조건범위,"조건식",계산범위) '처번째인자:조건식에 비교할 범위 | 두번째인자:조건식 | 세번째인자:계산할 범위(없으면 첫번째 인자를 계산)

SUMIFS(계산범위,조건범위,"조건식",조건범위,"조건식"....) '첫번째인자로 계산할 범위를 정하고 뒤에 조건범위와 조건식을 넣는다

' MAX,MIN,

' RANK,MEDIAN
'RANK 셀한테 순위를 반환. 
RANK(비교셀,비교범위(절대참조),[0 or 1]) ' 0 or 생략 : 높은 숫자가 1위 | 1 : 낮은 숫자가 1위
EX) RANK("A1","$B$1:$B$10",1)

'MEDIAN 셀들의 중간값을 반환
MEDIAN(범위) ' 셀들의 개수가 홀수일때는 중간값 짝수일때는 중앙 셀 2개의 평균을 반환

