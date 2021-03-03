' DAYS
' YEAR,MONTH,DAY,
' DATE
' TODAY,NOW
' DATEDIF
' EDATE,EOMONTH,WEEKDAY
' WORKDAY
' NETWORKDAY

DAYS(END_DAY,START_DAY) '완료일 - 시작일 반환
YEAR(TODAY()) '날짜의 년
MONTH(TODAY()) ' 날짜의 달
DAY(TODAY())  ' 날짜의 일

DATE(YEAR_NUM,MONTH_NUM,DAY_NUM) ' YEAR-MONTH-DAY 반환

TODYA() '오늘 날짜 반환
NOW() '오늘 날짜 시간

DATEDFI(START_DAY,END_DAY,OPTION)
' OPTION
' Y(년차이),M(달차이),D(날차이),YM(년차이무시,달차이),MD(달차이,날차이)

EDATE(START_DAY,MONTH_NUM) 'START_DAY에 MONTH_NUM날짜를 더한다
EOMONTH(START_DAY,MONTH_NUM) 'TART_DAY에 MONTH_NUM날짜를 더한후 그 달의 마지막 날을 리턴 (EO END-OF-MONTH)

WEEKDAY(날짜,2) '월1 ~ 일7 숫자 반환

' WORKDAY N일후의 일해야하는 날짜 반환
' NETWORKDAYS 두 날짜에서 일한 날짜의 수를 반환
WORKDAY(start_date, days, [holidays]) ' 계산을 START_DATE+1 + DAY를 하므로 -1을 빼줘야 DAYS이 지난다음의 날짜이 나온다.(토,일,공휴일제외)
EX) WORKDAY(TODAY(),10)-1

NETWORKDAYS(START_DATE,END_DATE,[holidays]) ' 시작일과 완료일사이에서 일하는 날의 수를 반환