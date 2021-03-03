' ■  정적 배열 변수의 선언형태 (크기를 변경할수없음)					
' 	   Dim Rev(10) As Variant 모든 데이터의 배열
' 	   Dim Rev(10) As Integer 숫자형 배열
' 	   Dim Rev(10) As String  문자형 배열

'               범위지정 가능
'      Dim Rev(2 to 10) As Variant 2 ~ 10 까지 총 9개
'      Dim Rev(10) As Variant 0 ~ 10 까지 총 11개	
'      Dim Rev(-1 to 10) As Variant -1 ~ 0 ~ 10 까지 총 12개	


' ■  동적 배열 변수의 선언형태 (크기를 변경할수있음)				
'     Dim Rev( ) As Variant			

'  배열의 크기 바꾸기 (0으로 초기화되어 바껴짐)
'  ReDim arr(n)

'  원래의 값을 유지한채로 배열의 크기 바꾸기 
'  ReDim Preserve arr(n)
'  ReDim Preserver arr(n to m)
'  ReDim Preserve arr(n,a to b) 2차원배열 바꾸기
'  ReDim Preserve arr(n to m,a to b) 2차원배열 바꾸기

Sub 동적배열변수사용_2()

  Dim i() As Integer

  Dim r As Integer

     ReDim i(3)
    
    For r = 1 To 3
        i(r) = r * 100
        Cells(r, 1) = i(r)
    Next r
    
     ReDim i(5)
    
    For r = 1 To 5
        Cells(r, 2) = i(r)
    Next r

End Sub

' =====결과값====
' 100   0
' 200   0
' 300   0
'       0
'       0

' 배열 함수
' Split(string,"구분자") 배열 반환
' LBound(배열) 배열의 시작 인덱스
' UBound(배열) 배열긔 끝 인덱스
' 위치.resize(row,column) = Application.Transpose(배열) 배열 삽입

str = "1-2-3-4-5-6-7-8-9" 
strArr = Split(str,"-")  ' ["1","2","3","4","5","6","7","8","9"]

UBound(strArr) ' 8  

ReDim Arr1(10) as Variant
ReDim Arr2(1 to 20) as Variant
ReDim Arr3(-10 to 30) as Variant

UBound(Arr1) '10 범위 : 0 ~ 10 (11개) 
UBound(Arr2) '20 범위 : 1 ~ 20 (20개)
UBound(Arr3) '30 범위 : -10 ~ 0 ~ 30 (41개)

' 2차원 배열사용하기
' 범위를 사용할땐 dataRange(행,열) ex) range("a1:c3")(2,2)
' 2차원배열을 사용할땐 arr2(열,행) 

Sub count()
    Dim copyRange As Range
    Set copyRange = Sheets("Sheet1").Range("a6")
    Dim dataRange, arr() As Variant
    dataRange = Sheets("Sheet1").Range("a1").CurrentRegion
    Dim i, n As Integer
    For i = 1 To UBound(dataRange)
        ReDim Preserve arr(0 To 2, n)
        arr(0, n) = dataRange(i, 1)
        arr(1, n) = dataRange(i, 2)
        arr(2, n) = dataRange(i, 3)
        n = n + 1
    Next i
    
    copyRange.Resize(n, 3) = Application.Transpose(arr)
    
End Sub

' 배열 넣기 가로/세로
Dim arr(1 To 30) As Variant
    Dim rng As Range
    For i = 1 To 30
        arr(i) = i
    Next i
      
    Set rng = Sheets("Sheet1").Range("a1")
    
    rng.Resize(1, 30) = arr ' 가로
    rng.Resize(30) = Application.Transpose(arr) ' 세로
