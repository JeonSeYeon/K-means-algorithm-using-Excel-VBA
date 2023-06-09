' 1. Avg_Centroids 버튼 클릭 시,
'군집화된 데이터의 속성과 중심점 확인하기

Sub Avg_Centroids()

    Dim wkSheet As Worksheet
    Set wkSheet = ActiveWorkbook.Worksheets("Data") '실행 페이지

    Dim temp As Double
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    '-----군집화된 데이터의 중심점 값 출력하는 단계
    Dim start_col As Integer
    start_col = 9  '군집1의 속성값 시작 셀
    
    Dim result As Integer
    result = 0
    
    Do Until IsEmpty(ActiveWorkbook.Worksheets("Result").Cells(start_col, 2))
        start_col = start_col + 1 '군집 분류 개수 파악
    Loop
    
    
    '-----데이터 개수 파악하는 단계
    Dim col As Integer
    col = 2

    Dim cnt As Integer
    cnt = 0
    
    Do Until IsEmpty(wkSheet.Cells(col, 3).Value) '배열 크기 탐색 (데이터 개수 구하기)
        cnt = cnt + 1
        col = col + 1
    Loop

    
    '-----군집화 후 각 중심점거리의 평균을 출력하는 단계
    Dim num As Variant: ReDim num(cnt) '파악된 개수만큼 배열 크기 정하기
    
    'K1군집 중심점까지의 거리 차이의 평균
    temp = 0
    j = 0
    
        For i = 0 To cnt Step 1
            If IsEmpty(ActiveWorkbook.Worksheets("Result").Cells(9, 2 + j).Value) = False Then
                num(i) = (wkSheet.Cells(i + 2, 3 + j).Value - ActiveWorkbook.Worksheets("Result").Cells(9, 2 + j)) ^ 2
                temp = temp + Sqr(num(i))
            End If
        Next i
    
        '거리의 차이의 평균을 소수점 3자리까지 출력
        ActiveWorkbook.Worksheets("Result").Cells(20, 5 + j) = Round(temp / cnt, 3)
                
        '각 속성별 평균을 위해, 속성 초기화 작업
        temp = 0
        For k = 0 To cnt Step 1
            num(k) = 0
        Next k


    
    'K2군집 중심점까지의 거리 차이의 평균
    temp = 0
    j = 0
    
        For i = 0 To cnt Step 1
            If IsEmpty(ActiveWorkbook.Worksheets("Result").Cells(10, 2 + j).Value) = False Then
                num(i) = (wkSheet.Cells(i + 2, 3 + j).Value - ActiveWorkbook.Worksheets("Result").Cells(10, 2 + j)) ^ 2
                temp = temp + Sqr(num(i))
            End If
        Next i
        
        '거리의 차이의 평균을 소수점 3자리까지 출력
        ActiveWorkbook.Worksheets("Result").Cells(21, 5 + j) = Round(temp / cnt, 3)

        '각 속성별 평균을 위해, 속성 초기화 작업
        temp = 0
        For k = 0 To cnt Step 1
                num(k) = 0
        Next k

        
        
    'K3군집 중심점까지의 거리 차이의 평균
    temp = 0
    j = 0
    
        For i = 0 To cnt Step 1
            If IsEmpty(ActiveWorkbook.Worksheets("Result").Cells(11, 2 + j).Value) = False Then
                num(i) = (wkSheet.Cells(i + 2, 3 + j).Value - ActiveWorkbook.Worksheets("Result").Cells(11, 2 + j)) ^ 2
                temp = temp + Sqr(num(i))
            End If
        Next i
        
         '거리의 차이의 평균을 소수점 3자리까지 출력
        ActiveWorkbook.Worksheets("Result").Cells(22, 5 + j) = Round(temp / cnt, 3)

        '각 속성별 평균을 위해, 속성 초기화 작업
        temp = 0
        For k = 0 To cnt Step 1
            num(k) = 0
        Next k
    
       
    'K4군집 중심점까지의 거리 차이의 평균
    temp = 0
    j = 0
    
        For i = 0 To cnt Step 1
            If IsEmpty(ActiveWorkbook.Worksheets("Result").Cells(12, 2 + j).Value) = False Then
                num(i) = (wkSheet.Cells(i + 2, 3 + j).Value - ActiveWorkbook.Worksheets("Result").Cells(12, 2 + j)) ^ 2
                temp = temp + Sqr(num(i))
            End If
        Next i
        
         '거리의 차이의 평균을 소수점 3자리까지 출력
        ActiveWorkbook.Worksheets("Result").Cells(23, 5 + j) = Round(temp / cnt, 3)

        '각 속성별 평균을 위해, 속성 초기화 작업
        temp = 0
        For k = 0 To cnt Step 1
            num(k) = 0
        Next k


    'K5군집 중심점까지의 거리 차이의 평균
    temp = 0
    j = 0
    
        For i = 0 To cnt Step 1
            If IsEmpty(ActiveWorkbook.Worksheets("Result").Cells(13, 2 + j).Value) = False Then
                num(i) = (wkSheet.Cells(i + 2, 3 + j).Value - ActiveWorkbook.Worksheets("Result").Cells(13, 2 + j)) ^ 2
                temp = temp + Sqr(num(i))
            End If
        Next i
        
         '거리의 차이의 평균을 소수점 3자리까지 출력
        ActiveWorkbook.Worksheets("Result").Cells(24, 5 + j) = Round(temp / cnt, 3)

        '각 속성별 평균을 위해, 속성 초기화 작업
        temp = 0
        For k = 0 To cnt Step 1
            num(k) = 0
        Next k
    
    
End Sub
