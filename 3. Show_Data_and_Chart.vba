' 2. Update Data 버튼 클릭 시,
'군집화된 데이터 화면에 출력하기

Sub Update_Data()

    Dim wkSheet As Worksheet
    Set wkSheet = ActiveWorkbook.Worksheets("Data") '차트실행 페이지

    Dim temp As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim line As Integer
    line = 2 '시작될 데이터 셀
    
    Dim cnt As Integer
    cnt = 0 '데이터 개수
    
    '-----데이터 개수 파악하는 단계
    temp = wkSheet.Cells(line, 1).Value '군집화 시작 후 배정받은 군집번호
    
    Do Until IsEmpty(wkSheet.Cells(line, 1).Value) '배열 크기 탐색 (데이터 개수 구하기)
        cnt = cnt + 1
        line = line + 1
    Loop
    

    '-----군집화 후 군집분류된 데이터 가져오고 출력하는 단계
    Dim num As Variant: ReDim num(cnt) '파악된 개수만큼 배열 크기 정하기
    
    line = 2 '시작될 데이터 셀
     
    For i = 0 To cnt Step 1
       num(i) = wkSheet.Cells(line, 1).Value '군집분류 데이터 가져오기
       wkSheet.Cells(line, 10).Value = num(i) '가져온 데이터(배열)을 J열에 출력하기
       line = line + 1
    Next i
    
    
    '-----차트에 기준이 될 가로축, 세로축 설정하는 단계
    line = 2 '시작될 데이터 셀
    
    Dim data1 As Variant: ReDim data1(cnt, 6) '최대 군집분류 개수: 5개
    
    '차트에 보여질 가로축과 새로축 데이터를 각각 K, L열에 출력
    For i = 0 To cnt Step 1
        For j = 0 To 1 Step 1
            data1(i, j) = wkSheet.Cells(i + 2, j + 3).Value
        Next j
    Next i
    
    '군집화된 분류에 따라 각각 M, N, O, P, Q 셀에 출력하기
    For i = 0 To cnt Step 1
        For j = 0 To 6 Step 1 '군집 분류 최대 개수: 5개
        
                If num(i) = "5" Then '군집 분류가 5일 때
                    data1(i, 2) = ""
                    data1(i, 3) = ""
                    data1(i, 4) = ""
                    data1(i, 5) = ""
                    data1(i, 6) = data1(i, 1)
                    wkSheet.Cells(i + 2, j + 11).Value = data1(i, j) 'Q셀에 출력
                
                ElseIf num(i) = "4" Then '군집 분류가 4일 때
                    data1(i, 2) = ""
                    data1(i, 3) = ""
                    data1(i, 4) = ""
                    data1(i, 5) = data1(i, 1)
                    data1(i, 6) = ""
                    wkSheet.Cells(i + 2, j + 11).Value = data1(i, j) 'P셀에 출력
                
                ElseIf num(i) = "3" Then '군집 분류가 3일 때
                    data1(i, 2) = ""
                    data1(i, 3) = ""
                    data1(i, 4) = data1(i, 1)
                    data1(i, 5) = ""
                    data1(i, 6) = ""
                    wkSheet.Cells(i + 2, j + 11).Value = data1(i, j) 'O셀에 출력
                    
                ElseIf num(i) = "2" Then '군집 분류가 2일 때
                    data1(i, 2) = ""
                    data1(i, 3) = data1(i, 1)
                    data1(i, 4) = ""
                    data1(i, 5) = ""
                    data1(i, 6) = ""
                    wkSheet.Cells(i + 2, j + 11).Value = data1(i, j) 'N셀에 출력
                
                 ElseIf num(i) = "1" Then '군집 분류가 1일 때
                    data1(i, 2) = data1(i, 1)
                    data1(i, 3) = ""
                    data1(i, 4) = ""
                    data1(i, 5) = ""
                    data1(i, 6) = ""
                    wkSheet.Cells(i + 2, j + 11).Value = data1(i, j) 'M셀에 출력
                
                End If
                
        Next j
    Next i
        
End Sub

' 3. Show Chart 버튼 클릭 시,
'군집화된 데이터 화면에 출력하기
Sub Show_Chart()

   '-----차트의 기본 설정 값 정하는 단계
    Columns("M:N").Select '차트의 가로축과 세로축 설정하기
    Range("K:K,M:Q").Select '표시할 데이터 범위 선택하기
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select '차트의 모양은 분산형으로 정하기
    ActiveChart.SetSourceData Source:=Range("Data!$K:$K,Data!$M:$Q")
    
    ActiveSheet.ChartObjects(1).Activate
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    
    ActiveChart.SetSourceData Source:=Range("Result!$B$9:$C$13") '군집화된 데이터들의 각 중심점 출력
    
    ActiveChart.Legend.Select
    Selection.Delete
    
    ActiveChart.Parent.Cut
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.Paste
   
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.SetSourceData Source:=Range("Result!$B$9:$C$13")
    ActiveChart.Parent.Cut
    
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.ClearToMatchStyle '눈금선 표시
    ActiveChart.ChartStyle = 242 '차트 스타일 선택
    
    
    '-----중심점 디자인 수정하는 단계 (가시적으로 확인하기 위해)
    ActiveChart.FullSeriesCollection(6).Select '최대 군집 분류 개수 + 중심점 개수 = Max 6
    Selection.MarkerSize = 13 '중심점 크기
    With Selection.Format.Fill '색상 및 디자인
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .BackColor.RGB = RGB(255, 0, 0)
    End With
    Selection.Format.line.Visible = msoFalse
    
    ActiveChart.FullSeriesCollection(4).Select
    ActiveChart.FullSeriesCollection(4).Points(1).Select
    Selection.Format.line.Visible = msoFalse
    
    
    '-----차트 가로축, 세로축 범위 수정 및 차트 디자인 수정 단계
    ActiveChart.Axes(xlValue).MinimumScale = 2
    ActiveChart.Axes(xlValue).MaximumScale = 4.5
    ActiveChart.Axes(xlCategory).MinimumScale = 4
    ActiveChart.Axes(xlCategory).MaximumScale = 8
    
    ActiveChart.ChartArea.Select
    
    With ActiveSheet.ChartObjects(1)
            .Chart.ChartType = xlXYScatter
            .Width = 700
            .Height = 500
            With .Chart.ChartArea.Font
                .Size = 15
                .Bold = True
            End With

            With ActiveSheet.ChartObjects(1).Chart
           .HasTitle = True
                  With .ChartTitle
                      .Text = "최종 군집분석(Clustering) 결과"
                      .Characters.Font.Size = 16
                      .Characters.Font.Bold = True
                  End With
            End With
    End With
   
End Sub
