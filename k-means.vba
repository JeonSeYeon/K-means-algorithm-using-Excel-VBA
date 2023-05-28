Option Base 1

Option Explicit


'* Original Code
'** https://github.com/gpolic/kmeans-excel

'*** Modify Code
'**** https://github.com/JeonSeYeon/K-means-algorithm-using-Excel-VBA


' Start Clustering 버튼 클릭 시,
' K-means 알고리즘 실시


Public Sub kmeans()
    Dim wkSheet As Worksheet
    Set wkSheet = ActiveWorkbook.Worksheets("Start") '실행 페이지

    Dim MaxIt As Integer: MaxIt = wkSheet.Range("MaxIt").Value  '반복횟수 입력칸
    Dim DataSht As String: DataSht = wkSheet.Range("InputSheet").Value   '데이터 입력칸(시트이름입력)
    Dim DataRange As String: DataRange = wkSheet.Range("InputRange").Value '실행할 데이터 범위 선택
    Dim DataRecords As Variant: DataRecords = Worksheets(DataSht).Range(DataRange)  '결과저장 시
    Dim ClusterIndexes As Variant
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)   '저장된 데이터를 배열의 크기로 저장(1차원배열)
    Dim NUMCLUSTERS As Integer: NUMCLUSTERS = wkSheet.Range("Clusters").Value '입력된 클러스터 숫자
    Dim Centroids As Variant
    Dim counter As Integer
    Dim ClustersUpdated As Integer: ClustersUpdated = 1
   
    
    Application.StatusBar = "   [ 초기화 중..   ]"  '초기화 중 상태바에 표현

    Dim StartTime, SecondsElapsed As Double
    StartTime = Timer '시작 초
    
    Dim InitialCentroidsCalc As Variant
    InitialCentroidsCalc = ComputeInitialCentroidsCalc(DataRecords, NUMCLUSTERS) ' [k-means 함수] 호출
    
    Static minDistSquared As Variant
    
    Application.StatusBar = "   [ 시작 중..     ]"

    counter = FindClosestCentroid(DataRecords, InitialCentroidsCalc, ClusterIndexes) '[중심점 할당 함수] 호출
    counter = 1
    
    While counter <= MaxIt And ClustersUpdated > 0   '지정한 최대반복횟수만큼 정규화
        Application.StatusBar = "   [ 반복 실행 중 .. : " + CStr(counter) + "     ]" '반복횟수 상태바에 표시
        Centroids = ComputeCentroids(DataRecords, ClusterIndexes, NUMCLUSTERS) '[가까운 중심점 찾기 함수] 호출
        ClustersUpdated = FindClosestCentroid(DataRecords, Centroids, ClusterIndexes)
        counter = counter + 1
    Wend
    
    '결과 출력
    Dim ClusterOutputSht As String: ClusterOutputSht = wkSheet.Range("OutputSheet").Value
    Dim ClusterOutputRange As String: ClusterOutputRange = wkSheet.Range("OutputRange").Value
    Worksheets(ClusterOutputSht).Range(ClusterOutputRange).Resize(NUMRECORDS, 1).Value = WorksheetFunction.Transpose(ClusterIndexes)
    
    '[Result 시트에서 결과 출력 함수] 호출
    Call ShowResult(DataRecords, ClusterIndexes, Centroids, NUMCLUSTERS)
    
    '군집화 완료 시간 출력
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox "[실행완료!] 실행경과시간:   " & SecondsElapsed & " seconds", vbInformation

End Sub

'[k-means 함수]

Function ComputeInitialCentroidsCalc(ByRef DataRecords As Variant, NUMCLUSTERS As Integer) As Variant

    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)
    Dim NUMCOLUMNS As Integer: NUMCOLUMNS = UBound(DataRecords, 2)
    Dim Taken() As Variant: ReDim Taken(NUMRECORDS) '중심점 표시 기록
    Dim InitialCentroidsCalc As Variant: ReDim InitialCentroidsCalc(NUMCLUSTERS, NUMCOLUMNS) As Variant
    Dim counter As Integer
    Dim CentroidsFound As Integer
    Dim dist As Double
    Dim preventLoop As Boolean: preventLoop = True
    Dim minDistSquared As Variant: ReDim minDistSquared(NUMRECORDS)
    Dim FirstCentroid As Variant: ReDim FirstCentroid(NUMCOLUMNS)
    Dim FirstCentroidIndex As Integer
    
    FirstCentroidIndex = Int(Rnd * NUMRECORDS) + 1 '초기 중심점 랜덤으로 설정
    
    For counter = 1 To NUMCOLUMNS
        FirstCentroid(counter) = DataRecords(FirstCentroidIndex, counter)
        InitialCentroidsCalc(1, counter) = FirstCentroid(counter)
    Next counter
    
    Taken(FirstCentroidIndex) = 1
    CentroidsFound = 1
    
    For counter = 1 To NUMRECORDS
        If Not counter = FirstCentroidIndex Then
            dist = EuclideanDistance(FirstCentroid, Application.Index(DataRecords, counter, 0), NUMCOLUMNS) '[유클리디안 거리 계산 함수] 호출
            minDistSquared(counter) = dist * dist
        End If
    Next counter


    'main 실행
    Do While CentroidsFound < NUMCLUSTERS And preventLoop = True
    
        Dim distSqSum As Variant: distSqSum = 0
    
        For counter = 1 To NUMRECORDS
            If Not Taken(counter) = 1 Then
            distSqSum = distSqSum + minDistSquared(counter)
            End If
        Next counter
    
        Dim r As Variant
        r = Rnd * distSqSum
    
        Dim nextpoint As Integer
        nextpoint = -1
        
        Dim sum As Variant
        
        For counter = 1 To NUMRECORDS
            If Not Taken(counter) = 1 Then
                sum = sum + minDistSquared(counter)
                If sum > r Then
                    nextpoint = counter
                    Exit For
                End If
            End If
        Next counter
        
        '새로운 중심점을 찾지 못했을 때, 마지막 값으로 대체
        If nextpoint = -1 Then
            For counter = NUMRECORDS To 1
                If Not Taken(counter) = 1 Then
                    nextpoint = counter
                End If
            Next counter
        End If
        
        '새로운 중심점을 찾았을 때
        If nextpoint >= 0 Then
            CentroidsFound = CentroidsFound + 1
            Taken(nextpoint) = 1
            
            For counter = 1 To NUMCOLUMNS
                InitialCentroidsCalc(CentroidsFound, counter) = DataRecords(nextpoint, counter)
            Next counter
                
            If CentroidsFound < NUMCLUSTERS Then
                For counter = 1 To NUMRECORDS
                
                    If Not Taken(counter) = 1 Then
                    
                        Dim dist2 As Variant
                        dist2 = EuclideanDistance(Application.Index(InitialCentroidsCalc, CentroidsFound, 0), Application.Index(DataRecords, counter, 0), NUMCOLUMNS)
                                                  
                        Dim d2 As Variant
                        d2 = dist2 * dist2
                        
                        If d2 < minDistSquared(counter) Then
                            minDistSquared(counter) = d2
                        End If
                        
                    End If
                Next counter
            End If
        Else
            preventLoop = False
        End If
    Loop
    
    ComputeInitialCentroidsCalc = InitialCentroidsCalc

End Function
    
'[유클리디안 거리 계산 함수]

Public Function EuclideanDistance(X As Variant, Y As Variant, NumberOfObservations As Integer) As Double
    Dim counter As Integer
    Dim RunningSumSqr As Double: RunningSumSqr = 0
    
    For counter = 1 To NumberOfObservations
        RunningSumSqr = RunningSumSqr + ((X(counter) - Y(counter)) ^ 2)
    Next counter
    
    EuclideanDistance = Sqr(RunningSumSqr)
    
End Function

'[중심점 할당 함수]

Public Function FindClosestCentroid(ByRef DataRecords As Variant, ByRef Centroids As Variant, ByRef Cluster_Indexes As Variant) As Integer
    Dim NUMCLUSTERS As Integer: NUMCLUSTERS = UBound(Centroids, 1)
    Dim NUMCOLUMNS As Integer: NUMCOLUMNS = UBound(Centroids, 2)
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)
    Dim idx() As Variant: ReDim idx(NUMRECORDS) As Variant
    Dim recordsCounter, clusterCounter As Integer
    Dim changeCounter As Integer: changeCounter = 0

    For recordsCounter = 1 To NUMRECORDS '전체 데이터
    
        Dim MinimumDistance As Double: MinimumDistance = 99999999
        Dim MinCluster As Variant
        Dim dist As Double: dist = 0
        
        For clusterCounter = 1 To NUMCLUSTERS
            '[유클리디안 거리 계산함수] 호출
            dist = EuclideanDistance(Application.Index(DataRecords, recordsCounter, 0), Application.Index(Centroids, clusterCounter, 0), NUMCOLUMNS)
            
            If dist < MinimumDistance Then
                MinCluster = clusterCounter
                MinimumDistance = dist '최소의 거리 계산
            End If
            
        Next clusterCounter
        
        idx(recordsCounter) = MinCluster

        If Not (IsEmpty(Cluster_Indexes)) Then
            If Not (Cluster_Indexes(recordsCounter) = idx(recordsCounter)) Then
                changeCounter = changeCounter + 1
            End If
        End If
        
    Next recordsCounter
    
    FindClosestCentroid = changeCounter
 
    Cluster_Indexes = idx()
    
End Function

'[Result 시트에서 결과 출력 함수]

Public Sub ShowResult(ByRef DataRecords As Variant, ByRef Cluster_Indexes As Variant, ByRef Centroids, NUMCLUSTERS As Integer)

    Dim resultSheet As Worksheet
    Dim lRowLast, lColLast As Integer
    Dim Rng As Range
    Dim ClusterObjects() As Variant: ReDim ClusterObjects(NUMCLUSTERS) As Variant
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)
    
    '시트 초기화
    With resultSheet
        lRowLast = .UsedRange.row + .UsedRange.Rows.Count - 1
        lColLast = .UsedRange.Column + .UsedRange.Columns.Count - 1
        Set Rng = .Range(.Range("B4"), .Cells(lRowLast, lColLast))
    End With
    
    Rng.ClearContents
    
    Dim cluster As Integer
    For cluster = 1 To NUMCLUSTERS
        ClusterObjects(cluster) = 0
        resultSheet.Cells(4, 1 + cluster).Value = cluster
    Next cluster

    Dim counter As Integer
    For counter = 1 To NUMRECORDS
        ClusterObjects(Cluster_Indexes(counter)) = ClusterObjects(Cluster_Indexes(counter)) + 1
    Next counter
    
    '군집별 개수 출력
    resultSheet.Range("B5").Resize(1, NUMCLUSTERS).Value = ClusterObjects
    
    '속성 출력 (변경 가능)
    resultSheet.Range("B8") = "Sepal length(꽃받침의 길이)"
    resultSheet.Range("C8") = "Sepal width(꽃받침의 너비)"
    resultSheet.Range("D8") = "Petal length(꽃잎의 길이)"
    resultSheet.Range("E8") = "Petal width(꽃잎의 너비)"
    
    'Excel 버전에 따라 속성 출력 필요
    resultSheet.Range("D19") = "군집"
    resultSheet.Range("E19") = "응집도"
    resultSheet.Range("D20") = "K1"
    resultSheet.Range("D21") = "K2"
    resultSheet.Range("D22") = "K3"
    resultSheet.Range("D23") = "K4"
    resultSheet.Range("D24") = "K5"
    
    resultSheet.Range("B9").Resize(UBound(Centroids, 1), UBound(Centroids, 2)).Value = Centroids
    
End Sub

'[가까운 중심점 찾기 함수]

Public Function ComputeCentroids(DataRecords As Variant, ClusterIdx As Variant, NoOfClusters As Variant) As Variant
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)
    Dim RecordSize As Integer: RecordSize = UBound(DataRecords, 2)
    Dim ii As Integer: ii = 1
    Dim cc As Integer: cc = 1
    Dim bb As Integer: bb = 1
    Dim counter As Integer: counter = 0
    Dim tempSum() As Variant: ReDim tempSum(NoOfClusters, RecordSize) As Variant
    Dim Centroids() As Variant: ReDim Centroids(NoOfClusters, RecordSize) As Variant
    
    For ii = 1 To NoOfClusters
        For bb = 1 To RecordSize
        
            counter = 0
            
            For cc = 1 To NUMRECORDS
            
                If ClusterIdx(cc) = ii Then
                    Centroids(ii, bb) = Centroids(ii, bb) + DataRecords(cc, bb)
                    counter = counter + 1
                End If
                
            Next cc
            
            If counter > 0 Then
                Centroids(ii, bb) = Centroids(ii, bb) / counter
            Else
                Centroids(ii, bb) = 0
            End If
            
        Next bb
        
    Next ii
    
    ComputeCentroids = Centroids '새로운 중심점
    
End Function
