Attribute VB_Name = "DashboardModule"
'==============================================
' FUNCTIONAL DASHBOARD MODULE - DashboardModule
' Purpose: Generating and updating summary dashboard with real data
' State: FIXED TABLE DISPLAY ISSUE - CORRECT POSITIONING AND FORMATTING
' Version: 1.9.0
' Date: 04.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' Global module variables
Private dashboardTimer As Double
Private isAutoUpdateEnabled As Boolean
Private lastUpdateTime As Date
Private cachedData As Collection
Private cacheTimestamp As Date
Private previousTotalDocs As Long

' Configuration constants
Private Const UPDATE_INTERVAL_SECONDS As Integer = 30
Private Const CACHE_LIFETIME_MINUTES As Integer = 5
Private Const MAX_TOP_SERVICES As Integer = 5
Private Const CRITICAL_DELAY_DAYS As Integer = 7
Private Const WARNING_DELAY_DAYS As Integer = 3

' Correct column names based on translations
Private Const STATUS_COLUMN_NAME As String = "Confirmation Status"
Private Const DATE_COLUMN_NAME As String = "Date Inc.FRP"
Private Const SERVICE_COLUMN_NAME As String = "Service"

' ===============================================
' MAIN DASHBOARD GENERATION PROCEDURE
' ===============================================

Public Sub GenerateDashboard()
    ' Main variables declaration
    Dim wsDashboard As Worksheet
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim StartTime As Double
    Dim processedRecords As Long
    Dim errorMessage As String
    
    ' Filter parameters variables
    Dim dateFrom As Date
    Dim dateTo As Date
    Dim serviceFilter As String
    Dim statusFilter As String
    
    ' Data variables
    Dim filteredData As Collection
    Dim kpiMetrics As Collection
    
    On Error GoTo GenerateError
    
    StartTime = Timer
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Get worksheet and table references
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    Set wsData = CommonUtilities.GetWorksheetSafe("IncOut")
    
    If wsDashboard Is Nothing Or wsData Is Nothing Then
        MsgBox LocalizationManager.GetText("Error: Required sheets not found for Dashboard generation"), vbCritical, LocalizationManager.GetText("Error")
        GoTo CleanupAndExit
    End If
    
    Set tblData = CommonUtilities.GetListObjectSafe(wsData, "TableIncOut")
    
    If tblData Is Nothing Then
        MsgBox LocalizationManager.GetText("Error: Table 'TableIncOut' not found for Dashboard generation"), vbCritical, LocalizationManager.GetText("Error")
        GoTo CleanupAndExit
    End If
    
    ' Save previous value for trend calculation
    On Error Resume Next
    previousTotalDocs = CLng(wsDashboard.Range("A6").value)
    If Err.Number <> 0 Then previousTotalDocs = 0
    Err.Clear
    On Error GoTo GenerateError
    
    ' Get filter parameters
    Call GetFilterParameters(dateFrom, dateTo, serviceFilter, statusFilter)
    
    ' Update generation status
    wsDashboard.Range("I3").value = LocalizationManager.GetText("Generating...")
    wsDashboard.Range("F3").value = Now
    
    Debug.Print "=== DASHBOARD GENERATION ==="
    Debug.Print "Date from: " & dateFrom & " to: " & dateTo
    Debug.Print "Service: " & serviceFilter & ", Status: " & statusFilter
    
    ' Update helper calculations BEFORE data processing
    Call UpdateHelperCalculations(wsDashboard, tblData, dateFrom, dateTo, serviceFilter, statusFilter)
    
    ' Get filtered data with correct column names
    Set filteredData = GetFilteredDataFixed(tblData, dateFrom, dateTo, serviceFilter, statusFilter)
    processedRecords = filteredData.Count
    
    Debug.Print "Filtered records: " & processedRecords
    
    ' Calculate KPI metrics
    Set kpiMetrics = CalculateKPIMetricsFixed(filteredData, tblData)
    
    Debug.Print "=== KPI METRICS ==="
    Debug.Print "Total: " & GetMetricValue(kpiMetrics, "TotalDocs")
    Debug.Print "Confirmed: " & GetMetricValue(kpiMetrics, "ConfirmedDocs")
    Debug.Print "In progress: " & GetMetricValue(kpiMetrics, "InProgressDocs")
    Debug.Print "Overdue: " & GetMetricValue(kpiMetrics, "OverdueDocs")
    
    ' Update all Dashboard sections
    Call UpdateKPIBlocks(wsDashboard, kpiMetrics, previousTotalDocs)
    
    ' CRITICAL FIX: Update tables with proper formatting
    Call UpdateDataTablesWithFormatting(wsDashboard, filteredData, tblData)
    
    Call UpdateTrendsData(wsDashboard, filteredData)
    Call ApplyConditionalFormatting(wsDashboard, kpiMetrics)
    Call UpdateParametersPanel(wsDashboard, dateFrom, dateTo, processedRecords)
    
    ' Create charts with fixed data
    Call CreateChartsFixed(wsDashboard)
    
    ' Record last update time
    lastUpdateTime = Now
    wsDashboard.Range("Q5").value = lastUpdateTime
    wsDashboard.Range("Q5").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    
    ' Update completion status
    wsDashboard.Range("I3").value = LocalizationManager.GetText("Active")
    
CleanupAndExit:
    ' Re-enable updates
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Inform about completion
    Dim processingTime As Double
    processingTime = Timer - StartTime
    
    MsgBox LocalizationManager.GetText("Dashboard updated successfully!") & vbCrLf & vbCrLf & _
           LocalizationManager.GetText("Processed records: ") & processedRecords & vbCrLf & _
           LocalizationManager.GetText("Processing time: ") & Format(processingTime, "0.00") & LocalizationManager.GetText(" sec.") & vbCrLf & _
           LocalizationManager.GetText("Last update: ") & Format(lastUpdateTime, "dd.mm.yyyy hh:mm:ss"), _
           vbInformation, LocalizationManager.GetText("Dashboard Updated")
    
    Call StartAutoUpdate
    
    Exit Sub
    
GenerateError:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    errorMessage = LocalizationManager.GetText("Dashboard generation error: ") & Err.description & vbCrLf & _
                   LocalizationManager.GetText("Error number: ") & Err.Number & vbCrLf & _
                   LocalizationManager.GetText("Processed records: ") & processedRecords
    MsgBox errorMessage, vbCritical, LocalizationManager.GetText("Critical Error")
End Sub

' ===============================================
' HELPER CALCULATIONS UPDATE PROCEDURE
' ===============================================

Private Sub UpdateHelperCalculations(wsDashboard As Worksheet, tblData As ListObject, dateFrom As Date, dateTo As Date, serviceFilter As String, statusFilter As String)
    Dim totalDocs As Long
    Dim confirmedDocs As Long
    Dim sentDocs As Long
    Dim inProgressDocs As Long
    Dim emptyStatusDocs As Long
    Dim overdueDocs As Long
    Dim completionRate As Double
    Dim avgProcessingDays As Double
    Dim periodDays As Long
    
    Dim i As Long
    Dim recordDate As Date
    Dim recordService As String
    Dim recordStatus As String
    Dim daysDiff As Long
    Dim totalProcessingDays As Long
    Dim processedRecords As Long
    
    totalDocs = 0
    confirmedDocs = 0
    sentDocs = 0
    inProgressDocs = 0
    emptyStatusDocs = 0
    overdueDocs = 0
    totalProcessingDays = 0
    processedRecords = 0
    
    Debug.Print "=== UPDATING HELPER CALCULATIONS ==="
    
    If tblData.DataBodyRange Is Nothing Then
        Debug.Print "WARNING: Table is empty"
        GoTo WriteResults
    End If
    
    Dim statusColumnIndex As Integer
    Dim dateColumnIndex As Integer
    Dim serviceColumnIndex As Integer
    
    statusColumnIndex = FindColumnIndex(tblData, STATUS_COLUMN_NAME, 19)
    dateColumnIndex = FindColumnIndex(tblData, DATE_COLUMN_NAME, 8)
    serviceColumnIndex = FindColumnIndex(tblData, SERVICE_COLUMN_NAME, 2)
    
    Debug.Print "Used columns: Status=" & statusColumnIndex & ", Date=" & dateColumnIndex & ", Service=" & serviceColumnIndex
    
    For i = 1 To tblData.ListRows.Count
        On Error Resume Next
        recordDate = CDate(tblData.DataBodyRange.Cells(i, dateColumnIndex).value)
        If Err.Number <> 0 Then recordDate = 0
        Err.Clear
        On Error GoTo 0
        
        If recordDate > 0 And recordDate >= dateFrom And recordDate <= dateTo Then
            recordService = CStr(tblData.DataBodyRange.Cells(i, serviceColumnIndex).value)
            
            If serviceFilter = "All services" Or serviceFilter = "" Or recordService = serviceFilter Then
                recordStatus = CStr(tblData.DataBodyRange.Cells(i, statusColumnIndex).value)
                
                If statusFilter = "All statuses" Or statusFilter = "" Or recordStatus = statusFilter Or _
                   (statusFilter = "(empty)" And recordStatus = "") Then
                    
                    totalDocs = totalDocs + 1
                    
                    If InStr(UCase(recordStatus), "CONFIRM") > 0 Then
                        confirmedDocs = confirmedDocs + 1
                    ElseIf InStr(UCase(recordStatus), "SENT") > 0 And InStr(UCase(recordStatus), "OUT") > 0 Then
                        sentDocs = sentDocs + 1
                    ElseIf recordStatus = "" Or InStr(UCase(recordStatus), "NO ") > 0 Then
                        emptyStatusDocs = emptyStatusDocs + 1
                    End If
                    
                    daysDiff = Date - recordDate
                    If daysDiff > CRITICAL_DELAY_DAYS And Not (InStr(UCase(recordStatus), "CONFIRM") > 0 Or InStr(UCase(recordStatus), "SENT") > 0) Then
                        overdueDocs = overdueDocs + 1
                    End If
                    
                    If InStr(UCase(recordStatus), "CONFIRM") > 0 Or InStr(UCase(recordStatus), "SENT") > 0 Then
                        totalProcessingDays = totalProcessingDays + daysDiff
                        processedRecords = processedRecords + 1
                    End If
                End If
            End If
        End If
    Next i
    
WriteResults:
    inProgressDocs = totalDocs - confirmedDocs - sentDocs - emptyStatusDocs
    If inProgressDocs < 0 Then inProgressDocs = 0
    
    If totalDocs > 0 Then
        completionRate = (confirmedDocs + sentDocs) / totalDocs
    Else
        completionRate = 0
    End If
    
    If processedRecords > 0 Then
        avgProcessingDays = totalProcessingDays / processedRecords
    Else
        avgProcessingDays = 0
    End If
    
    periodDays = dateTo - dateFrom + 1
    
    On Error Resume Next
    wsDashboard.Range("T3").value = totalDocs
    wsDashboard.Range("T4").value = confirmedDocs
    wsDashboard.Range("T5").value = sentDocs
    wsDashboard.Range("T6").value = inProgressDocs
    wsDashboard.Range("T7").value = emptyStatusDocs
    wsDashboard.Range("T8").value = overdueDocs
    
    wsDashboard.Range("T9").value = completionRate
    wsDashboard.Range("T10").value = avgProcessingDays
    wsDashboard.Range("T11").value = periodDays
    wsDashboard.Range("T12").value = Date
    
    wsDashboard.Range("T9").NumberFormat = "0.0%"
    wsDashboard.Range("T10").NumberFormat = "0.0"
    wsDashboard.Range("T12").NumberFormat = "dd.mm.yyyy"
    On Error GoTo 0
    
    Debug.Print "Helper calculations updated:"
    Debug.Print "- Total docs: " & totalDocs
    Debug.Print "- Confirmed: " & confirmedDocs
    Debug.Print "- In progress: " & inProgressDocs
    Debug.Print "- Overdue: " & overdueDocs
End Sub

' ===============================================
' HELPER FUNCTION: FIND COLUMN INDEX
' ===============================================

Private Function FindColumnIndex(tblData As ListObject, ColumnName As String, defaultIndex As Integer) As Integer
    Dim i As Integer
    Dim foundIndex As Integer
    
    foundIndex = defaultIndex
    
    For i = 1 To tblData.ListColumns.Count
        If InStr(UCase(tblData.ListColumns(i).Name), UCase(ColumnName)) > 0 Then
            foundIndex = i
            Exit For
        End If
    Next i
    
    If foundIndex = defaultIndex Then
        Select Case ColumnName
            Case STATUS_COLUMN_NAME
                For i = 1 To tblData.ListColumns.Count
                    If InStr(UCase(tblData.ListColumns(i).Name), "STATUS") > 0 Then
                        foundIndex = i
                        Exit For
                    End If
                Next i
            Case DATE_COLUMN_NAME
                For i = 1 To tblData.ListColumns.Count
                    If InStr(UCase(tblData.ListColumns(i).Name), "DATE") > 0 And InStr(UCase(tblData.ListColumns(i).Name), "FRP") > 0 Then
                        foundIndex = i
                        Exit For
                    End If
                Next i
            Case SERVICE_COLUMN_NAME
                For i = 1 To tblData.ListColumns.Count
                    If InStr(UCase(tblData.ListColumns(i).Name), "SERVICE") > 0 Then
                        foundIndex = i
                        Exit For
                    End If
                Next i
        End Select
    End If
    
    FindColumnIndex = foundIndex
End Function

' ===============================================
' GET FILTER PARAMETERS PROCEDURE
' ===============================================

Private Sub GetFilterParameters(ByRef dateFrom As Date, ByRef dateTo As Date, ByRef serviceFilter As String, ByRef statusFilter As String)
    Dim wsReports As Worksheet
    
    On Error GoTo DefaultParameters
    
    Set wsReports = CommonUtilities.GetWorksheetSafe("Reports")
    
    If Not wsReports Is Nothing Then
        On Error Resume Next
        dateFrom = CDate(wsReports.Range("B18").value)
        If Err.Number <> 0 Then dateFrom = Date - 30
        Err.Clear
        
        dateTo = CDate(wsReports.Range("B19").value)
        If Err.Number <> 0 Then dateTo = Date
        Err.Clear
        
        serviceFilter = CStr(wsReports.Range("B20").value)
        statusFilter = CStr(wsReports.Range("B21").value)
        On Error GoTo DefaultParameters
    Else
        GoTo DefaultParameters
    End If
    
    Exit Sub
    
DefaultParameters:
    dateFrom = Date - 365
    dateTo = Date
    serviceFilter = "All services"
    statusFilter = "All statuses"
    
    Debug.Print "USING DEFAULT PARAMETERS"
    
    On Error GoTo 0
End Sub

' ===============================================
' GET FILTERED DATA FUNCTION
' ===============================================

Private Function GetFilteredDataFixed(tblData As ListObject, dateFrom As Date, dateTo As Date, serviceFilter As String, statusFilter As String) As Collection
    Dim filteredData As Collection
    Dim i As Long
    Dim recordDate As Date
    Dim recordService As String
    Dim recordStatus As String
    Dim record As Collection
    Dim fieldNames As Variant
    Dim j As Integer
    Dim TotalRows As Long
    Dim validDates As Long
    Dim matchedRecords As Long
    
    Dim statusColumnIndex As Integer
    Dim dateColumnIndex As Integer
    Dim serviceColumnIndex As Integer
    
    Set filteredData = New Collection
    
    If tblData.DataBodyRange Is Nothing Then
        Debug.Print "ERROR: Table has no data"
        Set GetFilteredDataFixed = filteredData
        Exit Function
    End If
    
    TotalRows = tblData.ListRows.Count
    validDates = 0
    matchedRecords = 0
    
    statusColumnIndex = FindColumnIndex(tblData, STATUS_COLUMN_NAME, 19)
    dateColumnIndex = FindColumnIndex(tblData, DATE_COLUMN_NAME, 8)
    serviceColumnIndex = FindColumnIndex(tblData, SERVICE_COLUMN_NAME, 2)
    
    Debug.Print "=== PROCESSING DATA ==="
    Debug.Print "Total rows: " & TotalRows & ", Indices: S=" & statusColumnIndex & ", D=" & dateColumnIndex & ", SRV=" & serviceColumnIndex
    
    ' Translated internal array keys
    fieldNames = Array("SeqNo", "Service", "DocGroup", "DocType", "DocNumber", _
                      "DocAmount", "FRPNumber", "FRPDate", "ReceivedFrom", "TransferDate", _
                      "Executor", "OutServiceNo", "OutServiceDate", "ReturnNo", _
                      "ReturnDate", "OutEnvelopeNo", "OutEnvelopeDate", "ExecutionMark", "Status")
    
    For i = 1 To TotalRows
        On Error Resume Next
        recordDate = CDate(tblData.DataBodyRange.Cells(i, dateColumnIndex).value)
        If Err.Number <> 0 Then
            recordDate = 0
        Else
            validDates = validDates + 1
        End If
        Err.Clear
        On Error GoTo 0
        
        If recordDate > 0 Then
            recordService = CStr(tblData.DataBodyRange.Cells(i, serviceColumnIndex).value)
            
            If serviceFilter = "All services" Or serviceFilter = "" Or recordService = serviceFilter Then
                recordStatus = CStr(tblData.DataBodyRange.Cells(i, statusColumnIndex).value)
                
                Dim statusMatch As Boolean
                statusMatch = False
                
                If statusFilter = "All statuses" Or statusFilter = "" Then
                    statusMatch = True
                ElseIf statusFilter = "(empty)" And recordStatus = "" Then
                    statusMatch = True
                ElseIf recordStatus = statusFilter Then
                    statusMatch = True
                End If
                
                If statusMatch Then
                    If recordDate >= dateFrom And recordDate <= dateTo Then
                        matchedRecords = matchedRecords + 1
                        
                        Set record = New Collection
                        
                        For j = 0 To UBound(fieldNames)
                            If j + 1 <= tblData.ListColumns.Count Then
                                record.Add tblData.DataBodyRange.Cells(i, j + 1).value, CStr(fieldNames(j))
                            End If
                        Next j
                        
                        filteredData.Add record
                    End If
                End If
            End If
        End If
    Next i
    
    Debug.Print "Valid dates: " & validDates & ", Passed filter: " & matchedRecords
    
    Set GetFilteredDataFixed = filteredData
End Function

' ===============================================
' KPI METRICS CALCULATION FUNCTION
' ===============================================

Private Function CalculateKPIMetricsFixed(filteredData As Collection, tblData As ListObject) As Collection
    Dim metrics As Collection
    Dim totalDocs As Long
    Dim confirmedDocs As Long
    Dim sentDocs As Long
    Dim inProgressDocs As Long
    Dim overdueDocs As Long
    Dim emptyStatusDocs As Long
    Dim completionRate As Double
    Dim avgProcessingDays As Double
    Dim record As Collection
    Dim recordDate As Date
    Dim recordStatus As String
    Dim daysDiff As Long
    Dim totalDays As Long
    Dim processedRecords As Long
    
    Dim maxDelayDays As Long
    Dim avgInProgressDays As Double
    Dim totalInProgressDays As Long
    Dim inProgressRecords As Long
    
    Set metrics = New Collection
    
    totalDocs = filteredData.Count
    confirmedDocs = 0
    sentDocs = 0
    emptyStatusDocs = 0
    overdueDocs = 0
    totalDays = 0
    processedRecords = 0
    maxDelayDays = 0
    totalInProgressDays = 0
    inProgressRecords = 0
    
    Debug.Print "=== STATUS ANALYSIS (FIXED) ==="
    
    For Each record In filteredData
        On Error Resume Next
        recordStatus = CStr(record.item("Status"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        On Error GoTo 0
        
        If InStr(UCase(recordStatus), "CONFIRM") > 0 Then
            confirmedDocs = confirmedDocs + 1
        ElseIf InStr(UCase(recordStatus), "SENT") > 0 And InStr(UCase(recordStatus), "OUT") > 0 Then
            sentDocs = sentDocs + 1
        ElseIf recordStatus = "" Or InStr(UCase(recordStatus), "NO ") > 0 Then
            emptyStatusDocs = emptyStatusDocs + 1
        End If
        
        On Error Resume Next
        recordDate = CDate(record.item("FRPDate"))
        If Err.Number = 0 And recordDate > 0 Then
            daysDiff = Date - recordDate
            
            If daysDiff > maxDelayDays Then
                maxDelayDays = daysDiff
            End If
            
            If daysDiff > CRITICAL_DELAY_DAYS And Not (InStr(UCase(recordStatus), "CONFIRM") > 0 Or InStr(UCase(recordStatus), "SENT") > 0) Then
                overdueDocs = overdueDocs + 1
            End If
            
            If InStr(UCase(recordStatus), "CONFIRM") > 0 Or InStr(UCase(recordStatus), "SENT") > 0 Then
                totalDays = totalDays + daysDiff
                processedRecords = processedRecords + 1
            End If
            
            If Not (InStr(UCase(recordStatus), "CONFIRM") > 0 Or InStr(UCase(recordStatus), "SENT") > 0) Then
                totalInProgressDays = totalInProgressDays + daysDiff
                inProgressRecords = inProgressRecords + 1
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next record
    
    inProgressDocs = totalDocs - confirmedDocs - sentDocs - emptyStatusDocs
    If inProgressDocs < 0 Then inProgressDocs = 0
    
    If totalDocs > 0 Then
        completionRate = (confirmedDocs + sentDocs) / totalDocs
    Else
        completionRate = 0
    End If
    
    If processedRecords > 0 Then
        avgProcessingDays = totalDays / processedRecords
    Else
        avgProcessingDays = 0
    End If
    
    If inProgressRecords > 0 Then
        avgInProgressDays = totalInProgressDays / inProgressRecords
    Else
        avgInProgressDays = 0
    End If
    
    On Error Resume Next
    metrics.Add totalDocs, "TotalDocs"
    metrics.Add confirmedDocs, "ConfirmedDocs"
    metrics.Add sentDocs, "SentDocs"
    metrics.Add inProgressDocs, "InProgressDocs"
    metrics.Add overdueDocs, "OverdueDocs"
    metrics.Add emptyStatusDocs, "EmptyStatusDocs"
    metrics.Add completionRate, "CompletionRate"
    metrics.Add avgProcessingDays, "AvgProcessingDays"
    metrics.Add maxDelayDays, "MaxDelayDays"
    metrics.Add avgInProgressDays, "AvgInProgressDays"
    On Error GoTo 0
    
    Set CalculateKPIMetricsFixed = metrics
End Function

' ===============================================
' KPI BLOCKS UPDATE PROCEDURE
' ===============================================

Private Sub UpdateKPIBlocks(wsDashboard As Worksheet, metrics As Collection, previousTotal As Long)
    Dim totalDocs As Long
    Dim confirmedDocs As Long
    Dim inProgressDocs As Long
    Dim overdueDocs As Long
    Dim completionRate As Double
    Dim maxDelayDays As Long
    Dim avgInProgressDays As Double
    Dim dynamicChange As Long
    Dim dynamicText As String
    Dim dynamicIcon As String
    
    On Error Resume Next
    
    totalDocs = GetMetricValue(metrics, "TotalDocs")
    confirmedDocs = GetMetricValue(metrics, "ConfirmedDocs")
    inProgressDocs = GetMetricValue(metrics, "InProgressDocs")
    overdueDocs = GetMetricValue(metrics, "OverdueDocs")
    completionRate = GetMetricValue(metrics, "CompletionRate")
    maxDelayDays = GetMetricValue(metrics, "MaxDelayDays")
    avgInProgressDays = GetMetricValue(metrics, "AvgInProgressDays")
    
    wsDashboard.Range("A6").value = totalDocs
    wsDashboard.Range("E6").value = confirmedDocs
    wsDashboard.Range("E8").value = completionRate
    wsDashboard.Range("I6").value = inProgressDocs
    wsDashboard.Range("M6").value = overdueDocs
    
    dynamicChange = totalDocs - previousTotal
    If dynamicChange > 0 Then
        dynamicIcon = "?"
        dynamicText = "+" & dynamicChange & LocalizationManager.GetText(" vs previous")
    ElseIf dynamicChange < 0 Then
        dynamicIcon = "?"
        dynamicText = dynamicChange & LocalizationManager.GetText(" vs previous")
    Else
        dynamicIcon = ">"
        dynamicText = LocalizationManager.GetText("No changes")
    End If
    wsDashboard.Range("A8").value = dynamicIcon & " " & dynamicText
    
    If avgInProgressDays > 0 Then
        wsDashboard.Range("I8").value = LocalizationManager.GetText("Avg. time: ") & Format(avgInProgressDays, "0.0") & LocalizationManager.GetText(" days")
    Else
        wsDashboard.Range("I8").value = LocalizationManager.GetText("Avg. time: ") & LocalizationManager.GetText("N/A")
    End If
    
    If maxDelayDays > 0 Then
        wsDashboard.Range("M8").value = LocalizationManager.GetText("Max delay: ") & maxDelayDays & LocalizationManager.GetText(" days")
    Else
        wsDashboard.Range("M8").value = LocalizationManager.GetText("No delays")
    End If
    
    With wsDashboard.Range("A8,I8,M8")
        .Font.Size = 9
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With
    
    If dynamicChange > 0 Then
        wsDashboard.Range("A8").Font.Color = RGB(0, 150, 0)
    ElseIf dynamicChange < 0 Then
        wsDashboard.Range("A8").Font.Color = RGB(200, 0, 0)
    Else
        wsDashboard.Range("A8").Font.Color = RGB(100, 100, 100)
    End If
    
    If avgInProgressDays > 7 Then
        wsDashboard.Range("I8").Font.Color = RGB(200, 100, 0)
    ElseIf avgInProgressDays > 3 Then
        wsDashboard.Range("I8").Font.Color = RGB(150, 150, 0)
    Else
        wsDashboard.Range("I8").Font.Color = RGB(0, 150, 0)
    End If
    
    If maxDelayDays > CRITICAL_DELAY_DAYS Then
        wsDashboard.Range("M8").Font.Color = RGB(200, 0, 0)
    ElseIf maxDelayDays > WARNING_DELAY_DAYS Then
        wsDashboard.Range("M8").Font.Color = RGB(200, 100, 0)
    Else
        wsDashboard.Range("M8").Font.Color = RGB(0, 150, 0)
    End If
    
    On Error GoTo 0
End Sub

' ===============================================
' HELPER FUNCTION: SAFE GET FROM COLLECTION
' ===============================================

Private Function GetMetricValue(metrics As Collection, keyName As String) As Variant
    Dim result As Variant
    
    On Error Resume Next
    result = metrics.item(keyName)
    If Err.Number <> 0 Then
        result = 0
    End If
    Err.Clear
    On Error GoTo 0
    
    GetMetricValue = result
End Function

' ===============================================
' NEW PROCEDURE: UPDATE TABLES WITH FORMATTING
' ===============================================

Private Sub UpdateDataTablesWithFormatting(wsDashboard As Worksheet, filteredData As Collection, tblData As ListObject)
    Debug.Print "=== UPDATING TABLES WITH FORMATTING ==="
    
    ' CRITICAL FIX: First clear table areas
    Call ClearTableAreas(wsDashboard)
    
    ' Create table headers
    Call CreateTableHeaders(wsDashboard)
    
    ' Update data tables
    Call UpdateTopServicesTableFormatted(wsDashboard, filteredData)
    Call UpdateProblemDocsTableFormatted(wsDashboard, filteredData)
    Call UpdateChartDataTablesFixed(wsDashboard, filteredData)
    
    ' Apply formatting
    Call ApplyTableFormatting(wsDashboard)
End Sub

' ===============================================
' CLEAR TABLE AREAS PROCEDURE
' ===============================================

Private Sub ClearTableAreas(wsDashboard As Worksheet)
    On Error Resume Next
    
    ' Clear top services table area (A27:F35)
    wsDashboard.Range("A27:F35").ClearContents
    wsDashboard.Range("A27:F35").Interior.ColorIndex = xlNone
    wsDashboard.Range("A27:F35").Borders.LineStyle = xlNone
    
    ' Clear problem docs table area (H27:M35)
    wsDashboard.Range("H27:M35").ClearContents
    wsDashboard.Range("H27:M35").Interior.ColorIndex = xlNone
    wsDashboard.Range("H27:M35").Borders.LineStyle = xlNone
    
    ' Clear hidden chart data
    wsDashboard.Range("V4:W8").ClearContents
    wsDashboard.Range("V11:W17").ClearContents
    
    On Error GoTo 0
    
    Debug.Print "Table areas cleared"
End Sub

' ===============================================
' CREATE TABLE HEADERS PROCEDURE
' ===============================================

Private Sub CreateTableHeaders(wsDashboard As Worksheet)
    On Error Resume Next
    
    ' Top services headers
    wsDashboard.Range("A27").value = LocalizationManager.GetText("No.")
    wsDashboard.Range("B27").value = LocalizationManager.GetText("MOST ACTIVE SERVICES")
    wsDashboard.Range("C27").value = LocalizationManager.GetText("Total")
    wsDashboard.Range("D27").value = LocalizationManager.GetText("Completed")
    wsDashboard.Range("E27").value = "%"
    wsDashboard.Range("F27").value = LocalizationManager.GetText("Trend")
    
    ' Problem docs headers
    wsDashboard.Range("H27").value = LocalizationManager.GetText("DOCUMENTS REQUIRING ATTENTION")
    wsDashboard.Range("I27").value = LocalizationManager.GetText("Service")
    wsDashboard.Range("J27").value = LocalizationManager.GetText("Doc No.")
    wsDashboard.Range("K27").value = LocalizationManager.GetText("Date")
    wsDashboard.Range("L27").value = LocalizationManager.GetText("Days")
    wsDashboard.Range("M27").value = LocalizationManager.GetText("Status")
    
    ' Headers formatting
    With wsDashboard.Range("A27:F27")
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(230, 230, 230)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With wsDashboard.Range("H27:M27")
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(255, 230, 230)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    On Error GoTo 0
    
    Debug.Print "Table headers created"
End Sub

' ===============================================
' FIXED UPDATE TOP SERVICES TABLE WITH FORMATTING
' ===============================================

Private Sub UpdateTopServicesTableFormatted(wsDashboard As Worksheet, filteredData As Collection)
    Dim serviceStats As Object
    Dim serviceName As String
    Dim record As Collection
    Dim sortedServices As Collection
    Dim i As Integer
    Dim topCount As Integer
    Dim recordStatus As String
    Dim totalCount As Long
    Dim completedCount As Long
    Dim serviceKey As Variant
    
    Set serviceStats = CreateObject("Scripting.Dictionary")
    
    Debug.Print "=== UPDATING TOP SERVICES TABLE WITH FORMATTING ==="
    
    ' Collect service stats
    For Each record In filteredData
        On Error Resume Next
        serviceName = CStr(record.item("Service"))
        If Err.Number <> 0 Then serviceName = ""
        Err.Clear
        On Error GoTo 0
        
        If serviceName <> "" Then
            On Error Resume Next
            recordStatus = CStr(record.item("Status"))
            If Err.Number <> 0 Then recordStatus = ""
            Err.Clear
            On Error GoTo 0
            
            If Not serviceStats.Exists(serviceName) Then
                If InStr(UCase(recordStatus), "CONFIRM") > 0 Or InStr(UCase(recordStatus), "SENT") > 0 Then
                    serviceStats.Add serviceName, Array(1, 1)
                Else
                    serviceStats.Add serviceName, Array(1, 0)
                End If
            Else
                Dim currentData As Variant
                currentData = serviceStats(serviceName)
                totalCount = currentData(0) + 1
                If InStr(UCase(recordStatus), "CONFIRM") > 0 Or InStr(UCase(recordStatus), "SENT") > 0 Then
                    completedCount = currentData(1) + 1
                Else
                    completedCount = currentData(1)
                End If
                serviceStats(serviceName) = Array(totalCount, completedCount)
            End If
        End If
    Next record
    
    Debug.Print "Unique services found: " & serviceStats.Count
    
    ' Create sorted list
    Set sortedServices = New Collection
    For Each serviceKey In serviceStats.Keys
        sortedServices.Add serviceKey
    Next serviceKey
    
    topCount = IIf(sortedServices.Count < MAX_TOP_SERVICES, sortedServices.Count, MAX_TOP_SERVICES)
    
    ' Fill table with formatting
    For i = 1 To MAX_TOP_SERVICES
        Dim RowIndex As Integer
        RowIndex = 28 + i - 1  ' Rows 28-32
        
        If i <= topCount Then
            serviceName = sortedServices.item(i)
            Dim serviceData As Variant
            serviceData = serviceStats(serviceName)
            
            ' CRITICAL FIX: Fill cells safely
            On Error Resume Next
            wsDashboard.Cells(RowIndex, 1).value = i
            wsDashboard.Cells(RowIndex, 2).value = serviceName
            wsDashboard.Cells(RowIndex, 3).value = serviceData(0) ' Total
            wsDashboard.Cells(RowIndex, 4).value = serviceData(1) ' Completed
            If serviceData(0) > 0 Then
                wsDashboard.Cells(RowIndex, 5).value = serviceData(1) / serviceData(0)
                wsDashboard.Cells(RowIndex, 5).NumberFormat = "0.0%"
            Else
                wsDashboard.Cells(RowIndex, 5).value = 0
                wsDashboard.Cells(RowIndex, 5).NumberFormat = "0.0%"
            End If
            wsDashboard.Cells(RowIndex, 6).value = "?"
            On Error GoTo 0
            
            ' Format row
            With wsDashboard.Range("A" & RowIndex & ":F" & RowIndex)
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.Color = RGB(245, 245, 245)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.Color = RGB(200, 200, 200)
            End With
            
        Else
            ' Clear empty rows
            wsDashboard.Range("A" & RowIndex & ":F" & RowIndex).ClearContents
        End If
    Next i
    
    Debug.Print "Top services table updated"
End Sub

' ===============================================
' FIXED UPDATE PROBLEM DOCS TABLE WITH FORMATTING
' ===============================================

Private Sub UpdateProblemDocsTableFormatted(wsDashboard As Worksheet, filteredData As Collection)
    Dim problemDocs As Collection
    Dim record As Collection
    Dim recordDate As Date
    Dim daysDiff As Long
    Dim problemRecord As Collection
    Dim i As Integer
    Dim maxProblems As Integer
    Dim recordStatus As String
    Dim serviceName As String
    Dim docNumber As String
    
    Set problemDocs = New Collection
    maxProblems = 5
    
    Debug.Print "=== UPDATING PROBLEM DOCS TABLE WITH FORMATTING ==="
    
    ' Find problem docs
    For Each record In filteredData
        On Error Resume Next
        recordDate = CDate(record.item("FRPDate"))
        If Err.Number <> 0 Then recordDate = 0
        Err.Clear
        
        recordStatus = CStr(record.item("Status"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        
        serviceName = CStr(record.item("Service"))
        If Err.Number <> 0 Then serviceName = ""
        Err.Clear
        
        docNumber = CStr(record.item("DocNumber"))
        If Err.Number <> 0 Then docNumber = ""
        Err.Clear
        On Error GoTo 0
        
        If recordDate > 0 Then
            daysDiff = Date - recordDate
            
            Dim isProblematic As Boolean
            isProblematic = False
            
            If daysDiff > WARNING_DELAY_DAYS And recordStatus = "" Then
                isProblematic = True
            ElseIf daysDiff > CRITICAL_DELAY_DAYS And Not (InStr(UCase(recordStatus), "CONFIRM") > 0 Or InStr(UCase(recordStatus), "SENT") > 0) Then
                isProblematic = True
            End If
            
            If isProblematic Then
                Set problemRecord = New Collection
                problemRecord.Add IIf(daysDiff > CRITICAL_DELAY_DAYS, LocalizationManager.GetText("Critical"), LocalizationManager.GetText("Warning")), "Type"
                problemRecord.Add serviceName, "Service"
                problemRecord.Add docNumber, "DocNumber"
                problemRecord.Add recordDate, "Date"
                problemRecord.Add daysDiff, "Days"
                problemRecord.Add recordStatus, "Status"
                
                problemDocs.Add problemRecord
                
                If problemDocs.Count >= maxProblems Then Exit For
            End If
        End If
    Next record
    
    Debug.Print "Problem docs found: " & problemDocs.Count
    
    ' Fill table with formatting
    For i = 1 To maxProblems
        Dim RowIndex As Integer
        RowIndex = 28 + i - 1  ' Rows 28-32
        
        If i <= problemDocs.Count Then
            Set problemRecord = problemDocs.item(i)
            
            ' CRITICAL FIX: Safe filling
            On Error Resume Next
            wsDashboard.Cells(RowIndex, 8).value = problemRecord.item("Type")
            wsDashboard.Cells(RowIndex, 9).value = problemRecord.item("Service")
            wsDashboard.Cells(RowIndex, 10).value = problemRecord.item("DocNumber")
            wsDashboard.Cells(RowIndex, 11).value = problemRecord.item("Date")
            wsDashboard.Cells(RowIndex, 11).NumberFormat = "dd.mm.yyyy"
            wsDashboard.Cells(RowIndex, 12).value = problemRecord.item("Days")
            wsDashboard.Cells(RowIndex, 13).value = IIf(problemRecord.item("Status") = "", LocalizationManager.GetText("Not set"), problemRecord.item("Status"))
            On Error GoTo 0
            
            ' Conditional formatting by criticality
            If problemRecord.item("Type") = LocalizationManager.GetText("Critical") Then
                With wsDashboard.Range("H" & RowIndex & ":M" & RowIndex)
                    .Interior.Color = RGB(255, 200, 200)
                    .Font.Color = RGB(150, 0, 0)
                    .Font.Bold = True
                End With
            Else
                With wsDashboard.Range("H" & RowIndex & ":M" & RowIndex)
                    .Interior.Color = RGB(255, 255, 200)
                    .Font.Color = RGB(150, 100, 0)
                End With
            End If
            
            ' General formatting
            With wsDashboard.Range("H" & RowIndex & ":M" & RowIndex)
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.Color = RGB(200, 200, 200)
            End With
            
        Else
            ' Clear empty rows
            wsDashboard.Range("H" & RowIndex & ":M" & RowIndex).ClearContents
        End If
    Next i
    
    Debug.Print "Problem docs table updated"
End Sub

' ===============================================
' GENERAL TABLE FORMATTING PROCEDURE
' ===============================================

Private Sub ApplyTableFormatting(wsDashboard As Worksheet)
    On Error Resume Next
    
    ' Top services
    With wsDashboard.Range("A27:F32")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(150, 150, 150)
    End With
    
    wsDashboard.Columns("A:A").ColumnWidth = 3
    wsDashboard.Columns("B:B").ColumnWidth = 20
    wsDashboard.Columns("C:C").ColumnWidth = 8
    wsDashboard.Columns("D:D").ColumnWidth = 10
    wsDashboard.Columns("E:E").ColumnWidth = 8
    wsDashboard.Columns("F:F").ColumnWidth = 6
    
    ' Problem docs
    With wsDashboard.Range("H27:M32")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(150, 150, 150)
    End With
    
    wsDashboard.Columns("H:H").ColumnWidth = 12
    wsDashboard.Columns("I:I").ColumnWidth = 15
    wsDashboard.Columns("J:J").ColumnWidth = 12
    wsDashboard.Columns("K:K").ColumnWidth = 10
    wsDashboard.Columns("L:L").ColumnWidth = 6
    wsDashboard.Columns("M:M").ColumnWidth = 12
    
    ' Vertical align
    wsDashboard.Range("A27:F32").VerticalAlignment = xlCenter
    wsDashboard.Range("H27:M32").VerticalAlignment = xlCenter
    
    On Error GoTo 0
End Sub

' ===============================================
' UPDATE CHART DATA PROCEDURE
' ===============================================

Private Sub UpdateChartDataTablesFixed(wsDashboard As Worksheet, filteredData As Collection)
    Dim confirmedCount As Long, sentCount As Long, inProgressCount As Long, emptyCount As Long
    Dim record As Collection
    Dim recordStatus As String
    
    confirmedCount = 0
    sentCount = 0
    inProgressCount = 0
    emptyCount = 0
    
    Debug.Print "=== PREPARING CHART DATA (FIXED) ==="
    
    For Each record In filteredData
        On Error Resume Next
        recordStatus = CStr(record.item("Status"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        On Error GoTo 0
        
        If InStr(UCase(recordStatus), "CONFIRM") > 0 Then
            confirmedCount = confirmedCount + 1
        ElseIf InStr(UCase(recordStatus), "SENT") > 0 And InStr(UCase(recordStatus), "OUT") > 0 Then
            sentCount = sentCount + 1
        ElseIf recordStatus = "" Or InStr(UCase(recordStatus), "NO ") > 0 Then
            emptyCount = emptyCount + 1
        Else
            inProgressCount = inProgressCount + 1
        End If
    Next record
    
    On Error Resume Next
    wsDashboard.Range("V4:W8").ClearContents
    wsDashboard.Range("V11:W17").ClearContents
    
    wsDashboard.Range("V4").value = LocalizationManager.GetText("Status")
    wsDashboard.Range("W4").value = LocalizationManager.GetText("Count")
    
    wsDashboard.Range("V5").value = LocalizationManager.GetText("Confirmed")
    wsDashboard.Range("W5").value = confirmedCount
    
    wsDashboard.Range("V6").value = LocalizationManager.GetText("Sent Out.")
    wsDashboard.Range("W6").value = sentCount
    
    wsDashboard.Range("V7").value = LocalizationManager.GetText("In progress")
    wsDashboard.Range("W7").value = inProgressCount
    
    wsDashboard.Range("V8").value = LocalizationManager.GetText("No status")
    wsDashboard.Range("W8").value = emptyCount
    
    wsDashboard.Range("V11").value = LocalizationManager.GetText("Service")
    wsDashboard.Range("W11").value = LocalizationManager.GetText("Count")
    
    Dim i As Integer
    Dim hasServiceData As Boolean
    hasServiceData = False
    
    For i = 1 To 6
        If wsDashboard.Cells(27 + i, 2).value <> "" And IsNumeric(wsDashboard.Cells(27 + i, 3).value) And wsDashboard.Cells(27 + i, 3).value > 0 Then
            wsDashboard.Cells(11 + i, 22).value = wsDashboard.Cells(27 + i, 2).value
            wsDashboard.Cells(11 + i, 23).value = wsDashboard.Cells(27 + i, 3).value
            hasServiceData = True
        Else
            wsDashboard.Cells(11 + i, 22).ClearContents
            wsDashboard.Cells(11 + i, 23).ClearContents
        End If
    Next i
    
    On Error GoTo 0
End Sub

' ===============================================
' CREATE CHARTS PROCEDURE
' ===============================================

Private Sub CreateChartsFixed(wsDashboard As Worksheet)
    Call CreatePieChartFixed(wsDashboard)
    Call CreateColumnChartFixed(wsDashboard)
End Sub

' ===============================================
' CREATE PIE CHART
' ===============================================

Private Sub CreatePieChartFixed(wsDashboard As Worksheet)
    Dim chartRange As Range
    Dim chartData As Range
    Dim chartObj As ChartObject
    Dim totalValues As Long
    
    On Error GoTo PieChartError
    
    On Error Resume Next
    wsDashboard.ChartObjects("StatusPieChart").Delete
    On Error GoTo PieChartError
    
    totalValues = CLng(wsDashboard.Range("W5").value) + CLng(wsDashboard.Range("W6").value) + _
                  CLng(wsDashboard.Range("W7").value) + CLng(wsDashboard.Range("W8").value)
    
    If totalValues = 0 Then
        wsDashboard.Range("C15").value = LocalizationManager.GetText("No data for pie chart")
        wsDashboard.Range("C15").Font.Color = RGB(150, 150, 150)
        wsDashboard.Range("C15").Font.Italic = True
        Exit Sub
    End If
    
    wsDashboard.Range("C15").ClearContents
    
    Set chartData = wsDashboard.Range("V4:W8")
    Set chartRange = wsDashboard.Range("A11:G25")
    
    Set chartObj = wsDashboard.ChartObjects.Add(Left:=chartRange.Left + 5, Top:=chartRange.Top + 5, _
                                                Width:=chartRange.Width - 10, Height:=chartRange.Height - 10)
    chartObj.Name = "StatusPieChart"
    
    With chartObj.Chart
        .SetSourceData Source:=chartData, PlotBy:=xlColumns
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = LocalizationManager.GetText("Status Distribution")
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        
        On Error Resume Next
        If .SeriesCollection.Count > 0 Then
            .SeriesCollection(1).HasDataLabels = True
            .SeriesCollection(1).DataLabels.ShowPercentage = True
            
            If .SeriesCollection(1).Points.Count >= 4 Then
                .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(200, 255, 200)
                .SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(200, 200, 255)
                .SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = RGB(255, 255, 200)
                .SeriesCollection(1).Points(4).Format.Fill.ForeColor.RGB = RGB(255, 200, 200)
            End If
        End If
        On Error GoTo PieChartError
    End With
    
    Exit Sub
    
PieChartError:
    wsDashboard.Range("C15").value = LocalizationManager.GetText("Chart error: ") & Err.description
    wsDashboard.Range("C15").Font.Color = RGB(255, 0, 0)
End Sub

' ===============================================
' CREATE COLUMN CHART
' ===============================================

Private Sub CreateColumnChartFixed(wsDashboard As Worksheet)
    Dim chartRange As Range
    Dim chartData As Range
    Dim chartObj As ChartObject
    Dim hasValidData As Boolean
    Dim i As Integer
    
    On Error GoTo ColumnChartError
    
    On Error Resume Next
    wsDashboard.ChartObjects("ServiceColumnChart").Delete
    On Error GoTo ColumnChartError
    
    hasValidData = False
    For i = 12 To 17
        If wsDashboard.Cells(i, 22).value <> "" And IsNumeric(wsDashboard.Cells(i, 23).value) And wsDashboard.Cells(i, 23).value > 0 Then
            hasValidData = True
            Exit For
        End If
    Next i
    
    If Not hasValidData Then
        wsDashboard.Range("L15").value = LocalizationManager.GetText("No service data")
        wsDashboard.Range("L15").Font.Color = RGB(150, 150, 150)
        wsDashboard.Range("L15").Font.Italic = True
        Exit Sub
    End If
    
    wsDashboard.Range("L15").ClearContents
    
    Set chartData = wsDashboard.Range("V11:W17")
    Set chartRange = wsDashboard.Range("I11:O25")
    
    Set chartObj = wsDashboard.ChartObjects.Add(Left:=chartRange.Left + 5, Top:=chartRange.Top + 5, _
                                                Width:=chartRange.Width - 10, Height:=chartRange.Height - 10)
    chartObj.Name = "ServiceColumnChart"
    
    With chartObj.Chart
        .SetSourceData Source:=chartData, PlotBy:=xlColumns
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = LocalizationManager.GetText("Activity by Services")
        .HasLegend = False
        
        On Error Resume Next
        If .Axes.Count >= 2 Then
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = LocalizationManager.GetText("Service")
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = LocalizationManager.GetText("Count")
        End If
        
        If .SeriesCollection.Count > 0 Then
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(100, 150, 255)
            .SeriesCollection(1).HasDataLabels = True
        End If
        On Error GoTo ColumnChartError
    End With
    
    Exit Sub
    
ColumnChartError:
    wsDashboard.Range("L15").value = LocalizationManager.GetText("Chart error: ") & Err.description
    wsDashboard.Range("L15").Font.Color = RGB(255, 0, 0)
End Sub

' ===============================================
' OTHER PROCEDURES
' ===============================================

Private Sub UpdateTrendsData(wsDashboard As Worksheet, filteredData As Collection)
    ' Empty for now
End Sub

Private Sub ApplyConditionalFormatting(wsDashboard As Worksheet, metrics As Collection)
    Dim overdueDocs As Long
    Dim totalDocs As Long
    Dim completionRate As Double
    
    overdueDocs = GetMetricValue(metrics, "OverdueDocs")
    totalDocs = GetMetricValue(metrics, "TotalDocs")
    completionRate = GetMetricValue(metrics, "CompletionRate")
    
    If overdueDocs > 0 Then
        wsDashboard.Range("M5:O8").Interior.Color = RGB(255, 150, 150)
    Else
        wsDashboard.Range("M5:O8").Interior.Color = RGB(255, 180, 180)
    End If
    
    If completionRate >= 0.8 Then
        wsDashboard.Range("E5:G8").Interior.Color = RGB(180, 255, 180)
    ElseIf completionRate >= 0.5 Then
        wsDashboard.Range("E5:G8").Interior.Color = RGB(255, 255, 180)
    Else
        wsDashboard.Range("E5:G8").Interior.Color = RGB(255, 200, 200)
    End If
End Sub

Private Sub UpdateParametersPanel(wsDashboard As Worksheet, dateFrom As Date, dateTo As Date, processedRecords As Long)
    wsDashboard.Range("Q5").value = Now
    
    If processedRecords > 0 Then
        wsDashboard.Range("Q6").value = "OK (" & processedRecords & ")"
    Else
        wsDashboard.Range("Q6").value = "N/A"
    End If
    
    wsDashboard.Range("Q8").value = Format(dateFrom, "dd.mm.yyyy") & " - " & Format(dateTo, "dd.mm.yyyy")
End Sub

Public Sub StartAutoUpdate()
    Dim wsDashboard As Worksheet
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    
    If Not wsDashboard Is Nothing Then
        If UCase(CStr(wsDashboard.Range("Q3").value)) = "ON" Then
            isAutoUpdateEnabled = True
            dashboardTimer = UPDATE_INTERVAL_SECONDS
            Application.OnTime Now + TimeSerial(0, 0, UPDATE_INTERVAL_SECONDS), "DashboardModule.AutoUpdateDashboard"
        End If
    End If
End Sub

Public Sub StopAutoUpdate()
    isAutoUpdateEnabled = False
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, UPDATE_INTERVAL_SECONDS), "DashboardModule.AutoUpdateDashboard", , False
    On Error GoTo 0
End Sub

Public Sub AutoUpdateDashboard()
    If isAutoUpdateEnabled Then
        Call GenerateDashboard
        Application.OnTime Now + TimeSerial(0, 0, UPDATE_INTERVAL_SECONDS), "DashboardModule.AutoUpdateDashboard"
    End If
End Sub

Public Sub RefreshDashboard()
    Call GenerateDashboard
End Sub

Public Sub ToggleAutoUpdate()
    Dim wsDashboard As Worksheet
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    
    If Not wsDashboard Is Nothing Then
        If UCase(CStr(wsDashboard.Range("Q3").value)) = "ON" Then
            wsDashboard.Range("Q3").value = "OFF"
            Call StopAutoUpdate
            MsgBox LocalizationManager.GetText("Auto-update disabled"), vbInformation
        Else
            wsDashboard.Range("Q3").value = "ON"
            Call StartAutoUpdate
            MsgBox LocalizationManager.GetText("Auto-update enabled"), vbInformation
        End If
    End If
End Sub

Public Sub ExportDashboard()
    MsgBox LocalizationManager.GetText("Export function will be implemented in future versions"), vbInformation
End Sub

Public Sub DiagnoseDashboardData()
    Dim wsDashboard As Worksheet
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    Set wsData = CommonUtilities.GetWorksheetSafe("IncOut")
    Set tblData = CommonUtilities.GetListObjectSafe(wsData, "TableIncOut")
    
    Debug.Print "=== DASHBOARD DIAGNOSTICS ==="
    Debug.Print "Dashboard sheet found: " & Not (wsDashboard Is Nothing)
    Debug.Print "IncOut sheet found: " & Not (wsData Is Nothing)
    Debug.Print "Table found: " & Not (tblData Is Nothing)
    
    MsgBox LocalizationManager.GetText("Dashboard diagnostics completed. Check Immediate window (Ctrl+G)"), vbInformation
End Sub

Public Sub AddProvodkaIntegrationControls()
    Dim wsDashboard As Worksheet
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    
    If wsDashboard Is Nothing Then Exit Sub
    
    With wsDashboard
        .Range("A35").value = LocalizationManager.GetText("1C INTEGRATION")
        .Range("A35").Font.Bold = True
        .Range("A35").Font.Size = 12
        
        .Range("A37").value = LocalizationManager.GetText("Mass Processing")
        .Hyperlinks.Add .Range("A37"), "", "ProvodkaIntegrationModule.MassProcessWithFileSelection", , LocalizationManager.GetText("Mass Processing")
        
        .Range("C37").value = LocalizationManager.GetText("Statistics")
        .Hyperlinks.Add .Range("C37"), "", "ProvodkaIntegrationModule.ShowMatchingStatistics", , LocalizationManager.GetText("Statistics")
        
        .Range("E37").value = LocalizationManager.GetText("Clear Marks")
        .Hyperlinks.Add .Range("E37"), "", "ProvodkaIntegrationModule.ClearAllProvodkaMarks", , LocalizationManager.GetText("Clear Marks")
        
        .Range("A37:E37").Font.Bold = True
        .Range("A37:E37").Interior.Color = RGB(200, 220, 255)
        .Range("A37:E37").Borders.LineStyle = xlContinuous
    End With
End Sub

