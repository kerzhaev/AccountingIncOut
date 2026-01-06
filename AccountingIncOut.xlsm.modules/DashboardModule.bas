Attribute VB_Name = "DashboardModule"
'==============================================
' МОДУЛЬ ФУНКЦИОНАЛЬНОГО DASHBOARD - DashboardModule
' Назначение: Генерация и обновление сводного dashboard с реальными данными
' Состояние: ИСПРАВЛЕНА ПРОБЛЕМА ОТОБРАЖЕНИЯ ТАБЛИЦ - ПРАВИЛЬНОЕ ПОЗИЦИОНИРОВАНИЕ И ФОРМАТИРОВАНИЕ
' Версия: 1.9.0
' Дата: 04.08.2025
' Автор: VBA Solution
'==============================================

Option Explicit

' Глобальные переменные модуля
Private dashboardTimer As Double
Private isAutoUpdateEnabled As Boolean
Private lastUpdateTime As Date
Private cachedData As Collection
Private cacheTimestamp As Date
Private previousTotalDocs As Long

' Константы для настроек
Private Const UPDATE_INTERVAL_SECONDS As Integer = 30
Private Const CACHE_LIFETIME_MINUTES As Integer = 5
Private Const MAX_TOP_SERVICES As Integer = 5
Private Const CRITICAL_DELAY_DAYS As Integer = 7
Private Const WARNING_DELAY_DAYS As Integer = 3

' Правильные имена столбцов из диагностики
Private Const STATUS_COLUMN_NAME As String = "СтатусПодтверждения"
Private Const DATE_COLUMN_NAME As String = "ДатаВх. ФРП"
Private Const SERVICE_COLUMN_NAME As String = "Служба"

' ===============================================
' ОСНОВНАЯ ПРОЦЕДУРА ГЕНЕРАЦИИ DASHBOARD
' ===============================================

Public Sub GenerateDashboard()
    ' Объявление основных переменных
    Dim wsDashboard As Worksheet
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim StartTime As Double
    Dim processedRecords As Long
    Dim errorMessage As String
    
    ' Переменные для параметров фильтрации
    Dim dateFrom As Date
    Dim dateTo As Date
    Dim serviceFilter As String
    Dim statusFilter As String
    
    ' Переменные для данных
    Dim filteredData As Collection
    Dim kpiMetrics As Collection
    
    On Error GoTo GenerateError
    
    StartTime = Timer
    
    ' Отключаем обновление экрана для производительности
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Получаем ссылки на листы и таблицу
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    Set wsData = CommonUtilities.GetWorksheetSafe("ВхИсх")
    
    If wsDashboard Is Nothing Or wsData Is Nothing Then
        MsgBox "Ошибка: Не найдены необходимые листы для генерации Dashboard", vbCritical, "Ошибка"
        GoTo CleanupAndExit
    End If
    
    Set tblData = CommonUtilities.GetListObjectSafe(wsData, "ВходящиеИсходящие")
    
    If tblData Is Nothing Then
        MsgBox "Ошибка: Не найдена таблица 'ВходящиеИсходящие' для генерации Dashboard", vbCritical, "Ошибка"
        GoTo CleanupAndExit
    End If
    
    ' Сохраняем предыдущее значение для расчета динамики
    On Error Resume Next
    previousTotalDocs = CLng(wsDashboard.Range("A6").value)
    If Err.Number <> 0 Then previousTotalDocs = 0
    Err.Clear
    On Error GoTo GenerateError
    
    ' Получаем параметры фильтрации
    Call GetFilterParameters(dateFrom, dateTo, serviceFilter, statusFilter)
    
    ' Обновляем статус генерации
    wsDashboard.Range("I3").value = "Генерация..."
    wsDashboard.Range("F3").value = Now
    
    Debug.Print "=== ГЕНЕРАЦИЯ DASHBOARD ==="
    Debug.Print "Дата с: " & dateFrom & " до: " & dateTo
    Debug.Print "Служба: " & serviceFilter & ", Статус: " & statusFilter
    
    ' Обновляем вспомогательные расчёты ПЕРЕД обработкой данных
    Call UpdateHelperCalculations(wsDashboard, tblData, dateFrom, dateTo, serviceFilter, statusFilter)
    
    ' Получаем отфильтрованные данные с правильными именами столбцов
    Set filteredData = GetFilteredDataFixed(tblData, dateFrom, dateTo, serviceFilter, statusFilter)
    processedRecords = filteredData.Count
    
    Debug.Print "Отфильтровано записей: " & processedRecords
    
    ' Рассчитываем метрики KPI
    Set kpiMetrics = CalculateKPIMetricsFixed(filteredData, tblData)
    
    Debug.Print "=== МЕТРИКИ KPI ==="
    Debug.Print "Всего: " & GetMetricValue(kpiMetrics, "TotalDocs")
    Debug.Print "Подтверждено: " & GetMetricValue(kpiMetrics, "ConfirmedDocs")
    Debug.Print "В работе: " & GetMetricValue(kpiMetrics, "InProgressDocs")
    Debug.Print "Просрочено: " & GetMetricValue(kpiMetrics, "OverdueDocs")
    
    ' Обновляем все секции Dashboard
    Call UpdateKPIBlocks(wsDashboard, kpiMetrics, previousTotalDocs)
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Обновляем таблицы с правильным форматированием
    Call UpdateDataTablesWithFormatting(wsDashboard, filteredData, tblData)
    
    Call UpdateTrendsData(wsDashboard, filteredData)
    Call ApplyConditionalFormatting(wsDashboard, kpiMetrics)
    Call UpdateParametersPanel(wsDashboard, dateFrom, dateTo, processedRecords)
    
    ' Создаем диаграммы с исправленными данными
    Call CreateChartsFixed(wsDashboard, filteredData)
    
    ' Записываем время последнего обновления
    lastUpdateTime = Now
    wsDashboard.Range("Q5").value = lastUpdateTime
    wsDashboard.Range("Q5").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    
    ' Обновляем статус завершения
    wsDashboard.Range("I3").value = "Активен"
    
CleanupAndExit:
    ' Включаем обратно обновления
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Информируем о завершении
    Dim processingTime As Double
    processingTime = Timer - StartTime
    
    MsgBox "Dashboard успешно обновлен!" & vbCrLf & vbCrLf & _
           "Обработано записей: " & processedRecords & vbCrLf & _
           "Время обработки: " & Format(processingTime, "0.00") & " сек." & vbCrLf & _
           "Последнее обновление: " & Format(lastUpdateTime, "dd.mm.yyyy hh:mm:ss"), _
           vbInformation, "Dashboard обновлен"
    
    Call StartAutoUpdate
    
    Exit Sub
    
GenerateError:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    errorMessage = "Ошибка генерации Dashboard: " & Err.description & vbCrLf & _
                   "Номер ошибки: " & Err.Number & vbCrLf & _
                   "Обработано записей: " & processedRecords
    MsgBox errorMessage, vbCritical, "Критическая ошибка"
End Sub

' ===============================================
' ПРОЦЕДУРА ОБНОВЛЕНИЯ ВСПОМОГАТЕЛЬНЫХ РАСЧЁТОВ
' ===============================================

Private Sub UpdateHelperCalculations(wsDashboard As Worksheet, tblData As ListObject, dateFrom As Date, dateTo As Date, serviceFilter As String, statusFilter As String)
    ' Объявление переменных для расчётов
    Dim totalDocs As Long
    Dim confirmedDocs As Long
    Dim sentDocs As Long
    Dim inProgressDocs As Long
    Dim emptyStatusDocs As Long
    Dim overdueDocs As Long
    Dim completionRate As Double
    Dim avgProcessingDays As Double
    Dim periodDays As Long
    
    ' Переменные для обработки данных
    Dim i As Long
    Dim recordDate As Date
    Dim recordService As String
    Dim recordStatus As String
    Dim daysDiff As Long
    Dim totalProcessingDays As Long
    Dim processedRecords As Long
    
    ' Инициализация счётчиков
    totalDocs = 0
    confirmedDocs = 0
    sentDocs = 0
    inProgressDocs = 0
    emptyStatusDocs = 0
    overdueDocs = 0
    totalProcessingDays = 0
    processedRecords = 0
    
    Debug.Print "=== ОБНОВЛЕНИЕ ВСПОМОГАТЕЛЬНЫХ РАСЧЁТОВ ==="
    
    ' Проверяем наличие данных в таблице
    If tblData.DataBodyRange Is Nothing Then
        Debug.Print "ПРЕДУПРЕЖДЕНИЕ: Таблица пуста"
        GoTo WriteResults
    End If
    
    ' Находим правильные индексы столбцов
    Dim statusColumnIndex As Integer
    Dim dateColumnIndex As Integer
    Dim serviceColumnIndex As Integer
    
    statusColumnIndex = FindColumnIndex(tblData, STATUS_COLUMN_NAME, 19)
    dateColumnIndex = FindColumnIndex(tblData, DATE_COLUMN_NAME, 8)
    serviceColumnIndex = FindColumnIndex(tblData, SERVICE_COLUMN_NAME, 2)
    
    Debug.Print "Используемые столбцы: Статус=" & statusColumnIndex & ", Дата=" & dateColumnIndex & ", Служба=" & serviceColumnIndex
    
    ' Обрабатываем каждую строку таблицы
    For i = 1 To tblData.ListRows.Count
        ' Получаем дату записи
        On Error Resume Next
        recordDate = CDate(tblData.DataBodyRange.Cells(i, dateColumnIndex).value)
        If Err.Number <> 0 Then recordDate = 0
        Err.Clear
        On Error GoTo 0
        
        ' Фильтр по дате
        If recordDate > 0 And recordDate >= dateFrom And recordDate <= dateTo Then
            ' Получаем службу
            recordService = CStr(tblData.DataBodyRange.Cells(i, serviceColumnIndex).value)
            
            ' Фильтр по службе
            If serviceFilter = "Все службы" Or serviceFilter = "" Or recordService = serviceFilter Then
                ' Получаем статус
                recordStatus = CStr(tblData.DataBodyRange.Cells(i, statusColumnIndex).value)
                
                ' Фильтр по статусу
                If statusFilter = "Все статусы" Or statusFilter = "" Or recordStatus = statusFilter Or _
                   (statusFilter = "(пустой)" And recordStatus = "") Then
                    
                    ' Увеличиваем общий счётчик
                    totalDocs = totalDocs + 1
                    
                    ' Классифицируем по статусам
                    If InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Then
                        confirmedDocs = confirmedDocs + 1
                    ElseIf InStr(UCase(recordStatus), "ОТПР") > 0 And InStr(UCase(recordStatus), "ИСХ") > 0 Then
                        sentDocs = sentDocs + 1
                    ElseIf recordStatus = "" Or InStr(UCase(recordStatus), "НЕТ") > 0 Then
                        emptyStatusDocs = emptyStatusDocs + 1
                    End If
                    
                    ' Проверяем на просрочку
                    daysDiff = Date - recordDate
                    If daysDiff > CRITICAL_DELAY_DAYS And Not (InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Or InStr(UCase(recordStatus), "ОТПР") > 0) Then
                        overdueDocs = overdueDocs + 1
                    End If
                    
                    ' Для расчёта среднего времени обработки
                    If InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Or InStr(UCase(recordStatus), "ОТПР") > 0 Then
                        totalProcessingDays = totalProcessingDays + daysDiff
                        processedRecords = processedRecords + 1
                    End If
                End If
            End If
        End If
    Next i
    
WriteResults:
    ' Рассчитываем производные показатели
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
    
    ' Записываем вспомогательные расчёты в правильные ячейки
    On Error Resume Next
    
    ' Основные показатели
    wsDashboard.Range("T3").value = totalDocs              ' Всего_Документов
    wsDashboard.Range("T4").value = confirmedDocs          ' Подтвержденных
    wsDashboard.Range("T5").value = sentDocs               ' Отправленных
    wsDashboard.Range("T6").value = inProgressDocs         ' В_Работе
    wsDashboard.Range("T7").value = emptyStatusDocs        ' Без_Статуса
    wsDashboard.Range("T8").value = overdueDocs            ' Просроченных
    
    ' Производные показатели
    wsDashboard.Range("T9").value = completionRate         ' Процент_Завершения
    wsDashboard.Range("T10").value = avgProcessingDays     ' Средний_Срок
    wsDashboard.Range("T11").value = periodDays            ' Период_Дней
    wsDashboard.Range("T12").value = Date                  ' Последняя_Дата
    
    ' Форматирование
    wsDashboard.Range("T9").NumberFormat = "0.0%"          ' Процент
    wsDashboard.Range("T10").NumberFormat = "0.0"          ' Дни с десятичными
    wsDashboard.Range("T12").NumberFormat = "dd.mm.yyyy"   ' Дата
    
    On Error GoTo 0
    
    Debug.Print "Вспомогательные расчёты обновлены:"
    Debug.Print "- Всего документов: " & totalDocs
    Debug.Print "- Подтверждено: " & confirmedDocs
    Debug.Print "- В работе: " & inProgressDocs
    Debug.Print "- Просрочено: " & overdueDocs
End Sub

' ===============================================
' ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: ПОИСК ИНДЕКСА СТОЛБЦА
' ===============================================

Private Function FindColumnIndex(tblData As ListObject, ColumnName As String, defaultIndex As Integer) As Integer
    Dim i As Integer
    Dim foundIndex As Integer
    
    foundIndex = defaultIndex ' Значение по умолчанию
    
    ' Ищем столбец по имени
    For i = 1 To tblData.ListColumns.Count
        If InStr(UCase(tblData.ListColumns(i).Name), UCase(ColumnName)) > 0 Then
            foundIndex = i
            Debug.Print "Найден столбец '" & ColumnName & "' с индексом " & i
            Exit For
        End If
    Next i
    
    ' Если не найден, используем альтернативные варианты
    If foundIndex = defaultIndex Then
        Select Case ColumnName
            Case STATUS_COLUMN_NAME
                ' Ищем альтернативные названия для статуса
                For i = 1 To tblData.ListColumns.Count
                    If InStr(UCase(tblData.ListColumns(i).Name), "СТАТУС") > 0 Then
                        foundIndex = i
                        Debug.Print "Найден альтернативный столбец статуса: " & tblData.ListColumns(i).Name & " (индекс " & i & ")"
                        Exit For
                    End If
                Next i
            Case DATE_COLUMN_NAME
                ' Ищем альтернативные названия для даты
                For i = 1 To tblData.ListColumns.Count
                    If InStr(UCase(tblData.ListColumns(i).Name), "ДАТА") > 0 And InStr(UCase(tblData.ListColumns(i).Name), "ФРП") > 0 Then
                        foundIndex = i
                        Debug.Print "Найден альтернативный столбец даты: " & tblData.ListColumns(i).Name & " (индекс " & i & ")"
                        Exit For
                    End If
                Next i
            Case SERVICE_COLUMN_NAME
                ' Ищем альтернативные названия для службы
                For i = 1 To tblData.ListColumns.Count
                    If InStr(UCase(tblData.ListColumns(i).Name), "СЛУЖБ") > 0 Then
                        foundIndex = i
                        Debug.Print "Найден альтернативный столбец службы: " & tblData.ListColumns(i).Name & " (индекс " & i & ")"
                        Exit For
                    End If
                Next i
        End Select
    End If
    
    FindColumnIndex = foundIndex
End Function

' ===============================================
' ПРОЦЕДУРА ПОЛУЧЕНИЯ ПАРАМЕТРОВ ФИЛЬТРАЦИИ
' ===============================================

Private Sub GetFilterParameters(ByRef dateFrom As Date, ByRef dateTo As Date, ByRef serviceFilter As String, ByRef statusFilter As String)
    Dim wsReports As Worksheet
    
    On Error GoTo DefaultParameters
    
    Set wsReports = CommonUtilities.GetWorksheetSafe("Отчеты")
    
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
    serviceFilter = "Все службы"
    statusFilter = "Все статусы"
    
    Debug.Print "ИСПОЛЬЗУЮТСЯ ПАРАМЕТРЫ ПО УМОЛЧАНИЮ"
    
    On Error GoTo 0
End Sub

' ===============================================
' ФУНКЦИЯ ПОЛУЧЕНИЯ ОТФИЛЬТРОВАННЫХ ДАННЫХ
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
    
    ' Определяем правильные индексы столбцов
    Dim statusColumnIndex As Integer
    Dim dateColumnIndex As Integer
    Dim serviceColumnIndex As Integer
    
    Set filteredData = New Collection
    
    If tblData.DataBodyRange Is Nothing Then
        Debug.Print "ОШИБКА: Таблица не содержит данных"
        Set GetFilteredDataFixed = filteredData
        Exit Function
    End If
    
    TotalRows = tblData.ListRows.Count
    validDates = 0
    matchedRecords = 0
    
    statusColumnIndex = FindColumnIndex(tblData, STATUS_COLUMN_NAME, 19)
    dateColumnIndex = FindColumnIndex(tblData, DATE_COLUMN_NAME, 8)
    serviceColumnIndex = FindColumnIndex(tblData, SERVICE_COLUMN_NAME, 2)
    
    Debug.Print "=== ОБРАБОТКА ДАННЫХ ==="
    Debug.Print "Всего строк: " & TotalRows & ", Индексы: С=" & statusColumnIndex & ", Д=" & dateColumnIndex & ", СЛ=" & serviceColumnIndex
    
    fieldNames = Array("НомерПП", "Служба", "ВидДокумента", "ТипДокумента", "НомерДокумента", _
                      "СуммаДокумента", "НомерФРП", "ДатаФРП", "ОтКогоПоступил", "ДатаПередачи", _
                      "Исполнитель", "НомерИсхВСлужбу", "ДатаИсхВСлужбу", "НомерВозврата", _
                      "ДатаВозврата", "НомерИсхКонверт", "ДатаИсхКонверт", "ОтметкаИсполнения", "Статус")
    
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
            
            If serviceFilter = "Все службы" Or serviceFilter = "" Or recordService = serviceFilter Then
                recordStatus = CStr(tblData.DataBodyRange.Cells(i, statusColumnIndex).value)
                
                Dim statusMatch As Boolean
                statusMatch = False
                
                If statusFilter = "Все статусы" Or statusFilter = "" Then
                    statusMatch = True
                ElseIf statusFilter = "(пустой)" And recordStatus = "" Then
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
    
    Debug.Print "Валидных дат: " & validDates & ", Прошло фильтрацию: " & matchedRecords
    
    Set GetFilteredDataFixed = filteredData
End Function

' ===============================================
' ФУНКЦИЯ РАСЧЕТА МЕТРИК KPI
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
    
    Debug.Print "=== АНАЛИЗ СТАТУСОВ (ИСПРАВЛЕННЫЙ) ==="
    
    For Each record In filteredData
        On Error Resume Next
        recordStatus = CStr(record.item("Статус"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        On Error GoTo 0
        
        If InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Then
            confirmedDocs = confirmedDocs + 1
        ElseIf InStr(UCase(recordStatus), "ОТПР") > 0 And InStr(UCase(recordStatus), "ИСХ") > 0 Then
            sentDocs = sentDocs + 1
        ElseIf recordStatus = "" Or InStr(UCase(recordStatus), "НЕТ") > 0 Then
            emptyStatusDocs = emptyStatusDocs + 1
        End If
        
        On Error Resume Next
        recordDate = CDate(record.item("ДатаФРП"))
        If Err.Number = 0 And recordDate > 0 Then
            daysDiff = Date - recordDate
            
            If daysDiff > maxDelayDays Then
                maxDelayDays = daysDiff
            End If
            
            If daysDiff > CRITICAL_DELAY_DAYS And Not (InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Or InStr(UCase(recordStatus), "ОТПР") > 0) Then
                overdueDocs = overdueDocs + 1
            End If
            
            If InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Or InStr(UCase(recordStatus), "ОТПР") > 0 Then
                totalDays = totalDays + daysDiff
                processedRecords = processedRecords + 1
            End If
            
            If Not (InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Or InStr(UCase(recordStatus), "ОТПР") > 0) Then
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
' ПРОЦЕДУРА ОБНОВЛЕНИЯ KPI БЛОКОВ
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
        dynamicText = "+" & dynamicChange & " к предыдущему"
    ElseIf dynamicChange < 0 Then
        dynamicIcon = "?"
        dynamicText = dynamicChange & " к предыдущему"
    Else
        dynamicIcon = ">"
        dynamicText = "Без изменений"
    End If
    wsDashboard.Range("A8").value = dynamicIcon & " " & dynamicText
    
    If avgInProgressDays > 0 Then
        wsDashboard.Range("I8").value = "Ср. время: " & Format(avgInProgressDays, "0.0") & " дн."
    Else
        wsDashboard.Range("I8").value = "Ср. время: н/д"
    End If
    
    If maxDelayDays > 0 Then
        wsDashboard.Range("M8").value = "Макс. задержка: " & maxDelayDays & " дн."
    Else
        wsDashboard.Range("M8").value = "Задержек нет"
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
' ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: БЕЗОПАСНОЕ ПОЛУЧЕНИЕ ЗНАЧЕНИЙ ИЗ КОЛЛЕКЦИИ
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
' НОВАЯ ПРОЦЕДУРА: ОБНОВЛЕНИЕ ТАБЛИЦ С ФОРМАТИРОВАНИЕМ
' ===============================================

Private Sub UpdateDataTablesWithFormatting(wsDashboard As Worksheet, filteredData As Collection, tblData As ListObject)
    Debug.Print "=== ОБНОВЛЕНИЕ ТАБЛИЦ С ФОРМАТИРОВАНИЕМ ==="
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Сначала очищаем области таблиц
    Call ClearTableAreas(wsDashboard)
    
    ' Создаем заголовки таблиц
    Call CreateTableHeaders(wsDashboard)
    
    ' Обновляем таблицы с данными
    Call UpdateTopServicesTableFormatted(wsDashboard, filteredData)
    Call UpdateProblemDocsTableFormatted(wsDashboard, filteredData)
    Call UpdateChartDataTablesFixed(wsDashboard, filteredData)
    
    ' Применяем форматирование к таблицам
    Call ApplyTableFormatting(wsDashboard)
End Sub

' ===============================================
' ПРОЦЕДУРА ОЧИСТКИ ОБЛАСТЕЙ ТАБЛИЦ
' ===============================================

Private Sub ClearTableAreas(wsDashboard As Worksheet)
    On Error Resume Next
    
    ' Очищаем область таблицы топ-служб (A27:F35)
    wsDashboard.Range("A27:F35").ClearContents
    wsDashboard.Range("A27:F35").Interior.ColorIndex = xlNone
    wsDashboard.Range("A27:F35").Borders.LineStyle = xlNone
    
    ' Очищаем область таблицы проблемных документов (H27:M35)
    wsDashboard.Range("H27:M35").ClearContents
    wsDashboard.Range("H27:M35").Interior.ColorIndex = xlNone
    wsDashboard.Range("H27:M35").Borders.LineStyle = xlNone
    
    ' Очищаем скрытые данные для диаграмм
    wsDashboard.Range("V4:W8").ClearContents
    wsDashboard.Range("V11:W17").ClearContents
    
    On Error GoTo 0
    
    Debug.Print "Области таблиц очищены"
End Sub

' ===============================================
' ПРОЦЕДУРА СОЗДАНИЯ ЗАГОЛОВКОВ ТАБЛИЦ
' ===============================================

Private Sub CreateTableHeaders(wsDashboard As Worksheet)
    On Error Resume Next
    
    ' Заголовки таблицы топ-служб
    wsDashboard.Range("A27").value = "№"
    wsDashboard.Range("B27").value = "СЛУЖБЫ С НАИБОЛЬШЕЙ АКТИВНОСТЬЮ"
    wsDashboard.Range("C27").value = "Всего"
    wsDashboard.Range("D27").value = "Завершено"
    wsDashboard.Range("E27").value = "%"
    wsDashboard.Range("F27").value = "Тренд"
    
    ' Заголовки таблицы проблемных документов
    wsDashboard.Range("H27").value = "ДОКУМЕНТЫ, ТРЕБУЮЩИЕ ВНИМАНИЯ"
    wsDashboard.Range("I27").value = "Служба"
    wsDashboard.Range("J27").value = "№ Документа"
    wsDashboard.Range("K27").value = "Дата"
    wsDashboard.Range("L27").value = "Дней"
    wsDashboard.Range("M27").value = "Статус"
    
    ' Форматирование заголовков
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
    
    Debug.Print "Заголовки таблиц созданы"
End Sub

' ===============================================
' ИСПРАВЛЕННАЯ ПРОЦЕДУРА ОБНОВЛЕНИЯ ТАБЛИЦЫ ТОП-СЛУЖБ С ФОРМАТИРОВАНИЕМ
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
    
    Debug.Print "=== ОБНОВЛЕНИЕ ТАБЛИЦЫ ТОП-СЛУЖБ С ФОРМАТИРОВАНИЕМ ==="
    
    ' Собираем статистику по службам
    For Each record In filteredData
        On Error Resume Next
        serviceName = CStr(record.item("Служба"))
        If Err.Number <> 0 Then serviceName = ""
        Err.Clear
        On Error GoTo 0
        
        If serviceName <> "" Then
            On Error Resume Next
            recordStatus = CStr(record.item("Статус"))
            If Err.Number <> 0 Then recordStatus = ""
            Err.Clear
            On Error GoTo 0
            
            If Not serviceStats.Exists(serviceName) Then
                If InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Or InStr(UCase(recordStatus), "ОТПР") > 0 Then
                    serviceStats.Add serviceName, Array(1, 1)
                Else
                    serviceStats.Add serviceName, Array(1, 0)
                End If
            Else
                Dim currentData As Variant
                currentData = serviceStats(serviceName)
                totalCount = currentData(0) + 1
                If InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Or InStr(UCase(recordStatus), "ОТПР") > 0 Then
                    completedCount = currentData(1) + 1
                Else
                    completedCount = currentData(1)
                End If
                serviceStats(serviceName) = Array(totalCount, completedCount)
            End If
        End If
    Next record
    
    Debug.Print "Найдено уникальных служб: " & serviceStats.Count
    
    ' Создаем отсортированный список служб
    Set sortedServices = New Collection
    For Each serviceKey In serviceStats.Keys
        sortedServices.Add serviceKey
    Next serviceKey
    
    topCount = IIf(sortedServices.Count < MAX_TOP_SERVICES, sortedServices.Count, MAX_TOP_SERVICES)
    
    ' Заполняем таблицу с форматированием
    For i = 1 To MAX_TOP_SERVICES
        Dim RowIndex As Integer
        RowIndex = 28 + i - 1  ' Строки 28-32
        
        If i <= topCount Then
            serviceName = sortedServices.item(i)
            Dim serviceData As Variant
            serviceData = serviceStats(serviceName)
            
            ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Заполняем ячейки с проверкой
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
            
            ' Форматирование строки
            With wsDashboard.Range("A" & RowIndex & ":F" & RowIndex)
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.Color = RGB(245, 245, 245)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.Color = RGB(200, 200, 200)
            End With
            
            Debug.Print "Топ-" & i & ": " & serviceName & " (" & serviceData(0) & "/" & serviceData(1) & ")"
        Else
            ' Очищаем пустые строки
            wsDashboard.Range("A" & RowIndex & ":F" & RowIndex).ClearContents
        End If
    Next i
    
    Debug.Print "Таблица топ-служб обновлена с форматированием"
End Sub

' ===============================================
' ИСПРАВЛЕННАЯ ПРОЦЕДУРА ОБНОВЛЕНИЯ ТАБЛИЦЫ ПРОБЛЕМНЫХ ДОКУМЕНТОВ С ФОРМАТИРОВАНИЕМ
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
    
    Debug.Print "=== ОБНОВЛЕНИЕ ТАБЛИЦЫ ПРОБЛЕМНЫХ ДОКУМЕНТОВ С ФОРМАТИРОВАНИЕМ ==="
    
    ' Находим проблемные документы
    For Each record In filteredData
        On Error Resume Next
        recordDate = CDate(record.item("ДатаФРП"))
        If Err.Number <> 0 Then recordDate = 0
        Err.Clear
        
        recordStatus = CStr(record.item("Статус"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        
        serviceName = CStr(record.item("Служба"))
        If Err.Number <> 0 Then serviceName = ""
        Err.Clear
        
        docNumber = CStr(record.item("НомерДокумента"))
        If Err.Number <> 0 Then docNumber = ""
        Err.Clear
        On Error GoTo 0
        
        If recordDate > 0 Then
            daysDiff = Date - recordDate
            
            Dim isProblematic As Boolean
            isProblematic = False
            
            If daysDiff > WARNING_DELAY_DAYS And recordStatus = "" Then
                isProblematic = True
            ElseIf daysDiff > CRITICAL_DELAY_DAYS And Not (InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Or InStr(UCase(recordStatus), "ОТПР") > 0) Then
                isProblematic = True
            End If
            
            If isProblematic Then
                Set problemRecord = New Collection
                problemRecord.Add IIf(daysDiff > CRITICAL_DELAY_DAYS, "Критично", "Внимание"), "Type"
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
    
    Debug.Print "Найдено проблемных документов: " & problemDocs.Count
    
    ' Заполняем таблицу с форматированием
    For i = 1 To maxProblems
        Dim RowIndex As Integer
        RowIndex = 28 + i - 1  ' Строки 28-32
        
        If i <= problemDocs.Count Then
            Set problemRecord = problemDocs.item(i)
            
            ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Заполняем ячейки с проверкой
            On Error Resume Next
            wsDashboard.Cells(RowIndex, 8).value = problemRecord.item("Type")
            wsDashboard.Cells(RowIndex, 9).value = problemRecord.item("Service")
            wsDashboard.Cells(RowIndex, 10).value = problemRecord.item("DocNumber")
            wsDashboard.Cells(RowIndex, 11).value = problemRecord.item("Date")
            wsDashboard.Cells(RowIndex, 11).NumberFormat = "dd.mm.yyyy"
            wsDashboard.Cells(RowIndex, 12).value = problemRecord.item("Days")
            wsDashboard.Cells(RowIndex, 13).value = IIf(problemRecord.item("Status") = "", "Не задан", problemRecord.item("Status"))
            On Error GoTo 0
            
            ' Условное форматирование по критичности
            If problemRecord.item("Type") = "Критично" Then
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
            
            ' Общее форматирование строки
            With wsDashboard.Range("H" & RowIndex & ":M" & RowIndex)
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.Color = RGB(200, 200, 200)
            End With
            
            Debug.Print "Проблемный документ " & i & ": " & problemRecord.item("DocNumber") & " (" & problemRecord.item("Days") & " дней)"
        Else
            ' Очищаем пустые строки
            wsDashboard.Range("H" & RowIndex & ":M" & RowIndex).ClearContents
        End If
    Next i
    
    Debug.Print "Таблица проблемных документов обновлена с форматированием"
End Sub

' ===============================================
' ПРОЦЕДУРА ПРИМЕНЕНИЯ ОБЩЕГО ФОРМАТИРОВАНИЯ ТАБЛИЦ
' ===============================================

Private Sub ApplyTableFormatting(wsDashboard As Worksheet)
    On Error Resume Next
    
    ' Общее форматирование таблицы топ-служб
    With wsDashboard.Range("A27:F32")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(150, 150, 150)
    End With
    
    ' Автоширина столбцов таблицы топ-служб
    wsDashboard.Columns("A:A").ColumnWidth = 3
    wsDashboard.Columns("B:B").ColumnWidth = 20
    wsDashboard.Columns("C:C").ColumnWidth = 8
    wsDashboard.Columns("D:D").ColumnWidth = 10
    wsDashboard.Columns("E:E").ColumnWidth = 8
    wsDashboard.Columns("F:F").ColumnWidth = 6
    
    ' Общее форматирование таблицы проблемных документов
    With wsDashboard.Range("H27:M32")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(150, 150, 150)
    End With
    
    ' Автоширина столбцов таблицы проблемных документов
    wsDashboard.Columns("H:H").ColumnWidth = 12
    wsDashboard.Columns("I:I").ColumnWidth = 15
    wsDashboard.Columns("J:J").ColumnWidth = 12
    wsDashboard.Columns("K:K").ColumnWidth = 10
    wsDashboard.Columns("L:L").ColumnWidth = 6
    wsDashboard.Columns("M:M").ColumnWidth = 12
    
    ' Вертикальное выравнивание для всех таблиц
    wsDashboard.Range("A27:F32").VerticalAlignment = xlCenter
    wsDashboard.Range("H27:M32").VerticalAlignment = xlCenter
    
    On Error GoTo 0
    
    Debug.Print "Форматирование таблиц применено"
End Sub

' ===============================================
' ПРОЦЕДУРА ОБНОВЛЕНИЯ ДАННЫХ ДЛЯ ДИАГРАММ
' ===============================================

Private Sub UpdateChartDataTablesFixed(wsDashboard As Worksheet, filteredData As Collection)
    Dim confirmedCount As Long, sentCount As Long, inProgressCount As Long, emptyCount As Long
    Dim record As Collection
    Dim recordStatus As String
    
    confirmedCount = 0
    sentCount = 0
    inProgressCount = 0
    emptyCount = 0
    
    Debug.Print "=== ПОДГОТОВКА ДАННЫХ ДЛЯ ДИАГРАММ (ИСПРАВЛЕННАЯ) ==="
    
    For Each record In filteredData
        On Error Resume Next
        recordStatus = CStr(record.item("Статус"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        On Error GoTo 0
        
        If InStr(UCase(recordStatus), "ПОДТВЕРЖДЕНО") > 0 Then
            confirmedCount = confirmedCount + 1
        ElseIf InStr(UCase(recordStatus), "ОТПР") > 0 And InStr(UCase(recordStatus), "ИСХ") > 0 Then
            sentCount = sentCount + 1
        ElseIf recordStatus = "" Or InStr(UCase(recordStatus), "НЕТ") > 0 Then
            emptyCount = emptyCount + 1
        Else
            inProgressCount = inProgressCount + 1
        End If
    Next record
    
    Debug.Print "Данные для диаграммы статусов (исправленные):"
    Debug.Print "- Подтверждено: " & confirmedCount
    Debug.Print "- Отпр.Исх.: " & sentCount
    Debug.Print "- В работе: " & inProgressCount
    Debug.Print "- Без статуса: " & emptyCount
    
    On Error Resume Next
    wsDashboard.Range("V4:W8").ClearContents
    wsDashboard.Range("V11:W17").ClearContents
    
    wsDashboard.Range("V4").value = "Статус"
    wsDashboard.Range("W4").value = "Количество"
    
    wsDashboard.Range("V5").value = "Подтверждено"
    wsDashboard.Range("W5").value = confirmedCount
    
    wsDashboard.Range("V6").value = "Отпр.Исх."
    wsDashboard.Range("W6").value = sentCount
    
    wsDashboard.Range("V7").value = "В работе"
    wsDashboard.Range("W7").value = inProgressCount
    
    wsDashboard.Range("V8").value = "Без статуса"
    wsDashboard.Range("W8").value = emptyCount
    
    wsDashboard.Range("V11").value = "Служба"
    wsDashboard.Range("W11").value = "Количество"
    
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
    
    Debug.Print "Данные для диаграмм подготовлены. Есть данные служб: " & hasServiceData
    
    On Error GoTo 0
End Sub

' ===============================================
' ПРОЦЕДУРА СОЗДАНИЯ ДИАГРАММ
' ===============================================

Private Sub CreateChartsFixed(wsDashboard As Worksheet, filteredData As Collection)
    Debug.Print "=== СОЗДАНИЕ ДИАГРАММ (ИСПРАВЛЕННАЯ ВЕРСИЯ) ==="
    Debug.Print "Количество записей для диаграмм: " & filteredData.Count
    
    Debug.Print "Проверка данных в столбце V:"
    Debug.Print "V4=" & wsDashboard.Range("V4").value & ", W4=" & wsDashboard.Range("W4").value
    Debug.Print "V5=" & wsDashboard.Range("V5").value & ", W5=" & wsDashboard.Range("W5").value
    Debug.Print "V6=" & wsDashboard.Range("V6").value & ", W6=" & wsDashboard.Range("W6").value
    Debug.Print "V7=" & wsDashboard.Range("V7").value & ", W7=" & wsDashboard.Range("W7").value
    Debug.Print "V8=" & wsDashboard.Range("V8").value & ", W8=" & wsDashboard.Range("W8").value
    
    Call CreatePieChartFixed(wsDashboard)
    Call CreateColumnChartFixed(wsDashboard)
End Sub

' ===============================================
' СОЗДАНИЕ КРУГОВОЙ ДИАГРАММЫ
' ===============================================

Private Sub CreatePieChartFixed(wsDashboard As Worksheet)
    Dim chartRange As Range
    Dim chartData As Range
    Dim chartObj As ChartObject
    Dim totalValues As Long
    
    Debug.Print "Создание круговой диаграммы (исправленная версия)..."
    
    On Error GoTo PieChartError
    
    On Error Resume Next
    wsDashboard.ChartObjects("StatusPieChart").Delete
    On Error GoTo PieChartError
    
    totalValues = CLng(wsDashboard.Range("W5").value) + CLng(wsDashboard.Range("W6").value) + _
                  CLng(wsDashboard.Range("W7").value) + CLng(wsDashboard.Range("W8").value)
    
    Debug.Print "Общее количество для круговой диаграммы: " & totalValues
    
    If totalValues = 0 Then
        wsDashboard.Range("C15").value = "Нет данных для круговой диаграммы"
        wsDashboard.Range("C15").Font.Color = RGB(150, 150, 150)
        wsDashboard.Range("C15").Font.Italic = True
        Debug.Print "Отменено создание круговой диаграммы - нет данных"
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
        .ChartTitle.Text = "Распределение по статусам"
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
    
    Debug.Print "Круговая диаграмма создана успешно"
    Exit Sub
    
PieChartError:
    Debug.Print "Ошибка создания круговой диаграммы: " & Err.description
    wsDashboard.Range("C15").value = "Ошибка диаграммы: " & Err.description
    wsDashboard.Range("C15").Font.Color = RGB(255, 0, 0)
End Sub

' ===============================================
' СОЗДАНИЕ СТОЛБЧАТОЙ ДИАГРАММЫ
' ===============================================

Private Sub CreateColumnChartFixed(wsDashboard As Worksheet)
    Dim chartRange As Range
    Dim chartData As Range
    Dim chartObj As ChartObject
    Dim hasValidData As Boolean
    Dim i As Integer
    
    Debug.Print "Создание столбчатой диаграммы (исправленная версия)..."
    
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
    
    Debug.Print "Есть данные служб для столбчатой диаграммы: " & hasValidData
    
    If Not hasValidData Then
        wsDashboard.Range("L15").value = "Нет данных по службам"
        wsDashboard.Range("L15").Font.Color = RGB(150, 150, 150)
        wsDashboard.Range("L15").Font.Italic = True
        Debug.Print "Отменено создание столбчатой диаграммы - нет данных"
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
        .ChartTitle.Text = "Активность по службам"
        .HasLegend = False
        
        On Error Resume Next
        If .Axes.Count >= 2 Then
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Службы"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Количество документов"
        End If
        
        If .SeriesCollection.Count > 0 Then
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(100, 150, 255)
            .SeriesCollection(1).HasDataLabels = True
        End If
        On Error GoTo ColumnChartError
    End With
    
    Debug.Print "Столбчатая диаграмма создана успешно"
    Exit Sub
    
ColumnChartError:
    Debug.Print "Ошибка создания столбчатой диаграммы: " & Err.description
    wsDashboard.Range("L15").value = "Ошибка диаграммы: " & Err.description
    wsDashboard.Range("L15").Font.Color = RGB(255, 0, 0)
End Sub

' ===============================================
' ОСТАЛЬНЫЕ ПРОЦЕДУРЫ ОСТАЮТСЯ БЕЗ ИЗМЕНЕНИЙ
' ===============================================

Private Sub UpdateTrendsData(wsDashboard As Worksheet, filteredData As Collection)
    ' Пока оставляем пустой
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
        wsDashboard.Range("Q6").value = "ОК (" & processedRecords & " записей)"
    Else
        wsDashboard.Range("Q6").value = "Нет данных"
    End If
    
    wsDashboard.Range("Q8").value = Format(dateFrom, "dd.mm.yyyy") & " - " & Format(dateTo, "dd.mm.yyyy")
End Sub

Public Sub StartAutoUpdate()
    Dim wsDashboard As Worksheet
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    
    If Not wsDashboard Is Nothing Then
        If UCase(CStr(wsDashboard.Range("Q3").value)) = "ВКЛ" Then
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

' Функции GetWorksheetSafe и GetListObjectSafe перенесены в CommonUtilities.bas

Public Sub RefreshDashboard()
    Call GenerateDashboard
End Sub

Public Sub ToggleAutoUpdate()
    Dim wsDashboard As Worksheet
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    
    If Not wsDashboard Is Nothing Then
        If UCase(CStr(wsDashboard.Range("Q3").value)) = "ВКЛ" Then
            wsDashboard.Range("Q3").value = "Выкл"
            Call StopAutoUpdate
            MsgBox "Автообновление отключено", vbInformation
        Else
            wsDashboard.Range("Q3").value = "Вкл"
            Call StartAutoUpdate
            MsgBox "Автообновление включено", vbInformation
        End If
    End If
End Sub

Public Sub ExportDashboard()
    MsgBox "Функция экспорта будет реализована в следующих версиях", vbInformation
End Sub

Public Sub DiagnoseDashboardData()
    Dim wsDashboard As Worksheet
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    Set wsData = CommonUtilities.GetWorksheetSafe("ВхИсх")
    Set tblData = CommonUtilities.GetListObjectSafe(wsData, "ВходящиеИсходящие")
    
    Debug.Print "=== ДИАГНОСТИКА DASHBOARD ==="
    Debug.Print "Лист Dashboard найден: " & Not (wsDashboard Is Nothing)
    Debug.Print "Лист ВхИсх найден: " & Not (wsData Is Nothing)
    Debug.Print "Таблица найдена: " & Not (tblData Is Nothing)
    
    If Not tblData Is Nothing Then
        Debug.Print "Строк в таблице: " & IIf(tblData.DataBodyRange Is Nothing, 0, tblData.ListRows.Count)
        Debug.Print "Столбцов в таблице: " & tblData.ListColumns.Count
        
        If tblData.ListColumns.Count >= 19 Then
            Debug.Print "Столбец статуса (19): " & tblData.ListColumns(19).Name
        End If
        If tblData.ListColumns.Count >= 8 Then
            Debug.Print "Столбец даты (8): " & tblData.ListColumns(8).Name
        End If
        If tblData.ListColumns.Count >= 2 Then
            Debug.Print "Столбец службы (2): " & tblData.ListColumns(2).Name
        End If
    End If
    
    If Not wsDashboard Is Nothing Then
        Debug.Print "Данные в таблицах:"
        Debug.Print "A28=" & wsDashboard.Range("A28").value & " (номер строки топ-служб)"
        Debug.Print "B28=" & wsDashboard.Range("B28").value & " (название службы)"
        Debug.Print "H28=" & wsDashboard.Range("H28").value & " (тип проблемного документа)"
        Debug.Print "I28=" & wsDashboard.Range("I28").value & " (служба проблемного документа)"
    End If
    
    MsgBox "Диагностика завершена. Смотрите результаты в окне Immediate (Ctrl+G)", vbInformation
End Sub


' В модуле DashboardModule добавить процедуры:

Public Sub AddProvodkaIntegrationControls()
    Dim wsDashboard As Worksheet
    Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
    
    If wsDashboard Is Nothing Then Exit Sub
    
    ' Добавляем кнопки интеграции с 1С
    With wsDashboard
        .Range("A35").value = "ИНТЕГРАЦИЯ С 1С"
        .Range("A35").Font.Bold = True
        .Range("A35").Font.Size = 12
        
        ' Кнопка массовой обработки
        .Range("A37").value = "Массовая обработка"
        .Hyperlinks.Add .Range("A37"), "", "ProvodkaIntegrationModule.MassProcessWithFileSelection", , "Автоматическое сопоставление всех записей с выгрузкой 1С"
        
        ' Кнопка статистики
        .Range("C37").value = "Статистика"
        .Hyperlinks.Add .Range("C37"), "", "ProvodkaIntegrationModule.ShowMatchingStatistics", , "Показать статистику сопоставления"
        
        ' Кнопка очистки
        .Range("E37").value = "Очистить отметки"
        .Hyperlinks.Add .Range("E37"), "", "ProvodkaIntegrationModule.ClearAllProvodkaMarks", , "Очистить все отметки об исполнении"
        
        ' Форматирование кнопок
        .Range("A37:E37").Font.Bold = True
        .Range("A37:E37").Interior.Color = RGB(200, 220, 255)
        .Range("A37:E37").Borders.LineStyle = xlContinuous
    End With
End Sub




