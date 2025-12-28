Attribute VB_Name = "Module1"
'==============================================
' МОДУЛЬ ФУНКЦИОНАЛЬНОГО DASHBOARD - DashboardModule
' Назначение: Генерация и обновление сводного dashboard с реальными данными
' Состояние: ИСПРАВЛЕНА КРИТИЧЕСКАЯ ОШИБКА "Недопустимый параметр" - ПОЛНАЯ ВАЛИДАЦИЯ ДАННЫХ ДЛЯ ДИАГРАММ
' Версия: 1.4.0
' Дата: 03.08.2025
' Автор: VBA Solution
'==============================================

Option Explicit

' Глобальные переменные модуля
Private dashboardTimer As Double
Private isAutoUpdateEnabled As Boolean
Private lastUpdateTime As Date
Private cachedData As Collection
Private cacheTimestamp As Date

' Константы для настроек
Private Const UPDATE_INTERVAL_SECONDS As Integer = 30
Private Const CACHE_LIFETIME_MINUTES As Integer = 5
Private Const MAX_TOP_SERVICES As Integer = 5
Private Const CRITICAL_DELAY_DAYS As Integer = 7
Private Const WARNING_DELAY_DAYS As Integer = 3

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
    Set wsDashboard = GetWorksheetSafe("Dashboard")
    Set wsData = GetWorksheetSafe("ВхИсх")
    
    If wsDashboard Is Nothing Or wsData Is Nothing Then
        MsgBox "Ошибка: Не найдены необходимые листы для генерации Dashboard", vbCritical, "Ошибка"
        GoTo CleanupAndExit
    End If
    
    Set tblData = GetListObjectSafe(wsData, "ВходящиеИсходящие")
    
    If tblData Is Nothing Then
        MsgBox "Ошибка: Не найдена таблица 'ВходящиеИсходящие' для генерации Dashboard", vbCritical, "Ошибка"
        GoTo CleanupAndExit
    End If
    
    ' Получаем параметры фильтрации
    Call GetFilterParameters(dateFrom, dateTo, serviceFilter, statusFilter)
    
    ' Обновляем статус генерации
    wsDashboard.Range("I3").value = "Генерация..."
    wsDashboard.Range("F3").value = Now
    
    ' Получаем отфильтрованные данные
    Set filteredData = GetFilteredData(tblData, dateFrom, dateTo, serviceFilter, statusFilter)
    processedRecords = filteredData.Count
    
    ' Рассчитываем метрики KPI
    Set kpiMetrics = CalculateKPIMetrics(filteredData, tblData)
    
    ' Обновляем все секции Dashboard
    Call UpdateKPIBlocks(wsDashboard, kpiMetrics)
    Call UpdateDataTables(wsDashboard, filteredData, tblData)
    Call UpdateTrendsData(wsDashboard, filteredData)
    Call ApplyConditionalFormatting(wsDashboard, kpiMetrics)
    Call UpdateParametersPanel(wsDashboard, dateFrom, dateTo, processedRecords)
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Создаем диаграммы с полной валидацией данных
    Call CreateChartsWithValidation(wsDashboard, filteredData)
    
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
    
    ' Запускаем автообновление если включено
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
' ПРОЦЕДУРА ПОЛУЧЕНИЯ ПАРАМЕТРОВ ФИЛЬТРАЦИИ
' ===============================================

Private Sub GetFilterParameters(ByRef dateFrom As Date, ByRef dateTo As Date, ByRef serviceFilter As String, ByRef statusFilter As String)
    ' Объявление локальных переменных
    Dim wsReports As Worksheet
    
    On Error GoTo DefaultParameters
    
    ' Пытаемся получить параметры с листа "Отчеты"
    Set wsReports = GetWorksheetSafe("Отчеты")
    
    If Not wsReports Is Nothing Then
        ' Получаем даты - безопасно
        On Error Resume Next
        dateFrom = CDate(wsReports.Range("B18").value)
        If Err.Number <> 0 Then dateFrom = Date - 30
        Err.Clear
        
        dateTo = CDate(wsReports.Range("B19").value)
        If Err.Number <> 0 Then dateTo = Date
        Err.Clear
        
        ' Получаем фильтры - безопасно
        serviceFilter = CStr(wsReports.Range("B20").value)
        statusFilter = CStr(wsReports.Range("B21").value)
        On Error GoTo DefaultParameters
    Else
        GoTo DefaultParameters
    End If
    
    Exit Sub
    
DefaultParameters:
    ' Устанавливаем параметры по умолчанию
    dateFrom = Date - 30
    dateTo = Date
    serviceFilter = "Все службы"
    statusFilter = "Все статусы"
    
    On Error GoTo 0
End Sub

' ===============================================
' ФУНКЦИЯ ПОЛУЧЕНИЯ ОТФИЛЬТРОВАННЫХ ДАННЫХ
' ===============================================

Private Function GetFilteredData(tblData As ListObject, dateFrom As Date, dateTo As Date, serviceFilter As String, statusFilter As String) As Collection
    ' Объявление переменных
    Dim filteredData As Collection
    Dim i As Long
    Dim recordDate As Date
    Dim recordService As String
    Dim recordStatus As String
    Dim record As Collection
    Dim fieldNames As Variant
    Dim j As Integer
    
    Set filteredData = New Collection
    
    ' Проверяем наличие данных
    If tblData.DataBodyRange Is Nothing Then
        Set GetFilteredData = filteredData
        Exit Function
    End If
    
    ' Инициализируем названия полей
    fieldNames = Array("НомерПП", "Служба", "ВидДокумента", "ТипДокумента", "НомерДокумента", _
                      "СуммаДокумента", "НомерФРП", "ДатаФРП", "ОтКогоПоступил", "ДатаПередачи", _
                      "Исполнитель", "НомерИсхВСлужбу", "ДатаИсхВСлужбу", "НомерВозврата", _
                      "ДатаВозврата", "НомерИсхКонверт", "ДатаИсхКонверт", "ОтметкаИсполнения", "Статус")
    
    ' Обрабатываем каждую строку таблицы
    For i = 1 To tblData.ListRows.Count
        ' Получаем дату записи (8-й столбец - Дата Вх.ФРП/Исх.ФРП)
        On Error Resume Next
        recordDate = CDate(tblData.DataBodyRange.Cells(i, 8).value)
        If Err.Number <> 0 Then recordDate = 0
        Err.Clear
        On Error GoTo 0
        
        ' Фильтр по дате - БОЛЕЕ ГИБКАЯ ПРОВЕРКА
        If recordDate > 0 And recordDate >= dateFrom And recordDate <= dateTo Then
            ' Получаем службу (2-й столбец)
            recordService = CStr(tblData.DataBodyRange.Cells(i, 2).value)
            
            ' Фильтр по службе
            If serviceFilter = "Все службы" Or serviceFilter = "" Or recordService = serviceFilter Then
                ' Получаем статус (19-й столбец)
                recordStatus = CStr(tblData.DataBodyRange.Cells(i, 19).value)
                
                ' Фильтр по статусу
                If statusFilter = "Все статусы" Or statusFilter = "" Or recordStatus = statusFilter Or _
                   (statusFilter = "(пустой)" And recordStatus = "") Then
                    
                    ' Создаем запись
                    Set record = New Collection
                    
                    For j = 0 To UBound(fieldNames)
                        If j + 1 <= tblData.ListColumns.Count Then
                            record.Add tblData.DataBodyRange.Cells(i, j + 1).value, CStr(fieldNames(j))
                        End If
                    Next j
                    
                    ' Добавляем в коллекцию
                    filteredData.Add record
                End If
            End If
        End If
    Next i
    
    Set GetFilteredData = filteredData
End Function

' ===============================================
' ФУНКЦИЯ РАСЧЕТА МЕТРИК KPI
' ===============================================

Private Function CalculateKPIMetrics(filteredData As Collection, tblData As ListObject) As Collection
    ' Объявление переменных
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
    
    Set metrics = New Collection
    
    ' Инициализация счетчиков
    totalDocs = filteredData.Count
    confirmedDocs = 0
    sentDocs = 0
    emptyStatusDocs = 0
    overdueDocs = 0
    totalDays = 0
    processedRecords = 0
    
    ' Обрабатываем каждую запись
    For Each record In filteredData
        ' Безопасное получение статуса
        On Error Resume Next
        recordStatus = CStr(record.item("Статус"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        On Error GoTo 0
        
        ' Подсчет по статусам
        Select Case recordStatus
            Case "Подтверждено"
                confirmedDocs = confirmedDocs + 1
            Case "Отпр.Исх."
                sentDocs = sentDocs + 1
            Case ""
                emptyStatusDocs = emptyStatusDocs + 1
        End Select
        
        ' Проверка на просрочку
        On Error Resume Next
        recordDate = CDate(record.item("ДатаФРП"))
        If Err.Number = 0 And recordDate > 0 Then
            daysDiff = Date - recordDate
            If daysDiff > CRITICAL_DELAY_DAYS And recordStatus <> "Подтверждено" And recordStatus <> "Отпр.Исх." Then
                overdueDocs = overdueDocs + 1
            End If
            
            ' Для расчета среднего времени обработки
            If recordStatus = "Подтверждено" Or recordStatus = "Отпр.Исх." Then
                totalDays = totalDays + daysDiff
                processedRecords = processedRecords + 1
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next record
    
    ' Расчет производных метрик
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
    
    ' Безопасное сохранение метрик
    On Error Resume Next
    metrics.Add totalDocs, "TotalDocs"
    metrics.Add confirmedDocs, "ConfirmedDocs"
    metrics.Add sentDocs, "SentDocs"
    metrics.Add inProgressDocs, "InProgressDocs"
    metrics.Add overdueDocs, "OverdueDocs"
    metrics.Add emptyStatusDocs, "EmptyStatusDocs"
    metrics.Add completionRate, "CompletionRate"
    metrics.Add avgProcessingDays, "AvgProcessingDays"
    On Error GoTo 0
    
    Set CalculateKPIMetrics = metrics
End Function

' ===============================================
' ПРОЦЕДУРА ОБНОВЛЕНИЯ KPI БЛОКОВ
' ===============================================

Private Sub UpdateKPIBlocks(wsDashboard As Worksheet, metrics As Collection)
    ' Объявление переменных
    Dim totalDocs As Long
    Dim confirmedDocs As Long
    Dim inProgressDocs As Long
    Dim overdueDocs As Long
    Dim completionRate As Double
    
    On Error Resume Next
    
    ' Безопасное получение значений метрик
    totalDocs = GetMetricValue(metrics, "TotalDocs")
    confirmedDocs = GetMetricValue(metrics, "ConfirmedDocs")
    inProgressDocs = GetMetricValue(metrics, "InProgressDocs")
    overdueDocs = GetMetricValue(metrics, "OverdueDocs")
    completionRate = GetMetricValue(metrics, "CompletionRate")
    
    ' Обновляем KPI блоки
    ' Блок 1 - Всего документов
    wsDashboard.Range("A6").value = totalDocs
    
    ' Блок 2 - Подтверждено
    wsDashboard.Range("E6").value = confirmedDocs
    wsDashboard.Range("E8").value = completionRate
    
    ' Блок 3 - В работе
    wsDashboard.Range("I6").value = inProgressDocs
    
    ' Блок 4 - Просрочены
    wsDashboard.Range("M6").value = overdueDocs
    
    On Error GoTo 0
End Sub

' ===============================================
' ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: БЕЗОПАСНОЕ ПОЛУЧЕНИЕ ЗНАЧЕНИЙ ИЗ КОЛЛЕКЦИИ
' ===============================================

Private Function GetMetricValue(metrics As Collection, keyName As String) As Variant
    ' Безопасное получение значения из коллекции по ключу
    Dim Result As Variant
    
    On Error Resume Next
    Result = metrics.item(keyName)
    If Err.Number <> 0 Then
        Result = 0
    End If
    Err.Clear
    On Error GoTo 0
    
    GetMetricValue = Result
End Function

' ===============================================
' ПРОЦЕДУРА ОБНОВЛЕНИЯ ТАБЛИЦ ДАННЫХ
' ===============================================

Private Sub UpdateDataTables(wsDashboard As Worksheet, filteredData As Collection, tblData As ListObject)
    ' Обновляем таблицу топ-служб
    Call UpdateTopServicesTable(wsDashboard, filteredData)
    
    ' Обновляем таблицу проблемных документов
    Call UpdateProblemDocsTable(wsDashboard, filteredData)
    
    ' Обновляем скрытые данные для диаграмм
    Call UpdateChartDataTables(wsDashboard, filteredData)
End Sub

' ===============================================
' ПРОЦЕДУРА ОБНОВЛЕНИЯ ТАБЛИЦЫ ТОП-СЛУЖБ
' ===============================================

Private Sub UpdateTopServicesTable(wsDashboard As Worksheet, filteredData As Collection)
    ' Объявление переменных
    Dim serviceStats As Object ' Dictionary для безопасной работы
    Dim serviceName As String
    Dim record As Collection
    Dim sortedServices As Collection
    Dim i As Integer
    Dim topCount As Integer
    Dim recordStatus As String
    Dim totalCount As Long
    Dim completedCount As Long
    Dim serviceKey As Variant
    
    ' Используем Dictionary для безопасной работы
    Set serviceStats = CreateObject("Scripting.Dictionary")
    
    For Each record In filteredData
        ' Безопасное получение службы
        On Error Resume Next
        serviceName = CStr(record.item("Служба"))
        If Err.Number <> 0 Then serviceName = ""
        Err.Clear
        On Error GoTo 0
        
        If serviceName <> "" Then
            ' Безопасное получение статуса
            On Error Resume Next
            recordStatus = CStr(record.item("Статус"))
            If Err.Number <> 0 Then recordStatus = ""
            Err.Clear
            On Error GoTo 0
            
            If Not serviceStats.Exists(serviceName) Then
                ' Создаем новую запись для службы
                If recordStatus = "Подтверждено" Or recordStatus = "Отпр.Исх." Then
                    serviceStats.Add serviceName, Array(1, 1) ' (Total, Completed)
                Else
                    serviceStats.Add serviceName, Array(1, 0)
                End If
            Else
                ' Обновляем существующую запись
                Dim currentData As Variant
                currentData = serviceStats(serviceName)
                totalCount = currentData(0) + 1
                If recordStatus = "Подтверждено" Or recordStatus = "Отпр.Исх." Then
                    completedCount = currentData(1) + 1
                Else
                    completedCount = currentData(1)
                End If
                serviceStats(serviceName) = Array(totalCount, completedCount)
            End If
        End If
    Next record
    
    ' Безопасная работа с Dictionary
    Set sortedServices = New Collection
    For Each serviceKey In serviceStats.Keys
        sortedServices.Add serviceKey
    Next serviceKey
    
    ' Заполняем таблицу топ-служб
    topCount = IIf(sortedServices.Count < MAX_TOP_SERVICES, sortedServices.Count, MAX_TOP_SERVICES)
    
    For i = 1 To MAX_TOP_SERVICES
        If i <= topCount Then
            ' Получаем данные службы
            serviceName = sortedServices.item(i)
            Dim serviceData As Variant
            serviceData = serviceStats(serviceName)
            
            ' Заполняем строку таблицы
            wsDashboard.Cells(29 + i, 1).value = i
            wsDashboard.Cells(29 + i, 2).value = serviceName
            wsDashboard.Cells(29 + i, 3).value = serviceData(0) ' Total
            wsDashboard.Cells(29 + i, 4).value = serviceData(1) ' Completed
            If serviceData(0) > 0 Then
                wsDashboard.Cells(29 + i, 5).value = serviceData(1) / serviceData(0)
            Else
                wsDashboard.Cells(29 + i, 5).value = 0
            End If
            wsDashboard.Cells(29 + i, 6).value = "?"
        Else
            ' Очищаем пустые строки
            wsDashboard.Range("A" & (29 + i) & ":F" & (29 + i)).ClearContents
        End If
    Next i
End Sub

' ===============================================
' ПРОЦЕДУРА ОБНОВЛЕНИЯ ПРОБЛЕМНЫХ ДОКУМЕНТОВ
' ===============================================

Private Sub UpdateProblemDocsTable(wsDashboard As Worksheet, filteredData As Collection)
    ' Объявление переменных
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
    
    ' Находим проблемные документы
    For Each record In filteredData
        ' Безопасное получение даты
        On Error Resume Next
        recordDate = CDate(record.item("ДатаФРП"))
        If Err.Number <> 0 Then recordDate = 0
        Err.Clear
        
        ' Безопасное получение статуса
        recordStatus = CStr(record.item("Статус"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        
        ' Безопасное получение службы
        serviceName = CStr(record.item("Служба"))
        If Err.Number <> 0 Then serviceName = ""
        Err.Clear
        
        ' Безопасное получение номера документа
        docNumber = CStr(record.item("НомерДокумента"))
        If Err.Number <> 0 Then docNumber = ""
        Err.Clear
        On Error GoTo 0
        
        If recordDate > 0 Then
            daysDiff = Date - recordDate
            
            ' Проверяем критерии проблемности
            If (daysDiff > WARNING_DELAY_DAYS And recordStatus = "") Or _
               (daysDiff > CRITICAL_DELAY_DAYS And recordStatus <> "Подтверждено" And recordStatus <> "Отпр.Исх.") Then
                
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
    
    ' Заполняем таблицу проблемных документов
    For i = 1 To maxProblems
        If i <= problemDocs.Count Then
            Set problemRecord = problemDocs.item(i)
            
            wsDashboard.Cells(29 + i, 8).value = problemRecord.item("Type")
            wsDashboard.Cells(29 + i, 9).value = problemRecord.item("Service")
            wsDashboard.Cells(29 + i, 10).value = problemRecord.item("DocNumber")
            wsDashboard.Cells(29 + i, 11).value = problemRecord.item("Date")
            wsDashboard.Cells(29 + i, 12).value = problemRecord.item("Days")
            wsDashboard.Cells(29 + i, 13).value = IIf(problemRecord.item("Status") = "", "Не задан", problemRecord.item("Status"))
            
            ' Условное форматирование критичных записей
            If problemRecord.item("Type") = "Критично" Then
                wsDashboard.Range("H" & (29 + i) & ":M" & (29 + i)).Interior.Color = RGB(255, 200, 200)
            Else
                wsDashboard.Range("H" & (29 + i) & ":M" & (29 + i)).Interior.Color = RGB(255, 255, 200)
            End If
        Else
            ' Очищаем пустые строки
            wsDashboard.Range("H" & (29 + i) & ":M" & (29 + i)).ClearContents
            wsDashboard.Range("H" & (29 + i) & ":M" & (29 + i)).Interior.ColorIndex = xlNone
        End If
    Next i
End Sub

' ===============================================
' ИСПРАВЛЕННАЯ ПРОЦЕДУРА ОБНОВЛЕНИЯ ДАННЫХ ДЛЯ ДИАГРАММ
' ===============================================

Private Sub UpdateChartDataTables(wsDashboard As Worksheet, filteredData As Collection)
    ' Объявление переменных для статистики статусов
    Dim confirmedCount As Long, sentCount As Long, inProgressCount As Long, emptyCount As Long
    Dim record As Collection
    Dim recordStatus As String
    
    ' Инициализация счетчиков
    confirmedCount = 0
    sentCount = 0
    inProgressCount = 0
    emptyCount = 0
    
    ' Подсчитываем статусы
    For Each record In filteredData
        ' Безопасное получение статуса
        On Error Resume Next
        recordStatus = CStr(record.item("Статус"))
        If Err.Number <> 0 Then recordStatus = ""
        Err.Clear
        On Error GoTo 0
        
        Select Case recordStatus
            Case "Подтверждено"
                confirmedCount = confirmedCount + 1
            Case "Отпр.Исх."
                sentCount = sentCount + 1
            Case ""
                emptyCount = emptyCount + 1
            Case Else
                inProgressCount = inProgressCount + 1
        End Select
    Next record
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Записываем данные в правильные ячейки с проверкой
    On Error Resume Next
    ' Очищаем старые данные
    wsDashboard.Range("V4:W8").ClearContents
    wsDashboard.Range("V11:W17").ClearContents
    
    ' Заполняем заголовки для диаграммы статусов
    wsDashboard.Range("V4").value = "Статус"
    wsDashboard.Range("W4").value = "Количество"
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Записываем данные только если есть хотя бы один непустой статус
    If confirmedCount + sentCount + inProgressCount + emptyCount > 0 Then
        ' Записываем данные построчно с проверкой
        wsDashboard.Range("V5").value = "Подтверждено"
        wsDashboard.Range("W5").value = confirmedCount
        
        wsDashboard.Range("V6").value = "Отпр.Исх."
        wsDashboard.Range("W6").value = sentCount
        
        wsDashboard.Range("V7").value = "В работе"
        wsDashboard.Range("W7").value = inProgressCount
        
        wsDashboard.Range("V8").value = "Без статуса"
        wsDashboard.Range("W8").value = emptyCount
        
        Debug.Print "Данные для диаграммы статусов: " & confirmedCount & ", " & sentCount & ", " & inProgressCount & ", " & emptyCount
    Else
        ' Если нет данных, устанавливаем минимальные значения для предотвращения ошибок
        wsDashboard.Range("V5").value = "Нет данных"
        wsDashboard.Range("W5").value = 1
        Debug.Print "Нет данных для диаграммы статусов"
    End If
    
    ' Заполняем заголовки для диаграммы служб
    wsDashboard.Range("V11").value = "Служба"
    wsDashboard.Range("W11").value = "Количество"
    
    ' Обновляем данные для столбчатой диаграммы служб (берем из таблицы топ-служб)
    Dim i As Integer
    Dim hasServiceData As Boolean
    hasServiceData = False
    
    For i = 1 To 6
        If wsDashboard.Cells(29 + i, 2).value <> "" And IsNumeric(wsDashboard.Cells(29 + i, 3).value) And wsDashboard.Cells(29 + i, 3).value > 0 Then
            wsDashboard.Cells(11 + i, 22).value = wsDashboard.Cells(29 + i, 2).value ' Название службы
            wsDashboard.Cells(11 + i, 23).value = wsDashboard.Cells(29 + i, 3).value ' Количество
            hasServiceData = True
        Else
            wsDashboard.Cells(11 + i, 22).ClearContents
            wsDashboard.Cells(11 + i, 23).ClearContents
        End If
    Next i
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Если нет данных по службам, добавляем заглушку
    If Not hasServiceData Then
        wsDashboard.Range("V12").value = "Нет данных"
        wsDashboard.Range("W12").value = 1
        Debug.Print "Нет данных для диаграммы служб"
    End If
    
    On Error GoTo 0
End Sub

' ===============================================
' НОВАЯ ПРОЦЕДУРА: СОЗДАНИЕ ДИАГРАММ С ПОЛНОЙ ВАЛИДАЦИЕЙ
' ===============================================

Private Sub CreateChartsWithValidation(wsDashboard As Worksheet, filteredData As Collection)
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Полная валидация данных перед созданием диаграмм
    
    ' Проверяем наличие данных
    If filteredData.Count = 0 Then
        Debug.Print "Нет данных для создания диаграмм"
        Call ShowNoDataMessage(wsDashboard)
        Exit Sub
    End If
    
    ' Создаем диаграммы с полной валидацией
    Call CreatePieChartSafe(wsDashboard)
    Call CreateColumnChartSafe(wsDashboard)
End Sub

' ===============================================
' БЕЗОПАСНАЯ ПРОЦЕДУРА СОЗДАНИЯ КРУГОВОЙ ДИАГРАММЫ
' ===============================================

Private Sub CreatePieChartSafe(wsDashboard As Worksheet)
    ' Объявление переменных
    Dim chartRange As Range
    Dim chartData As Range
    Dim chartObj As ChartObject
    Dim chartExists As Boolean
    Dim totalValues As Long
    Dim dataRange As Range
    Dim i As Integer
    
    On Error GoTo PieChartError
    
    ' Проверяем существование диаграммы
    Set chartObj = Nothing
    On Error Resume Next
    Set chartObj = wsDashboard.ChartObjects("StatusPieChart")
    chartExists = Not (chartObj Is Nothing)
    On Error GoTo PieChartError
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Валидация данных для диаграммы
    Set dataRange = wsDashboard.Range("V4:W8")
    Set chartRange = wsDashboard.Range("A11:G25")
    
    ' Проверяем наличие валидных данных
    totalValues = 0
    For i = 5 To 8
        If IsNumeric(wsDashboard.Cells(i, 23).value) Then
            totalValues = totalValues + CLng(wsDashboard.Cells(i, 23).value)
        End If
    Next i
    
    Debug.Print "Общие значения для круговой диаграммы: " & totalValues
    
    If totalValues <= 0 Then
        ' Если нет данных, удаляем диаграмму
        If chartExists Then
            chartObj.Delete
            Debug.Print "Удалена пустая круговая диаграмма"
        End If
        ' Показываем сообщение об отсутствии данных
        wsDashboard.Range("C15").value = "Нет данных для круговой диаграммы"
        wsDashboard.Range("C15").Font.Color = RGB(150, 150, 150)
        wsDashboard.Range("C15").Font.Italic = True
        Exit Sub
    End If
    
    ' Очищаем сообщение об отсутствии данных
    wsDashboard.Range("C15").ClearContents
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Очищаем старые диаграммы перед созданием новых
    If chartExists Then
        chartObj.Delete
        Set chartObj = Nothing
        chartExists = False
    End If
    
    ' Создаем новую диаграмму с валидными данными
    Set chartObj = wsDashboard.ChartObjects.Add(Left:=chartRange.Left, Top:=chartRange.Top, _
                                               Width:=chartRange.Width, Height:=chartRange.Height)
    chartObj.Name = "StatusPieChart"
    
    With chartObj.Chart
        ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Устанавливаем источник данных с проверкой
        .SetSourceData Source:=dataRange, PlotBy:=xlColumns
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Распределение по статусам"
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        
        ' Настройка цветов - безопасно
        On Error Resume Next
        If .SeriesCollection.Count > 0 And .SeriesCollection(1).Points.Count >= 4 Then
            .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(200, 255, 200) ' Подтверждено
            .SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(200, 200, 255) ' Отпр.Исх.
            .SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = RGB(255, 255, 200) ' В работе
            .SeriesCollection(1).Points(4).Format.Fill.ForeColor.RGB = RGB(255, 200, 200) ' Без статуса
        End If
        On Error GoTo PieChartError
        
        ' Показываем проценты
        If .SeriesCollection.Count > 0 Then
            .SeriesCollection(1).HasDataLabels = True
            .SeriesCollection(1).DataLabels.ShowPercentage = True
        End If
    End With
    
    Debug.Print "Круговая диаграмма создана успешно"
    Exit Sub
    
PieChartError:
    Debug.Print "Ошибка создания круговой диаграммы: " & Err.description
    ' При ошибке показываем сообщение пользователю
    If chartExists And Not chartObj Is Nothing Then
        chartObj.Delete
    End If
    wsDashboard.Range("C15").value = "Ошибка создания диаграммы: " & Err.description
    wsDashboard.Range("C15").Font.Color = RGB(255, 0, 0)
End Sub

' ===============================================
' БЕЗОПАСНАЯ ПРОЦЕДУРА СОЗДАНИЯ СТОЛБЧАТОЙ ДИАГРАММЫ
' ===============================================

Private Sub CreateColumnChartSafe(wsDashboard As Worksheet)
    ' Объявление переменных
    Dim chartRange As Range
    Dim chartData As Range
    Dim chartObj As ChartObject
    Dim chartExists As Boolean
    Dim hasValidData As Boolean
    Dim i As Integer
    
    On Error GoTo ColumnChartError
    
    ' Проверяем существование диаграммы
    Set chartObj = Nothing
    On Error Resume Next
    Set chartObj = wsDashboard.ChartObjects("ServiceColumnChart")
    chartExists = Not (chartObj Is Nothing)
    On Error GoTo ColumnChartError
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Валидация данных для диаграммы служб
    Set chartData = wsDashboard.Range("V11:W17")
    Set chartRange = wsDashboard.Range("I11:O25")
    
    ' Проверяем наличие валидных данных служб
    hasValidData = False
    For i = 12 To 17
        If wsDashboard.Cells(i, 22).value <> "" And IsNumeric(wsDashboard.Cells(i, 23).value) And wsDashboard.Cells(i, 23).value > 0 Then
            hasValidData = True
            Exit For
        End If
    Next i
    
    Debug.Print "Есть данные для столбчатой диаграммы: " & hasValidData
    
    If Not hasValidData Then
        ' Если нет данных, удаляем диаграмму
        If chartExists Then
            chartObj.Delete
            Debug.Print "Удалена пустая столбчатая диаграмма"
        End If
        ' Показываем сообщение об отсутствии данных
        wsDashboard.Range("L15").value = "Нет данных по службам"
        wsDashboard.Range("L15").Font.Color = RGB(150, 150, 150)
        wsDashboard.Range("L15").Font.Italic = True
        Exit Sub
    End If
    
    ' Очищаем сообщение об отсутствии данных
    wsDashboard.Range("L15").ClearContents
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Очищаем старые диаграммы перед созданием новых
    If chartExists Then
        chartObj.Delete
        Set chartObj = Nothing
        chartExists = False
    End If
    
    ' Создаем новую диаграмму с валидными данными
    Set chartObj = wsDashboard.ChartObjects.Add(Left:=chartRange.Left, Top:=chartRange.Top, _
                                               Width:=chartRange.Width, Height:=chartRange.Height)
    chartObj.Name = "ServiceColumnChart"
    
    With chartObj.Chart
        ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Устанавливаем источник данных с проверкой
        .SetSourceData Source:=chartData, PlotBy:=xlColumns
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Активность по службам"
        .HasLegend = False
        
        ' Настройка осей - безопасно
        On Error Resume Next
        If .Axes.Count >= 2 Then
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Службы"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Количество документов"
        End If
        On Error GoTo ColumnChartError
        
        ' Цвет столбцов и подписи данных
        If .SeriesCollection.Count > 0 Then
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(100, 150, 255)
            .SeriesCollection(1).HasDataLabels = True
        End If
    End With
    
    Debug.Print "Столбчатая диаграмма создана успешно"
    Exit Sub
    
ColumnChartError:
    Debug.Print "Ошибка создания столбчатой диаграммы: " & Err.description
    ' При ошибке показываем сообщение пользователю
    If chartExists And Not chartObj Is Nothing Then
        chartObj.Delete
    End If
    wsDashboard.Range("L15").value = "Ошибка создания диаграммы: " & Err.description
    wsDashboard.Range("L15").Font.Color = RGB(255, 0, 0)
End Sub

' ===============================================
' ПРОЦЕДУРА ПОКАЗА СООБЩЕНИЯ ОБ ОТСУТСТВИИ ДАННЫХ
' ===============================================

Private Sub ShowNoDataMessage(wsDashboard As Worksheet)
    ' Показываем сообщения об отсутствии данных
    wsDashboard.Range("C15").value = "Нет данных за выбранный период"
    wsDashboard.Range("C15").Font.Color = RGB(150, 150, 150)
    wsDashboard.Range("C15").Font.Size = 12
    wsDashboard.Range("C15").Font.Italic = True
    
    wsDashboard.Range("L15").value = "Нет данных за выбранный период"
    wsDashboard.Range("L15").Font.Color = RGB(150, 150, 150)
    wsDashboard.Range("L15").Font.Size = 12
    wsDashboard.Range("L15").Font.Italic = True
    
    ' Удаляем существующие диаграммы
    On Error Resume Next
    wsDashboard.ChartObjects("StatusPieChart").Delete
    wsDashboard.ChartObjects("ServiceColumnChart").Delete
    On Error GoTo 0
End Sub

' ===============================================
' ПРОЦЕДУРА ОБНОВЛЕНИЯ ТРЕНДОВ
' ===============================================

Private Sub UpdateTrendsData(wsDashboard As Worksheet, filteredData As Collection)
    ' Здесь можно добавить логику для анализа трендов по дням недели
    ' Пока оставляем существующие случайные данные
    
    ' В будущем можно добавить:
    ' - Анализ поступления документов по дням недели
    ' - Расчет средней эффективности по дням
    ' - Создание линейного графика динамики
End Sub

' ===============================================
' ПРОЦЕДУРА ПРИМЕНЕНИЯ УСЛОВНОГО ФОРМАТИРОВАНИЯ
' ===============================================

Private Sub ApplyConditionalFormatting(wsDashboard As Worksheet, metrics As Collection)
    ' Объявление переменных
    Dim overdueDocs As Long
    Dim totalDocs As Long
    Dim completionRate As Double
    
    ' Безопасное получение метрик
    overdueDocs = GetMetricValue(metrics, "OverdueDocs")
    totalDocs = GetMetricValue(metrics, "TotalDocs")
    completionRate = GetMetricValue(metrics, "CompletionRate")
    
    ' Форматирование KPI блоков в зависимости от критичности
    ' Блок просроченных документов
    If overdueDocs > 0 Then
        wsDashboard.Range("M5:O8").Interior.Color = RGB(255, 150, 150) ' Ярко-красный
    Else
        wsDashboard.Range("M5:O8").Interior.Color = RGB(255, 180, 180) ' Обычный красный
    End If
    
    ' Блок завершенности
    If completionRate >= 0.8 Then
        wsDashboard.Range("E5:G8").Interior.Color = RGB(180, 255, 180) ' Зеленый
    ElseIf completionRate >= 0.5 Then
        wsDashboard.Range("E5:G8").Interior.Color = RGB(255, 255, 180) ' Желтый
    Else
        wsDashboard.Range("E5:G8").Interior.Color = RGB(255, 200, 200) ' Красный
    End If
End Sub

' ===============================================
' ПРОЦЕДУРА ОБНОВЛЕНИЯ ПАНЕЛИ ПАРАМЕТРОВ
' ===============================================

Private Sub UpdateParametersPanel(wsDashboard As Worksheet, dateFrom As Date, dateTo As Date, processedRecords As Long)
    ' Обновляем информацию о последнем обновлении
    wsDashboard.Range("Q5").value = Now
    
    ' Обновляем статус данных
    If processedRecords > 0 Then
        wsDashboard.Range("Q6").value = "ОК (" & processedRecords & " записей)"
    Else
        wsDashboard.Range("Q6").value = "Нет данных"
    End If
    
    ' Обновляем период фильтрации
    wsDashboard.Range("Q8").value = Format(dateFrom, "dd.mm.yyyy") & " - " & Format(dateTo, "dd.mm.yyyy")
End Sub

' ===============================================
' СИСТЕМА АВТОМАТИЧЕСКОГО ОБНОВЛЕНИЯ
' ===============================================

Public Sub StartAutoUpdate()
    ' Проверяем настройку автообновления
    Dim wsDashboard As Worksheet
    Set wsDashboard = GetWorksheetSafe("Dashboard")
    
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
        ' Планируем следующее обновление
        Application.OnTime Now + TimeSerial(0, 0, UPDATE_INTERVAL_SECONDS), "DashboardModule.AutoUpdateDashboard"
    End If
End Sub

' ===============================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ===============================================

Private Function GetWorksheetSafe(sheetName As String) As Worksheet
    Dim Ws As Worksheet
    On Error Resume Next
    Set Ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    Set GetWorksheetSafe = Ws
End Function

Private Function GetListObjectSafe(Ws As Worksheet, tableName As String) As ListObject
    Dim tbl As ListObject
    On Error Resume Next
    If Not Ws Is Nothing Then
        Set tbl = Ws.ListObjects(tableName)
    End If
    On Error GoTo 0
    Set GetListObjectSafe = tbl
End Function

' ===============================================
' ПРОЦЕДУРЫ УПРАВЛЕНИЯ DASHBOARD
' ===============================================

Public Sub RefreshDashboard()
    Call GenerateDashboard
End Sub

Public Sub ToggleAutoUpdate()
    Dim wsDashboard As Worksheet
    Set wsDashboard = GetWorksheetSafe("Dashboard")
    
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


