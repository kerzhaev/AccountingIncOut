Attribute VB_Name = "SystemLogger"
'==============================================
' МОДУЛЬ ЖУРНАЛИРОВАНИЯ - SystemLogger
' Назначение: Ведение журнала операций системы
' Состояние: ЭТАП 1 - БАЗОВАЯ ИНФРАСТРУКТУРА
' Версия: 1.0.0
' Дата: 23.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

Private Const LOG_SHEET_NAME As String = "ЖурналОпераций"

' Инициализация журнала
Public Sub InitializeLog()
    Dim wsLog As Worksheet
    
    On Error GoTo CreateLog
    
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    Exit Sub
    
CreateLog:
    Set wsLog = ThisWorkbook.Worksheets.Add
    wsLog.Name = LOG_SHEET_NAME
    wsLog.Visible = xlSheetVeryHidden
    
    ' Создаем заголовки журнала
    With wsLog
        .Range("A1").value = "Дата/Время"
        .Range("B1").value = "Операция"
        .Range("C1").value = "Описание"
        .Range("D1").value = "Статус"
        .Range("E1").value = "Время выполнения (сек)"
        .Range("F1").value = "Пользователь"
        
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 220, 255)
    End With
End Sub

' Логирование операции
Public Sub LogOperation(operationName As String, description As String, Status As String, executionTime As Double)
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim maxRecords As Long
    
    On Error GoTo LogError
    
    ' Проверяем, включено ли журналирование
    If Not SystemSettings.GetSetting("LogEnabled", True) Then
        Exit Sub
    End If
    
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    maxRecords = SystemSettings.GetSetting("MaxLogRecords", 100)
    
    ' Находим следующую свободную строку
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Записываем данные
    With wsLog
        .Cells(nextRow, 1).value = Now()
        .Cells(nextRow, 2).value = operationName
        .Cells(nextRow, 3).value = description
        .Cells(nextRow, 4).value = Status
        .Cells(nextRow, 5).value = Round(executionTime, 2)
        .Cells(nextRow, 6).value = Application.UserName
    End With
    
    ' Применяем условное форматирование по статусу
    Call FormatLogEntry(wsLog.Range(wsLog.Cells(nextRow, 1), wsLog.Cells(nextRow, 6)), Status)
    
    ' Очищаем старые записи, если превышен лимит
    If nextRow - 1 > maxRecords Then
        Call CleanOldLogEntries(wsLog, maxRecords)
    End If
    
    Exit Sub
    
LogError:
    ' В случае ошибки логирования, записываем в Debug
    Debug.Print "LOG ERROR: " & Now() & " | " & operationName & " | " & description & " | " & Status
End Sub

' Форматирование записи журнала по статусу
Private Sub FormatLogEntry(entryRange As Range, Status As String)
    Select Case UCase(Status)
        Case "SUCCESS"
            entryRange.Interior.Color = RGB(200, 255, 200) ' Светло-зеленый
        Case "ERROR"
            entryRange.Interior.Color = RGB(255, 200, 200) ' Светло-красный
        Case "WARNING"
            entryRange.Interior.Color = RGB(255, 255, 200) ' Светло-желтый
        Case "START"
            entryRange.Interior.Color = RGB(200, 200, 255) ' Светло-синий
        Case Else
            entryRange.Interior.ColorIndex = xlNone
    End Select
End Sub

' Очистка старых записей журнала
Private Sub CleanOldLogEntries(wsLog As Worksheet, maxRecords As Long)
    Dim TotalRows As Long
    Dim RowsToDelete As Long
    
    TotalRows = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row - 1 ' -1 для заголовка
    RowsToDelete = TotalRows - maxRecords
    
    If RowsToDelete > 0 Then
        ' Удаляем старые строки (начиная с 2-й, так как 1-я - заголовок)
        wsLog.Rows("2:" & (RowsToDelete + 1)).Delete
    End If
End Sub

' Показ диалога журнала
Public Sub ShowLogDialog()
    Dim wsLog As Worksheet
    
    On Error GoTo ShowLogError
    
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    
    ' Временно делаем лист видимым
    wsLog.Visible = xlSheetVisible
    wsLog.Activate
    
    ' Автоподбор ширины столбцов
    wsLog.Columns("A:F").AutoFit
    
    ' Применяем фильтр
    If wsLog.Range("A1").AutoFilter = False Then
        wsLog.Range("A1:F1").AutoFilter
    End If
    
    MsgBox "Журнал операций открыт." & vbCrLf & _
           "Используйте фильтры для поиска нужных записей." & vbCrLf & _
           "После просмотра нажмите OK для скрытия журнала.", _
           vbInformation, "Журнал операций"
    
    ' Скрываем лист обратно
    wsLog.Visible = xlSheetVeryHidden
    
    Exit Sub
    
ShowLogError:
    MsgBox "Ошибка открытия журнала: " & Err.description, vbExclamation, "Ошибка"
End Sub

' Очистка журнала
Public Sub ClearLog()
    Dim wsLog As Worksheet
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Очистить весь журнал операций?" & vbCrLf & _
                     "Это действие нельзя отменить!", _
                     vbYesNo + vbExclamation + vbDefaultButton2, "Подтверждение очистки")
    
    If response = vbYes Then
        Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
        wsLog.Range("2:" & wsLog.Rows.Count).ClearContents
        MsgBox "Журнал операций очищен.", vbInformation, "Очистка завершена"
    End If
End Sub


