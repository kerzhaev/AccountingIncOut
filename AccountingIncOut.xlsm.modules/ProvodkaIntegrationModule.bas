Attribute VB_Name = "ProvodkaIntegrationModule"
'==============================================
' МОДУЛЬ ИНТЕГРАЦИИ С 1С - ProvodkaIntegrationModule
' Назначение: Автоматическое сопоставление проводок 1С с документами ВхИсх
' Состояние: НОВЫЙ МОДУЛЬ ДЛЯ ИНТЕГРАЦИИ С ВЫГРУЗКОЙ 1С
' Версия: 1.0.0
' Дата: 21.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' Структура для результата поиска
Public Type MatchResult
    Found As Boolean
    ProvodkaNumber As String
    ProvodkaDate As Date
    MatchCount As Long
    StatusMessage As String
    candidatesList As String
End Type

' Массовая обработка всех записей ВхИсх с выгрузкой из файла 1С
Public Sub MassProcessWithFileSelection()
    Dim filePath As String
    Dim wb1C As Workbook
    Dim ws1C As Worksheet
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    Dim ProcessedCount As Long
    Dim FoundCount As Long
    Dim SkippedCount As Long
    Dim MultipleCount As Long
    Dim errorCount As Long
    
    On Error GoTo MassProcessError
    
    ' Выбор файла выгрузки 1С
    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        , "Выберите файл выгрузки из 1С")
    
    If filePath = "False" Then Exit Sub
    
    ' Открываем файл выгрузки 1С
    Application.StatusBar = "Открытие файла выгрузки 1С..."
    Set wb1C = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws1C = wb1C.Worksheets(1) ' Первый лист файла
    
    ' Получаем таблицу ВхИсх
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    Application.StatusBar = "Начало массовой обработки..."
    Application.ScreenUpdating = False
    
    ' Обрабатываем каждую запись ВхИсх
    Dim i As Long
    Dim currentSuma As Double
    Dim currentCorrespondent As String
    Dim currentOtmetka As String
    Dim MatchResult As MatchResult
    
    For i = 1 To tblData.ListRows.Count
        
        ' Получаем данные текущей записи ВхИсх
        On Error Resume Next
        currentSuma = CDbl(tblData.DataBodyRange.Cells(i, 6).value)          ' Сумма документа
        currentCorrespondent = CStr(tblData.DataBodyRange.Cells(i, 9).value) ' От кого поступил
        currentOtmetka = Trim(CStr(tblData.DataBodyRange.Cells(i, 18).value)) ' Отметка об исполнении
        On Error GoTo MassProcessError
        
        ' Пропускаем записи с уже заполненной отметкой об исполнении
        If currentOtmetka = "" Then
            
            ' Ищем соответствие в выгрузке 1С
            MatchResult = FindMatchInFile(currentSuma, currentCorrespondent, ws1C)
            
            If MatchResult.Found Then
                ' Записываем номер проводки в отметку об исполнении
                tblData.DataBodyRange.Cells(i, 18).value = MatchResult.ProvodkaNumber
                FoundCount = FoundCount + 1
                
            ElseIf MatchResult.MatchCount > 1 Then
                ' Множественные совпадения - требуют ручной обработки
                MultipleCount = MultipleCount + 1
                
            End If
            
            ProcessedCount = ProcessedCount + 1
        Else
            SkippedCount = SkippedCount + 1
        End If
        
        ' Обновляем прогресс каждые 25 записей
        If (ProcessedCount + SkippedCount) Mod 25 = 0 Then
            Application.StatusBar = "Обработано " & (ProcessedCount + SkippedCount) & " из " & tblData.ListRows.Count & " записей"
        End If
        
    Next i
    
    ' Закрываем файл выгрузки 1С
    wb1C.Close False
    
    Application.ScreenUpdating = True
    
    ' Показываем результаты обработки
    MsgBox "МАССОВАЯ ОБРАБОТКА ЗАВЕРШЕНА:" & vbCrLf & vbCrLf & _
           "?? СТАТИСТИКА:" & vbCrLf & _
           "Всего записей в ВхИсх: " & tblData.ListRows.Count & vbCrLf & _
           "Обработано (без отметки): " & ProcessedCount & vbCrLf & _
           "Пропущено (уже заполнены): " & SkippedCount & vbCrLf & vbCrLf & _
           "? РЕЗУЛЬТАТЫ:" & vbCrLf & _
           "Найдено автоматически: " & FoundCount & vbCrLf & _
           "Множественные совпадения: " & MultipleCount & vbCrLf & _
           "Не найдено: " & (ProcessedCount - FoundCount - MultipleCount) & vbCrLf & vbCrLf & _
           "?? Процент успеха: " & Format(IIf(ProcessedCount > 0, FoundCount / ProcessedCount, 0), "0.0%"), _
           vbInformation, "Результаты интеграции с 1С"
           
    Application.StatusBar = "Интеграция завершена. Найдено " & FoundCount & " соответствий из " & ProcessedCount & " записей."
    
    Exit Sub
    
MassProcessError:
    Application.ScreenUpdating = True
    If Not wb1C Is Nothing Then wb1C.Close False
    
    MsgBox "Ошибка массовой обработки:" & vbCrLf & _
           "Ошибка: " & Err.Number & " - " & Err.description & vbCrLf & _
           "Обработано записей: " & ProcessedCount, _
           vbCritical, "Критическая ошибка"
           
    Application.StatusBar = "Ошибка массовой обработки"
End Sub

' Поиск соответствующей проводки в файле выгрузки 1С
Private Function FindMatchInFile(suma As Double, Correspondent As String, ws1C As Worksheet) As MatchResult
    Dim Result As MatchResult
    Dim LastRow As Long
    Dim i As Long
    
    ' Данные из выгрузки 1С
    Dim currentStatus As String
    Dim currentSuma As Double
    Dim currentCorrespondent As String
    Dim CurrentNumber As String
    Dim CurrentDate As Date
    
    Dim CandidatesCount As Long
    Dim candidatesList As String
    Dim bestCandidate As String
    Dim bestCandidateDate As Date
    
    On Error GoTo FindError
    
    ' Инициализация результата
    Result.Found = False
    Result.MatchCount = 0
    Result.StatusMessage = "Не найдено"
    Result.candidatesList = ""
    
    ' Определяем последнюю строку с данными
    LastRow = ws1C.Cells(ws1C.Rows.Count, 1).End(xlUp).Row
    
    ' Проверяем наличие данных
    If LastRow < 2 Then
        Result.StatusMessage = "Файл выгрузки пустой"
        FindMatchInFile = Result
        Exit Function
    End If
    
    ' Поиск по всем строкам выгрузки 1С (начиная со 2-й строки, первая - заголовки)
    For i = 2 To LastRow
        
        On Error Resume Next
        ' Читаем данные из выгрузки 1С
        currentStatus = CStr(ws1C.Cells(i, 1).value)        ' Столбец A - Статус
        currentSuma = CDbl(ws1C.Cells(i, 5).value)          ' Столбец E - Сумма (предположительно)
        currentCorrespondent = CStr(ws1C.Cells(i, 6).value) ' Столбец F - Корреспондент (предположительно)
        CurrentNumber = CStr(ws1C.Cells(i, 3).value)        ' Столбец C - Номер
        CurrentDate = CDate(ws1C.Cells(i, 2).value)         ' Столбец B - Дата
        On Error GoTo FindError
        
' ПРОВЕРЯЕМ КРИТЕРИИ СООТВЕТСТВИЯ
' Исключаем непроведенные документы (статус = 1)
' Точное совпадение суммы (с допуском 0.01)
' Частичное вхождение корреспондента
If (currentStatus <> "1") And _
   (Abs(currentSuma - suma) < 0.01) And _
   (InStr(UCase(currentCorrespondent), UCase(Correspondent)) > 0) Then
   
    CandidatesCount = CandidatesCount + 1
    
    ' Сохраняем информацию о кандидатах
    If CandidatesCount = 1 Then
        bestCandidate = CurrentNumber
        bestCandidateDate = CurrentDate
        candidatesList = CurrentNumber & " (" & Format(CurrentDate, "dd.mm.yyyy") & ")"
    Else
        candidatesList = candidatesList & "; " & CurrentNumber & " (" & Format(CurrentDate, "dd.mm.yyyy") & ")"
        
        ' Выбираем более раннюю дату как лучший кандидат
        If CurrentDate < bestCandidateDate Then
            bestCandidate = CurrentNumber
            bestCandidateDate = CurrentDate
        End If
    End If
    
End If

        
    Next i
    
    ' Определяем результат поиска
    Result.MatchCount = CandidatesCount
    Result.candidatesList = candidatesList
    
    If CandidatesCount = 1 Then
        Result.Found = True
        Result.ProvodkaNumber = bestCandidate
        Result.ProvodkaDate = bestCandidateDate
        Result.StatusMessage = "Найдено единственное соответствие"
        
    ElseIf CandidatesCount > 1 Then
        Result.Found = False ' Требует ручного выбора
        Result.ProvodkaNumber = bestCandidate
        Result.ProvodkaDate = bestCandidateDate
        Result.StatusMessage = "Найдено " & CandidatesCount & " вариантов (выбран по дате)"
        
    Else
        Result.Found = False
        Result.StatusMessage = "Соответствие не найдено"
    End If
    
    FindMatchInFile = Result
    Exit Function
    
FindError:
    Result.Found = False
    Result.MatchCount = 0
    Result.StatusMessage = "Ошибка поиска: " & Err.description
    FindMatchInFile = Result
End Function

' Обработка одной записи (для ручного поиска из формы)
Public Sub ProcessSingleRecord(RowIndex As Long)
    Dim filePath As String
    Dim wb1C As Workbook
    Dim ws1C As Worksheet
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    Dim currentSuma As Double
    Dim currentCorrespondent As String
    Dim MatchResult As MatchResult
    
    On Error GoTo SingleProcessError
    
    ' Получаем таблицу ВхИсх
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    ' Проверяем корректность номера строки
    If RowIndex < 1 Or RowIndex > tblData.ListRows.Count Then
        MsgBox "Неверный номер записи: " & RowIndex, vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' Получаем данные записи
    currentSuma = CDbl(tblData.DataBodyRange.Cells(RowIndex, 6).value)
    currentCorrespondent = CStr(tblData.DataBodyRange.Cells(RowIndex, 9).value)
    
    ' Выбор файла выгрузки 1С
    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        , "Выберите файл выгрузки из 1С для поиска проводки")
    
    If filePath = "False" Then Exit Sub
    
    ' Открываем файл выгрузки 1С
    Application.StatusBar = "Поиск проводки в файле 1С..."
    Set wb1C = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws1C = wb1C.Worksheets(1)
    
    ' Ищем соответствие
    MatchResult = FindMatchInFile(currentSuma, currentCorrespondent, ws1C)
    
    ' Закрываем файл
    wb1C.Close False
    
    ' Обрабатываем результат
    If MatchResult.Found Then
        ' Автоматически заполняем поле
        tblData.DataBodyRange.Cells(RowIndex, 18).value = MatchResult.ProvodkaNumber
        
        MsgBox "? ПРОВОДКА НАЙДЕНА!" & vbCrLf & vbCrLf & _
               "Номер проводки: " & MatchResult.ProvodkaNumber & vbCrLf & _
               "Дата проводки: " & Format(MatchResult.ProvodkaDate, "dd.mm.yyyy") & vbCrLf & _
               "Сумма: " & currentSuma & vbCrLf & _
               "Корреспондент: " & currentCorrespondent & vbCrLf & vbCrLf & _
               "Номер проводки записан в 'Отметка об исполнении'", _
               vbInformation, "Поиск проводки"
               
        ' Если форма открыта, обновляем отображение
        If TableEventHandler.IsFormOpen("UserFormVhIsh") Then
            UserFormVhIsh.txtOtmetkaIspolnenie.Text = MatchResult.ProvodkaNumber
            UserFormVhIsh.txtOtmetkaIspolnenie.BackColor = RGB(200, 255, 200) ' Светло-зеленый
        End If
        
    ElseIf MatchResult.MatchCount > 1 Then
        ' Показываем диалог выбора из множественных вариантов
        Call ShowMultipleChoiceDialog(RowIndex, MatchResult, currentSuma, currentCorrespondent)
        
    Else
        MsgBox "? ПРОВОДКА НЕ НАЙДЕНА" & vbCrLf & vbCrLf & _
               "Критерии поиска:" & vbCrLf & _
               "Сумма: " & currentSuma & vbCrLf & _
               "Корреспондент: " & currentCorrespondent & vbCrLf & vbCrLf & _
               "Возможные причины:" & vbCrLf & _
               "• Документ еще не проведен в 1С" & vbCrLf & _
               "• Отличается сумма или название корреспондента" & vbCrLf & _
               "• Документ сторнирован в 1С", _
               vbExclamation, "Поиск проводки"
    End If
    
    Application.StatusBar = "Поиск проводки завершен"
    Exit Sub
    
SingleProcessError:
    If Not wb1C Is Nothing Then wb1C.Close False
    
    MsgBox "Ошибка поиска проводки:" & vbCrLf & _
           "Ошибка: " & Err.Number & " - " & Err.description, _
           vbCritical, "Ошибка"
           
    Application.StatusBar = "Ошибка поиска проводки"
End Sub

' Диалог выбора из множественных вариантов
Private Sub ShowMultipleChoiceDialog(RowIndex As Long, MatchResult As MatchResult, suma As Double, Correspondent As String)
    Dim userChoice As String
    Dim selectedProvodka As String
    
    userChoice = InputBox( _
        "?? НАЙДЕНО НЕСКОЛЬКО ВАРИАНТОВ ПРОВОДОК:" & vbCrLf & vbCrLf & _
        "Критерии поиска:" & vbCrLf & _
        "Сумма: " & suma & vbCrLf & _
        "Корреспондент: " & Correspondent & vbCrLf & vbCrLf & _
        "Найденные варианты:" & vbCrLf & _
        MatchResult.candidatesList & vbCrLf & vbCrLf & _
        "Введите номер проводки для записи" & vbCrLf & _
        "(или оставьте пустым для отмены):", _
        "Выбор проводки", _
        MatchResult.ProvodkaNumber)
    
    If Trim(userChoice) <> "" Then
        ' Записываем выбранную проводку
        Dim wsData As Worksheet
        Dim tblData As ListObject
        
        Set wsData = ThisWorkbook.Worksheets("ВхИсх")
        Set tblData = wsData.ListObjects("ВходящиеИсходящие")
        
        tblData.DataBodyRange.Cells(RowIndex, 18).value = Trim(userChoice)
        
        MsgBox "? Проводка записана: " & Trim(userChoice), vbInformation, "Выбор сохранен"
        
        ' Обновляем форму, если она открыта
        If TableEventHandler.IsFormOpen("UserFormVhIsh") Then
            UserFormVhIsh.txtOtmetkaIspolnenie.Text = Trim(userChoice)
            UserFormVhIsh.txtOtmetkaIspolnenie.BackColor = RGB(255, 255, 200) ' Светло-желтый (ручной ввод)
        End If
    End If
End Sub

' Поиск проводки для текущей записи из формы
Public Sub FindProvodkaForCurrentRecord()
    If DataManager.CurrentRecordRow > 0 Then
        Call ProcessSingleRecord(DataManager.CurrentRecordRow)
    Else
        MsgBox "Выберите запись в таблице или откройте форму с записью", vbExclamation, "Поиск проводки"
    End If
End Sub

' Получение статистики сопоставления
Public Sub ShowMatchingStatistics()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    
    Dim totalRecords As Long
    Dim filledRecords As Long
    Dim emptyRecords As Long
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    totalRecords = tblData.ListRows.Count
    
    For i = 1 To totalRecords
        If Trim(CStr(tblData.DataBodyRange.Cells(i, 18).value)) <> "" Then
            filledRecords = filledRecords + 1
        Else
            emptyRecords = emptyRecords + 1
        End If
    Next i
    
    MsgBox "?? СТАТИСТИКА СОПОСТАВЛЕНИЯ С 1С:" & vbCrLf & vbCrLf & _
           "Всего записей в системе: " & totalRecords & vbCrLf & _
           "С отметкой об исполнении: " & filledRecords & " (" & Format(filledRecords / totalRecords, "0.0%") & ")" & vbCrLf & _
           "Без отметки об исполнении: " & emptyRecords & " (" & Format(emptyRecords / totalRecords, "0.0%") & ")" & vbCrLf & vbCrLf & _
           "Рекомендация: Используйте 'Массовую обработку' для" & vbCrLf & _
           "автоматического заполнения пустых отметок.", _
           vbInformation, "Статистика интеграции"
End Sub

' Очистка всех отметок об исполнении (для повторной обработки)
Public Sub ClearAllProvodkaMarks()
    Dim Response As VbMsgBoxResult
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    Dim clearedCount As Long
    
    Response = MsgBox("?? ВНИМАНИЕ!" & vbCrLf & vbCrLf & _
                      "Вы собираетесь очистить ВСЕ отметки об исполнении" & vbCrLf & _
                      "в таблице ВходящиеИсходящие." & vbCrLf & vbCrLf & _
                      "Это действие нельзя отменить!" & vbCrLf & _
                      "Продолжить?", _
                      vbYesNo + vbExclamation + vbDefaultButton2, "Подтверждение очистки")
    
    If Response = vbNo Then Exit Sub
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    Application.StatusBar = "Очистка отметок об исполнении..."
    
    For i = 1 To tblData.ListRows.Count
        If Trim(CStr(tblData.DataBodyRange.Cells(i, 18).value)) <> "" Then
            tblData.DataBodyRange.Cells(i, 18).value = ""
            clearedCount = clearedCount + 1
        End If
    Next i
    
    MsgBox "? Очистка завершена!" & vbCrLf & _
           "Очищено записей: " & clearedCount, _
           vbInformation, "Очистка выполнена"
           
    Application.StatusBar = "Очистка отметок завершена. Очищено " & clearedCount & " записей."
End Sub


