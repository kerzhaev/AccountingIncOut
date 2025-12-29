Attribute VB_Name = "CommonUtilities"
'==============================================
' МОДУЛЬ ОБЩИХ УТИЛИТАРНЫХ ФУНКЦИЙ - CommonUtilities
' Назначение: Централизованное хранение общих утилитарных функций
' Состояние: СОЗДАН В РЕЗУЛЬТАТЕ РЕФАКТОРИНГА
' Версия: 1.0.0
' Дата: 10.01.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' =============================================
' БЕЗОПАСНОЕ ПОЛУЧЕНИЕ ЛИСТА
' =============================================
' @description Получает лист по имени с обработкой ошибок
' @param   sheetName [String] Имя листа
' @return  [Worksheet] Объект листа или Nothing если не найден
' =============================================
Public Function GetWorksheetSafe(sheetName As String) As Worksheet
    Dim Ws As Worksheet
    
    On Error Resume Next
    Set Ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    Set GetWorksheetSafe = Ws
End Function

' =============================================
' БЕЗОПАСНОЕ ПОЛУЧЕНИЕ ТАБЛИЦЫ
' =============================================
' @description Получает таблицу (ListObject) по имени с обработкой ошибок
' @param   Ws [Worksheet] Лист, на котором находится таблица
' @param   tableName [String] Имя таблицы
' @return  [ListObject] Объект таблицы или Nothing если не найдена
' =============================================
Public Function GetListObjectSafe(Ws As Worksheet, tableName As String) As ListObject
    Dim tbl As ListObject
    
    On Error Resume Next
    If Not Ws Is Nothing Then
        Set tbl = Ws.ListObjects(tableName)
    End If
    On Error GoTo 0
    
    Set GetListObjectSafe = tbl
End Function

' =============================================
' ВАЛИДАЦИЯ ФОРМАТА ДАТЫ
' =============================================
' @description Проверяет корректность формата даты ДД.ММ.ГГ
' @param   DateText [String] Текст даты в формате ДД.ММ.ГГ
' @return  [Boolean] True если формат корректен и дата не в будущем
' =============================================
Public Function IsValidDateFormat(DateText As String) As Boolean
    On Error GoTo DateError
    
    ' Проверка формата ДД.ММ.ГГ
    If Len(DateText) = 8 And Mid(DateText, 3, 1) = "." And Mid(DateText, 6, 1) = "." Then
        Dim TestDate As Date
        ' Преобразуем ГГ в полный год (20ГГ)
        Dim fullDateText As String
        fullDateText = Left(DateText, 6) & "20" & Right(DateText, 2)
        TestDate = CDate(fullDateText)
        
        ' Проверка, что дата не в будущем
        If TestDate > Date Then
            IsValidDateFormat = False
        Else
            IsValidDateFormat = True
        End If
    Else
        IsValidDateFormat = False
    End If
    
    Exit Function
    
DateError:
    IsValidDateFormat = False
End Function

' =============================================
' ФОРМАТИРОВАНИЕ ДАТЫ ИЗ ЯЧЕЙКИ
' =============================================
' @description Форматирует дату из ячейки в строку ДД.ММ.ГГ
' @param   cell [Range] Ячейка с датой
' @return  [String] Отформатированная дата или пустая строка
' =============================================
Public Function FormatDateCell(cell As Range) As String
    On Error GoTo FormatError
    
    If Not IsEmpty(cell.value) And IsDate(cell.value) Then
        FormatDateCell = Format(cell.value, "dd.mm.yy")
    Else
        FormatDateCell = ""
    End If
    
    Exit Function
    
FormatError:
    FormatDateCell = ""
End Function

' =============================================
' ЗАПИСЬ ДАТЫ В ЯЧЕЙКУ
' =============================================
' @description Записывает дату в формате ДД.ММ.ГГ в ячейку
' @param   cell [Range] Ячейка для записи
' @param   DateText [String] Текст даты в формате ДД.ММ.ГГ
' =============================================
Public Sub WriteDateToCell(cell As Range, DateText As String)
    If Trim(DateText) <> "" And IsValidDateFormat(DateText) Then
        ' Преобразуем ГГ в полный год для записи
        Dim fullDateText As String
        fullDateText = Left(DateText, 6) & "20" & Right(DateText, 2)
        cell.value = CDate(fullDateText)
    Else
        cell.value = ""
    End If
End Sub

