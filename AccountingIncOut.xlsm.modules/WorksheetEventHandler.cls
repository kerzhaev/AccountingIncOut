VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================
' КЛАСС ОБРАБОТКИ СОБЫТИЙ ЛИСТА - Class Module WorksheetEventHandler
' Назначение: Класс для обработки событий листа с таблицей
' Тип: Class Module
' Состояние: УДАЛЕНЫ ВСЕ ДУБЛИРУЮЩИЕ ФУНКЦИИ - ТОЛЬКО УНИКАЛЬНЫЙ КОД
' Версия: 1.3.0
' Дата: 10.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' Объявление переменной с событиями (только в классе!)
Private WithEvents wsData As Worksheet
Attribute wsData.VB_VarHelpID = -1
Private tblVhIsh As ListObject

' Свойство для установки листа
Public Property Set Worksheet(Ws As Worksheet)
    Set wsData = Ws
    On Error Resume Next
    Set tblVhIsh = Ws.ListObjects("ВходящиеИсходящие")
    On Error GoTo 0
End Property

' Обработчик события изменения выделения
Private Sub wsData_SelectionChange(ByVal Target As Range)
    Dim intersectRange As Range
    Dim RowNumber As Long
    Dim columnNumber As Long
    Dim fieldToActivate As String
    
    On Error GoTo EventError
    
    ' Проверяем, что tblVhIsh инициализирован
    If tblVhIsh Is Nothing Then Exit Sub
    
    ' Проверяем, что DataBodyRange существует
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = "Таблица пуста. Добавьте данные для работы с системой."
        Exit Sub
    End If
    
    ' Проверяем, что щелчок произошел в области таблицы (исключая заголовки)
    Set intersectRange = Intersect(Target, tblVhIsh.DataBodyRange)
    
    If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
        ' Получаем номер строки в таблице
        RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
        
        ' Получаем номер столбца в таблице
        columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
        
        ' ИСПРАВЛЕНО: Используем функцию из TableEventHandler
        fieldToActivate = TableEventHandler.GetFieldNameByColumn(columnNumber)
        
        ' Показываем информацию в статус-баре
        Application.StatusBar = "Строка " & RowNumber & ", столбец " & columnNumber & _
                               " | Ctrl+E для редактирования | " & _
                               TableEventHandler.GetFieldDisplayName(fieldToActivate)
    Else
        ' Сбрасываем если вышли из таблицы
        Application.StatusBar = "Система интерактивных форм активна. Выберите ячейку таблицы."
    End If
    
    Exit Sub
    
EventError:
    ' Игнорируем ошибки обработки события
    Debug.Print "Ошибка в WorksheetEventHandler.wsData_SelectionChange: " & Err.description
End Sub

' ИСПРАВЛЕНО: Публичный метод для открытия формы через TableEventHandler
Public Sub OpenFormForSelection()
    On Error GoTo OpenError
    
    If wsData Is Nothing Or tblVhIsh Is Nothing Then Exit Sub
    
    ' Проверяем, что DataBodyRange существует
    If tblVhIsh.DataBodyRange Is Nothing Then
        MsgBox "Таблица пуста! Добавьте данные для работы с формой.", vbExclamation, "Пустая таблица"
        Exit Sub
    End If
    
    ' ИСПРАВЛЕНО: Используем метод из TableEventHandler вместо дублирования
    Call TableEventHandler.OpenFormForCurrentSelection
    
    Exit Sub
    
OpenError:
    MsgBox "Ошибка открытия формы: " & Err.description, vbExclamation, "Ошибка"
End Sub

' Публичный метод для получения информации о классе
Public Function GetClassInfo() As String
    GetClassInfo = "WorksheetEventHandler | " & _
                  "Лист: " & IIf(wsData Is Nothing, "НЕ УСТАНОВЛЕН", wsData.Name) & " | " & _
                  "Таблица: " & IIf(tblVhIsh Is Nothing, "НЕ НАЙДЕНА", tblVhIsh.Name) & " | " & _
                  "Строк: " & IIf(tblVhIsh Is Nothing Or tblVhIsh.DataBodyRange Is Nothing, 0, tblVhIsh.ListRows.Count)
End Function

' Метод для безопасной очистки объектов
Public Sub Cleanup()
    On Error Resume Next
    Set wsData = Nothing
    Set tblVhIsh = Nothing
    On Error GoTo 0
End Sub

' Публичный метод для программного вызова события SelectionChange
Public Sub TriggerSelectionChange(Target As Range)
    Call wsData_SelectionChange(Target)
End Sub

' Метод проверки активности класса
Public Function IsActive() As Boolean
    IsActive = (Not wsData Is Nothing) And (Not tblVhIsh Is Nothing)
End Function

' Метод получения текущего листа
Public Property Get CurrentWorksheet() As Worksheet
    Set CurrentWorksheet = wsData
End Property

' Метод получения текущей таблицы
Public Property Get CurrentTable() As ListObject
    Set CurrentTable = tblVhIsh
End Property


