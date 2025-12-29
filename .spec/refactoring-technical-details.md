# Технические детали рефакторинга
**Дополнение к:** refactoring-plan.md  
**Дата:** 10.01.2025

## 1. Структура CommonUtilities.bas

### Полный код модуля CommonUtilities.bas

```vba
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
' @param sheetName [String] Имя листа
' @return [Worksheet] Объект листа или Nothing если не найден
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
' @param Ws [Worksheet] Лист, на котором находится таблица
' @param tableName [String] Имя таблицы
' @return [ListObject] Объект таблицы или Nothing если не найдена
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
' @param DateText [String] Текст даты в формате ДД.ММ.ГГ
' @return [Boolean] True если формат корректен и дата не в будущем
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
' @param cell [Range] Ячейка с датой
' @return [String] Отформатированная дата или пустая строка
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
' @param cell [Range] Ячейка для записи
' @param DateText [String] Текст даты в формате ДД.ММ.ГГ
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
```

## 2. Изменения в DashboardModule.bas

### Удалить функции (строки 1446-1462):
```vba
' УДАЛИТЬ ЭТИ ФУНКЦИИ:
Private Function GetWorksheetSafe(sheetName As String) As Worksheet
    ' ... код ...
End Function

Private Function GetListObjectSafe(Ws As Worksheet, tableName As String) As ListObject
    ' ... код ...
End Function
```

### Заменить вызовы:
```vba
' БЫЛО:
Set wsDashboard = GetWorksheetSafe("Dashboard")
Set wsData = GetWorksheetSafe("ВхИсх")
Set tblData = GetListObjectSafe(wsData, "ВходящиеИсходящие")

' СТАЛО:
Set wsDashboard = CommonUtilities.GetWorksheetSafe("Dashboard")
Set wsData = CommonUtilities.GetWorksheetSafe("ВхИсх")
Set tblData = CommonUtilities.GetListObjectSafe(wsData, "ВходящиеИсходящие")
```

## 3. Изменения в DataManager.bas

### Удалить функцию (строки 186-212):
```vba
' УДАЛИТЬ:
Private Function IsValidDateFormat(DateText As String) As Boolean
    ' ... код ...
End Function
```

### Заменить вызовы:
```vba
' БЫЛО:
If Not IsValidDateFormat(.txtDataVhFRP.Text) Then

' СТАЛО:
If Not CommonUtilities.IsValidDateFormat(.txtDataVhFRP.Text) Then
```

### Обновить WriteDateToCell (строка 268):
```vba
' БЫЛО:
Private Sub WriteDateToCell(cell As Range, DateText As String)
    ' ... код ...
End Sub

' СТАЛО:
' УДАЛИТЬ эту функцию, использовать CommonUtilities.WriteDateToCell()

' БЫЛО:
Call WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 8), .txtDataVhFRP.Text)

' СТАЛО:
Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 8), .txtDataVhFRP.Text)
```

## 4. Изменения в NavigationModule.bas

### Удалить функцию (строки 265-278):
```vba
' УДАЛИТЬ:
Private Function FormatDateCell(cell As Range) As String
    ' ... код ...
End Function
```

### Заменить вызовы (если есть):
```vba
' БЫЛО:
FormatDateCell = FormatDateCell(cell)

' СТАЛО:
FormatDateCell = CommonUtilities.FormatDateCell(cell)
```

## 5. Изменения в UserFormVhIsh.frm

### Удалить функции:
```vba
' УДАЛИТЬ (строка 914-961):
Private Function ValidateDates() As Boolean
    ' ... код ...
End Function

' УДАЛИТЬ (строка 963-986):
Private Function IsValidDateFormat(DateText As String) As Boolean
    ' ... код ...
End Function

' УДАЛИТЬ (строка 1025-1032):
Private Sub WriteDateToCell(cell As Range, DateText As String)
    ' ... код ...
End Sub

' УДАЛИТЬ (строка 1107-1119):
Private Function FormatDateCell(cell As Range) As String
    ' ... код ...
End Function

' УДАЛИТЬ (строка 1142-1191):
Private Sub ClearForm()
    ' ... код ...
End Sub

' УДАЛИТЬ (строка 1236):
Private Sub UpdateStatusBar()
    ' ... код ...
End Sub
```

### Заменить вызовы:
```vba
' БЫЛО:
If Not ValidateDates() Then
    Cancel = True
    Exit Sub
End If

' СТАЛО:
If Not DataManager.ValidateDates() Then
    Cancel = True
    Exit Sub
End If

' БЫЛО:
If Not IsValidDateFormat(Me.txtDataVhFRP.Text) Then

' СТАЛО:
If Not CommonUtilities.IsValidDateFormat(Me.txtDataVhFRP.Text) Then

' БЫЛО:
Call WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 8), Me.txtDataVhFRP.Text)

' СТАЛО:
Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 8), Me.txtDataVhFRP.Text)

' БЫЛО:
Me.txtDataVhFRP.Text = FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 8))

' СТАЛО:
Me.txtDataVhFRP.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 8))

' БЫЛО:
Call ClearForm

' СТАЛО:
Call DataManager.ClearForm

' БЫЛО:
Call UpdateStatusBar

' СТАЛО:
Call NavigationModule.UpdateStatusBar
```

## 6. Изменения в Module1.bas

### Проверка использования

**Команда для поиска вызовов:**
```bash
# Найти все вызовы функций из Module1
grep -r "Module1\." AccountingIncOut.xlsm.modules/
grep -r "GenerateDashboard" AccountingIncOut.xlsm.modules/
```

**Если Module1.bas не используется:**
- Удалить файл полностью

**Если Module1.bas используется:**
- Определить, какие функции используются
- Перенести их в соответствующие модули
- Удалить Module1.bas

## 7. Карта зависимостей функций

### GetWorksheetSafe
**Используется в:**
- DashboardModule.bas (множественные вызовы)
- Module1.bas (если используется)

**Заменяется на:**
- CommonUtilities.GetWorksheetSafe()

### GetListObjectSafe
**Используется в:**
- DashboardModule.bas (множественные вызовы)
- Module1.bas (если используется)

**Заменяется на:**
- CommonUtilities.GetListObjectSafe()

### IsValidDateFormat
**Используется в:**
- DataManager.bas: ValidateDates() (строка 140, 149, 158, 167, 176)
- UserFormVhIsh.frm: ValidateDates() (строка 918, 927, 936, 945, 954)
- DataManager.bas: WriteDateToCell() (строка 269)

**Заменяется на:**
- CommonUtilities.IsValidDateFormat()

### FormatDateCell
**Используется в:**
- NavigationModule.bas: FormatDateCell() (строка 265) - не используется напрямую
- UserFormVhIsh.frm: LoadRecordToForm() (строки 1076, 1078, 1081, 1083, 1085)

**Заменяется на:**
- CommonUtilities.FormatDateCell()

### WriteDateToCell
**Используется в:**
- DataManager.bas: WriteFormDataToTable() (строки 236, 240, 245, 249, 253)
- UserFormVhIsh.frm: WriteFormDataToTable() (строки 1005, 1007, 1010, 1012, 1014)

**Заменяется на:**
- CommonUtilities.WriteDateToCell()

### ClearForm
**Используется в:**
- DataManager.bas: ClearForm() - публичная функция
- UserFormVhIsh.frm: ClearForm() - приватная функция (дубликат)
- NavigationModule.bas: NavigateToRecord() (строка 24) - вызывает DataManager.ClearForm
- NavigationModule.bas: RefreshNavigation() (строка 255) - вызывает DataManager.ClearForm

**Решение:**
- Удалить ClearForm() из UserFormVhIsh.frm
- Использовать DataManager.ClearForm() везде

### UpdateStatusBar
**Используется в:**
- NavigationModule.bas: UpdateStatusBar() - публичная функция
- UserFormVhIsh.frm: UpdateStatusBar() - приватная функция (дубликат)
- DataManager.bas: MarkFormAsChanged() (строка 281) - вызывает UpdateStatusBar
- DataManager.bas: ClearForm() (строка 354) - вызывает UpdateStatusBar
- UserFormVhIsh.frm: множественные вызовы

**Решение:**
- Удалить UpdateStatusBar() из UserFormVhIsh.frm
- Использовать NavigationModule.UpdateStatusBar() везде

### ValidateDates
**Используется в:**
- DataManager.bas: ValidateDates() - публичная функция
- DataManager.bas: SaveCurrentRecord() (строка 28) - вызывает ValidateDates
- UserFormVhIsh.frm: ValidateDates() - приватная функция (дубликат)
- UserFormVhIsh.frm: btnSave_Click() (строка 791) - вызывает ValidateDates

**Решение:**
- Удалить ValidateDates() из UserFormVhIsh.frm
- Использовать DataManager.ValidateDates() везде

## 8. Последовательность замены

### Рекомендуемый порядок:

1. **Создать CommonUtilities.bas** (новый файл)
2. **Заменить GetWorksheetSafe/GetListObjectSafe в DashboardModule**
3. **Заменить IsValidDateFormat в DataManager**
4. **Заменить WriteDateToCell в DataManager**
5. **Заменить FormatDateCell в UserFormVhIsh**
6. **Удалить дубликаты из UserFormVhIsh**
7. **Заменить вызовы ClearForm/UpdateStatusBar/ValidateDates**
8. **Удалить функции из Module1.bas (если не используются)**
9. **Тестирование после каждого шага**

## 9. Чек-лист проверки после рефакторинга

### Компиляция
- [ ] Проект компилируется без ошибок
- [ ] Нет предупреждений о неиспользуемых переменных
- [ ] Все функции имеют корректные сигнатуры

### Функциональность
- [ ] Dashboard генерируется корректно
- [ ] Форма UserFormVhIsh открывается
- [ ] Создание новой записи работает
- [ ] Редактирование записи работает
- [ ] Валидация дат работает
- [ ] Навигация по записям работает
- [ ] Сохранение данных работает
- [ ] Поиск работает

### Код
- [ ] Нет дублирующихся функций
- [ ] Все вызовы обновлены
- [ ] Комментарии актуальны
- [ ] Код соответствует стандартам проекта

---

**Примечание:** Этот документ содержит конкретные примеры кода для каждого шага рефакторинга. Используйте его вместе с refactoring-plan.md для пошагового выполнения.

