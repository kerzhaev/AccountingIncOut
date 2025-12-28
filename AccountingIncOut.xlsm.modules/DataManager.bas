Attribute VB_Name = "DataManager"
'==============================================
' МОДУЛЬ УПРАВЛЕНИЯ ДАННЫМИ ТАБЛИЦЫ - DataManager
' Назначение: Функции для работы с таблицей ВходящиеИсходящие
' Состояние: ДОБАВЛЕНА ФУНКЦИЯ ДУБЛИРОВАНИЯ ЗАПИСЕЙ С ИСКЛЮЧЕНИЯМИ ПОЛЕЙ
' Версия: 2.0.0
' Дата: 09.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

Public CurrentRecordRow As Long
Public IsNewRecord As Boolean
Public FormDataChanged As Boolean

Public Sub SaveCurrentRecord()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim newRow As ListRow
    
    ' Валидация обязательных полей
    If Not ValidateRequiredFields() Then
        Exit Sub
    End If
    
    ' Валидация дат
    If Not ValidateDates() Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Правильное название листа
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    If IsNewRecord Then
        ' Добавление новой записи
        Set newRow = tblData.ListRows.Add
        CurrentRecordRow = newRow.Index
        
        ' Автозаполнение номера П/П
        UserFormVhIsh.txtNomerPP.Text = CStr(tblData.ListRows.Count)
    End If
    
    ' Сохранение данных в таблицу
    Call WriteFormDataToTable(tblData, CurrentRecordRow)
    
    ' Обновление статуса
    IsNewRecord = False
    FormDataChanged = False
    UserFormVhIsh.lblStatusBar.Caption = "Запись сохранена успешно"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при сохранении данных: " & Err.description, vbCritical, "Ошибка сохранения"
    UserFormVhIsh.lblStatusBar.Caption = "Ошибка сохранения данных"
End Sub

Public Function ValidateRequiredFields() As Boolean
    ValidateRequiredFields = True
    
    ' Проверка обязательных полей
    With UserFormVhIsh
        ' Используем Value для ComboBox
        ' Служба
        If Trim(CStr(.cmbSlujba.value)) = "" Then
            MsgBox "Поле 'Служба' обязательно для заполнения!", vbExclamation, "Проверка данных"
            .cmbSlujba.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Вид документа (ComboBox) - 4-й столбец
        If Trim(CStr(.cmbVidDoc.value)) = "" Then
            MsgBox "Поле 'Тип документа' обязательно для заполнения!", vbExclamation, "Проверка данных"
            .cmbVidDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Вид документа (Вх./Исх.) - 3-й столбец - cmbVidDocumenta
        If Trim(CStr(.cmbVidDocumenta.value)) = "" Then
            MsgBox "Поле 'Вид документа (Вх./Исх.)' обязательно для заполнения!", vbExclamation, "Проверка данных"
            .cmbVidDocumenta.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Номер документа
        If Trim(.txtNomerDoc.Text) = "" Then
            MsgBox "Поле 'Номер документа' обязательно для заполнения!", vbExclamation, "Проверка данных"
            .txtNomerDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Сумма документа
        If Trim(.txtSummaDoc.Text) = "" Then
            MsgBox "Поле 'Сумма документа' обязательно для заполнения!", vbExclamation, "Проверка данных"
            .txtSummaDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Проверка, что сумма - число
        If Not IsNumeric(.txtSummaDoc.Text) Then
            MsgBox "Поле 'Сумма документа' должно содержать числовое значение!", vbExclamation, "Проверка данных"
            .txtSummaDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Вх.ФРП/Исх.ФРП
        If Trim(.txtVhFRP.Text) = "" Then
            MsgBox "Поле 'Вх.ФРП/Исх.ФРП' обязательно для заполнения!", vbExclamation, "Проверка данных"
            .txtVhFRP.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Дата Вх.ФРП/Исх.ФРП
        If Trim(.txtDataVhFRP.Text) = "" Then
            MsgBox "Поле 'Дата Вх.ФРП/Исх.ФРП' обязательно для заполнения!", vbExclamation, "Проверка данных"
            .txtDataVhFRP.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
    End With
End Function

Public Function ValidateDates() As Boolean
    ValidateDates = True
    
    With UserFormVhIsh
        ' Проверка всех полей дат
        If Trim(.txtDataVhFRP.Text) <> "" Then
            If Not IsValidDateFormat(.txtDataVhFRP.Text) Then
                MsgBox "Введите корректную дату в формате ДД.ММ.ГГ в поле 'Дата Вх.ФРП/Исх.ФРП'", vbExclamation, "Проверка данных"
                .txtDataVhFRP.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDataPeredachi.Text) <> "" Then
            If Not IsValidDateFormat(.txtDataPeredachi.Text) Then
                MsgBox "Введите корректную дату в формате ДД.ММ.ГГ в поле 'Дата передачи исполнителю'", vbExclamation, "Проверка данных"
                .txtDataPeredachi.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDataIshVSlujbu.Text) <> "" Then
            If Not IsValidDateFormat(.txtDataIshVSlujbu.Text) Then
                MsgBox "Введите корректную дату в формате ДД.ММ.ГГ в поле 'Дата исх. в службу'", vbExclamation, "Проверка данных"
                .txtDataIshVSlujbu.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDataVozvrata.Text) <> "" Then
            If Not IsValidDateFormat(.txtDataVozvrata.Text) Then
                MsgBox "Введите корректную дату в формате ДД.ММ.ГГ в поле 'Дата возврата со службы'", vbExclamation, "Проверка данных"
                .txtDataVozvrata.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDataIshKonvert.Text) <> "" Then
            If Not IsValidDateFormat(.txtDataIshKonvert.Text) Then
                MsgBox "Введите корректную дату в формате ДД.ММ.ГГ в поле 'Дата исх. конверт'", vbExclamation, "Проверка данных"
                .txtDataIshKonvert.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
    End With
End Function

Private Function IsValidDateFormat(DateText As String) As Boolean
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
            MsgBox "Дата не может быть позднее текущей даты!", vbExclamation, "Проверка данных"
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

Public Sub WriteFormDataToTable(tbl As ListObject, RowIndex As Long)
    On Error GoTo WriteError
    
    With UserFormVhIsh
        tbl.DataBodyRange.Cells(RowIndex, 1).value = RowIndex ' Номер записи = номер строки
        tbl.DataBodyRange.Cells(RowIndex, 2).value = .cmbSlujba.value
        ' cmbVidDocumenta идет в 3-й столбец
        tbl.DataBodyRange.Cells(RowIndex, 3).value = .cmbVidDocumenta.value
        ' cmbVidDoc идет в 4-й столбец
        tbl.DataBodyRange.Cells(RowIndex, 4).value = .cmbVidDoc.value
        tbl.DataBodyRange.Cells(RowIndex, 5).value = .txtNomerDoc.Text
        
        ' Безопасное преобразование суммы
        If IsNumeric(.txtSummaDoc.Text) Then
            tbl.DataBodyRange.Cells(RowIndex, 6).value = CDbl(.txtSummaDoc.Text)
        Else
            tbl.DataBodyRange.Cells(RowIndex, 6).value = 0
        End If
        
        tbl.DataBodyRange.Cells(RowIndex, 7).value = .txtVhFRP.Text
        
        ' Безопасная запись даты
        Call WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 8), .txtDataVhFRP.Text)
        
        tbl.DataBodyRange.Cells(RowIndex, 9).value = .cmbOtKogoPostupil.value
        
        Call WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 10), .txtDataPeredachi.Text)
        
        tbl.DataBodyRange.Cells(RowIndex, 11).value = .cmbIspolnitel.value
        tbl.DataBodyRange.Cells(RowIndex, 12).value = .txtNomerIshVSlujbu.Text
        
        Call WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 13), .txtDataIshVSlujbu.Text)
        
        tbl.DataBodyRange.Cells(RowIndex, 14).value = .txtNomerVozvrata.Text
        
        Call WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 15), .txtDataVozvrata.Text)
        
        tbl.DataBodyRange.Cells(RowIndex, 16).value = .txtNomerIshKonvert.Text
        
        Call WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 17), .txtDataIshKonvert.Text)
        
        tbl.DataBodyRange.Cells(RowIndex, 18).value = .txtOtmetkaIspolnenie.Text
        tbl.DataBodyRange.Cells(RowIndex, 19).value = .cmbStatusPodtverjdenie.value
        
        ' Информация о наряде - 20-й столбец
        tbl.DataBodyRange.Cells(RowIndex, 20).value = .txtNaryadInfo.Text
    End With
    
    Exit Sub
    
WriteError:
    MsgBox "Ошибка записи данных в таблицу: " & Err.description, vbCritical, "Ошибка"
End Sub

Private Sub WriteDateToCell(cell As Range, DateText As String)
    If Trim(DateText) <> "" And IsValidDateFormat(DateText) Then
        ' Преобразуем ГГ в полный год для записи
        Dim fullDateText As String
        fullDateText = Left(DateText, 6) & "20" & Right(DateText, 2)
        cell.value = CDate(fullDateText)
    Else
        cell.value = ""
    End If
End Sub

Public Sub MarkFormAsChanged()
    FormDataChanged = True
    Call UpdateStatusBar
End Sub

Public Function HasUnsavedChanges() As Boolean
    HasUnsavedChanges = FormDataChanged
End Function

Public Sub CancelChanges()
    If IsNewRecord Then
        Call ClearForm
    Else
        Call NavigateToRecord(CurrentRecordRow)
    End If
    
    FormDataChanged = False
    UserFormVhIsh.lblStatusBar.Caption = "Изменения отменены"
End Sub

Public Sub ClearForm()
    IsNewRecord = True
    CurrentRecordRow = 0
    FormDataChanged = False
    
    ' Устанавливаем следующий номер П/П для новой записи
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim nextNumber As Long
    
    On Error Resume Next
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Правильное название листа
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    If Not wsData Is Nothing And Not tblData Is Nothing Then
        nextNumber = tblData.ListRows.Count + 1
    Else
        nextNumber = 1
    End If
    On Error GoTo 0
    
    With UserFormVhIsh
        ' Очистка всех полей
        .txtNomerPP.Text = CStr(nextNumber) ' Устанавливаем следующий номер
        .cmbVidDocumenta.value = ""
        .txtNomerDoc.Text = ""
        .txtSummaDoc.Text = ""
        .txtVhFRP.Text = ""
        .txtDataVhFRP.Text = ""
        .cmbOtKogoPostupil.value = ""
        .txtDataPeredachi.Text = ""
        .txtNomerIshVSlujbu.Text = ""
        .txtDataIshVSlujbu.Text = ""
        .txtNomerVozvrata.Text = ""
        .txtDataVozvrata.Text = ""
        .txtNomerIshKonvert.Text = ""
        .txtDataIshKonvert.Text = ""
        .txtOtmetkaIspolnenie.Text = ""
        .cmbStatusPodtverjdenie.listIndex = 0 ' Устанавливаем пустое значение
        
        ' Очистка остальных комбобоксов
        .cmbSlujba.value = ""
        .cmbVidDoc.value = ""
        .cmbIspolnitel.value = ""
        
        ' Очистка поля наряда
        .txtNaryadInfo.Text = ""
        
        ' Очистка поиска
        .txtSearch.Text = ""
        .lstSearchResults.Clear
        .lstSearchResults.Visible = False
    End With
    
    Call UpdateStatusBar
End Sub

' НОВАЯ ФУНКЦИЯ: Дублирование записи с исключением определенных полей
Public Function DuplicateRecord(sourceRowNumber As Long) As Long
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim newRow As ListRow
    Dim newRowIndex As Long
    Dim i As Long
    
    On Error GoTo DuplicateError
    
    ' Получаем ссылки на таблицу
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    ' Проверяем корректность исходной строки
    If sourceRowNumber < 1 Or sourceRowNumber > tblData.ListRows.Count Then
        MsgBox "Неверный номер записи для дублирования: " & sourceRowNumber, vbExclamation, "Ошибка"
        DuplicateRecord = 0
        Exit Function
    End If
    
    ' Добавляем новую строку
    Set newRow = tblData.ListRows.Add
    newRowIndex = newRow.Index
    
    ' Копируем данные из исходной строки, исключая определенные поля
    For i = 1 To tblData.ListColumns.Count
        Select Case i
            Case 1  ' № П/П - устанавливаем новый номер
                tblData.DataBodyRange.Cells(newRowIndex, i).value = newRowIndex
                
            Case 5  ' txtNomerDoc - НЕ копируем (номер документа)
                tblData.DataBodyRange.Cells(newRowIndex, i).value = ""
                
            Case 6  ' txtSummaDoc - НЕ копируем (сумма документа)
                tblData.DataBodyRange.Cells(newRowIndex, i).value = 0
                
            Case 20 ' txtNaryadInfo - НЕ копируем (информация о наряде)
                tblData.DataBodyRange.Cells(newRowIndex, i).value = ""
                
            Case Else ' Все остальные поля - копируем
                tblData.DataBodyRange.Cells(newRowIndex, i).value = _
                    tblData.DataBodyRange.Cells(sourceRowNumber, i).value
        End Select
    Next i
    
    ' Возвращаем номер новой записи
    DuplicateRecord = newRowIndex
    
    Exit Function
    
DuplicateError:
    MsgBox "Ошибка дублирования записи: " & Err.description, vbCritical, "Критическая ошибка"
    DuplicateRecord = 0
    
    ' Пытаемся удалить созданную строку при ошибке
    On Error Resume Next
    If Not newRow Is Nothing Then
        newRow.Delete
    End If
    On Error GoTo 0
End Function

' НОВАЯ ФУНКЦИЯ: Получение информации о записи для дублирования
Public Function GetRecordInfo(RowNumber As Long) As String
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim recordInfo As String
    
    On Error GoTo InfoError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    If RowNumber < 1 Or RowNumber > tblData.ListRows.Count Then
        GetRecordInfo = "Неверный номер записи"
        Exit Function
    End If
    
    ' Формируем информационную строку
    recordInfo = "Запись №" & RowNumber & ": "
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 2).value) & " - " ' Служба
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 3).value) & " " ' Вид документа
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 4).value) & " " ' Тип документа
    recordInfo = recordInfo & "№" & CStr(tblData.DataBodyRange.Cells(RowNumber, 5).value) ' Номер документа
    
    GetRecordInfo = recordInfo
    
    Exit Function
    
InfoError:
    GetRecordInfo = "Ошибка получения информации о записи"
End Function

Public Sub SetupGroupBoxes()
    ' Функция-заглушка для настройки группировки
    ' Будет реализована при создании интерфейса формы
End Sub

' НОВАЯ ФУНКЦИЯ: Проверка возможности дублирования записи
Public Function CanDuplicateRecord(RowNumber As Long) As Boolean
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo CannotDuplicate
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    ' Проверяем существование таблицы и записи
    If tblData Is Nothing Then
        CanDuplicateRecord = False
        Exit Function
    End If
    
    If tblData.DataBodyRange Is Nothing Then
        CanDuplicateRecord = False
        Exit Function
    End If
    
    If RowNumber < 1 Or RowNumber > tblData.ListRows.Count Then
        CanDuplicateRecord = False
        Exit Function
    End If
    
    CanDuplicateRecord = True
    Exit Function
    
CannotDuplicate:
    CanDuplicateRecord = False
End Function


