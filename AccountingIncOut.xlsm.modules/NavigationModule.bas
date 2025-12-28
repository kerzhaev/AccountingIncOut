Attribute VB_Name = "NavigationModule"
'==============================================
' МОДУЛЬ НАВИГАЦИИ ПО ЗАПИСЯМ NavigationModule
' Назначение: Функции для перемещения между записями таблицы
' Состояние: ИСПРАВЛЕНЫ ДУБЛИРУЮЩИЕ ФУНКЦИИ И ОБЪЯВЛЕНИЯ ПЕРЕМЕННЫХ
' Версия: 1.6.1
' Дата: 10.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

Public Sub NavigateToRecord(RowNumber As Long)
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo NavigationError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    ' Проверка границ
    If tblData.ListRows.Count = 0 Then
        Call DataManager.ClearForm
        Exit Sub
    End If
    
    If RowNumber < 1 Then RowNumber = 1
    If RowNumber > tblData.ListRows.Count Then RowNumber = tblData.ListRows.Count
    
    ' Проверка несохранённых изменений
    If DataManager.HasUnsavedChanges() And DataManager.CurrentRecordRow <> RowNumber And DataManager.CurrentRecordRow > 0 Then
        Dim Response As VbMsgBoxResult
        Response = MsgBox("Сохранить изменения в текущей записи перед переходом?", vbYesNo + vbQuestion, "Несохранённые изменения")
        If Response = vbYes Then
            Call DataManager.SaveCurrentRecord
        End If
    End If
    
    DataManager.CurrentRecordRow = RowNumber
    DataManager.IsNewRecord = False
    
    ' Вызываем процедуру из формы
    Call UserFormVhIsh.LoadRecordToForm(RowNumber)
    Call UpdateNavigationButtons
    Call UpdateStatusBar
    
    DataManager.FormDataChanged = False
    
    Exit Sub
    
NavigationError:
    MsgBox "Ошибка навигации: " & Err.description, vbCritical, "Ошибка"
End Sub

Public Sub NavigateToPrevious()
    Call NavigateToRecord(DataManager.CurrentRecordRow - 1)
End Sub

Public Sub NavigateToNext()
    Call NavigateToRecord(DataManager.CurrentRecordRow + 1)
End Sub

Public Sub NavigateToFirst()
    Call NavigateToRecord(1)
End Sub

Public Sub NavigateToLast()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo LastError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    Call NavigateToRecord(tblData.ListRows.Count)
    
    Exit Sub
    
LastError:
    Call NavigateToRecord(1)
End Sub

Public Sub NavigateToRecordByNumber(recordNumber As Long)
    ' Навигация к записи по номеру П/П
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    
    On Error GoTo FindError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    ' Поиск записи с указанным номером П/П
    For i = 1 To tblData.ListRows.Count
        If CLng(tblData.DataBodyRange.Cells(i, 1).value) = recordNumber Then
            Call NavigateToRecord(i)
            Exit Sub
        End If
    Next i
    
    MsgBox "Запись с номером П/П " & recordNumber & " не найдена!", vbExclamation, "Поиск записи"
    
    Exit Sub
    
FindError:
    MsgBox "Ошибка поиска записи: " & Err.description, vbExclamation, "Ошибка"
End Sub

Private Sub UpdateNavigationButtons()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo NavigationError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    With UserFormVhIsh
        .btnFirst.Enabled = (DataManager.CurrentRecordRow > 1)
        .btnPrevious.Enabled = (DataManager.CurrentRecordRow > 1)
        .btnNext.Enabled = (DataManager.CurrentRecordRow < tblData.ListRows.Count)
        .btnLast.Enabled = (DataManager.CurrentRecordRow < tblData.ListRows.Count)
        
        ' Обновляем статус-бар с информацией о навигации
        If tblData.ListRows.Count > 0 Then
            Dim navInfo As String
            navInfo = "Навигация: "
            
            If DataManager.CurrentRecordRow > 1 Then
                navInfo = navInfo & "Пред.(" & (DataManager.CurrentRecordRow - 1) & ") "
            End If
            
            navInfo = navInfo & "Тек.(" & DataManager.CurrentRecordRow & ") "
            
            If DataManager.CurrentRecordRow < tblData.ListRows.Count Then
                navInfo = navInfo & "След.(" & (DataManager.CurrentRecordRow + 1) & ") "
            End If
            
            navInfo = navInfo & "Посл.(" & tblData.ListRows.Count & ")"
            
            ' Временно показываем информацию в статус-баре
            .lblStatusBar.Caption = navInfo
        End If
    End With
    
    Exit Sub
    
NavigationError:
    ' При ошибке деактивируем все кнопки навигации
    With UserFormVhIsh
        .btnFirst.Enabled = False
        .btnPrevious.Enabled = False
        .btnNext.Enabled = False
        .btnLast.Enabled = False
    End With
End Sub

Public Sub UpdateStatusBar()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim statusText As String
    
    On Error GoTo StatusError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    If DataManager.IsNewRecord Then
        statusText = "Новая запись"
    ElseIf tblData.ListRows.Count = 0 Then
        statusText = "Нет записей в таблице"
    Else
        statusText = "Запись " & DataManager.CurrentRecordRow & " из " & tblData.ListRows.Count
        
        ' Добавляем информацию о служебных данных текущей записи
        If DataManager.CurrentRecordRow > 0 And DataManager.CurrentRecordRow <= tblData.ListRows.Count Then
            Dim serviceName As String
            Dim docNumber As String
            serviceName = CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 2).value)
            docNumber = CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 5).value)
            
            If Trim(serviceName) <> "" And Trim(docNumber) <> "" Then
                statusText = statusText & " | " & serviceName & " | Док.№" & docNumber
            End If
        End If
    End If
    
    If DataManager.FormDataChanged Then
        statusText = statusText & " (изменено)"
    End If
    
    UserFormVhIsh.lblStatusBar.Caption = statusText
    
    Exit Sub
    
StatusError:
    UserFormVhIsh.lblStatusBar.Caption = "Ошибка обновления статуса"
End Sub

Public Function GetCurrentRecordInfo() As String
    ' Возвращает информацию о текущей записи
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim recordInfo As String
    
    On Error GoTo InfoError
    
    If DataManager.IsNewRecord Then
        GetCurrentRecordInfo = "Новая запись"
        Exit Function
    End If
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    If DataManager.CurrentRecordRow > 0 And DataManager.CurrentRecordRow <= tblData.ListRows.Count Then
        recordInfo = "Запись №" & DataManager.CurrentRecordRow & ": " & _
                     CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 2).value) & " - " & _
                     CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 4).value) & " №" & _
                     CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 5).value)
    Else
        recordInfo = "Некорректная запись"
    End If
    
    GetCurrentRecordInfo = recordInfo
    
    Exit Function
    
InfoError:
    GetCurrentRecordInfo = "Ошибка получения информации о записи"
End Function

Public Sub RefreshNavigation()
    ' Обновление навигации после изменений в таблице
    Call UpdateNavigationButtons
    Call UpdateStatusBar
    
    ' Проверяем, что текущая запись все еще существует
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo RefreshError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    If DataManager.CurrentRecordRow > tblData.ListRows.Count Then
        ' Если текущая запись была удалена, переходим к последней
        If tblData.ListRows.Count > 0 Then
            Call NavigateToRecord(tblData.ListRows.Count)
        Else
            Call DataManager.ClearForm
        End If
    End If
    
    Exit Sub
    
RefreshError:
    Call DataManager.ClearForm
End Sub

Private Function FormatDateCell(cell As Range) As String
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

Public Function CanNavigate() As Boolean
    ' Проверка возможности навигации
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo CannotNavigate
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    CanNavigate = (tblData.ListRows.Count > 0)
    
    Exit Function
    
CannotNavigate:
    CanNavigate = False
End Function

Public Sub ShowNavigationSummary()
    ' Показывает сводку по навигации
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim summaryText As String
    
    On Error GoTo SummaryError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    summaryText = "Сводка по записям:" & vbCrLf & _
                  "Всего записей: " & tblData.ListRows.Count & vbCrLf
    
    If tblData.ListRows.Count > 0 Then
        summaryText = summaryText & _
                      "Текущая запись: " & DataManager.CurrentRecordRow & vbCrLf & _
                      "Первая запись: 1" & vbCrLf & _
                      "Последняя запись: " & tblData.ListRows.Count & vbCrLf
        
        If DataManager.IsNewRecord Then
            summaryText = summaryText & "Статус: Новая запись"
        Else
            summaryText = summaryText & "Статус: Существующая запись"
            If DataManager.FormDataChanged Then
                summaryText = summaryText & " (изменена)"
            End If
        End If
    Else
        summaryText = summaryText & "Статус: Таблица пустая"
    End If
    
    MsgBox summaryText, vbInformation, "Сводка навигации"
    
    Exit Sub
    
SummaryError:
    MsgBox "Ошибка получения сводки навигации", vbExclamation, "Ошибка"
End Sub

Public Sub ShowNavigationHelp()
    ' Показывает справку по навигации
    Dim helpText As String
    
    helpText = "Справка по навигации:" & vbCrLf & vbCrLf & _
               "|< Первая - переход к первой записи" & vbCrLf & _
               "< Пред. - переход к предыдущей записи" & vbCrLf & _
               "След. > - переход к следующей записи" & vbCrLf & _
               "Послед. >| - переход к последней записи" & vbCrLf & vbCrLf & _
               "Горячие клавиши:" & vbCrLf & _
               "Ctrl+S - сохранить запись" & vbCrLf & _
               "Ctrl+N - новая запись" & vbCrLf & _
               "Ctrl+F - поиск" & vbCrLf & _
               "F3 - следующий результат поиска" & vbCrLf & _
               "Esc - отмена изменений"
    
    MsgBox helpText, vbInformation, "Справка по навигации"
End Sub


