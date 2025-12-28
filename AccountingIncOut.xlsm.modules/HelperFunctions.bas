Attribute VB_Name = "HelperFunctions"
'==============================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ HelperFunctions
' Назначение: Общие функции для работы формы
' Состояние: УДАЛЕНЫ ДУБЛИРУЮЩИЕ ФУНКЦИИ - ОСТАВЛЕНЫ ТОЛЬКО УНИКАЛЬНЫЕ
' Версия: 1.4.1
' Дата: 10.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' УДАЛЕНЫ ДУБЛИРУЮЩИЕ ФУНКЦИИ CancelChanges() и MarkFormAsChanged()
' Эти функции остались только в модуле DataManager

Public Sub SetupFormAppearance()
    Call SetupRequiredFieldsHighlight
    Call SetupFieldAlignment
    Call SetupNavigationButtons
    Call SetupStatusBar
End Sub

Private Sub SetupRequiredFieldsHighlight()
    ' Подсветка обязательных полей светло-желтым цветом
    Dim requiredColor As Long
    requiredColor = RGB(255, 255, 224)
    
    With UserFormVhIsh
        .cmbSlujba.BackColor = requiredColor
        .cmbVidDoc.BackColor = requiredColor
        .cmbVidDocumenta.BackColor = requiredColor
        .txtNomerDoc.BackColor = requiredColor
        .txtSummaDoc.BackColor = requiredColor
        .txtVhFRP.BackColor = requiredColor
        .txtDataVhFRP.BackColor = requiredColor
        
        ' Поле номера П/П остается серым (недоступно для редактирования)
        .txtNomerPP.BackColor = RGB(240, 240, 240)
    End With
End Sub

Private Sub SetupFieldAlignment()
    With UserFormVhIsh
        ' Выравнивание числовых полей
        .txtSummaDoc.TextAlign = fmTextAlignRight
        .txtNomerPP.TextAlign = fmTextAlignCenter
        .txtNomerDoc.TextAlign = fmTextAlignCenter
        .txtVhFRP.TextAlign = fmTextAlignCenter
        
        ' Ограничение длины для полей дат
        .txtDataVhFRP.MaxLength = 8
        .txtDataPeredachi.MaxLength = 8
        .txtDataIshVSlujbu.MaxLength = 8
        .txtDataVozvrata.MaxLength = 8
        .txtDataIshKonvert.MaxLength = 8
    End With
End Sub

Private Sub SetupNavigationButtons()
    With UserFormVhIsh
        ' Улучшаем подписи кнопок навигации
        .btnFirst.Caption = "|< Первая"
        .btnPrevious.Caption = "< Пред."
        .btnNext.Caption = "След. >"
        .btnLast.Caption = "Послед. >|"
        
        ' Делаем кнопки чуть больше
        .btnFirst.Width = 70
        .btnPrevious.Width = 70
        .btnNext.Width = 70
        .btnLast.Width = 70
    End With
End Sub

Private Sub SetupStatusBar()
    With UserFormVhIsh.lblStatusBar
        .BackColor = RGB(245, 245, 245)
        .borderStyle = fmBorderStyleSingle
        .Font.Size = 9
        .Font.Name = "Segoe UI"
    End With
End Sub

Public Sub ShowFieldTooltip(fieldName As String)
    ' Подсказки для полей
    Dim tooltip As String
    
    Select Case fieldName
        Case "txtSummaDoc"
            tooltip = "Введите сумму документа в рублях (только цифры)"
        Case "txtDataVhFRP"
            tooltip = "Формат даты: ДД.ММ.ГГ (например: 15.07.25)"
        Case "cmbSlujba"
            tooltip = "Выберите службу из списка или введите новую"
        Case "txtVhFRP"
            tooltip = "Номер входящего/исходящего документа ФРП"
        Case "cmbOtKogoPostupil"
            tooltip = "От кого поступил документ или куда направлен"
        Case "cmbIspolnitel"
            tooltip = "Выберите исполнителя из списка или введите нового"
        Case Else
            tooltip = "Поле для ввода данных"
    End Select
    
    UserFormVhIsh.lblStatusBar.Caption = tooltip
End Sub

Public Sub SetupGroupBoxes()
    ' Функция для настройки группировки полей
    ' Группировка полей по логическим блокам для улучшения UX
    
    ' Основная информация документа (поля 1-6)
    ' Данные ФРП (поля 7-8)
    ' Движение документа (поля 9-11)
    ' Работа со службой (поля 12-15)
    ' Конвертация (поля 16-17)
    ' Статусы (поля 18-19)
    
    ' Эта функция может быть расширена при необходимости группировки элементов формы
End Sub

Public Sub ValidateFormData()
    ' Дополнительная валидация данных перед сохранением
    Dim IsValid As Boolean
    IsValid = True
    
    With UserFormVhIsh
        ' Проверка логической целостности данных
        
        ' Если есть дата передачи исполнителю, должен быть указан исполнитель
        If Trim(.txtDataPeredachi.Text) <> "" And Trim(.cmbIspolnitel.Text) = "" Then
            MsgBox "Если указана дата передачи исполнителю, необходимо выбрать исполнителя!", vbExclamation, "Проверка данных"
            .cmbIspolnitel.SetFocus
            IsValid = False
        End If
        
        ' Если есть номер исходящего в службу, должна быть дата
        If Trim(.txtNomerIshVSlujbu.Text) <> "" And Trim(.txtDataIshVSlujbu.Text) = "" Then
            MsgBox "Если указан номер исходящего в службу, необходимо указать дату!", vbExclamation, "Проверка данных"
            .txtDataIshVSlujbu.SetFocus
            IsValid = False
        End If
        
        ' Если есть номер возврата, должна быть дата
        If Trim(.txtNomerVozvrata.Text) <> "" And Trim(.txtDataVozvrata.Text) = "" Then
            MsgBox "Если указан номер возврата, необходимо указать дату возврата!", vbExclamation, "Проверка данных"
            .txtDataVozvrata.SetFocus
            IsValid = False
        End If
        
        ' Если есть номер исходящего конверта, должна быть дата
        If Trim(.txtNomerIshKonvert.Text) <> "" And Trim(.txtDataIshKonvert.Text) = "" Then
            MsgBox "Если указан номер исходящего конверта, необходимо указать дату!", vbExclamation, "Проверка данных"
            .txtDataIshKonvert.SetFocus
            IsValid = False
        End If
    End With
    
    If Not IsValid Then
        DataManager.FormDataChanged = True
    End If
End Sub

Public Sub FormatCurrencyField()
    ' Форматирование поля суммы как валюты
    With UserFormVhIsh.txtSummaDoc
        If IsNumeric(.Text) And Trim(.Text) <> "" Then
            Dim amount As Double
            amount = CDbl(.Text)
            .Text = Format(amount, "#,##0.00")
        End If
    End With
End Sub

Public Sub ClearSearchResults()
    ' Очистка результатов поиска
    With UserFormVhIsh
        .txtSearch.Text = ""
        .lstSearchResults.Clear
        .lstSearchResults.Visible = False
        .lblStatusBar.Caption = "Поиск очищен"
    End With
End Sub

Public Function GetNextRecordNumber() As Long
    ' Получение следующего номера записи
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo ErrorHandler
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    GetNextRecordNumber = tblData.ListRows.Count + 1
    
    Exit Function
    
ErrorHandler:
    GetNextRecordNumber = 1
End Function

Public Sub ResetFormToDefaults()
    ' Сброс формы к настройкам по умолчанию
    DataManager.CurrentRecordRow = 0
    DataManager.IsNewRecord = True
    DataManager.FormDataChanged = False
    
    Call DataManager.ClearForm
    Call SetupFormAppearance
    
    UserFormVhIsh.lblStatusBar.Caption = "Форма сброшена к настройкам по умолчанию"
End Sub

