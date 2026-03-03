VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormVhIsh 
   Caption         =   "Система учёта входящих и исходящих"
   ClientHeight    =   12045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17940
   OleObjectBlob   =   "UserFormVhIsh.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormVhIsh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'==============================================
' МОДУЛЬ УПРАВЛЕНИЯ ФОРМОЙ "ВходящиеИсходящие" - UserFormVhIsh
' Назначение: Полнофункциональная форма для добавления, редактирования и поиска записей
' Состояние: ИСПРАВЛЕНА ОШИБКА "Expected array" В ФУНКЦИИ IsNaryadOnlyMode
' Версия: 3.2.1
' Дата: 09.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' API объявления для работы с разрешением экрана
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

' Константы для GetSystemMetrics
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

' Глобальные переменные модуля формы
Private CurrentRecordRow As Long
Private IsNewRecord As Boolean
Private FormDataChanged As Boolean

' Кнопка очистки всех отметок об исполнении
Private Sub btnClearMarks_Click()
    Call ProvodkaIntegrationModule.ClearAllProvodkaMarks
    
    ' Обновляем текущую запись в форме
    If DataManager.CurrentRecordRow > 0 Then
        Call Me.LoadRecordToForm(DataManager.CurrentRecordRow)
    End If
End Sub


Private Sub btnFindProvodka_Click()
    ' Кнопка поиска проводки для текущей записи
    Call ProvodkaIntegrationModule.FindProvodkaForCurrentRecord
End Sub

' Кнопка массовой проверки всех записей с выгрузкой 1С
Private Sub btnMassCheck_Click()
    Dim response As VbMsgBoxResult
    Dim StartTime As Double
    
    On Error GoTo MassCheckError
    
    ' Предупреждение пользователя
    response = MsgBox("?? МАССОВАЯ ПРОВЕРКА ЗАПИСЕЙ" & vbCrLf & vbCrLf & _
                     "Будет выполнена проверка ВСЕХ записей в таблице" & vbCrLf & _
                     "ВходящиеИсходящие на соответствие с выгрузкой 1С." & vbCrLf & vbCrLf & _
                     "Это может занять некоторое время." & vbCrLf & _
                     "Продолжить?", _
                     vbYesNo + vbQuestion, "Подтверждение массовой проверки")
    
    If response = vbNo Then Exit Sub
    
    ' Обновляем статус в форме
    Me.lblStatusBar.Caption = "Выполняется массовая проверка..."
    Me.btnMassCheck.Enabled = False
    
    StartTime = Timer
    
    ' Запускаем массовую обработку
    Call ProvodkaIntegrationModule.MassProcessWithFileSelection
    
    ' Восстанавливаем интерфейс
    Me.btnMassCheck.Enabled = True
    
    ' Если текущая запись открыта, обновляем её отображение
    If DataManager.CurrentRecordRow > 0 Then
        Call Me.LoadRecordToForm(DataManager.CurrentRecordRow)
    End If
    
    ' Обновляем статус
    Me.lblStatusBar.Caption = "Массовая проверка завершена за " & _
                             Format(Timer - StartTime, "0.0") & " сек."
    
    Exit Sub
    
MassCheckError:
    Me.btnMassCheck.Enabled = True
    Me.lblStatusBar.Caption = "Ошибка массовой проверки: " & Err.description
    MsgBox "Ошибка при выполнении массовой проверки:" & vbCrLf & Err.description, vbCritical, "Ошибка"
End Sub


' Кнопка показа статистики сопоставления
Private Sub btnShowStats_Click()
    Call ProvodkaIntegrationModule.ShowMatchingStatistics
End Sub


Private Sub CommandButton1_Click()

End Sub

' ===============================================
' СОБЫТИЯ ФОРМЫ
' ===============================================

Private Sub UserForm_Initialize()
    ' Масштабирование и центрирование формы ДО инициализации остальных элементов
    Call ResizeAndCenterForm
    
    Call InitializeForm
    Call LoadSettings
    Call NavigationModule.UpdateStatusBar
'    Me.KeyPreview = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If HasUnsavedChanges() Then
        Dim response As VbMsgBoxResult
        response = MsgBox("У вас есть несохранённые изменения. Сохранить перед закрытием?", _
                         vbYesNoCancel + vbQuestion, "Несохранённые изменения")
        
        Select Case response
            Case vbYes
                Call SaveCurrentRecord
                Call SaveSettings
            Case vbCancel
                Cancel = True
                Exit Sub
        End Select
    End If
    
    Call SaveSettings
    
    ' Возвращаем фокус к таблице
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    wsData.Activate
    Application.StatusBar = "Форма закрыта. Система интерактивных таблиц активна."
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Горячие клавиши
    If Shift = 2 Then ' Ctrl нажат
        Select Case KeyCode
            Case vbKeyS ' Ctrl+S = Сохранить
                Call SaveCurrentRecord
            Case vbKeyN ' Ctrl+N = Новая запись
                Call DataManager.ClearForm
            Case vbKeyF ' Ctrl+F = Фокус на поиск
                Me.txtSearch.SetFocus
        End Select
    ElseIf KeyCode = vbKeyF3 Then ' F3 = Следующий результат поиска
        If Me.lstSearchResults.Visible And Me.lstSearchResults.ListCount > 0 Then
            With Me.lstSearchResults
                If .listIndex < .ListCount - 1 Then
                    .listIndex = .listIndex + 1
                Else
                    .listIndex = 0
                End If
            End With
        End If
    End If
    
' Проверка нажатия клавиши Escape (код 27)
    If KeyCode = vbKeyEscape Then
        ' Закрытие формы
        Me.Hide
        ' Или для полной выгрузки используйте: Unload Me
    End If
    
    
End Sub

Private Sub InitializeForm()
    ' Инициализация переменных
    CurrentRecordRow = 0
    IsNewRecord = True
    FormDataChanged = False
    
    ' Инициализация элементов формы
    Call LoadComboBoxData
    Call SetupNewComboBoxes
    Call DataManager.ClearForm
    
    ' Настройка элементов поиска
    Me.txtSearch.Text = ""
    Me.lstSearchResults.Clear
    Me.lstSearchResults.Visible = False
    
    ' Настройка фиксированного размера для lstSearchResults
    Call SetupSearchResultsList
    
    ' Делаем поле номера П/П недоступным для редактирования
    Me.txtNomerPP.Enabled = False
    Me.txtNomerPP.BackColor = RGB(240, 240, 240) ' Серый цвет для недоступного поля
    Me.txtNomerPP.Locked = True
    
    ' Настройка поля наряда
    Call SetupNaryadField
    
    ' Настройка строки состояния
    Me.lblStatusBar.Caption = "Готов к работе"
    
    ' Деактивация кнопки удаления
    Me.btnDelete.Enabled = False
    
    ' Улучшения интерфейса
    Call SetupFormAppearance
    
    ' Настройка поиска
    Me.txtSearch.Width = 300
    Me.txtSearch.Height = 24
    
    ' Подсказка для поиска
    Me.lblStatusBar.Caption = "Готов к работе. Горячие клавиши: Ctrl+S (сохранить), Ctrl+N (новая), Ctrl+F (поиск)"
End Sub

' Настройка поля наряда (теперь необязательное)
Private Sub SetupNaryadField()
    With Me.txtNaryadInfo
        .MaxLength = 100 ' Ограничение длины
        .BackColor = RGB(255, 255, 255) ' Белый цвет - необязательное поле
        .Text = ""
    End With
End Sub

' Настройка новых ComboBox'ов
Private Sub SetupNewComboBoxes()
    ' Настройка cmbVidDocumenta (Вх./Исх.)
    With Me.cmbVidDocumenta
        .Clear
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        .AddItem "Вх."
        .AddItem "Исх."
        .listIndex = -1
    End With
    
    ' Настройка cmbStatusPodtverjdenie
    With Me.cmbStatusPodtverjdenie
        .Clear
        .Style = fmStyleDropDownList
        .MatchRequired = False
        .AddItem ""
        .AddItem "Подтверждено"
        .AddItem "Отпр.Исх."
        .listIndex = 0
    End With
End Sub

' Настройка списка результатов поиска
Private Sub SetupSearchResultsList()
    With Me.lstSearchResults
        .ColumnCount = 1
        .columnWidths = "550"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Width = 420
        .Height = 120
        .Visible = False
        .MultiSelect = fmMultiSelectSingle
    End With
End Sub

Private Sub LoadComboBoxData()
    Dim wsSettings As Worksheet
    
    ' Проверка существования листа "Настройки"
    On Error GoTo SettingsError
    Set wsSettings = ThisWorkbook.Worksheets("Настройки")
    On Error GoTo 0
    
    ' Загрузка служб
    Call LoadComboData(Me.cmbSlujba, wsSettings, "A:A", "Службы")
    
    ' Загрузка видов документов в cmbVidDoc (4-й столбец)
    Call LoadComboData(Me.cmbVidDoc, wsSettings, "C:C", "Виды документов")
    
    ' Загрузка исполнителей
    Call LoadComboData(Me.cmbIspolnitel, wsSettings, "E:E", "Исполнители")
    
    ' Загрузка данных для автодополнения cmbOtKogoPostupil
    Call LoadAutoCompleteData
    
    ' Загрузка данных для автодополнения txtNaryadInfo
    Call LoadNaryadAutoCompleteData
    
    Exit Sub
    
SettingsError:
    MsgBox "Лист 'Настройки' не найден. Создайте лист для справочников.", vbExclamation, "Ошибка"
End Sub

' Загрузка данных для автодополнения поля наряда
Private Sub LoadNaryadAutoCompleteData()
    ' В будущем можно добавить автодополнение для нарядов
    ' из отдельной таблицы или из уже введенных данных
End Sub

' Загрузка данных для автодополнения
Private Sub LoadAutoCompleteData()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    Dim cellValue As String
    Dim uniqueValues As Collection
    Dim CurrentValue As String
    
    On Error GoTo LoadAutoError
    
    CurrentValue = CStr(Me.cmbOtKogoPostupil.value)
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    Set uniqueValues = New Collection
    
    ' Собираем уникальные значения из 9-го столбца
    If tblData.ListRows.Count > 0 Then
        For i = 1 To tblData.ListRows.Count
            cellValue = Trim(CStr(tblData.DataBodyRange.Cells(i, 9).value))
            If cellValue <> "" Then
                On Error Resume Next
                uniqueValues.Add cellValue, cellValue
                On Error GoTo LoadAutoError
            End If
        Next i
    End If
    
    ' Настройка автодополнения для cmbOtKogoPostupil
    With Me.cmbOtKogoPostupil
        .Clear
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        
        Dim item As Variant
        For Each item In uniqueValues
            .AddItem item
        Next item
        
        If Trim(CurrentValue) <> "" Then
            .value = CurrentValue
        Else
            .listIndex = -1
        End If
    End With
    
    Exit Sub
    
LoadAutoError:
    If Trim(CurrentValue) <> "" Then
        Me.cmbOtKogoPostupil.value = CurrentValue
    End If
End Sub

Private Sub LoadComboData(ComboBox As ComboBox, Ws As Worksheet, columnRange As String, headerText As String)
    Dim cell As Range
    Dim StartRow As Long
    
    On Error GoTo LoadError
    
    With ComboBox
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        .Clear
    End With
    
    ' Найти строку с заголовком
    Set cell = Ws.Range(columnRange).Find(headerText, LookIn:=xlValues, LookAt:=xlWhole)
    If Not cell Is Nothing Then
        StartRow = cell.Row + 1
        
        ' Загрузить данные, пропуская пустые ячейки
        Set cell = Ws.Cells(StartRow, cell.Column)
        Do While cell.value <> "" And cell.Row <= Ws.UsedRange.Rows.Count + Ws.UsedRange.Row
            ComboBox.AddItem cell.value
            Set cell = cell.Offset(1, 0)
        Loop
    End If
    
    ComboBox.listIndex = -1
    
    Exit Sub
    
LoadError:
    ' Просто пропускаем ошибки загрузки данных
End Sub

' ===============================================
' ОБРАБОТЧИКИ КНОПОК НАВИГАЦИИ
' ===============================================

Private Sub btnFirst_Click()
    Call NavigateToRecord(1)
End Sub

Private Sub btnPrevious_Click()
    Call NavigateToPrevious
End Sub

Private Sub btnNext_Click()
    Call NavigateToNext
End Sub

Private Sub btnLast_Click()
    Call NavigateToLast
End Sub

' ===============================================
' ОБРАБОТЧИКИ ОСНОВНЫХ КНОПОК
' ===============================================

Private Sub btnSave_Click()
    Call SaveCurrentRecord
End Sub

Private Sub btnCancel_Click()
    Call CancelChanges
End Sub

Private Sub btnClear_Click()
    Call DataManager.ClearForm
End Sub

Private Sub btnNew_Click()
    If HasUnsavedChanges() Then
        Dim response As VbMsgBoxResult
        response = MsgBox("У вас есть несохранённые изменения. Сохранить перед созданием новой записи?", _
                         vbYesNoCancel + vbQuestion, "Несохранённые изменения")
        
        Select Case response
            Case vbYes
                Call SaveCurrentRecord
            Case vbCancel
                Exit Sub
        End Select
    End If
    
    Call DataManager.ClearForm
    Me.lblStatusBar.Caption = "Создание новой записи"
End Sub

Private Sub btnDelete_Click()
    MsgBox "Функция удаления временно отключена", vbInformation, "Информация"
End Sub

' ===============================================
' ПОИСК В РЕАЛЬНОМ ВРЕМЕНИ
' ===============================================

Private Sub txtSearch_Change()
    Call PerformSearch
End Sub

Private Sub lstSearchResults_Click()
    Call SelectSearchResult
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call ClearSearch
    ElseIf KeyCode = vbKeyReturn Then
        If Me.lstSearchResults.Visible And Me.lstSearchResults.ListCount > 0 Then
            Me.lstSearchResults.listIndex = 0
            Call SelectSearchResult
        End If
    ElseIf KeyCode = vbKeyDown Then
        If Me.lstSearchResults.Visible And Me.lstSearchResults.ListCount > 0 Then
            Call NavigateSearchResults("DOWN")
        End If
    ElseIf KeyCode = vbKeyUp Then
        If Me.lstSearchResults.Visible And Me.lstSearchResults.ListCount > 0 Then
            Call NavigateSearchResults("UP")
        End If
    End If
End Sub

Private Sub lstSearchResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call SelectSearchResult
        Case vbKeyEscape
            Me.txtSearch.SetFocus
        Case vbKeyDelete
            Call ClearSearch
            Me.txtSearch.SetFocus
        Case vbKeyF3
            Call NavigateSearchResults("DOWN")
        Case vbKeyF3 + 256
            Call NavigateSearchResults("UP")
    End Select
End Sub

Private Sub lstSearchResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call SelectSearchResult
    Call HighlightSearchTermInForm
End Sub

' ===============================================
' ОБРАБОТЧИКИ ИЗМЕНЕНИЯ ПОЛЕЙ
' ===============================================

Private Sub cmbVidDocumenta_Change()
    Call MarkFormAsChanged
    Call UpdateFieldAppearanceBasedOnNaryad ' Обновляем подсветку полей
End Sub

Private Sub txtNomerDoc_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtSummaDoc_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtVhFRP_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataVhFRP_Change()
    Call MarkFormAsChanged
End Sub

Private Sub cmbOtKogoPostupil_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataPeredachi_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtNomerIshVSlujbu_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataIshVSlujbu_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtNomerVozvrata_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataVozvrata_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtNomerIshKonvert_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataIshKonvert_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtOtmetkaIspolnenie_Change()
    Call MarkFormAsChanged
End Sub

Private Sub cmbStatusPodtverjdenie_Change()
    Call MarkFormAsChanged
End Sub

Private Sub cmbSlujba_Change()
    Call MarkFormAsChanged
End Sub

Private Sub cmbVidDoc_Change()
    Call MarkFormAsChanged
End Sub

Private Sub cmbIspolnitel_Change()
    Call MarkFormAsChanged
End Sub

' Изменение поля наряда с обновлением подсветки
Private Sub txtNaryadInfo_Change()
    Call MarkFormAsChanged
    Call UpdateFieldAppearanceBasedOnNaryad ' Обновляем подсветку полей при изменении наряда
End Sub

' Обновление внешнего вида полей в зависимости от режима наряда
Private Sub UpdateFieldAppearanceBasedOnNaryad()
    On Error Resume Next
    Dim isNaryadOnly As Boolean
    Dim requiredColor As Long
    Dim optionalColor As Long
    
    requiredColor = RGB(255, 255, 224) ' Светло-желтый для обязательных
    optionalColor = RGB(255, 255, 255) ' Белый для необязательных
    
    ' ВАЖНО: получаем результат функции в ДРУГУЮ переменную
    isNaryadOnly = IsNaryadOnlyMode()
    
    If isNaryadOnly Then
        ' Режим "только наряд" - делаем основные поля необязательными
        Me.cmbVidDoc.BackColor = optionalColor
        Me.txtNomerDoc.BackColor = optionalColor
        Me.txtSummaDoc.BackColor = optionalColor
        Me.txtVhFRP.BackColor = optionalColor
        Me.txtDataVhFRP.BackColor = optionalColor
        
        Me.lblStatusBar.Caption = "РЕЖИМ НАРЯДА: Основные поля документа необязательны"
    Else
        ' Обычный режим - основные поля обязательны
        Me.cmbVidDoc.BackColor = requiredColor
        Me.txtNomerDoc.BackColor = requiredColor
        Me.txtSummaDoc.BackColor = requiredColor
        Me.txtVhFRP.BackColor = requiredColor
        Me.txtDataVhFRP.BackColor = requiredColor
        
        If Me.lblStatusBar.Caption = "РЕЖИМ НАРЯДА: Основные поля документа необязательны" Then
            Me.lblStatusBar.Caption = "Обычный режим: все обязательные поля должны быть заполнены"
        End If
End If

End Sub

' КРИТИЧЕСКИ ИСПРАВЛЕННАЯ ФУНКЦИЯ: Проверка режима "только наряд"
Private Function IsNaryadOnlyMode() As Boolean
    ' Режим "только наряд" активен если:
    ' 1. Заполнено поле наряда
    ' 2. Заполнена служба
    ' 3. НЕ заполнены основные поля: тип документа, номер, сумма, ВхФРП, дата ВхФРП
    
    On Error GoTo ModeError
    
    Dim naryadText As String
Dim slujbaText As String
Dim vidDocText As String
Dim nomerDocText As String
Dim summaDocText As String
Dim vhFrpText As String
Dim dataVhFrpText As String

' Безопасное чтение значений элементов формы в локальные переменные
naryadText = ""
If Not Me.txtNaryadInfo Is Nothing Then naryadText = Trim(CStr(Me.txtNaryadInfo.Text))

slujbaText = ""
If Not Me.cmbSlujba Is Nothing Then slujbaText = Trim(CStr(Me.cmbSlujba.Text))

vidDocText = ""
If Not Me.cmbVidDoc Is Nothing Then vidDocText = Trim(CStr(Me.cmbVidDoc.Text))

nomerDocText = ""
If Not Me.txtNomerDoc Is Nothing Then nomerDocText = Trim(CStr(Me.txtNomerDoc.Text))

summaDocText = ""
If Not Me.txtSummaDoc Is Nothing Then summaDocText = Trim(CStr(Me.txtSummaDoc.Text))

vhFrpText = ""
If Not Me.txtVhFRP Is Nothing Then vhFrpText = Trim(CStr(Me.txtVhFRP.Text))

dataVhFrpText = ""
If Not Me.txtDataVhFRP Is Nothing Then dataVhFrpText = Trim(CStr(Me.txtDataVhFRP.Text))

' 1) Наряд должен быть заполнен
If naryadText = "" Then
    IsNaryadOnlyMode = False
    Exit Function
End If

' 2) Служба должна быть заполнена
If slujbaText = "" Then
    IsNaryadOnlyMode = False
    Exit Function
End If

' 3) Основные поля должны быть пустыми
If vidDocText <> "" Then
    IsNaryadOnlyMode = False
    Exit Function
End If

If nomerDocText <> "" Then
    IsNaryadOnlyMode = False
    Exit Function
End If

If summaDocText <> "" Then
    IsNaryadOnlyMode = False
    Exit Function
End If

If vhFrpText <> "" Then
    IsNaryadOnlyMode = False
    Exit Function
End If

If dataVhFrpText <> "" Then
    IsNaryadOnlyMode = False
    Exit Function
End If

' Все условия выполнены — режим "только наряд" активен
IsNaryadOnlyMode = True
Exit Function

ModeError:
' При любой ошибке возвращаем False
IsNaryadOnlyMode = False

End Function



' ===============================================
' ВАЛИДАЦИЯ ПОЛЕЙ ПРИ ВЫХОДЕ
' ===============================================

Private Sub txtSummaDoc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Me.txtSummaDoc.Text) > 0 And Not IsNumeric(Me.txtSummaDoc.Text) Then
        MsgBox "Сумма должна быть числом!", vbExclamation, "Проверка данных"
        Cancel = True
        Me.txtSummaDoc.BackColor = RGB(255, 200, 200)
    Else
        ' Восстанавливаем правильный цвет в зависимости от режима
        Call UpdateFieldAppearanceBasedOnNaryad
    End If
End Sub

Private Sub txtDataVhFRP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataVhFRP, Cancel)
End Sub

Private Sub txtDataPeredachi_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataPeredachi, Cancel)
End Sub

Private Sub txtDataIshVSlujbu_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataIshVSlujbu, Cancel)
End Sub

Private Sub txtDataVozvrata_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataVozvrata, Cancel)
End Sub

Private Sub txtDataIshKonvert_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataIshKonvert, Cancel)
End Sub

' Проверка поля наряда при выходе
Private Sub txtNaryadInfo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Валидация формата наряда (необязательное поле)
    If Len(Trim(Me.txtNaryadInfo.Text)) > 0 Then
        Me.txtNaryadInfo.BackColor = RGB(255, 255, 255) ' Белый цвет для необязательного поля
    End If
    
    ' Обновляем подсветку других полей при выходе из поля наряда
    Call UpdateFieldAppearanceBasedOnNaryad
End Sub

Private Sub ValidateAndFormatDate(dateField As MSForms.TextBox, Cancel As MSForms.ReturnBoolean)
    If Len(dateField.Text) > 0 Then
        ' Автоматическое добавление точек при вводе даты
        Dim DateText As String
        DateText = Replace(dateField.Text, ".", "")
        
        If Len(DateText) = 6 And IsNumeric(DateText) Then
            dateField.Text = Left(DateText, 2) & "." & Mid(DateText, 3, 2) & "." & Right(DateText, 2)
        End If
        
        ' Валидация формата
        If Not CommonUtilities.IsValidDateFormat(dateField.Text) Then
            MsgBox "Введите дату в формате ДД.ММ.ГГ", vbExclamation, "Проверка данных"
            Cancel = True
            dateField.BackColor = RGB(255, 200, 200)
        Else
            ' Восстанавливаем правильный цвет в зависимости от режима
            Call UpdateFieldAppearanceBasedOnNaryad
        End If
    End If
End Sub

' ===============================================
' ФУНКЦИИ УПРАВЛЕНИЯ ДАННЫМИ
' ===============================================

Private Sub SaveCurrentRecord()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim newRow As ListRow
    Dim targetRowIndex As Long
    
    ' Условная проверка обязательных полей
    If Not ValidateRequiredFieldsConditional() Then
        Exit Sub
    End If
    
    ' Валидация дат
    If Not DataManager.ValidateDates() Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    If IsNewRecord Then
        Set newRow = tblData.ListRows.Add
        targetRowIndex = newRow.Index
        CurrentRecordRow = targetRowIndex
        
        Me.txtNomerPP.Text = CStr(targetRowIndex)
        Me.lblStatusBar.Caption = "Создание новой записи №" & targetRowIndex
    Else
        If CurrentRecordRow < 1 Or CurrentRecordRow > tblData.ListRows.Count Then
            MsgBox "Ошибка: неверный номер записи для обновления!", vbCritical, "Ошибка сохранения"
            Exit Sub
        End If
        
        targetRowIndex = CurrentRecordRow
        Me.lblStatusBar.Caption = "Обновление существующей записи №" & targetRowIndex
    End If
    
    ' Сохранение данных в указанную строку таблицы
    Call WriteFormDataToTable(tblData, targetRowIndex)
    
    ' Обновление статуса
    IsNewRecord = False
    FormDataChanged = False
    Me.lblStatusBar.Caption = "Запись №" & targetRowIndex & " сохранена успешно"
    
    If targetRowIndex = tblData.ListRows.Count Then
        Call LoadAutoCompleteData
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при сохранении данных: " & Err.description, vbCritical, "Ошибка сохранения"
    Me.lblStatusBar.Caption = "Ошибка сохранения данных"
End Sub

' Условная валидация обязательных полей
Private Function ValidateRequiredFieldsConditional() As Boolean
    ValidateRequiredFieldsConditional = True
    
    ' Служба всегда обязательна
    If Trim(CStr(Me.cmbSlujba.value)) = "" Then
        MsgBox "Поле 'Служба' обязательно для заполнения!", vbExclamation, "Проверка данных"
        Me.cmbSlujba.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    ' Проверяем режим "только наряд"
    If IsNaryadOnlyMode() Then
        ' В режиме "только наряд" основные поля необязательны
        ' Но должен быть заполнен хотя бы наряд
        If Trim(Me.txtNaryadInfo.Text) = "" Then
            MsgBox "В режиме наряда должно быть заполнено поле 'Информация о наряде'!", vbExclamation, "Проверка данных"
            Me.txtNaryadInfo.SetFocus
            ValidateRequiredFieldsConditional = False
            Exit Function
        End If
        
        ' Успешная валидация для режима "только наряд"
        Exit Function
    End If
    
    ' Обычный режим - проверяем все обязательные поля
    If Trim(CStr(Me.cmbVidDoc.value)) = "" Then
        MsgBox "Поле 'Тип документа' обязательно для заполнения!", vbExclamation, "Проверка данных"
        Me.cmbVidDoc.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(CStr(Me.cmbVidDocumenta.value)) = "" Then
        MsgBox "Поле 'Вид документа (Вх./Исх.)' обязательно для заполнения!", vbExclamation, "Проверка данных"
        Me.cmbVidDocumenta.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(Me.txtNomerDoc.Text) = "" Then
        MsgBox "Поле 'Номер документа' обязательно для заполнения!", vbExclamation, "Проверка данных"
        Me.txtNomerDoc.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(Me.txtSummaDoc.Text) = "" Then
        MsgBox "Поле 'Сумма документа' обязательно для заполнения!", vbExclamation, "Проверка данных"
        Me.txtSummaDoc.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Not IsNumeric(Me.txtSummaDoc.Text) Then
        MsgBox "Поле 'Сумма документа' должно содержать числовое значение!", vbExclamation, "Проверка данных"
        Me.txtSummaDoc.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(Me.txtVhFRP.Text) = "" Then
        MsgBox "Поле 'Вх.ФРП/Исх.ФРП' обязательно для заполнения!", vbExclamation, "Проверка данных"
        Me.txtVhFRP.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(Me.txtDataVhFRP.Text) = "" Then
        MsgBox "Поле 'Дата Вх.ФРП/Исх.ФРП' обязательно для заполнения!", vbExclamation, "Проверка данных"
        Me.txtDataVhFRP.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
End Function

' Функции ValidateDates и IsValidDateFormat перенесены в DataManager.bas и CommonUtilities.bas

' Запись данных формы в таблицу
Private Sub WriteFormDataToTable(tbl As ListObject, RowIndex As Long)
    On Error GoTo WriteError
    
    tbl.DataBodyRange.Cells(RowIndex, 1).value = RowIndex
    tbl.DataBodyRange.Cells(RowIndex, 2).value = Me.cmbSlujba.value
    tbl.DataBodyRange.Cells(RowIndex, 3).value = Me.cmbVidDocumenta.value
    tbl.DataBodyRange.Cells(RowIndex, 4).value = Me.cmbVidDoc.value
    tbl.DataBodyRange.Cells(RowIndex, 5).value = Me.txtNomerDoc.Text
    
    If IsNumeric(Me.txtSummaDoc.Text) Then
        tbl.DataBodyRange.Cells(RowIndex, 6).value = CDbl(Me.txtSummaDoc.Text)
    Else
        tbl.DataBodyRange.Cells(RowIndex, 6).value = 0
    End If
    
    tbl.DataBodyRange.Cells(RowIndex, 7).value = Me.txtVhFRP.Text
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 8), Me.txtDataVhFRP.Text)
    tbl.DataBodyRange.Cells(RowIndex, 9).value = Me.cmbOtKogoPostupil.value
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 10), Me.txtDataPeredachi.Text)
    tbl.DataBodyRange.Cells(RowIndex, 11).value = Me.cmbIspolnitel.value
    tbl.DataBodyRange.Cells(RowIndex, 12).value = Me.txtNomerIshVSlujbu.Text
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 13), Me.txtDataIshVSlujbu.Text)
    tbl.DataBodyRange.Cells(RowIndex, 14).value = Me.txtNomerVozvrata.Text
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 15), Me.txtDataVozvrata.Text)
    tbl.DataBodyRange.Cells(RowIndex, 16).value = Me.txtNomerIshKonvert.Text
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 17), Me.txtDataIshKonvert.Text)
    tbl.DataBodyRange.Cells(RowIndex, 18).value = Me.txtOtmetkaIspolnenie.Text
    tbl.DataBodyRange.Cells(RowIndex, 19).value = Me.cmbStatusPodtverjdenie.value
    tbl.DataBodyRange.Cells(RowIndex, 20).value = Me.txtNaryadInfo.Text
    
    Exit Sub
    
WriteError:
    MsgBox "Ошибка записи данных в таблицу: " & Err.description, vbCritical, "Ошибка"
End Sub

' Функция WriteDateToCell перенесена в CommonUtilities.bas

' ===============================================
' ЗАГРУЗКА ЗАПИСИ ИЗ ТАБЛИЦЫ В ФОРМУ
' ===============================================

Private Sub UpdateProvodkaIndicator()
    Dim otmetka As String
    otmetka = Trim(Me.txtOtmetkaIspolnenie.Text)
    
    If otmetka <> "" Then
        Me.txtOtmetkaIspolnenie.BackColor = RGB(200, 255, 200) ' Светло-зеленый
        Me.lblStatusBar.Caption = Me.lblStatusBar.Caption & " | + Проводка: " & otmetka
    Else
        Me.txtOtmetkaIspolnenie.BackColor = RGB(255, 255, 255) ' Белый
    End If
End Sub



Public Sub LoadRecordToForm(RowIndex As Long)
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo LoadError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    If RowIndex < 1 Or RowIndex > tblData.ListRows.Count Then
        Exit Sub
    End If
    
    IsNewRecord = False
    CurrentRecordRow = RowIndex
    
    Me.txtNomerPP.Text = CStr(RowIndex)
    Me.cmbSlujba.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 2).value)
    Me.cmbVidDocumenta.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 3).value)
    Me.cmbVidDoc.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 4).value)
    Me.txtNomerDoc.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 5).value)
    Me.txtSummaDoc.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 6).value)
    Me.txtVhFRP.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 7).value)
    Me.txtDataVhFRP.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 8))
    Me.cmbOtKogoPostupil.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 9).value)
    Me.txtDataPeredachi.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 10))
    Me.cmbIspolnitel.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 11).value)
    Me.txtNomerIshVSlujbu.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 12).value)
    Me.txtDataIshVSlujbu.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 13))
    Me.txtNomerVozvrata.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 14).value)
    Me.txtDataVozvrata.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 15))
    Me.txtNomerIshKonvert.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 16).value)
    Me.txtDataIshKonvert.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 17))
    Me.txtOtmetkaIspolnenie.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 18).value)
    Me.cmbStatusPodtverjdenie.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 19).value)
    
    If tblData.DataBodyRange.Columns.Count >= 20 Then
        Me.txtNaryadInfo.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 20).value)
    Else
        Me.txtNaryadInfo.Text = ""
    End If
    
    FormDataChanged = False
    
    ' Обновляем подсветку полей после загрузки записи
    Call UpdateFieldAppearanceBasedOnNaryad
    
    Exit Sub
    
LoadError:
    MsgBox "Ошибка загрузки записи: " & Err.description, vbExclamation, "Ошибка"
    Call DataManager.ClearForm
End Sub

' Функция FormatDateCell перенесена в CommonUtilities.bas

' ===============================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ===============================================

Private Function HasUnsavedChanges() As Boolean
    HasUnsavedChanges = FormDataChanged
End Function

Private Sub CancelChanges()
    If IsNewRecord Then
        Call DataManager.ClearForm
    Else
        Call NavigateToRecord(CurrentRecordRow)
    End If
    
    FormDataChanged = False
    Me.lblStatusBar.Caption = "Изменения отменены"
End Sub

' Функция ClearForm перенесена в DataManager.bas
' Используйте DataManager.ClearForm() для очистки формы

' Методы для работы с формой из таблицы
Public Sub ShowFromTable(RowNumber As Long, fieldName As String)
    Call Me.LoadRecordToForm(RowNumber)
    
    CurrentRecordRow = RowNumber
    IsNewRecord = False
    FormDataChanged = False
    
    Me.Show vbModeless
    
    Application.OnTime Now + TimeValue("00:00:01"), "'SetFieldFocus """ & fieldName & """'"
    
    Me.lblStatusBar.Caption = "Открыто из таблицы | Запись №" & RowNumber & " | Поле: " & TableEventHandler.GetFieldDisplayName(fieldName)
End Sub

Public Sub SetFieldFocus(fieldName As String)
    Call TableEventHandler.ActivateFormField(fieldName)
End Sub

Private Sub UpdateNavigationButtons()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo NavigationError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    Me.btnFirst.Enabled = (CurrentRecordRow > 1)
    Me.btnPrevious.Enabled = (CurrentRecordRow > 1)
    Me.btnNext.Enabled = (CurrentRecordRow < tblData.ListRows.Count)
    Me.btnLast.Enabled = (CurrentRecordRow < tblData.ListRows.Count)
    
    Exit Sub
    
NavigationError:
    Me.btnFirst.Enabled = False
    Me.btnPrevious.Enabled = False
    Me.btnNext.Enabled = False
    Me.btnLast.Enabled = False
End Sub

' Функция UpdateStatusBar перенесена в NavigationModule.bas
' Используйте NavigationModule.UpdateStatusBar() для обновления статус-бара

Private Sub MarkFormAsChanged()
    FormDataChanged = True
    Call NavigationModule.UpdateStatusBar
End Sub

' ===============================================
' ФУНКЦИИ НАСТРОЕК
' ===============================================

Private Sub SaveSettings()
    Dim wsSettings As Worksheet
    
    On Error GoTo SaveError
    
    Set wsSettings = GetSettingsWorksheet()
    
    With wsSettings
        .Cells(1, 1).value = "CurrentRecordRow"
        .Cells(1, 2).value = CurrentRecordRow
        
        .Cells(2, 1).value = "FormTop"
        .Cells(2, 2).value = Me.Top
        
        .Cells(3, 1).value = "FormLeft"
        .Cells(3, 2).value = Me.Left
        
        .Cells(4, 1).value = "LastSaved"
        .Cells(4, 2).value = Now()
    End With
    
    Exit Sub
    
SaveError:
    ' Игнорируем ошибки сохранения настроек
End Sub

Private Sub LoadSettings()
    Dim wsSettings As Worksheet
    
    On Error GoTo DefaultSettings
    
    Set wsSettings = GetSettingsWorksheet()
    
    With wsSettings
        If .Cells(1, 2).value <> "" Then
            CurrentRecordRow = .Cells(1, 2).value
        End If
        
        ' НЕ загружаем сохранённые позиции - форма всегда центрируется автоматически
        ' Размеры и позиция устанавливаются в ResizeAndCenterForm
    End With
    
    If CurrentRecordRow > 0 Then
        Call NavigateToRecord(CurrentRecordRow)
    Else
        Call DataManager.ClearForm
    End If
    
    Exit Sub
    
DefaultSettings:
    CurrentRecordRow = 0
    IsNewRecord = True
    FormDataChanged = False
End Sub

Private Function GetSettingsWorksheet() As Worksheet
    Dim Ws As Worksheet
    
    On Error GoTo CreateSheet
    
    Set Ws = ThisWorkbook.Worksheets("Настройки_Системы")
    Set GetSettingsWorksheet = Ws
    Exit Function
    
CreateSheet:
    Set Ws = ThisWorkbook.Worksheets.Add
    Ws.Name = "Настройки_Системы"
    Ws.Visible = xlSheetHidden
    Set GetSettingsWorksheet = Ws
End Function

' ===============================================
' ФУНКЦИИ ОФОРМЛЕНИЯ ИНТЕРФЕЙСА
' ===============================================

Private Sub SetupFormAppearance()
    Call SetupRequiredFieldsHighlight
    Call SetupFieldAlignment
    Call SetupNavigationButtons
    Call SetupStatusBar
End Sub

' Настройка подсветки с учетом режима наряда
Private Sub SetupRequiredFieldsHighlight()
    Dim requiredColor As Long
    requiredColor = RGB(255, 255, 224)
    
    With Me
        ' Поля, которые всегда обязательны
        .cmbSlujba.BackColor = requiredColor
        .cmbVidDocumenta.BackColor = requiredColor
        
        ' Поле наряда всегда необязательное (белый цвет)
        .txtNaryadInfo.BackColor = RGB(255, 255, 255)
        
        ' Поле номера П/П недоступно
        .txtNomerPP.BackColor = RGB(240, 240, 240)
    End With
    
    ' Условно обязательные поля - обновляем в зависимости от режима
    Call UpdateFieldAppearanceBasedOnNaryad
End Sub

Private Sub SetupFieldAlignment()
    With Me
        .txtSummaDoc.TextAlign = fmTextAlignRight
        .txtNomerPP.TextAlign = fmTextAlignCenter
        .txtNomerDoc.TextAlign = fmTextAlignCenter
        .txtVhFRP.TextAlign = fmTextAlignCenter
        
        .txtDataVhFRP.MaxLength = 8
        .txtDataPeredachi.MaxLength = 8
        .txtDataIshVSlujbu.MaxLength = 8
        .txtDataVozvrata.MaxLength = 8
        .txtDataIshKonvert.MaxLength = 8
        
        .txtNaryadInfo.MaxLength = 100
        .txtNaryadInfo.TextAlign = fmTextAlignLeft
    End With
End Sub

Private Sub SetupNavigationButtons()
    With Me
        .btnFirst.Caption = "|< Первая"
        .btnPrevious.Caption = "< Пред."
        .btnNext.Caption = "След. >"
        .btnLast.Caption = "Послед. >|"
        
        .btnFirst.Width = 70
        .btnPrevious.Width = 70
        .btnNext.Width = 70
        .btnLast.Width = 70
    End With
End Sub

Private Sub SetupStatusBar()
    With Me.lblStatusBar
        .BackColor = RGB(245, 245, 245)
        .borderStyle = fmBorderStyleSingle
        .Font.Size = 9
        .Font.Name = "Segoe UI"
    End With
End Sub

' ===============================================
' ФУНКЦИИ РАБОТЫ С РАЗРЕШЕНИЕМ ЭКРАНА И МАСШТАБИРОВАНИЕМ
' ===============================================

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Получает разрешение основного экрана в пикселях
' @return [String] Строка вида "ширинаxвысота"
' =============================================
Private Function GetScreenResolution() As String
    On Error GoTo ErrorHandler
    
    Dim screenWidth As Long
    Dim screenHeight As Long
    
    ' Получаем разрешение основного экрана
    screenWidth = GetSystemMetrics(SM_CXSCREEN)
    screenHeight = GetSystemMetrics(SM_CYSCREEN)
    
    GetScreenResolution = screenWidth & "x" & screenHeight
    
    Exit Function
    
ErrorHandler:
    ' В случае ошибки возвращаем стандартное разрешение
    GetScreenResolution = "1920x1080"
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Получает ширину экрана в пикселях
' @return [Long] Ширина экрана
' =============================================
Private Function GetScreenWidth() As Long
    On Error GoTo ErrorHandler
    
    GetScreenWidth = GetSystemMetrics(SM_CXSCREEN)
    Exit Function
    
ErrorHandler:
    GetScreenWidth = 1920 ' Значение по умолчанию
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Получает высоту экрана в пикселях
' @return [Long] Высота экрана
' =============================================
Private Function GetScreenHeight() As Long
    On Error GoTo ErrorHandler
    
    GetScreenHeight = GetSystemMetrics(SM_CYSCREEN)
    Exit Function
    
ErrorHandler:
    GetScreenHeight = 1080 ' Значение по умолчанию
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Масштабирует форму под разрешение экрана и центрирует её
' =============================================
Private Sub ResizeAndCenterForm()
    On Error GoTo ErrorHandler
    
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim formWidth As Single
    Dim formHeight As Single
    Dim scaleFactor As Double
    Dim minScaleFactor As Double
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim newLeft As Single
    Dim newTop As Single
    Dim usableScreenWidth As Single
    Dim usableScreenHeight As Single
    
    ' Получаем текущие (исходные) размеры формы из свойств
    ' Эти размеры уже установлены из файла .frm и гарантируют видимость всех элементов
    originalWidth = Me.Width
    originalHeight = Me.Height
    
    ' Получаем разрешение экрана
    screenWidth = GetScreenWidth()
    screenHeight = GetScreenHeight()
    
    ' Переводим пиксели экрана в точки (96 DPI = 72 точки на дюйм)
    Dim screenWidthPoints As Double
    Dim screenHeightPoints As Double
    screenWidthPoints = screenWidth * (72 / 96)
    screenHeightPoints = screenHeight * (72 / 96)
    
    ' Используем рабочую область экрана (90% от размера экрана, чтобы оставить место для панели задач)
    ' Но ограничиваем максимальный размер, чтобы форма не была слишком большой
    usableScreenWidth = screenWidthPoints * 0.9
    usableScreenHeight = screenHeightPoints * 0.9
    
    ' Если исходная форма помещается на экран, используем её размеры без изменений
    ' Иначе масштабируем только при необходимости
    If originalWidth <= usableScreenWidth And originalHeight <= usableScreenHeight Then
        ' Форма помещается на экран - используем исходные размеры
        formWidth = originalWidth
        formHeight = originalHeight
        scaleFactor = 1#
    Else
        ' Форма не помещается - вычисляем коэффициент масштабирования
        Dim scaleX As Double
        Dim scaleY As Double
        scaleX = usableScreenWidth / originalWidth
        scaleY = usableScreenHeight / originalHeight
        
        ' Используем минимальный коэффициент для сохранения пропорций
        scaleFactor = Application.WorksheetFunction.Min(scaleX, scaleY)
        
        ' Минимальный масштаб - 0.75 (75%), чтобы все элементы оставались видимыми
        minScaleFactor = 0.75
        If scaleFactor < minScaleFactor Then scaleFactor = minScaleFactor
        
        ' Вычисляем новые размеры формы
        formWidth = originalWidth * scaleFactor
        formHeight = originalHeight * scaleFactor
        
        ' Дополнительная проверка: убеждаемся, что форма помещается
        If formWidth > usableScreenWidth Then
            formWidth = usableScreenWidth
            scaleFactor = formWidth / originalWidth
            formHeight = originalHeight * scaleFactor
        End If
        If formHeight > usableScreenHeight Then
            formHeight = usableScreenHeight
            scaleFactor = formHeight / originalHeight
            formWidth = originalWidth * scaleFactor
        End If
    End If
    
    ' Устанавливаем размеры формы
    Me.Width = formWidth
    Me.Height = formHeight
    
    ' Центрируем форму на основном экране
    newLeft = (screenWidthPoints - formWidth) / 2
    newTop = (screenHeightPoints - formHeight) / 2
    
    ' Убеждаемся, что форма не выходит за границы экрана
    If newLeft < 0 Then newLeft = 0
    If newTop < 0 Then newTop = 0
    If newLeft + formWidth > screenWidthPoints Then newLeft = screenWidthPoints - formWidth
    If newTop + formHeight > screenHeightPoints Then newTop = screenHeightPoints - formHeight
    
    ' Устанавливаем позицию формы
    Me.Left = newLeft
    Me.Top = newTop
    
    Exit Sub
    
ErrorHandler:
    ' В случае ошибки оставляем исходные размеры и центрируем через StartUpPosition
    Me.StartUpPosition = 1 ' CenterOwner
End Sub


