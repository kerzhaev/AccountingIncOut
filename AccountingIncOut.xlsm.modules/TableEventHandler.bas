Attribute VB_Name = "TableEventHandler"
'==============================================
' МОДУЛЬ ОБРАБОТКИ СОБЫТИЙ ТАБЛИЦЫ - TableEventHandler
' Назначение: Обработка щелчков по таблице для открытия формы с активацией нужного поля
' Состояние: ИСПРАВЛЕНЫ ВСЕ ОБРАЩЕНИЯ К ГЛОБАЛЬНЫМ ПЕРЕМЕННЫМ
' Версия: 1.6.0
' Дата: 10.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' Глобальные переменные для работы с таблицей
Private wsData As Worksheet
Private tblVhIsh As ListObject
Private isSystemActive As Boolean

' Переменные для отслеживания последнего выбора
Private lastSelectedRow As Long
Private lastSelectedColumn As Long

' Переменные для контекстного меню
Private currentRightClickRow As Long
Private contextMenuInitialized As Boolean

' Обработчик событий Application через класс
Private AppEventsHandler As AppEventHandler

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Инициализация обработчика событий
Public Sub InitializeTableEvents()
    On Error GoTo InitError
    
    ' Деактивируем систему перед инициализацией
    isSystemActive = False
    contextMenuInitialized = False
    
    ' Проверка существования листа
    Dim wsExists As Boolean
    wsExists = False
    
    Dim Ws As Worksheet
    For Each Ws In ThisWorkbook.Worksheets
        If Ws.Name = "ВхИсх" Then
            wsExists = True
            Set wsData = Ws
            Exit For
        End If
    Next Ws
    
    If Not wsExists Then
        ' ОСТАВЛЕНО: Критические ошибки показываем
        MsgBox "Лист 'ВхИсх' не найден в книге!", vbCritical, "Ошибка инициализации"
        Exit Sub
    End If
    
    ' Проверка существования таблицы
    Dim tblExists As Boolean
    tblExists = False
    
    Dim tbl As ListObject
    For Each tbl In wsData.ListObjects
        If tbl.Name = "ВходящиеИсходящие" Then
            tblExists = True
            Set tblVhIsh = tbl
            Exit For
        End If
    Next tbl
    
    If Not tblExists Then
        ' ОСТАВЛЕНО: Критические ошибки показываем
        MsgBox "Таблица 'ВходящиеИсходящие' не найдена на листе 'ВхИсх'!" & vbCrLf & _
               "Убедитесь, что таблица существует и имеет правильное название.", vbCritical, "Ошибка инициализации"
        Exit Sub
    End If
    
    ' Проверка содержимого таблицы (только Debug, без MsgBox)
    If tblVhIsh.DataBodyRange Is Nothing Then
        Debug.Print "ПРЕДУПРЕЖДЕНИЕ: Таблица 'ВходящиеИсходящие' пустая!"
        Application.StatusBar = "Предупреждение: таблица пуста. Добавьте данные для работы."
    End If
    
    ' Инициализируем переменные отслеживания
    lastSelectedRow = 0
    lastSelectedColumn = 0
    currentRightClickRow = 0
    
    ' Инициализация обработчика событий Application через класс
    Set AppEventsHandler = New AppEventHandler
    Call AppEventsHandler.InitializeAppEvents
    
    ' Активируем систему
    isSystemActive = True
    contextMenuInitialized = True
    
    ' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Только статус-бар и Debug
    Application.StatusBar = "Система интерактивных форм активна (Excel 2021). Используйте Ctrl+Shift+M для меню."
    
    Debug.Print "Система интерактивных форм активирована для Excel 2021!"
    Debug.Print "Найдена таблица: " & tblVhIsh.Name
    Debug.Print "Строк данных: " & IIf(tblVhIsh.DataBodyRange Is Nothing, 0, tblVhIsh.ListRows.Count)
    Debug.Print "СПОСОБЫ РАБОТЫ: Двойной щелчок, Ctrl+E, Ctrl+Shift+M, Ctrl+D"
    
    Exit Sub
    
InitError:
    isSystemActive = False
    contextMenuInitialized = False
    ' ОСТАВЛЕНО: Критические ошибки показываем
    MsgBox "Критическая ошибка инициализации системы событий таблицы:" & vbCrLf & _
           "Ошибка: " & Err.Number & " - " & Err.description, vbCritical, "Критическая ошибка"
End Sub

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Показ меню действий для Excel 2021
Public Sub ShowActionMenuExcel2021()
    Dim userChoice As String
    Dim currentCell As Range
    Dim intersectRange As Range
    Dim RowNumber As Long
    
    On Error GoTo MenuError
    
    ' Проверяем, что находимся в правильной таблице
    If ActiveSheet.Name <> "ВхИсх" Then
        Application.StatusBar = "Перейдите на лист 'ВхИсх' для использования меню действий."
        Exit Sub
    End If
    
    Set currentCell = Selection
    
    If Not tblVhIsh Is Nothing Then
        If Not tblVhIsh.DataBodyRange Is Nothing Then
            Set intersectRange = Intersect(currentCell, tblVhIsh.DataBodyRange)
            
            If Not intersectRange Is Nothing And currentCell.Cells.Count = 1 Then
                RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
                
                ' ОСТАВЛЕНО: Меню через InputBox (основная функциональность)
                userChoice = InputBox( _
                    "МЕНЮ ДЕЙСТВИЙ ДЛЯ ЗАПИСИ №" & RowNumber & ":" & vbCrLf & vbCrLf & _
                    "Доступные команды:" & vbCrLf & _
                    "1 - Редактировать в форме" & vbCrLf & _
                    "2 - Дублировать запись" & vbCrLf & _
                    "3 - Показать информацию о записи" & vbCrLf & _
                    "0 - Отмена" & vbCrLf & vbCrLf & _
                    "Введите номер команды:", _
                    "Меню действий Excel 2021", "1")
                
                ' Обрабатываем выбор пользователя
                Select Case userChoice
                    Case "1"
                        Call OpenFormForCurrentSelection
                    Case "2"
                        Call DuplicateCurrentRecord
                    Case "3"
                        Call ShowRecordInfo(RowNumber)
                    Case "0", ""
                        ' Отмена - ничего не делаем
                    Case Else
                        Application.StatusBar = "Неверный выбор. Используйте числа 0-3."
                End Select
            Else
                Application.StatusBar = "Выберите ячейку внутри таблицы данных для использования меню."
            End If
        Else
            Application.StatusBar = "Таблица пуста!"
        End If
    Else
        Application.StatusBar = "Таблица не найдена!"
    End If
    
    Exit Sub
    
MenuError:
    Application.StatusBar = "Ошибка отображения меню: " & Err.description
End Sub

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Показ информации о записи
Private Sub ShowRecordInfo(RowNumber As Long)
    Dim info As String
    Dim serviceText As String
    Dim docTypeText As String
    Dim docNumberText As String
    Dim amountText As String
    
    On Error GoTo InfoError
    
    If tblVhIsh Is Nothing Or tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    If RowNumber < 1 Or RowNumber > tblVhIsh.ListRows.Count Then Exit Sub
    
    ' Получаем основную информацию о записи
    serviceText = CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 2).value)     ' Служба
    docTypeText = CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 4).value)     ' Тип документа
    docNumberText = CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 5).value)   ' Номер документа
    amountText = CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 6).value)      ' Сумма
    
    info = "ИНФОРМАЦИЯ О ЗАПИСИ №" & RowNumber & ":" & vbCrLf & vbCrLf
    info = info & "Служба: " & serviceText & vbCrLf
    info = info & "Тип документа: " & docTypeText & vbCrLf
    info = info & "Номер документа: " & docNumberText & vbCrLf
    info = info & "Сумма: " & amountText & " руб." & vbCrLf & vbCrLf
    info = info & "Нажмите OK для продолжения."
    
    ' ОСТАВЛЕНО: Информационные окна (основная функциональность)
    MsgBox info, vbInformation, "Информация о записи"
    
    Exit Sub
    
InfoError:
    Application.StatusBar = "Ошибка получения информации о записи: " & Err.description
End Sub

' Добавление контекстного меню
Public Sub AddContextMenuButton()
    On Error GoTo MenuError
    
    ' Пробуем стандартный подход для совместимости
    Call RemoveContextMenuButton
    
    Dim contextMenu As CommandBar
    Set contextMenu = Application.CommandBars("Cell")
    
    ' Кнопка редактирования
    Dim editButton As CommandBarButton
    Set editButton = contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
    
    With editButton
        .Caption = "?? Редактировать в форме"
        .OnAction = "TableEventHandler.OpenFormForCurrentSelection"
        .FaceId = 162
        .BeginGroup = True
        .Tag = "CustomEdit_VhIsh_2021"
        .Visible = True
        .Enabled = True
    End With
    
    ' Кнопка дублирования
    Dim duplicateButton As CommandBarButton
    Set duplicateButton = contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
    
    With duplicateButton
        .Caption = "?? Дублировать запись"
        .OnAction = "TableEventHandler.DuplicateCurrentRecord"
        .FaceId = 19
        .BeginGroup = False
        .Tag = "CustomDuplicate_VhIsh_2021"
        .Visible = True
        .Enabled = True
    End With
    
    ' Альтернативное меню для Excel 2021
    Dim altMenuButton As CommandBarButton
    Set altMenuButton = contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
    
    With altMenuButton
        .Caption = "?? Меню действий (Ctrl+Shift+M)"
        .OnAction = "TableEventHandler.ShowActionMenuExcel2021"
        .FaceId = 923
        .BeginGroup = False
        .Tag = "CustomAltMenu_VhIsh_2021"
        .Visible = True
        .Enabled = True
    End With
    
    ' Кнопка поиска проводки
Dim provodkaButton As CommandBarButton
Set provodkaButton = contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)

With provodkaButton
    .Caption = "Найти проводку в 1С"
    .OnAction = "ProvodkaIntegrationModule.FindProvodkaForCurrentRecord"
    .FaceId = 1219
    .BeginGroup = False
    .Tag = "CustomProvodka_VhIsh_2025"
    .Visible = True
    .Enabled = True
End With
    
    
    
    ' Принудительное обновление для Excel 2021
    contextMenu.Reset
    
    Exit Sub
    
MenuError:
    Debug.Print "Ошибка создания контекстного меню: " & Err.description
End Sub

' Удаление контекстного меню
Public Sub RemoveContextMenuButton()
    On Error GoTo RemoveMenuError
    
    Dim contextMenu As CommandBar
    Dim ctrl As CommandBarControl
    
    Set contextMenu = Application.CommandBars("Cell")
    
    ' Удаляем все наши кнопки по тегам
    For Each ctrl In contextMenu.Controls
        If InStr(ctrl.Tag, "Custom") > 0 And InStr(ctrl.Tag, "VhIsh") > 0 Then
            ctrl.Delete
        End If
    Next ctrl
    
    ' Дополнительная очистка по названиям
    For Each ctrl In contextMenu.Controls
        If InStr(ctrl.Caption, "?? Редактировать") > 0 Or _
           InStr(ctrl.Caption, "?? Дублировать") > 0 Or _
           InStr(ctrl.Caption, "?? Меню действий") > 0 Then
            ctrl.Delete
        End If
    Next ctrl
    
    Exit Sub
    
RemoveMenuError:
    Debug.Print "Ошибка удаления контекстного меню: " & Err.description
End Sub

' Деактивация обработчика событий
Public Sub DeactivateTableEvents()
    On Error Resume Next
    
    isSystemActive = False
    contextMenuInitialized = False
    
    ' Деактивируем обработчик событий Application через класс
    If Not AppEventsHandler Is Nothing Then
        Call AppEventsHandler.DeactivateAppEvents
        Set AppEventsHandler = Nothing
    End If
    
    ' Безопасно очищаем объекты
    Set wsData = Nothing
    Set tblVhIsh = Nothing
    
    lastSelectedRow = 0
    lastSelectedColumn = 0
    currentRightClickRow = 0
    
    Application.StatusBar = "Система интерактивных форм деактивирована"
    
    On Error GoTo 0
End Sub

' Функция определения имени поля по номеру столбца
Public Function GetFieldNameByColumn(columnIndex As Long) As String
    On Error GoTo FieldNameError
    
    Select Case columnIndex
        Case 1: GetFieldNameByColumn = "txtNomerPP"
        Case 2: GetFieldNameByColumn = "cmbSlujba"
        Case 3: GetFieldNameByColumn = "cmbVidDocumenta"
        Case 4: GetFieldNameByColumn = "cmbVidDoc"
        Case 5: GetFieldNameByColumn = "txtNomerDoc"
        Case 6: GetFieldNameByColumn = "txtSummaDoc"
        Case 7: GetFieldNameByColumn = "txtVhFRP"
        Case 8: GetFieldNameByColumn = "txtDataVhFRP"
        Case 9: GetFieldNameByColumn = "cmbOtKogoPostupil"
        Case 10: GetFieldNameByColumn = "txtDataPeredachi"
        Case 11: GetFieldNameByColumn = "cmbIspolnitel"
        Case 12: GetFieldNameByColumn = "txtNomerIshVSlujbu"
        Case 13: GetFieldNameByColumn = "txtDataIshVSlujbu"
        Case 14: GetFieldNameByColumn = "txtNomerVozvrata"
        Case 15: GetFieldNameByColumn = "txtDataVozvrata"
        Case 16: GetFieldNameByColumn = "txtNomerIshKonvert"
        Case 17: GetFieldNameByColumn = "txtDataIshKonvert"
        Case 18: GetFieldNameByColumn = "txtOtmetkaIspolnenie"
        Case 19: GetFieldNameByColumn = "cmbStatusPodtverjdenie"
        Case 20: GetFieldNameByColumn = "txtNaryadInfo"
        Case Else: GetFieldNameByColumn = "txtNomerDoc"
    End Select
    
    Exit Function
    
FieldNameError:
    GetFieldNameByColumn = "txtNomerDoc"
End Function

' Функция получения отображаемого имени поля
Public Function GetFieldDisplayName(fieldName As String) As String
    On Error GoTo DisplayNameError
    
    Select Case fieldName
        Case "txtNomerPP": GetFieldDisplayName = "№ П/П"
        Case "cmbSlujba": GetFieldDisplayName = "Служба"
        Case "cmbVidDocumenta": GetFieldDisplayName = "Вид документа"
        Case "cmbVidDoc": GetFieldDisplayName = "Тип документа"
        Case "txtNomerDoc": GetFieldDisplayName = "Номер документа"
        Case "txtSummaDoc": GetFieldDisplayName = "Сумма документа"
        Case "txtVhFRP": GetFieldDisplayName = "Вх.ФРП/Исх.ФРП"
        Case "txtDataVhFRP": GetFieldDisplayName = "Дата Вх.ФРП/Исх.ФРП"
        Case "cmbOtKogoPostupil": GetFieldDisplayName = "От кого поступил"
        Case "txtDataPeredachi": GetFieldDisplayName = "Дата передачи исполнителю"
        Case "cmbIspolnitel": GetFieldDisplayName = "Исполнитель"
        Case "txtNomerIshVSlujbu": GetFieldDisplayName = "№ исх. в службу"
        Case "txtDataIshVSlujbu": GetFieldDisplayName = "Дата исх. в службу"
        Case "txtNomerVozvrata": GetFieldDisplayName = "№ возврата"
        Case "txtDataVozvrata": GetFieldDisplayName = "Дата возврата"
        Case "txtNomerIshKonvert": GetFieldDisplayName = "№ исх. конверт"
        Case "txtDataIshKonvert": GetFieldDisplayName = "Дата исх. конверт"
        Case "txtOtmetkaIspolnenie": GetFieldDisplayName = "Отметка об исполнении"
        Case "cmbStatusPodtverjdenie": GetFieldDisplayName = "Статус подтверждения"
        Case "txtNaryadInfo": GetFieldDisplayName = "Информация о наряде"
        Case Else: GetFieldDisplayName = "Поле данных"
    End Select
    
    Exit Function
    
DisplayNameError:
    GetFieldDisplayName = "Поле данных"
End Function

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Открытие формы по команде
Public Sub OpenFormForCurrentSelection()
    Dim currentCell As Range
    Dim intersectRange As Range
    Dim RowNumber As Long
    Dim columnNumber As Long
    Dim fieldToActivate As String
    
    On Error GoTo OpenError
    
    ' Проверки состояния системы (только статус-бар)
    If Not isSystemActive Then
        Application.StatusBar = "Система интерактивных форм не активна!"
        Exit Sub
    End If
    
    If wsData Is Nothing Or tblVhIsh Is Nothing Then
        Application.StatusBar = "Ошибка: объекты системы не инициализированы!"
        Exit Sub
    End If
    
    ' Проверяем активный лист
    If ActiveSheet.Name <> "ВхИсх" Then
        Application.StatusBar = "Перейдите на лист 'ВхИсх' для работы с таблицей."
        Exit Sub
    End If
    
    ' Проверка: Таблица не пуста
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = "Таблица пуста! Добавьте данные для работы с формой."
        Exit Sub
    End If
    
    Set currentCell = Selection
    
    ' Проверяем, что выделена одна ячейка в области таблицы
    If currentCell.Cells.Count = 1 Then
        Set intersectRange = Intersect(currentCell, tblVhIsh.DataBodyRange)
        
        If Not intersectRange Is Nothing Then
            ' Получаем координаты в таблице
            RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
            columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
            fieldToActivate = GetFieldNameByColumn(columnNumber)
            
            ' Открываем форму
            Call OpenFormWithRecord(RowNumber, fieldToActivate, columnNumber)
        Else
            Application.StatusBar = "Выберите ячейку внутри таблицы данных."
        End If
    Else
        Application.StatusBar = "Выберите одну ячейку в таблице для редактирования."
    End If
    
    Exit Sub
    
OpenError:
    Application.StatusBar = "Ошибка открытия формы: " & Err.description
End Sub

' Процедура открытия формы с активацией нужного поля
Private Sub OpenFormWithRecord(RowNumber As Long, fieldName As String, columnNumber As Long)
    On Error GoTo OpenError
    
    ' Проверяем, открыта ли уже форма
    Dim formAlreadyOpen As Boolean
    formAlreadyOpen = IsFormOpen("UserFormVhIsh")
    
    If Not formAlreadyOpen Then
        ' Загружаем форму
        Load UserFormVhIsh
        ' Показываем форму
        UserFormVhIsh.Show vbModeless ' Показываем в немодальном режиме для работы с таблицей
    End If
    
    ' Загружаем нужную запись
    Call UserFormVhIsh.LoadRecordToForm(RowNumber)
    
    ' ИСПРАВЛЕНО: Правильные обращения к глобальным переменным DataManager
    DataManager.CurrentRecordRow = RowNumber
    DataManager.IsNewRecord = False
    DataManager.FormDataChanged = False
    
    ' Активируем нужное поле с небольшой задержкой
    Application.OnTime Now + TimeValue("00:00:01"), "'TableEventHandler.ActivateFormField """ & fieldName & """'"
    
    ' Обновляем статус-бар формы
    UserFormVhIsh.lblStatusBar.Caption = "Загружена запись №" & RowNumber & " | Активно поле: " & GetFieldDisplayName(fieldName)
    
    ' Подсвечиваем активное поле
    Call HighlightActiveField(fieldName, columnNumber)
    
    Exit Sub
    
OpenError:
    Application.StatusBar = "Ошибка открытия формы: " & Err.description
End Sub

' Функция проверки, открыта ли форма
Public Function IsFormOpen(formName As String) As Boolean
    Dim frm As Object
    IsFormOpen = False
    
    On Error GoTo FormNotOpen
    Set frm = VBA.UserForms(formName)
    IsFormOpen = True
    
FormNotOpen:
    On Error GoTo 0
End Function

' Процедура активации поля формы (вызывается через OnTime)
Public Sub ActivateFormField(fieldName As String)
    On Error GoTo ActivateError
    
    ' Проверяем, что форма все еще открыта
    If IsFormOpen("UserFormVhIsh") Then
        ' Устанавливаем фокус на нужное поле
        Select Case fieldName
            Case "txtNomerPP": UserFormVhIsh.txtNomerPP.SetFocus
            Case "cmbSlujba": UserFormVhIsh.cmbSlujba.SetFocus
            Case "cmbVidDocumenta": UserFormVhIsh.cmbVidDocumenta.SetFocus
            Case "cmbVidDoc": UserFormVhIsh.cmbVidDoc.SetFocus
            Case "txtNomerDoc": UserFormVhIsh.txtNomerDoc.SetFocus
            Case "txtSummaDoc": UserFormVhIsh.txtSummaDoc.SetFocus
            Case "txtVhFRP": UserFormVhIsh.txtVhFRP.SetFocus
            Case "txtDataVhFRP": UserFormVhIsh.txtDataVhFRP.SetFocus
            Case "cmbOtKogoPostupil": UserFormVhIsh.cmbOtKogoPostupil.SetFocus
            Case "txtDataPeredachi": UserFormVhIsh.txtDataPeredachi.SetFocus
            Case "cmbIspolnitel": UserFormVhIsh.cmbIspolnitel.SetFocus
            Case "txtNomerIshVSlujbu": UserFormVhIsh.txtNomerIshVSlujbu.SetFocus
            Case "txtDataIshVSlujbu": UserFormVhIsh.txtDataIshVSlujbu.SetFocus
            Case "txtNomerVozvrata": UserFormVhIsh.txtNomerVozvrata.SetFocus
            Case "txtDataVozvrata": UserFormVhIsh.txtDataVozvrata.SetFocus
            Case "txtNomerIshKonvert": UserFormVhIsh.txtNomerIshKonvert.SetFocus
            Case "txtDataIshKonvert": UserFormVhIsh.txtDataIshKonvert.SetFocus
            Case "txtOtmetkaIspolnenie": UserFormVhIsh.txtOtmetkaIspolnenie.SetFocus
            Case "cmbStatusPodtverjdenie": UserFormVhIsh.cmbStatusPodtverjdenie.SetFocus
            Case "txtNaryadInfo": UserFormVhIsh.txtNaryadInfo.SetFocus
            Case Else: UserFormVhIsh.txtNomerDoc.SetFocus ' По умолчанию
        End Select
    End If
    
    Exit Sub
    
ActivateError:
    ' Игнорируем ошибки активации поля
End Sub

' Процедура подсветки активного поля
Private Sub HighlightActiveField(fieldName As String, columnIndex As Long)
    On Error GoTo HighlightError
    
    ' Проверки существования таблицы и столбца
    If tblVhIsh Is Nothing Then Exit Sub
    If columnIndex < 1 Or columnIndex > tblVhIsh.ListColumns.Count Then Exit Sub
    If tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    
    ' Временно подсвечиваем соответствующий столбец в таблице
    Dim columnRange As Range
    Set columnRange = tblVhIsh.ListColumns(columnIndex).DataBodyRange
    
    ' Применяем временную подсветку
    columnRange.Interior.Color = RGB(255, 255, 200) ' Светло-желтый
    
    ' Снимаем подсветку через 2 секунды
    Application.OnTime Now + TimeValue("00:00:02"), "'TableEventHandler.RemoveColumnHighlight " & columnIndex & "'"
    
    Exit Sub
    
HighlightError:
    ' Игнорируем ошибки подсветки
End Sub

' Процедура снятия подсветки столбца
Public Sub RemoveColumnHighlight(columnIndex As Long)
    On Error GoTo RemoveError
    
    ' Проверки существования таблицы и столбца
    If tblVhIsh Is Nothing Then Exit Sub
    If columnIndex < 1 Or columnIndex > tblVhIsh.ListColumns.Count Then Exit Sub
    If tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    
    Dim columnRange As Range
    Set columnRange = tblVhIsh.ListColumns(columnIndex).DataBodyRange
    columnRange.Interior.ColorIndex = xlNone ' Убираем подсветку
    
    Exit Sub
    
RemoveError:
    ' Игнорируем ошибки снятия подсветки
End Sub

' ОСТАВЛЕНО ПОДТВЕРЖДЕНИЕ: Дублирование текущей записи
Public Sub DuplicateCurrentRecord()
    Dim currentCell As Range
    Dim intersectRange As Range
    Dim RowNumber As Long
    
    On Error GoTo DuplicateError
    
    ' Проверяем состояние системы (статус-бар)
    If Not isSystemActive Then
        Application.StatusBar = "Система интерактивных форм не активна!"
        Exit Sub
    End If
    
    If wsData Is Nothing Or tblVhIsh Is Nothing Then
        Application.StatusBar = "Ошибка: объекты системы не инициализированы!"
        Exit Sub
    End If
    
    ' Проверяем активный лист
    If ActiveSheet.Name <> "ВхИсх" Then
        Application.StatusBar = "Перейдите на лист 'ВхИсх' для работы с таблицей."
        Exit Sub
    End If
    
    ' Проверяем, что таблица не пуста
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = "Таблица пуста! Нет записей для дублирования."
        Exit Sub
    End If
    
    Set currentCell = Selection
    
    ' Проверяем, что выделена одна ячейка в области таблицы
    If currentCell.Cells.Count = 1 Then
        Set intersectRange = Intersect(currentCell, tblVhIsh.DataBodyRange)
        
        If Not intersectRange Is Nothing Then
            ' Получаем номер строки в таблице
            RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
            
            ' ОСТАВЛЕНО: Подтверждение дублирования (основная функциональность)
            Dim response As VbMsgBoxResult
            response = MsgBox("Дублировать запись №" & RowNumber & "?" & vbCrLf & vbCrLf & _
                             "Будут скопированы все данные кроме:" & vbCrLf & _
                             "• Номера документа" & vbCrLf & _
                             "• Суммы документа" & vbCrLf & _
                             "• Информации о наряде", _
                             vbYesNo + vbQuestion, "Подтверждение дублирования")
            
            If response = vbYes Then
                Call PerformDuplication(RowNumber)
            End If
        Else
            Application.StatusBar = "Выберите ячейку внутри таблицы данных для дублирования."
        End If
    Else
        Application.StatusBar = "Выберите одну ячейку в таблице для дублирования записи."
    End If
    
    Exit Sub
    
DuplicateError:
    Application.StatusBar = "Ошибка дублирования записи: " & Err.description
End Sub

' ОСТАВЛЕНО УВЕДОМЛЕНИЕ: Выполнение дублирования записи
Private Sub PerformDuplication(sourceRowNumber As Long)
    Dim newRow As ListRow
    Dim sourceRowIndex As Long
    Dim newRowIndex As Long
    Dim i As Long
    
    On Error GoTo PerformDuplicationError
    
    ' Добавляем новую строку в таблицу
    Set newRow = tblVhIsh.ListRows.Add
    newRowIndex = newRow.Index
    sourceRowIndex = sourceRowNumber
    
    ' Копируем данные из исходной строки, исключая определенные поля
    For i = 1 To tblVhIsh.ListColumns.Count
        Select Case i
            Case 1  ' № П/П - устанавливаем новый номер
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = newRowIndex
                
            Case 5  ' txtNomerDoc - НЕ копируем (номер документа)
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = ""
                
            Case 6  ' txtSummaDoc - НЕ копируем (сумма документа)
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = 0
                
            Case 20 ' txtNaryadInfo - НЕ копируем (информация о наряде)
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = ""
                
            Case Else ' Все остальные поля - копируем
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = _
                    tblVhIsh.DataBodyRange.Cells(sourceRowIndex, i).value
        End Select
    Next i
    
    ' ОСТАВЛЕНО: Уведомление об успешном дублировании (основная функциональность)
    MsgBox "Запись №" & sourceRowNumber & " успешно продублирована!" & vbCrLf & _
           "Новая запись: №" & newRowIndex & vbCrLf & vbCrLf & _
           "Обратите внимание:" & vbCrLf & _
           "• Номер документа не скопирован (требует заполнения)" & vbCrLf & _
           "• Сумма документа обнулена (требует заполнения)" & vbCrLf & _
           "• Информация о наряде не скопирована", _
           vbInformation, "Дублирование завершено"
    
    ' Переходим к новой записи
    Dim newCellRange As Range
    Set newCellRange = tblVhIsh.DataBodyRange.Cells(newRowIndex, 1)
    newCellRange.Select
    
    ' Обновляем статус-бар
    Application.StatusBar = "Создана дубликат записи №" & newRowIndex & " на основе записи №" & sourceRowNumber
    
    Exit Sub
    
PerformDuplicationError:
    ' ОСТАВЛЕНО: Критические ошибки показываем
    MsgBox "Ошибка выполнения дублирования: " & Err.description, vbCritical, "Критическая ошибка"
    
    ' Пытаемся удалить созданную строку при ошибке
    On Error Resume Next
    If Not newRow Is Nothing Then
        newRow.Delete
    End If
    On Error GoTo 0
End Sub

' Обработка щелчка по ячейке из события листа
Public Sub ProcessCellClick(Target As Range)
    Dim intersectRange As Range
    Dim RowNumber As Long
    Dim columnNumber As Long
    Dim fieldToActivate As String
    
    On Error GoTo ProcessError
    
    ' Проверяем активность системы
    If Not isSystemActive Then Exit Sub
    
    ' Проверяем существование объектов
    If wsData Is Nothing Or tblVhIsh Is Nothing Then Exit Sub
    
    ' Проверяем, что таблица не пуста
    If tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    
    ' Проверяем, что щелчок в области данных таблицы
    Set intersectRange = Intersect(Target, tblVhIsh.DataBodyRange)
    
    If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
        ' Получаем координаты в таблице
        RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
        columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
        fieldToActivate = GetFieldNameByColumn(columnNumber)
        
        ' Обновляем отслеживание выделения
        lastSelectedRow = RowNumber
        lastSelectedColumn = columnNumber
        
        ' Обновляем статус-бар с информацией о горячих клавишах
        Application.StatusBar = "Строка " & RowNumber & ", столбец " & columnNumber & _
                               " | Ctrl+Shift+M: меню | Ctrl+E: редактировать | Ctrl+D: дублировать | " & _
                               GetFieldDisplayName(fieldToActivate)
    Else
        ' Сбрасываем если вышли из таблицы
        If lastSelectedRow <> 0 Then
            lastSelectedRow = 0
            lastSelectedColumn = 0
            Application.StatusBar = "Система интерактивных форм активна (Excel 2021). Используйте горячие клавиши."
        End If
    End If
    
    Exit Sub
    
ProcessError:
    ' Логируем ошибку но не прерываем работу
    Debug.Print "Ошибка обработки щелчка: " & Err.description
End Sub

' Открытие формы для конкретной ячейки (из двойного щелчка)
Public Sub OpenFormForSpecificCell(Target As Range)
    Dim intersectRange As Range
    Dim RowNumber As Long
    Dim columnNumber As Long
    Dim fieldToActivate As String
    
    On Error GoTo OpenSpecificError
    
    ' Проверяем активность системы (статус-бар)
    If Not isSystemActive Then
        Application.StatusBar = "Система интерактивных форм не активна!"
        Exit Sub
    End If
    
    ' Проверяем существование объектов
    If wsData Is Nothing Or tblVhIsh Is Nothing Then
        Application.StatusBar = "Ошибка: объекты системы не инициализированы!"
        Exit Sub
    End If
    
    ' Проверяем, что таблица не пуста
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = "Таблица пуста! Добавьте данные для работы с формой."
        Exit Sub
    End If
    
    ' Проверяем, что щелчок в области данных таблицы
    Set intersectRange = Intersect(Target, tblVhIsh.DataBodyRange)
    
    If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
        ' Получаем координаты в таблице
        RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
        columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
        fieldToActivate = GetFieldNameByColumn(columnNumber)
        
        ' Открываем форму
        Call OpenFormWithRecord(RowNumber, fieldToActivate, columnNumber)
    Else
        Application.StatusBar = "Выберите ячейку внутри таблицы данных."
    End If
    
    Exit Sub
    
OpenSpecificError:
    Application.StatusBar = "Ошибка открытия формы: " & Err.description
End Sub


