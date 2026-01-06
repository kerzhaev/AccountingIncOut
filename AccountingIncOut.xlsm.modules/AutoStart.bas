Attribute VB_Name = "AutoStart"
'==============================================
' МОДУЛЬ АВТОЗАПУСКА СИСТЕМЫ - AutoStart
' Назначение: Автоматическая инициализация системы при открытии книги
' Состояние: ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА - ТОЛЬКО СТАТУС-БАР И DEBUG
' Версия: 2.3.0
' Дата: 10.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' Глобальная переменная для отслеживания состояния системы
Public SystemInitialized As Boolean

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Основная инициализация системы
Public Sub InitializeSystem()
    On Error GoTo InitError
    
    ' Сброс флага инициализации
    SystemInitialized = False
    
    ' Информируем пользователя через статус-бар (без MsgBox)
    Application.StatusBar = "Инициализация системы интерактивных форм..."
    
    ' 1. Инициализация обработчика событий таблицы
    Call TableEventHandler.InitializeTableEvents
    
    ' 2. Принудительная активация контекстного меню
    Call ForceActivateContextMenu
    
    ' 3. Проверка состояния системы (только в Debug)
    Call DiagnoseSystemState
    
    ' 4. Активация листа с таблицей
    On Error Resume Next
    ThisWorkbook.Worksheets("ВхИсх").Activate
    On Error GoTo InitError
    
    ' 5. Установка флага успешной инициализации
    SystemInitialized = True
    
    ' ОТКЛЮЧЕНО ВСПЛЫВАЮЩЕЕ ОКНО: Только статус-бар
    Application.StatusBar = "Система интерактивных форм активна. Готова к работе."
    
    ' ОТКЛЮЧЕНО: ShowUserInstructions - убираем всплывающие инструкции
    ' Debug-информация для разработчика
    Debug.Print "Система интерактивных форм успешно инициализирована"
    
    Exit Sub
    
InitError:
    SystemInitialized = False
    ' ОСТАВЛЕНО: Критические ошибки показываем через MsgBox
    MsgBox "КРИТИЧЕСКАЯ ОШИБКА инициализации системы!" & vbCrLf & _
           "Ошибка: " & Err.Number & " - " & Err.description, vbCritical, "Системная ошибка"
    
    Application.StatusBar = "ОШИБКА инициализации системы."
End Sub

' Принудительная активация контекстного меню
Public Sub ForceActivateContextMenu()
    On Error GoTo MenuError
    
    ' Удаляем существующие кнопки (если есть)
    Call TableEventHandler.RemoveContextMenuButton
    
    ' Небольшая пауза
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Добавляем контекстное меню
    Call TableEventHandler.AddContextMenuButton
    
    ' Проверяем успешность добавления
    If Not CheckContextMenuExists() Then
        ' Повторная попытка
        Application.Wait Now + TimeValue("00:00:01")
        Call TableEventHandler.AddContextMenuButton
    End If
    
    Exit Sub
    
MenuError:
    Debug.Print "Ошибка принудительной активации контекстного меню: " & Err.description
End Sub

' Проверка существования контекстного меню
Public Function CheckContextMenuExists() As Boolean
    Dim contextMenu As CommandBar
    Dim ctrl As CommandBarControl
    
    On Error GoTo CheckError
    
    Set contextMenu = Application.CommandBars("Cell")
    
    For Each ctrl In contextMenu.Controls
        If ctrl.Caption = "Дублировать запись" Or InStr(ctrl.Caption, "Дублировать") > 0 Then
            CheckContextMenuExists = True
            Exit Function
        End If
    Next ctrl
    
    CheckContextMenuExists = False
    Exit Function
    
CheckError:
    CheckContextMenuExists = False
End Function

' Диагностика состояния системы (только Debug, без MsgBox)
Public Sub DiagnoseSystemState()
    Dim diagResult As String
    
    diagResult = "ДИАГНОСТИКА СИСТЕМЫ:" & vbCrLf & vbCrLf
    
    ' Проверка листа
    On Error Resume Next
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("ВхИсх")
    If Ws Is Nothing Then
        diagResult = diagResult & "? Лист 'ВхИсх' НЕ найден!" & vbCrLf
    Else
        diagResult = diagResult & "? Лист 'ВхИсх' найден" & vbCrLf
        
        ' Проверка таблицы
        Dim tbl As ListObject
        Set tbl = Ws.ListObjects("ВходящиеИсходящие")
        If tbl Is Nothing Then
            diagResult = diagResult & "? Таблица 'ВходящиеИсходящие' НЕ найдена!" & vbCrLf
        Else
            diagResult = diagResult & "? Таблица найдена (" & tbl.ListRows.Count & " строк)" & vbCrLf
        End If
    End If
    On Error GoTo 0
    
    ' Проверка контекстного меню
    If CheckContextMenuExists() Then
        diagResult = diagResult & "? Контекстное меню активно" & vbCrLf
    Else
        diagResult = diagResult & "? Контекстное меню НЕ активно!" & vbCrLf
    End If
    
    ' ОТКЛЮЧЕНО ВСПЛЫВАЮЩЕЕ ОКНО: Только записываем результат в Debug
    Debug.Print diagResult
End Sub

' ОТКЛЮЧЕНО ВСПЛЫВАЮЩЕЕ ОКНО: Показ инструкций пользователю
' Public Sub ShowUserInstructions()
'     ' Эта процедура отключена для предотвращения всплывающих окон
'     Debug.Print "Инструкции отключены. Система готова к работе."
' End Sub

' Ручная инициализация системы (без всплывающих окон)
Public Sub ManualInitializeSystem()
    Application.StatusBar = "Ручная инициализация системы..."
    Call InitializeSystem
    Application.StatusBar = "Ручная инициализация завершена."
End Sub

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Принудительная перезагрузка контекстного меню
Public Sub ReloadContextMenu()
    On Error GoTo ReloadError
    
    Application.StatusBar = "Перезагрузка контекстного меню..."
    
    ' Удаляем
    Call TableEventHandler.RemoveContextMenuButton
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Добавляем
    Call TableEventHandler.AddContextMenuButton
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Проверяем (только статус-бар)
    If CheckContextMenuExists() Then
        Application.StatusBar = "? Контекстное меню успешно перезагружено!"
        Debug.Print "Контекстное меню успешно перезагружено"
    Else
        Application.StatusBar = "? Ошибка перезагрузки контекстного меню!"
        Debug.Print "Ошибка перезагрузки контекстного меню"
    End If
    
    Exit Sub
    
ReloadError:
    Application.StatusBar = "Ошибка перезагрузки: " & Err.description
    Debug.Print "Ошибка перезагрузки: " & Err.description
End Sub

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Тест контекстного меню
Public Sub TestContextMenu()
    Dim result As String
    
    result = "ТЕСТ КОНТЕКСТНОГО МЕНЮ:" & vbCrLf & vbCrLf
    
    ' Активируем лист
    On Error Resume Next
    ThisWorkbook.Worksheets("ВхИсх").Activate
    On Error GoTo 0
    
    ' Проверяем наличие меню
    If CheckContextMenuExists() Then
        result = result & "? Контекстное меню найдено в системе" & vbCrLf
        result = result & "Щелкните ПРАВОЙ кнопкой по ячейке ВНУТРИ таблицы"
        Application.StatusBar = "Контекстное меню найдено - готово к использованию"
    Else
        result = result & "? Контекстное меню НЕ найдено!" & vbCrLf
        result = result & "Выполняю принудительную активацию..."
        
        Call ForceActivateContextMenu
        
        If CheckContextMenuExists() Then
            result = result & vbCrLf & "? Активация успешна!"
            Application.StatusBar = "Контекстное меню активировано успешно"
        Else
            result = result & vbCrLf & "? Активация не удалась!"
            Application.StatusBar = "Ошибка активации контекстного меню"
        End If
    End If
    
    ' ОТКЛЮЧЕНО ВСПЛЫВАЮЩЕЕ ОКНО: Только Debug и статус-бар
    Debug.Print result
End Sub

' Деактивация системы (переименована из DeactivateSystem)
Public Sub ShutdownSystem()
    On Error Resume Next
    
    ' Деактивируем обработчик событий
    Call TableEventHandler.DeactivateTableEvents
    
    ' Удаляем контекстное меню
    Call TableEventHandler.RemoveContextMenuButton
    
    ' Сбрасываем флаг системы
    SystemInitialized = False
    
    ' Обновляем статус
    Application.StatusBar = "Система интерактивных форм завершена"
    Debug.Print "Система корректно завершена"
    
    On Error GoTo 0
End Sub

' Альтернативное имя для совместимости
Public Sub DeactivateSystem()
    ' Вызываем основную процедуру завершения
    Call ShutdownSystem
End Sub

' Функция проверки состояния системы
Public Function IsSystemReady() As Boolean
    IsSystemReady = SystemInitialized And CheckContextMenuExists()
End Function

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Полная диагностика для пользователя
Public Sub FullDiagnostic()
    Dim Report As String
    
    Report = "?? ПОЛНАЯ ДИАГНОСТИКА СИСТЕМЫ:" & vbCrLf & vbCrLf
    
    ' 1. Проверка инициализации
    Report = Report & "Система инициализирована: " & IIf(SystemInitialized, "? ДА", "? НЕТ") & vbCrLf
    
    ' 2. Проверка листа и таблицы
    On Error Resume Next
    Dim Ws As Worksheet
    Dim tbl As ListObject
    Set Ws = ThisWorkbook.Worksheets("ВхИсх")
    
    If Not Ws Is Nothing Then
        Report = Report & "Лист 'ВхИсх': ? Найден" & vbCrLf
        Set tbl = Ws.ListObjects("ВходящиеИсходящие")
        If Not tbl Is Nothing Then
            Report = Report & "Таблица: ? Найдена (" & tbl.ListRows.Count & " строк)" & vbCrLf
        Else
            Report = Report & "Таблица: ? НЕ найдена!" & vbCrLf
        End If
    Else
        Report = Report & "Лист 'ВхИсх': ? НЕ найден!" & vbCrLf
    End If
    On Error GoTo 0
    
    ' 3. Проверка контекстного меню
    Report = Report & "Контекстное меню: " & IIf(CheckContextMenuExists(), "? Активно", "? НЕ активно") & vbCrLf
    
    ' 4. Проверка активного листа
    Report = Report & "Активный лист: " & ActiveSheet.Name & vbCrLf
    
    ' 5. Версия Office
    Report = Report & "Версия Office: " & Application.Version & vbCrLf
    
    ' ОТКЛЮЧЕНО ВСПЛЫВАЮЩЕЕ ОКНО: Только Debug и статус-бар
    Debug.Print Report
    Application.StatusBar = "Диагностика завершена. См. результаты в Debug (Ctrl+G)"
End Sub

' Экстренное завершение системы (без всплывающих окон)
Public Sub EmergencyShutdown()
    On Error Resume Next
    
    ' Принудительная очистка всех ресурсов
    Call TableEventHandler.RemoveContextMenuButton
    Call TableEventHandler.DeactivateTableEvents
    
    ' Сброс всех переменных состояния
    SystemInitialized = False
    
    ' Очистка статус-бара
    Application.StatusBar = "Экстренное завершение системы выполнено"
    Debug.Print "Экстренное завершение системы выполнено. Все ресурсы освобождены."
    
    On Error GoTo 0
End Sub

' Проверка целостности системы
Public Function SystemIntegrityCheck() As Boolean
    Dim allGood As Boolean
    
    On Error GoTo IntegrityError
    
    allGood = True
    
    ' Проверка 1: Лист существует
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("ВхИсх")
    If Ws Is Nothing Then allGood = False
    
    ' Проверка 2: Таблица существует
    If Not Ws Is Nothing Then
        Dim tbl As ListObject
        Set tbl = Ws.ListObjects("ВходящиеИсходящие")
        If tbl Is Nothing Then allGood = False
    End If
    
    ' Проверка 3: Модули существуют
    ' (Если VBA дошел до этой точки, модули существуют)
    
    SystemIntegrityCheck = allGood
    Exit Function
    
IntegrityError:
    SystemIntegrityCheck = False
End Function

' ОТКЛЮЧЕНЫ ВСПЛЫВАЮЩИЕ ОКНА: Быстрый перезапуск системы
Public Sub RestartSystem()
    Application.StatusBar = "Перезапуск системы интерактивных форм..."
    Debug.Print "Начало перезапуска системы"
    
    ' Завершаем текущую систему
    Call ShutdownSystem
    
    ' Пауза
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Запускаем заново
    Call InitializeSystem
    
    Application.StatusBar = "Система успешно перезапущена!"
    Debug.Print "Система успешно перезапущена!"
End Sub


