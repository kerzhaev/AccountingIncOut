Attribute VB_Name = "RibbonCallbacks"
'==============================================
' МОДУЛЬ ОБРАБОТЧИКОВ ЛЕНТЫ - RibbonCallbacks
' Назначение: Обработка команд ленты Office для системы УчётВхИсх
' Состояние: ЭТАП 1 - БАЗОВАЯ ИНФРАСТРУКТУРА
' Версия: 1.1.0
' Дата: 05.09.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

Public ribMain As IRibbonUI

' Callback для инициализации ленты
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set ribMain = ribbon
    Call SystemLogger.LogOperation("RibbonLoad", "Лента УчётВхИсх загружена", "SUCCESS", 0)
End Sub

' ГРУППА: Интеграция с 1С
'==============================================

' Команда: Массовая обработка (существующая функция)
' ИСПРАВЛЕННАЯ версия обработчика массовой обработки
Public Sub btnMassProcess_Click(control As IRibbonControl)
    Dim StartTime As Double
    
    On Error GoTo ProcessError
    
    StartTime = Timer
    
    ' Логируем начало операции (если модуль SystemLogger существует)
    On Error Resume Next
    Call SystemLogger.LogOperation("MassProcess", "Начало массовой обработки", "START", StartTime)
    On Error GoTo ProcessError
    
    ' Создаем резервную копию (если модуль SystemBackup существует)
    On Error Resume Next
    If Not SystemBackup.CreateBackup("MassProcess") Then
        MsgBox "Предупреждение: Не удалось создать резервную копию." & vbCrLf & _
               "Операция будет выполнена без резервирования.", vbExclamation, "Предупреждение"
    End If
    On Error GoTo ProcessError
    
    ' Показываем прогресс-бар (если модуль ProgressManager существует)
    On Error Resume Next
    Call ProgressManager.ShowProgress("Массовая обработка записей", True)
    On Error GoTo ProcessError
    
    ' ОСНОВНОЙ ВЫЗОВ - вызываем существующую функцию
    Call ProvodkaIntegrationModule.MassProcessWithFileSelection
    
    ' Скрываем прогресс-бар
    On Error Resume Next
    Call ProgressManager.HideProgress
    On Error GoTo ProcessError
    
    ' Логируем успешное завершение
    On Error Resume Next
    Call SystemLogger.LogOperation("MassProcess", "Массовая обработка завершена", "SUCCESS", Timer - StartTime)
    On Error GoTo ProcessError
    
    Exit Sub
    
ProcessError:
    ' Скрываем прогресс-бар при ошибке
    On Error Resume Next
    Call ProgressManager.HideProgress
    On Error GoTo 0
    
    ' Детальная диагностика ошибки
    Dim errorMsg As String
    errorMsg = "Ошибка массовой обработки:" & vbCrLf & vbCrLf & _
               "Код ошибки: " & Err.Number & vbCrLf & _
               "Описание: " & Err.description & vbCrLf & _
               "Источник: " & Err.Source
    
    ' Логируем ошибку
    On Error Resume Next
    Call SystemLogger.LogOperation("MassProcess", "ОШИБКА: " & Err.description, "ERROR", Timer - StartTime)
    On Error GoTo 0
    
    ' Показываем детальную ошибку пользователю
    MsgBox errorMsg, vbCritical, "Критическая ошибка массовой обработки"
End Sub


' Команда: Статистика интеграции с 1С
Public Sub btnStatistics_Click(control As IRibbonControl)
    Call ProvodkaIntegrationModule.ShowMatchingStatistics
    Call SystemLogger.LogOperation("Statistics", "Показана статистика интеграции", "INFO", 0)
End Sub

' Команда: Очистка отметок об исполнении
Public Sub btnClearMarks_Click(control As IRibbonControl)
    On Error GoTo ClearError
    
    ' Создаем резервную копию перед очисткой
    If Not SystemBackup.CreateBackup("ClearMarks") Then
        MsgBox "Не удалось создать резервную копию. Операция отменена.", vbCritical, "Ошибка резервирования"
        Exit Sub
    End If
    
    Call ProvodkaIntegrationModule.ClearAllProvodkaMarks
    Call SystemLogger.LogOperation("ClearMarks", "Очищены отметки об исполнении", "SUCCESS", 0)
    
    Exit Sub
    
ClearError:
    Call SystemLogger.LogOperation("ClearMarks", "Ошибка очистки: " & Err.description, "ERROR", 0)
    MsgBox "Ошибка очистки отметок: " & Err.description, vbCritical, "Ошибка"
End Sub

' ГРУППА: Анализ доверенностей
'==============================================

' АКТИВИРОВАННАЯ команда: Сопоставление доверенностей
Public Sub btnMatchDoverennosti_Click(control As IRibbonControl)
    Dim StartTime As Double
    
    On Error GoTo MatchError
    
    StartTime = Timer
    
    ' Логируем начало
    Call SystemLogger.LogOperation("MatchDoverennosti", "Запуск сопоставления доверенностей", "START", StartTime)
    
    ' Запускаем мастер-сценарий (мастер-файл + выгрузка периода + операции 1С)
    Call DoverennostMatchingModule.ProcessDoverennostMasterMatching
    
    ' Логируем завершение
    Call SystemLogger.LogOperation("MatchDoverennosti", "Сопоставление доверенностей завершено", "SUCCESS", Timer - StartTime)
    
    Exit Sub
    
   
    
MatchError:
    Call SystemLogger.LogOperation("MatchDoverennosti", "Ошибка: " & Err.description, "ERROR", Timer - StartTime)
    MsgBox "Ошибка сопоставления доверенностей:" & vbCrLf & Err.description, vbCritical, "Ошибка"
End Sub



 ' Команда: Формирование писем по доверенностям v1.1.4
Public Sub btnGenerateWordLetters_Click(control As IRibbonControl)
    Dim StartTime As Double
    
    On Error GoTo GenerateError
    
    StartTime = Timer
    
    ' Логируем начало операции
    On Error Resume Next
    Call SystemLogger.LogOperation("GenerateLetters", "Начало формирования писем по доверенностям", "START", StartTime)
    On Error GoTo GenerateError
    
    ' Создаем резервную копию (опционально)
    On Error Resume Next
    If Not SystemBackup.CreateBackup("GenerateLetters") Then
        MsgBox "Предупреждение: Не удалось создать резервную копию." & vbCrLf & _
               "Операция будет выполнена без резервирования.", vbExclamation, "Предупреждение"
    End If
    On Error GoTo GenerateError
    
    ' Показываем прогресс-бар
    On Error Resume Next
    Call ProgressManager.ShowProgress("Формирование писем по доверенностям", True)
    On Error GoTo GenerateError
    
    ' ОСНОВНОЙ ВЫЗОВ - запускаем формирование писем
    Call GenerateWordLetters
    
    ' Скрываем прогресс-бар
    On Error Resume Next
    Call ProgressManager.HideProgress
    On Error GoTo GenerateError
    
    ' Логируем успешное завершение
    On Error Resume Next
    Call SystemLogger.LogOperation("GenerateLetters", "Формирование писем завершено успешно", "SUCCESS", Timer - StartTime)
    On Error GoTo GenerateError
    
    Exit Sub
    
GenerateError:
    ' Скрываем прогресс-бар при ошибке
    On Error Resume Next
    Call ProgressManager.HideProgress
    On Error GoTo 0
    
    ' Детальная диагностика ошибки
    Dim errorMsg As String
    errorMsg = "Ошибка формирования писем по доверенностям:" & vbCrLf & vbCrLf & _
               "Код ошибки: " & Err.Number & vbCrLf & _
               "Описание: " & Err.description & vbCrLf & _
               "Источник: " & Err.Source & vbCrLf & vbCrLf & _
               "Проверьте:" & vbCrLf & _
               "• Наличие файлов доверенностей и адресов" & vbCrLf & _
               "• Корректность шаблона Word" & vbCrLf & _
               "• Права на запись в выбранную папку"
    
    ' Логируем ошибку
    On Error Resume Next
    Call SystemLogger.LogOperation("GenerateLetters", "ОШИБКА: " & Err.description, "ERROR", Timer - StartTime)
    On Error GoTo 0
    
    ' Показываем детальную ошибку пользователю
    MsgBox errorMsg, vbCritical, "Ошибка формирования писем по доверенностям v1.1.4"
End Sub

    
    


' ГРУППА: Настройки и отчеты
'==============================================

' Команда: Настройки системы
Public Sub btnSettings_Click(control As IRibbonControl)
    Call SystemSettings.ShowSettingsDialog
    Call SystemLogger.LogOperation("Settings", "Открыты настройки системы", "INFO", 0)
End Sub

' Команда: Журнал операций
Public Sub btnShowLog_Click(control As IRibbonControl)
    Call SystemLogger.ShowLogDialog
    Call SystemLogger.LogOperation("ShowLog", "Открыт журнал операций", "INFO", 0)
End Sub

' Команда: Резервные копии
Public Sub btnBackupManager_Click(control As IRibbonControl)
    Call SystemBackup.ShowBackupDialog
    Call SystemLogger.LogOperation("BackupManager", "Открыт менеджер резервных копий", "INFO", 0)
End Sub

' Команда: Справка по системе
Public Sub btnHelp_Click(control As IRibbonControl)
    Call ShowSystemHelp
    Call SystemLogger.LogOperation("Help", "Открыта справка системы", "INFO", 0)
End Sub

' Показ справки по системе
Private Sub ShowSystemHelp()
    Dim helpText As String
    
    helpText = "СИСТЕМА УЧЁТА ВХОДЯЩИХ И ИСХОДЯЩИХ ДОКУМЕНТОВ" & vbCrLf & vbCrLf & _
               "?? ИНТЕГРАЦИЯ С 1С:" & vbCrLf & _
               "• Массовая обработка - автоматическое сопоставление проводок" & vbCrLf & _
               "• Статистика - отчет по результатам интеграции" & vbCrLf & _
               "• Очистка отметок - сброс для повторной обработки" & vbCrLf & vbCrLf & _
               "?? АНАЛИЗ ДОВЕРЕННОСТЕЙ: (Этап 2)" & vbCrLf & _
               "• Сопоставление доверенностей с операциями" & vbCrLf & _
               "• Анализ документооборота" & vbCrLf & vbCrLf & _
               "?? НАСТРОЙКИ И ОТЧЕТЫ:" & vbCrLf & _
               "• Настройки - конфигурация системы" & vbCrLf & _
               "• Журнал - лог всех операций" & vbCrLf & _
               "• Резервные копии - управление бэкапами" & vbCrLf & vbCrLf & _
               "Версия: 1.0.0 | Дата: 23.08.2025" & vbCrLf & _
               "Автор: Кержаев Евгений, ФКУ '95 ФЭС' МО РФ"
    
    MsgBox helpText, vbInformation, "Справка по системе УчётВхИсх"
End Sub

' Команда обновления ленты
Public Sub RefreshRibbon()
    If Not ribMain Is Nothing Then
        ribMain.Invalidate
    End If
End Sub


