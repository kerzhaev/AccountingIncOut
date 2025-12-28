Attribute VB_Name = "SystemSettings"
'==============================================
' МОДУЛЬ НАСТРОЕК СИСТЕМЫ - SystemSettings
' Назначение: Управление настройками системы на скрытом листе
' Состояние: ЭТАП 1 - БАЗОВАЯ ИНФРАСТРУКТУРА
' Версия: 1.0.0
' Дата: 23.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

Private Const SETTINGS_SHEET_NAME As String = "СистемныеНастройки"

' Инициализация листа настроек
Public Sub InitializeSettings()
    Dim wsSettings As Worksheet
    
    On Error GoTo CreateSettings
    
    ' Проверяем существование листа настроек
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    Exit Sub
    
CreateSettings:
    ' Создаем лист настроек
    Set wsSettings = ThisWorkbook.Worksheets.Add
    wsSettings.Name = SETTINGS_SHEET_NAME
    wsSettings.Visible = xlSheetVeryHidden
    
    ' Заполняем настройки по умолчанию
    Call SetDefaultSettings(wsSettings)
End Sub

' Установка настроек по умолчанию
Private Sub SetDefaultSettings(Ws As Worksheet)
    With Ws
        .Cells.Clear
        
        ' Заголовки
        .Range("A1").value = "Параметр"
        .Range("B1").value = "Значение"
        .Range("C1").value = "Описание"
        
        ' Настройки интеграции с 1С
        .Range("A3").value = "BackupEnabled"
        .Range("B3").value = True
        .Range("C3").value = "Создавать резервные копии перед операциями"
        
        .Range("A4").value = "MaxBackupCount"
        .Range("B4").value = 10
        .Range("C4").value = "Максимальное количество резервных копий"
        
        .Range("A5").value = "BackupPath"
        .Range("B5").value = ThisWorkbook.Path & "\Backup\"
        .Range("C5").value = "Путь для сохранения резервных копий"
        
        .Range("A6").value = "LogEnabled"
        .Range("B6").value = True
        .Range("C6").value = "Включить журналирование операций"
        
        .Range("A7").value = "MaxLogRecords"
        .Range("B7").value = 100
        .Range("C7").value = "Максимальное количество записей в журнале"
        
        .Range("A8").value = "ProgressUpdateInterval"
        .Range("B8").value = 100
        .Range("C8").value = "Интервал обновления прогресс-бара (записей)"
        
        .Range("A9").value = "DefaultFileFormat"
        .Range("B9").value = "*.xlsx,*.csv"
        .Range("C9").value = "Форматы файлов по умолчанию"
        
        ' Настройки сопоставления (для этапа 2)
        .Range("A11").value = "MatchThreshold"
        .Range("B11").value = 75
        .Range("C11").value = "Порог совпадения для автосопоставления (%)"
        
        .Range("A12").value = "DateTolerance"
        .Range("B12").value = 30
        .Range("C12").value = "Допустимая разница в датах (дней)"
        
        .Range("A13").value = "AutoSelectBestMatch"
        .Range("B13").value = True
        .Range("C13").value = "Автоматически выбирать лучшее совпадение"
        
        ' Форматирование
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C13").Borders.LineStyle = xlContinuous
        .Columns("A:C").AutoFit
    End With
End Sub

' Получение значения настройки
Public Function GetSetting(parameterName As String, Optional defaultValue As Variant = "") As Variant
    Dim wsSettings As Worksheet
    Dim foundRange As Range
    
    On Error GoTo GetError
    
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    Set foundRange = wsSettings.Columns("A:A").Find(parameterName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRange Is Nothing Then
        GetSetting = wsSettings.Cells(foundRange.Row, 2).value
    Else
        GetSetting = defaultValue
    End If
    
    Exit Function
    
GetError:
    GetSetting = defaultValue
End Function

' Установка значения настройки
Public Sub SetSetting(parameterName As String, value As Variant)
    Dim wsSettings As Worksheet
    Dim foundRange As Range
    
    On Error GoTo SetError
    
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    Set foundRange = wsSettings.Columns("A:A").Find(parameterName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRange Is Nothing Then
        wsSettings.Cells(foundRange.Row, 2).value = value
    Else
        ' Добавляем новую настройку
        Dim LastRow As Long
        LastRow = wsSettings.Cells(wsSettings.Rows.Count, 1).End(xlUp).Row + 1
        wsSettings.Cells(LastRow, 1).value = parameterName
        wsSettings.Cells(LastRow, 2).value = value
    End If
    
    Exit Sub
    
SetError:
    MsgBox "Ошибка сохранения настройки: " & Err.description, vbExclamation, "Ошибка настроек"
End Sub

' Показ диалога настроек
Public Sub ShowSettingsDialog()
    Dim wsSettings As Worksheet
    
    On Error GoTo ShowError
    
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    
    ' Временно делаем лист видимым для редактирования
    wsSettings.Visible = xlSheetVisible
    wsSettings.Activate
    
    MsgBox "Лист настроек открыт для редактирования." & vbCrLf & _
           "После внесения изменений нажмите OK для скрытия листа.", _
           vbInformation, "Редактирование настроек"
    
    ' Скрываем лист обратно
    wsSettings.Visible = xlSheetVeryHidden
    
    Exit Sub
    
ShowError:
    MsgBox "Ошибка открытия настроек: " & Err.description, vbExclamation, "Ошибка"
End Sub


