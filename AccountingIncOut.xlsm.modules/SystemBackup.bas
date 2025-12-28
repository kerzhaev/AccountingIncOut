Attribute VB_Name = "SystemBackup"
'==============================================
' МОДУЛЬ РЕЗЕРВНОГО КОПИРОВАНИЯ - SystemBackup
' Назначение: Управление резервными копиями системы
' Состояние: ЭТАП 1 - БАЗОВАЯ ИНФРАСТРУКТУРА
' Версия: 1.0.0
' Дата: 23.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' Создание резервной копии
Public Function CreateBackup(operationName As String) As Boolean
    Dim BackupPath As String
    Dim backupFileName As String
    Dim fullBackupPath As String
    
    On Error GoTo BackupError
    
    CreateBackup = False
    
    ' Проверяем, включено ли резервное копирование
    If Not SystemSettings.GetSetting("BackupEnabled", True) Then
        CreateBackup = True ' Если отключено, считаем успешным
        Exit Function
    End If
    
    ' Получаем путь для резервных копий
    BackupPath = SystemSettings.GetSetting("BackupPath", ThisWorkbook.Path & "\Backup\")
    
    ' Создаем папку, если не существует
    If dir(BackupPath, vbDirectory) = "" Then
        MkDir BackupPath
    End If
    
    ' Формируем имя файла резервной копии
    backupFileName = "УчетВхИсх_" & operationName & "_" & Format(Now(), "YYYYMMDD_HHMMSS") & ".xlsm"
    fullBackupPath = BackupPath & backupFileName
    
    ' Сохраняем копию
    ThisWorkbook.SaveCopyAs fullBackupPath
    
    ' Очищаем старые резервные копии
    Call CleanOldBackups(BackupPath)
    
    CreateBackup = True
    
    Call SystemLogger.LogOperation("Backup", "Создана резервная копия: " & backupFileName, "SUCCESS", 0)
    
    Exit Function
    
BackupError:
    Call SystemLogger.LogOperation("Backup", "Ошибка создания копии: " & Err.description, "ERROR", 0)
    CreateBackup = False
End Function

' Очистка старых резервных копий
Private Sub CleanOldBackups(BackupPath As String)
    Dim maxBackups As Long
    Dim FileName As String
    Dim backupFiles() As String
    Dim fileCount As Long
    Dim i As Long
    
    On Error GoTo CleanError
    
    maxBackups = SystemSettings.GetSetting("MaxBackupCount", 10)
    
    ' Получаем список файлов резервных копий
    FileName = dir(BackupPath & "УчетВхИсх_*.xlsm")
    ReDim backupFiles(0)
    fileCount = 0
    
    Do While FileName <> ""
        ReDim Preserve backupFiles(fileCount)
        backupFiles(fileCount) = FileName
        fileCount = fileCount + 1
        FileName = dir
    Loop
    
    ' Если файлов больше максимума, удаляем старые
    If fileCount > maxBackups Then
        ' Сортируем файлы по имени (дате)
        Call QuickSort(backupFiles, 0, fileCount - 1)
        
        ' Удаляем старые файлы
        For i = 0 To fileCount - maxBackups - 1
            Kill BackupPath & backupFiles(i)
        Next i
        
        Call SystemLogger.LogOperation("Backup", "Удалено старых копий: " & (fileCount - maxBackups), "INFO", 0)
    End If
    
    Exit Sub
    
CleanError:
    Call SystemLogger.LogOperation("Backup", "Ошибка очистки старых копий: " & Err.description, "WARNING", 0)
End Sub

' Быстрая сортировка массива строк
Private Sub QuickSort(arr() As String, low As Long, high As Long)
    If low < high Then
        Dim pivotIndex As Long
        pivotIndex = Partition(arr, low, high)
        Call QuickSort(arr, low, pivotIndex - 1)
        Call QuickSort(arr, pivotIndex + 1, high)
    End If
End Sub

Private Function Partition(arr() As String, low As Long, high As Long) As Long
    Dim pivot As String
    Dim i As Long, j As Long
    Dim temp As String
    
    pivot = arr(high)
    i = low - 1
    
    For j = low To high - 1
        If arr(j) <= pivot Then
            i = i + 1
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
        End If
    Next j
    
    temp = arr(i + 1)
    arr(i + 1) = arr(high)
    arr(high) = temp
    
    Partition = i + 1
End Function

' Показ диалога управления резервными копиями
Public Sub ShowBackupDialog()
    Dim BackupPath As String
    Dim backupList As String
    Dim FileName As String
    Dim fileCount As Long
    
    On Error GoTo ShowError
    
    BackupPath = SystemSettings.GetSetting("BackupPath", ThisWorkbook.Path & "\Backup\")
    
    ' Получаем список резервных копий
    FileName = dir(BackupPath & "УчетВхИсх_*.xlsm")
    backupList = "СПИСОК РЕЗЕРВНЫХ КОПИЙ:" & vbCrLf & vbCrLf
    fileCount = 0
    
    Do While FileName <> ""
        fileCount = fileCount + 1
        backupList = backupList & fileCount & ". " & FileName & vbCrLf
        FileName = dir
    Loop
    
    If fileCount = 0 Then
        backupList = backupList & "Резервные копии не найдены."
    Else
        backupList = backupList & vbCrLf & "Всего копий: " & fileCount & vbCrLf & _
                     "Путь: " & BackupPath
    End If
    
    MsgBox backupList, vbInformation, "Менеджер резервных копий"
    
    Exit Sub
    
ShowError:
    MsgBox "Ошибка просмотра резервных копий: " & Err.description, vbExclamation, "Ошибка"
End Sub


