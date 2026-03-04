Attribute VB_Name = "SystemBackup"
'==============================================
' SYSTEM BACKUP MODULE - SystemBackup
' Purpose: Managing system backups
' State: STAGE 1 - BASIC INFRASTRUCTURE
' Version: 1.0.0
' Date: 23.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' Create a backup
Public Function CreateBackup(operationName As String) As Boolean
    Dim BackupPath As String
    Dim backupFileName As String
    Dim fullBackupPath As String
    
    On Error GoTo BackupError
    
    CreateBackup = False
    
    ' Check if backup is enabled
    If Not SystemSettings.GetSetting("BackupEnabled", True) Then
        CreateBackup = True ' If disabled, consider it successful
        Exit Function
    End If
    
    ' Get backup path
    BackupPath = SystemSettings.GetSetting("BackupPath", ThisWorkbook.Path & "\Backup\")
    
    ' Create folder if it doesn't exist
    If dir(BackupPath, vbDirectory) = "" Then
        MkDir BackupPath
    End If
    
    ' Format backup file name
    backupFileName = "IncOut_" & operationName & "_" & Format(Now(), "YYYYMMDD_HHMMSS") & ".xlsm"
    fullBackupPath = BackupPath & backupFileName
    
    ' Save copy
    ThisWorkbook.SaveCopyAs fullBackupPath
    
    ' Clean old backups
    Call CleanOldBackups(BackupPath)
    
    CreateBackup = True
    
    Call SystemLogger.LogOperation("Backup", "Created backup: " & backupFileName, "SUCCESS", 0)
    
    Exit Function
    
BackupError:
    Call SystemLogger.LogOperation("Backup", "Backup creation error: " & Err.description, "ERROR", 0)
    CreateBackup = False
End Function

' Clean old backups
Private Sub CleanOldBackups(BackupPath As String)
    Dim maxBackups As Long
    Dim FileName As String
    Dim backupFiles() As String
    Dim fileCount As Long
    Dim i As Long
    
    On Error GoTo CleanError
    
    maxBackups = SystemSettings.GetSetting("MaxBackupCount", 10)
    
    ' Get list of backup files
    FileName = dir(BackupPath & "IncOut_*.xlsm")
    ReDim backupFiles(0)
    fileCount = 0
    
    Do While FileName <> ""
        ReDim Preserve backupFiles(fileCount)
        backupFiles(fileCount) = FileName
        fileCount = fileCount + 1
        FileName = dir
    Loop
    
    ' If files exceed maximum, delete oldest
    If fileCount > maxBackups Then
        ' Sort files by name (date)
        Call QuickSort(backupFiles, 0, fileCount - 1)
        
        ' Delete oldest files
        For i = 0 To fileCount - maxBackups - 1
            Kill BackupPath & backupFiles(i)
        Next i
        
        Call SystemLogger.LogOperation("Backup", "Deleted old copies: " & (fileCount - maxBackups), "INFO", 0)
    End If
    
    Exit Sub
    
CleanError:
    Call SystemLogger.LogOperation("Backup", "Error cleaning old copies: " & Err.description, "WARNING", 0)
End Sub

' Quick sort for string array
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

' Show backup manager dialog
Public Sub ShowBackupDialog()
    Dim BackupPath As String
    Dim backupList As String
    Dim FileName As String
    Dim fileCount As Long
    
    On Error GoTo ShowError
    
    BackupPath = SystemSettings.GetSetting("BackupPath", ThisWorkbook.Path & "\Backup\")
    
    ' Get backup list
    FileName = dir(BackupPath & "IncOut_*.xlsm")
    backupList = "LIST OF BACKUPS:" & vbCrLf & vbCrLf
    fileCount = 0
    
    Do While FileName <> ""
        fileCount = fileCount + 1
        backupList = backupList & fileCount & ". " & FileName & vbCrLf
        FileName = dir
    Loop
    
    If fileCount = 0 Then
        backupList = backupList & "No backups found."
    Else
        backupList = backupList & vbCrLf & "Total backups: " & fileCount & vbCrLf & _
                     "Path: " & BackupPath
    End If
    
    MsgBox backupList, vbInformation, "Backup Manager"
    
    Exit Sub
    
ShowError:
    MsgBox "Error viewing backups: " & Err.description, vbExclamation, "Error"
End Sub

