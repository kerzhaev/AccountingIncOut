Attribute VB_Name = "SystemSettings"
'==============================================
' SYSTEM SETTINGS MODULE - SystemSettings
' Purpose: Managing system settings on a hidden sheet
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.0.1
' Date: 23.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

Private Const SETTINGS_SHEET_NAME As String = "SystemSettings"

' Initialize settings sheet
Public Sub InitializeSettings()
    Dim wsSettings As Worksheet
    
    On Error GoTo CreateSettings
    
    ' Check if settings sheet exists
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    Exit Sub
    
CreateSettings:
    ' Create settings sheet
    Set wsSettings = ThisWorkbook.Worksheets.Add
    wsSettings.Name = SETTINGS_SHEET_NAME
    wsSettings.Visible = xlSheetVeryHidden
    
    ' Fill default settings
    Call SetDefaultSettings(wsSettings)
End Sub

' Set default settings
Private Sub SetDefaultSettings(Ws As Worksheet)
    With Ws
        .Cells.Clear
        
        ' Headers
        .Range("A1").value = LocalizationManager.GetText("Parameter")
        .Range("B1").value = LocalizationManager.GetText("Value")
        .Range("C1").value = LocalizationManager.GetText("Description")
        
        ' 1C Integration settings
        .Range("A3").value = "BackupEnabled"
        .Range("B3").value = True
        .Range("C3").value = LocalizationManager.GetText("Create backups before operations")
        
        .Range("A4").value = "MaxBackupCount"
        .Range("B4").value = 10
        .Range("C4").value = LocalizationManager.GetText("Maximum number of backup copies")
        
        .Range("A5").value = "BackupPath"
        .Range("B5").value = ThisWorkbook.Path & "\Backup\"
        .Range("C5").value = LocalizationManager.GetText("Path to save backup copies")
        
        .Range("A6").value = "LogEnabled"
        .Range("B6").value = True
        .Range("C6").value = LocalizationManager.GetText("Enable operation logging")
        
        .Range("A7").value = "MaxLogRecords"
        .Range("B7").value = 100
        .Range("C7").value = LocalizationManager.GetText("Maximum number of log records")
        
        .Range("A8").value = "ProgressUpdateInterval"
        .Range("B8").value = 100
        .Range("C8").value = LocalizationManager.GetText("Progress bar update interval (records)")
        
        .Range("A9").value = "DefaultFileFormat"
        .Range("B9").value = "*.xlsx,*.csv"
        .Range("C9").value = LocalizationManager.GetText("Default file formats")
        
        ' Matching settings (for stage 2)
        .Range("A11").value = "MatchThreshold"
        .Range("B11").value = 75
        .Range("C11").value = LocalizationManager.GetText("Match threshold for auto-matching (%)")
        
        .Range("A12").value = "DateTolerance"
        .Range("B12").value = 30
        .Range("C12").value = LocalizationManager.GetText("Allowed date difference (days)")
        
        .Range("A13").value = "AutoSelectBestMatch"
        .Range("B13").value = True
        .Range("C13").value = LocalizationManager.GetText("Automatically select best match")
        
        ' Formatting
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C13").Borders.LineStyle = xlContinuous
        .Columns("A:C").AutoFit
    End With
End Sub

' Get setting value
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

' Set setting value
Public Sub SetSetting(parameterName As String, value As Variant)
    Dim wsSettings As Worksheet
    Dim foundRange As Range
    
    On Error GoTo SetError
    
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    Set foundRange = wsSettings.Columns("A:A").Find(parameterName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRange Is Nothing Then
        wsSettings.Cells(foundRange.Row, 2).value = value
    Else
        ' Add new setting
        Dim LastRow As Long
        LastRow = wsSettings.Cells(wsSettings.Rows.Count, 1).End(xlUp).Row + 1
        wsSettings.Cells(LastRow, 1).value = parameterName
        wsSettings.Cells(LastRow, 2).value = value
    End If
    
    Exit Sub
    
SetError:
    MsgBox LocalizationManager.GetText("Error saving setting: ") & Err.description, vbExclamation, LocalizationManager.GetText("Settings Error")
End Sub

' Show settings dialog
Public Sub ShowSettingsDialog()
    Dim wsSettings As Worksheet
    
    On Error GoTo ShowError
    
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    
    ' Temporarily make sheet visible for editing
    wsSettings.Visible = xlSheetVisible
    wsSettings.Activate
    
    MsgBox LocalizationManager.GetText("Settings sheet opened for editing.") & vbCrLf & _
           LocalizationManager.GetText("After making changes, press OK to hide the sheet."), _
           vbInformation, LocalizationManager.GetText("Edit Settings")
    
    ' Hide sheet back
    wsSettings.Visible = xlSheetVeryHidden
    
    Exit Sub
    
ShowError:
    MsgBox LocalizationManager.GetText("Error opening settings: ") & Err.description, vbExclamation, LocalizationManager.GetText("Error")
End Sub

