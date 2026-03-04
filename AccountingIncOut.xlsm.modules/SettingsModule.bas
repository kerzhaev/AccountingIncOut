Attribute VB_Name = "SettingsModule"
'==============================================
' SYSTEM SETTINGS MANAGEMENT MODULE - SettingsModule
' Purpose: Saving and loading user settings
' State: FIXED REFERENCES TO GLOBAL VARIABLES
' Version: 1.1.0
' Date: 10.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

Public Sub SaveSettings()
    Dim wsSettings As Worksheet
    
    ' Create or get hidden settings sheet
    Set wsSettings = GetSettingsWorksheet()
    
    With wsSettings
        .Cells(1, 1).value = "CurrentRecordRow"
        .Cells(1, 2).value = DataManager.CurrentRecordRow
        
        .Cells(2, 1).value = "FormTop"
        .Cells(2, 2).value = UserFormVhIsh.Top
        
        .Cells(3, 1).value = "FormLeft"
        .Cells(3, 2).value = UserFormVhIsh.Left
        
        .Cells(4, 1).value = "LastSaved"
        .Cells(4, 2).value = Now()
    End With
End Sub

Public Sub LoadSettings()
    Dim wsSettings As Worksheet
    
    On Error GoTo DefaultSettings
    
    Set wsSettings = GetSettingsWorksheet()
    
    With wsSettings
        If .Cells(1, 2).value <> "" Then
            DataManager.CurrentRecordRow = .Cells(1, 2).value
        End If
        
        If .Cells(2, 2).value <> "" Then
            UserFormVhIsh.Top = .Cells(2, 2).value
        End If
        
        If .Cells(3, 2).value <> "" Then
            UserFormVhIsh.Left = .Cells(3, 2).value
        End If
    End With
    
    ' Load last record
    If DataManager.CurrentRecordRow > 0 Then
        Call NavigationModule.NavigateToRecord(DataManager.CurrentRecordRow)
    Else
        Call DataManager.ClearForm
    End If
    
    Exit Sub
    
DefaultSettings:
    DataManager.CurrentRecordRow = 0
    DataManager.IsNewRecord = True
    DataManager.FormDataChanged = False
End Sub

Private Function GetSettingsWorksheet() As Worksheet
    Dim Ws As Worksheet
    
    On Error GoTo CreateSheet
    
    Set Ws = ThisWorkbook.Worksheets("System_Settings")
    Set GetSettingsWorksheet = Ws
    Exit Function
    
CreateSheet:
    Set Ws = ThisWorkbook.Worksheets.Add
    Ws.Name = "System_Settings"
    Ws.Visible = xlSheetHidden
    Set GetSettingsWorksheet = Ws
End Function
