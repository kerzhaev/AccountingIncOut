Attribute VB_Name = "RibbonCallbacks"
'==============================================
' RIBBON CALLBACKS MODULE - RibbonCallbacks
' Purpose: Handling Office ribbon commands for IncOut system
' State: STAGE 1 - BASIC INFRASTRUCTURE
' Version: 1.1.0
' Date: 05.09.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

Public ribMain As IRibbonUI

' Callback for ribbon initialization
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set ribMain = ribbon
    Call SystemLogger.LogOperation("RibbonLoad", "IncOut ribbon loaded", "SUCCESS", 0)
End Sub

' GROUP: 1C Integration
'==============================================

' Command: Mass processing (existing function)
' FIXED version of mass processing handler
Public Sub btnMassProcess_Click(control As IRibbonControl)
    Dim StartTime As Double
    
    On Error GoTo ProcessError
    
    StartTime = Timer
    
    ' Log operation start (if SystemLogger module exists)
    On Error Resume Next
    Call SystemLogger.LogOperation("MassProcess", "Mass processing started", "START", StartTime)
    On Error GoTo ProcessError
    
    ' Create backup (if SystemBackup module exists)
    On Error Resume Next
    If Not SystemBackup.CreateBackup("MassProcess") Then
        MsgBox "Warning: Failed to create a backup." & vbCrLf & _
               "The operation will be executed without backup.", vbExclamation, "Warning"
    End If
    On Error GoTo ProcessError
    
    ' Show progress bar (if ProgressManager module exists)
    On Error Resume Next
    Call ProgressManager.ShowProgress("Mass processing of records", True)
    On Error GoTo ProcessError
    
    ' MAIN CALL - execute existing function
    Call ProvodkaIntegrationModule.MassProcessWithFileSelection
    
    ' Hide progress bar
    On Error Resume Next
    Call ProgressManager.HideProgress
    On Error GoTo ProcessError
    
    ' Log successful completion
    On Error Resume Next
    Call SystemLogger.LogOperation("MassProcess", "Mass processing completed", "SUCCESS", Timer - StartTime)
    On Error GoTo ProcessError
    
    Exit Sub
    
ProcessError:
    ' Hide progress bar on error
    On Error Resume Next
    Call ProgressManager.HideProgress
    On Error GoTo 0
    
    ' Detailed error diagnostics
    Dim errorMsg As String
    errorMsg = "Mass processing error:" & vbCrLf & vbCrLf & _
               "Error code: " & Err.Number & vbCrLf & _
               "Description: " & Err.description & vbCrLf & _
               "Source: " & Err.Source
    
    ' Log error
    On Error Resume Next
    Call SystemLogger.LogOperation("MassProcess", "ERROR: " & Err.description, "ERROR", Timer - StartTime)
    On Error GoTo 0
    
    ' Show detailed error to user
    MsgBox errorMsg, vbCritical, "Critical Mass Processing Error"
End Sub

' Command: 1C Integration Statistics
Public Sub btnStatistics_Click(control As IRibbonControl)
    Call ProvodkaIntegrationModule.ShowMatchingStatistics
    Call SystemLogger.LogOperation("Statistics", "Integration statistics shown", "INFO", 0)
End Sub

' Command: Clear execution marks
Public Sub btnClearMarks_Click(control As IRibbonControl)
    On Error GoTo ClearError
    
    ' Create backup before clearing
    If Not SystemBackup.CreateBackup("ClearMarks") Then
        MsgBox "Failed to create a backup. Operation cancelled.", vbCritical, "Backup Error"
        Exit Sub
    End If
    
    Call ProvodkaIntegrationModule.ClearAllProvodkaMarks
    Call SystemLogger.LogOperation("ClearMarks", "Execution marks cleared", "SUCCESS", 0)
    
    Exit Sub
    
ClearError:
    Call SystemLogger.LogOperation("ClearMarks", "Clear error: " & Err.description, "ERROR", 0)
    MsgBox "Error clearing marks: " & Err.description, vbCritical, "Error"
End Sub

' GROUP: Doverennosti Analysis
'==============================================

' ACTIVATED command: Doverennosti Matching
Public Sub btnMatchDoverennosti_Click(control As IRibbonControl)
    Dim StartTime As Double
    
    On Error GoTo MatchError
    
    StartTime = Timer
    
    ' Log start
    Call SystemLogger.LogOperation("MatchDoverennosti", "Starting Doverennosti matching", "START", StartTime)
    
    ' Run master scenario (master file + period export + 1C operations)
    Call DoverennostMatchingModule.ProcessDoverennostMasterMatching
    
    ' Log completion
    Call SystemLogger.LogOperation("MatchDoverennosti", "Doverennosti matching completed", "SUCCESS", Timer - StartTime)
    
    Exit Sub
    
MatchError:
    Call SystemLogger.LogOperation("MatchDoverennosti", "Error: " & Err.description, "ERROR", Timer - StartTime)
    MsgBox "Doverennosti matching error:" & vbCrLf & Err.description, vbCritical, "Error"
End Sub

' Command: Generate Word Letters v1.1.4
Public Sub btnGenerateWordLetters_Click(control As IRibbonControl)
    Dim StartTime As Double
    
    On Error GoTo GenerateError
    
    StartTime = Timer
    
    ' Log start
    On Error Resume Next
    Call SystemLogger.LogOperation("GenerateLetters", "Starting Word letters generation", "START", StartTime)
    On Error GoTo GenerateError
    
    ' Create backup (optional)
    On Error Resume Next
    If Not SystemBackup.CreateBackup("GenerateLetters") Then
        MsgBox "Warning: Failed to create a backup." & vbCrLf & _
               "The operation will be executed without backup.", vbExclamation, "Warning"
    End If
    On Error GoTo GenerateError
    
    ' Show progress bar
    On Error Resume Next
    Call ProgressManager.ShowProgress("Generating Word letters", True)
    On Error GoTo GenerateError
    
    ' MAIN CALL - execute letter generation
    Call ModuleWordLetters.GenerateWordLetters
    
    ' Hide progress bar
    On Error Resume Next
    Call ProgressManager.HideProgress
    On Error GoTo GenerateError
    
    ' Log successful completion
    On Error Resume Next
    Call SystemLogger.LogOperation("GenerateLetters", "Letters generation completed successfully", "SUCCESS", Timer - StartTime)
    On Error GoTo GenerateError
    
    Exit Sub
    
GenerateError:
    ' Hide progress bar on error
    On Error Resume Next
    Call ProgressManager.HideProgress
    On Error GoTo 0
    
    ' Detailed error diagnostics
    Dim errorMsg As String
    errorMsg = "Letter generation error:" & vbCrLf & vbCrLf & _
               "Error code: " & Err.Number & vbCrLf & _
               "Description: " & Err.description & vbCrLf & _
               "Source: " & Err.Source & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "- Presence of doverennosti and addresses files" & vbCrLf & _
               "- Word template correctness" & vbCrLf & _
               "- Write permissions for the selected folder"
    
    ' Log error
    On Error Resume Next
    Call SystemLogger.LogOperation("GenerateLetters", "ERROR: " & Err.description, "ERROR", Timer - StartTime)
    On Error GoTo 0
    
    ' Show detailed error to user
    MsgBox errorMsg, vbCritical, "Word Letters Generation Error v1.1.4"
End Sub

' GROUP: Settings and Reports
'==============================================

' Command: System Settings
Public Sub btnSettings_Click(control As IRibbonControl)
    Call SystemSettings.ShowSettingsDialog
    Call SystemLogger.LogOperation("Settings", "System settings opened", "INFO", 0)
End Sub

' Command: Operation Log
Public Sub btnShowLog_Click(control As IRibbonControl)
    Call SystemLogger.ShowLogDialog
    Call SystemLogger.LogOperation("ShowLog", "Operation log opened", "INFO", 0)
End Sub

' Command: Backups
Public Sub btnBackupManager_Click(control As IRibbonControl)
    Call SystemBackup.ShowBackupDialog
    Call SystemLogger.LogOperation("BackupManager", "Backup manager opened", "INFO", 0)
End Sub

' Command: System Help
Public Sub btnHelp_Click(control As IRibbonControl)
    Call ShowSystemHelp
    Call SystemLogger.LogOperation("Help", "System help opened", "INFO", 0)
End Sub

' Show system help
Private Sub ShowSystemHelp()
    Dim helpText As String
    
    helpText = "INCOMING AND OUTGOING DOCUMENTS SYSTEM" & vbCrLf & vbCrLf & _
               "[1C INTEGRATION]:" & vbCrLf & _
               "- Mass processing - automatic matching of postings" & vbCrLf & _
               "- Statistics - report on integration results" & vbCrLf & _
               "- Clear marks - reset for reprocessing" & vbCrLf & vbCrLf & _
               "[DOVERENNOSTI ANALYSIS]: (Stage 2)" & vbCrLf & _
               "- Matching doverennosti with operations" & vbCrLf & _
               "- Workflow analysis" & vbCrLf & vbCrLf & _
               "[SETTINGS AND REPORTS]:" & vbCrLf & _
               "- Settings - system configuration" & vbCrLf & _
               "- Log - journal of all operations" & vbCrLf & _
               "- Backups - backup management" & vbCrLf & vbCrLf & _
               "Version: 1.0.0 | Date: 23.08.2025" & vbCrLf & _
               "Author: Evgeniy Kerzhaev, FKU '95 FES' MO RF"
    
    MsgBox helpText, vbInformation, "IncOut System Help"
End Sub

' Command to refresh ribbon
Public Sub RefreshRibbon()
    If Not ribMain Is Nothing Then
        ribMain.Invalidate
    End If
End Sub

