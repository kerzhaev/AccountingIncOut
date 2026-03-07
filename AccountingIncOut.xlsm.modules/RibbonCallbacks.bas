Attribute VB_Name = "RibbonCallbacks"
'==============================================
' RIBBON CALLBACKS MODULE - RibbonCallbacks
' Purpose: Handling Office ribbon commands for IncOut system
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.1.2
' Date: 08.03.2026
'==============================================

Option Explicit

Public ribMain As IRibbonUI

' Callback for ribbon initialization
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set ribMain = ribbon
    Call SafeLogOperation("RibbonLoad", LocalizationManager.GetText("IncOut ribbon loaded"), "SUCCESS", 0)
End Sub

' GROUP: 1C Integration
'==============================================

Public Sub btnMassProcess_Click(control As IRibbonControl)
    Dim StartTime As Double

    On Error GoTo ProcessError

    StartTime = Timer
    Call SafeLogOperation("MassProcess", LocalizationManager.GetText("Mass processing started"), "START", StartTime)

    If Not SafeCreateBackup("MassProcess") Then
        MsgBox LocalizationManager.GetText("Warning: Failed to create a backup.") & vbCrLf & _
               LocalizationManager.GetText("The operation will be executed without backup."), vbExclamation, LocalizationManager.GetText("Warning")
    End If

    Call SafeShowProgress(LocalizationManager.GetText("Mass processing of records"), True)
    Call ProvodkaIntegrationModule.MassProcessWithFileSelection
    Call SafeHideProgress

    Call SafeLogOperation("MassProcess", LocalizationManager.GetText("Mass processing completed"), "SUCCESS", Timer - StartTime)
    Exit Sub

ProcessError:
    Call SafeHideProgress

    Dim errorMsg As String
    errorMsg = LocalizationManager.GetText("Mass processing error:") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Error code: ") & Err.Number & vbCrLf & _
               LocalizationManager.GetText("Description: ") & Err.description & vbCrLf & _
               LocalizationManager.GetText("Source: ") & Err.Source

    Call SafeLogOperation("MassProcess", "ERROR: " & Err.description, "ERROR", Timer - StartTime)
    MsgBox errorMsg, vbCritical, LocalizationManager.GetText("Critical Mass Processing Error")
End Sub

Public Sub btnStatistics_Click(control As IRibbonControl)
    Call ProvodkaIntegrationModule.ShowMatchingStatistics
    Call SafeLogOperation("Statistics", LocalizationManager.GetText("Integration statistics shown"), "INFO", 0)
End Sub

Public Sub btnClearMarks_Click(control As IRibbonControl)
    On Error GoTo ClearError

    If Not SystemBackup.CreateBackup("ClearMarks") Then
        MsgBox LocalizationManager.GetText("Failed to create a backup. Operation cancelled."), vbCritical, LocalizationManager.GetText("Backup Error")
        Exit Sub
    End If

    Call ProvodkaIntegrationModule.ClearAllProvodkaMarks
    Call SafeLogOperation("ClearMarks", LocalizationManager.GetText("Execution marks cleared"), "SUCCESS", 0)
    Exit Sub

ClearError:
    Call SafeLogOperation("ClearMarks", LocalizationManager.GetText("Clear error: ") & Err.description, "ERROR", 0)
    MsgBox LocalizationManager.GetText("Error clearing marks: ") & Err.description, vbCritical, LocalizationManager.GetText("Error")
End Sub

' GROUP: Doverennosti Analysis
'==============================================

Public Sub btnMatchDoverennosti_Click(control As IRibbonControl)
    Dim StartTime As Double

    On Error GoTo MatchError

    StartTime = Timer
    Call SafeLogOperation("MatchDoverennosti", LocalizationManager.GetText("Starting Doverennosti matching"), "START", StartTime)

    Call DoverennostMatchingModule.ProcessDoverennostMasterMatching

    Call SafeLogOperation("MatchDoverennosti", LocalizationManager.GetText("Doverennosti matching completed"), "SUCCESS", Timer - StartTime)
    Exit Sub

MatchError:
    Call SafeLogOperation("MatchDoverennosti", "Error: " & Err.description, "ERROR", Timer - StartTime)
    MsgBox LocalizationManager.GetText("Doverennosti matching error:") & vbCrLf & Err.description, vbCritical, LocalizationManager.GetText("Error")
End Sub

Public Sub btnGenerateWordLetters_Click(control As IRibbonControl)
    Dim StartTime As Double

    On Error GoTo GenerateError

    StartTime = Timer
    Call SafeLogOperation("GenerateLetters", LocalizationManager.GetText("Starting Word letters generation"), "START", StartTime)

    If Not SafeCreateBackup("GenerateLetters") Then
        MsgBox LocalizationManager.GetText("Warning: Failed to create a backup.") & vbCrLf & _
               LocalizationManager.GetText("The operation will be executed without backup."), vbExclamation, LocalizationManager.GetText("Warning")
    End If

    Call SafeShowProgress(LocalizationManager.GetText("Generating Word letters"), True)
    Call ModuleWordLetters.GenerateWordLetters
    Call SafeHideProgress

    Call SafeLogOperation("GenerateLetters", LocalizationManager.GetText("Letters generation completed successfully"), "SUCCESS", Timer - StartTime)
    Exit Sub

GenerateError:
    Call SafeHideProgress

    Dim errorMsg As String
    errorMsg = LocalizationManager.GetText("Letter generation error:") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Error code: ") & Err.Number & vbCrLf & _
               LocalizationManager.GetText("Description: ") & Err.description & vbCrLf & _
               LocalizationManager.GetText("Source: ") & Err.Source & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Please check:") & vbCrLf & _
               LocalizationManager.GetText("- Presence of doverennosti and addresses files") & vbCrLf & _
               LocalizationManager.GetText("- Word template correctness") & vbCrLf & _
               LocalizationManager.GetText("- Write permissions for the selected folder")

    Call SafeLogOperation("GenerateLetters", "ERROR: " & Err.description, "ERROR", Timer - StartTime)
    MsgBox errorMsg, vbCritical, LocalizationManager.GetText("Word Letters Generation Error")
End Sub

' GROUP: Settings and Reports
'==============================================

Public Sub btnSettings_Click(control As IRibbonControl)
    Call SystemSettings.ShowSettingsDialog
    Call SafeLogOperation("Settings", LocalizationManager.GetText("System settings opened"), "INFO", 0)
End Sub

Public Sub btnShowLog_Click(control As IRibbonControl)
    Call SystemLogger.ShowLogDialog
    Call SafeLogOperation("ShowLog", LocalizationManager.GetText("Operation log opened"), "INFO", 0)
End Sub

Public Sub btnBackupManager_Click(control As IRibbonControl)
    Call SystemBackup.ShowBackupDialog
    Call SafeLogOperation("BackupManager", LocalizationManager.GetText("Backup manager opened"), "INFO", 0)
End Sub

Public Sub btnHelp_Click(control As IRibbonControl)
    Call ShowSystemHelp
    Call SafeLogOperation("Help", LocalizationManager.GetText("System help opened"), "INFO", 0)
End Sub

Private Sub ShowSystemHelp()
    Dim helpText As String

    helpText = LocalizationManager.GetText("INCOMING AND OUTGOING DOCUMENTS SYSTEM") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("[1C INTEGRATION]:") & vbCrLf & _
               LocalizationManager.GetText("- Mass processing - automatic matching of postings") & vbCrLf & _
               LocalizationManager.GetText("- Statistics - report on integration results") & vbCrLf & _
               LocalizationManager.GetText("- Clear marks - reset for reprocessing") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("[DOVERENNOSTI ANALYSIS]: (Stage 2)") & vbCrLf & _
               LocalizationManager.GetText("- Matching doverennosti with operations") & vbCrLf & _
               LocalizationManager.GetText("- Workflow analysis") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("[SETTINGS AND REPORTS]:") & vbCrLf & _
               LocalizationManager.GetText("- Settings - system configuration") & vbCrLf & _
               LocalizationManager.GetText("- Log - journal of all operations") & vbCrLf & _
               LocalizationManager.GetText("- Backups - backup management") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Version: ") & "1.1.2" & " | " & LocalizationManager.GetText("Date: ") & "08.03.2026" & vbCrLf & _
               LocalizationManager.GetText("Author: Evgeniy Kerzhaev, FKU '95 FES' MO RF")

    MsgBox helpText, vbInformation, LocalizationManager.GetText("IncOut System Help")
End Sub

Public Sub btnAnalyzeOperations_Click(control As IRibbonControl)
    Call ProvodkaIntegrationModule.ShowMatchingStatistics
    Call SafeLogOperation("Analyze", LocalizationManager.GetText("Analysis operations shown"), "INFO", 0)
End Sub

Public Sub RefreshRibbon()
    If Not ribMain Is Nothing Then
        ribMain.Invalidate
    End If
End Sub

Private Sub SafeLogOperation(ByVal operationName As String, ByVal description As String, ByVal statusText As String, ByVal executionTime As Double)
    On Error GoTo LogError

    Call SystemLogger.LogOperation(operationName, description, statusText, executionTime)
    Exit Sub

LogError:
    Debug.Print "Ribbon log error in " & operationName & ": " & Err.description
End Sub

Private Function SafeCreateBackup(ByVal backupTag As String) As Boolean
    On Error GoTo BackupError

    SafeCreateBackup = SystemBackup.CreateBackup(backupTag)
    Exit Function

BackupError:
    Debug.Print "Ribbon backup error in " & backupTag & ": " & Err.description
    SafeCreateBackup = False
End Function

Private Sub SafeShowProgress(ByVal operationTitle As String, Optional ByVal allowCancel As Boolean = True)
    On Error GoTo ProgressError

    Call ProgressManager.ShowProgress(operationTitle, allowCancel)
    Exit Sub

ProgressError:
    Debug.Print "Ribbon progress show error: " & Err.description
End Sub

Private Sub SafeHideProgress()
    On Error GoTo ProgressError

    Call ProgressManager.HideProgress
    Exit Sub

ProgressError:
    Debug.Print "Ribbon progress hide error: " & Err.description
End Sub
