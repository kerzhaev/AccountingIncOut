Attribute VB_Name = "ProgressManager"
'==============================================
' PROGRESS MANAGEMENT MODULE - ProgressManager
' Purpose: Displaying progress of long operations
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.0.2
' Date: 23.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

Private ProgressForm As UserFormProgress
Private operationCancelled As Boolean
Private StartTime As Double

' Show progress bar
Public Sub ShowProgress(operationTitle As String, Optional allowCancel As Boolean = True)
    Set ProgressForm = New UserFormProgress
    
    With ProgressForm
        .Caption = LocalizationManager.GetText("Executing operation")
        .lblOperation.Caption = operationTitle
        .ProgressBar.value = 0
        .btnCancel.Visible = allowCancel
        .Show vbModeless
    End With
    
    operationCancelled = False
    StartTime = Timer
    
    ' Update Excel status bar
    Application.StatusBar = operationTitle & " - 0%"
End Sub

' Update progress
Public Sub UpdateProgress(CurrentValue As Long, MaxValue As Long, Optional statusText As String = "")
    Dim percentage As Double
    
    If ProgressForm Is Nothing Then Exit Sub
    
    percentage = (CurrentValue / MaxValue) * 100
    
    With ProgressForm
        .ProgressBar.value = percentage
        .lblStatus.Caption = LocalizationManager.GetText("Completed: ") & CurrentValue & LocalizationManager.GetText(" of ") & MaxValue & " (" & Format(percentage, "0.0") & "%)"
        
        If statusText <> "" Then
            .lblDetails.Caption = statusText
        End If
        
        ' Update execution time
        .lblTime.Caption = LocalizationManager.GetText("Time: ") & Format(Timer - StartTime, "0.0") & LocalizationManager.GetText(" sec.")
    End With
    
    ' Update Excel status bar
    Application.StatusBar = ProgressForm.lblOperation.Caption & " - " & Format(percentage, "0.0") & "%"
    
    ' Process Windows events
    DoEvents
End Sub

' Hide progress bar
Public Sub HideProgress()
    If Not ProgressForm Is Nothing Then
        Unload ProgressForm
        Set ProgressForm = Nothing
    End If
    
    Application.StatusBar = False
End Sub

' Check operation cancellation
Public Function IsOperationCancelled() As Boolean
    IsOperationCancelled = operationCancelled
End Function

' Cancel operation (called from form)
Public Sub CancelOperation()
    operationCancelled = True
    Call HideProgress
    Application.StatusBar = LocalizationManager.GetText("Operation cancelled by user")
End Sub

' FIXED: Renamed function for quick status bar update
Public Sub UpdateProgressStatusBar(Message As String, Optional percentage As Double = -1)
    If percentage >= 0 Then
        Application.StatusBar = Message & " - " & Format(percentage, "0.0") & "%"
    Else
        Application.StatusBar = Message
    End If
End Sub
