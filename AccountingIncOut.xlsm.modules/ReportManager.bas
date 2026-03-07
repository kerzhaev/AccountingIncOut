Attribute VB_Name = "ReportManager"
Option Explicit

Public Sub ShowReportsMenu()
    Dim menuText As String

    menuText = LocalizationManager.GetText("Reports are not implemented yet.") & vbCrLf & vbCrLf & _
               "1 - " & LocalizationManager.GetText("Status report") & vbCrLf & _
               "2 - " & LocalizationManager.GetText("Service report") & vbCrLf & _
               "3 - " & LocalizationManager.GetText("Delay report")

    MsgBox menuText, vbInformation, LocalizationManager.GetText("Reports")
    Call SafeLogReport("ShowReportsMenu", "Reports menu opened")
End Sub

Public Sub GenerateStatusReport()
    MsgBox LocalizationManager.GetText("Status report is not implemented yet."), vbInformation, LocalizationManager.GetText("Reports")
    Call SafeLogReport("GenerateStatusReport", "Status report requested")
End Sub

Public Sub GenerateServiceReport()
    MsgBox LocalizationManager.GetText("Service report is not implemented yet."), vbInformation, LocalizationManager.GetText("Reports")
    Call SafeLogReport("GenerateServiceReport", "Service report requested")
End Sub

Public Sub GenerateDelayReport()
    MsgBox LocalizationManager.GetText("Delay report is not implemented yet."), vbInformation, LocalizationManager.GetText("Reports")
    Call SafeLogReport("GenerateDelayReport", "Delay report requested")
End Sub

Private Sub SafeLogReport(ByVal operationName As String, ByVal description As String)
    On Error GoTo LogError

    Call SystemLogger.LogOperation(operationName, description, "INFO", 0)
    Exit Sub

LogError:
    Debug.Print "ReportManager log error in " & operationName & ": " & Err.description
End Sub
