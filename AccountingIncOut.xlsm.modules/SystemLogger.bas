Attribute VB_Name = "SystemLogger"
'==============================================
' LOGGING MODULE - SystemLogger
' Purpose: Logging system operations
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.0.1
' Date: 23.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

Private Const LOG_SHEET_NAME As String = "OperationLog"

' Initialize log
Public Sub InitializeLog()
    Dim wsLog As Worksheet
    
    On Error GoTo CreateLog
    
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    Exit Sub
    
CreateLog:
    Set wsLog = ThisWorkbook.Worksheets.Add
    wsLog.Name = LOG_SHEET_NAME
    wsLog.Visible = xlSheetVeryHidden
    
    ' Create log headers
    With wsLog
        .Range("A1").value = LocalizationManager.GetText("Date/Time")
        .Range("B1").value = LocalizationManager.GetText("Operation")
        .Range("C1").value = LocalizationManager.GetText("Description")
        .Range("D1").value = LocalizationManager.GetText("Status")
        .Range("E1").value = LocalizationManager.GetText("Execution Time (sec)")
        .Range("F1").value = LocalizationManager.GetText("User")
        
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 220, 255)
    End With
End Sub

' Log operation
Public Sub LogOperation(operationName As String, description As String, Status As String, executionTime As Double)
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim maxRecords As Long
    
    On Error GoTo LogError
    
    ' Check if logging is enabled
    If Not SystemSettings.GetSetting("LogEnabled", True) Then
        Exit Sub
    End If
    
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    maxRecords = SystemSettings.GetSetting("MaxLogRecords", 100)
    
    ' Find next empty row
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Write data
    With wsLog
        .Cells(nextRow, 1).value = Now()
        .Cells(nextRow, 2).value = operationName
        .Cells(nextRow, 3).value = description
        .Cells(nextRow, 4).value = Status
        .Cells(nextRow, 5).value = Round(executionTime, 2)
        .Cells(nextRow, 6).value = Application.UserName
    End With
    
    ' Apply conditional formatting by status
    Call FormatLogEntry(wsLog.Range(wsLog.Cells(nextRow, 1), wsLog.Cells(nextRow, 6)), Status)
    
    ' Clean old records if limit exceeded
    If nextRow - 1 > maxRecords Then
        Call CleanOldLogEntries(wsLog, maxRecords)
    End If
    
    Exit Sub
    
LogError:
    ' In case of logging error, write to Debug
    Debug.Print "LOG ERROR: " & Now() & " | " & operationName & " | " & description & " | " & Status
End Sub

' Format log entry by status
Private Sub FormatLogEntry(entryRange As Range, Status As String)
    Select Case UCase(Status)
        Case "SUCCESS"
            entryRange.Interior.Color = RGB(200, 255, 200) ' Light green
        Case "ERROR"
            entryRange.Interior.Color = RGB(255, 200, 200) ' Light red
        Case "WARNING"
            entryRange.Interior.Color = RGB(255, 255, 200) ' Light yellow
        Case "START"
            entryRange.Interior.Color = RGB(200, 200, 255) ' Light blue
        Case Else
            entryRange.Interior.ColorIndex = xlNone
    End Select
End Sub

' Clean old log entries
Private Sub CleanOldLogEntries(wsLog As Worksheet, maxRecords As Long)
    Dim TotalRows As Long
    Dim RowsToDelete As Long
    
    TotalRows = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row - 1 ' -1 for header
    RowsToDelete = TotalRows - maxRecords
    
    If RowsToDelete > 0 Then
        ' Delete old rows (starting from 2nd, as 1st is header)
        wsLog.Rows("2:" & (RowsToDelete + 1)).Delete
    End If
End Sub

' Show log dialog
Public Sub ShowLogDialog()
    Dim wsLog As Worksheet
    
    On Error GoTo ShowLogError
    
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    
    ' Temporarily make sheet visible
    wsLog.Visible = xlSheetVisible
    wsLog.Activate
    
    ' Autofit column width
    wsLog.Columns("A:F").AutoFit
    
    ' Apply filter
    If wsLog.Range("A1").AutoFilter = False Then
        wsLog.Range("A1:F1").AutoFilter
    End If
    
    MsgBox LocalizationManager.GetText("Operation log opened.") & vbCrLf & _
           LocalizationManager.GetText("Use filters to find necessary records.") & vbCrLf & _
           LocalizationManager.GetText("Press OK to hide the log after viewing."), _
           vbInformation, LocalizationManager.GetText("Operation Log")
    
    ' Hide sheet back
    wsLog.Visible = xlSheetVeryHidden
    
    Exit Sub
    
ShowLogError:
    MsgBox LocalizationManager.GetText("Error opening log: ") & Err.description, vbExclamation, LocalizationManager.GetText("Error")
End Sub

' Clear log
Public Sub ClearLog()
    Dim wsLog As Worksheet
    Dim response As VbMsgBoxResult
    
    response = MsgBox(LocalizationManager.GetText("Clear the entire operation log?") & vbCrLf & _
                      LocalizationManager.GetText("This action cannot be undone!"), _
                      vbYesNo + vbExclamation + vbDefaultButton2, LocalizationManager.GetText("Clear Confirmation"))
    
    If response = vbYes Then
        Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
        wsLog.Range("2:" & wsLog.Rows.Count).ClearContents
        MsgBox LocalizationManager.GetText("Operation log cleared."), vbInformation, LocalizationManager.GetText("Clearing Completed")
    End If
End Sub

