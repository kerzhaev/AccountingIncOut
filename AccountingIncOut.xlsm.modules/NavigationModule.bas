Attribute VB_Name = "NavigationModule"
'==============================================
' RECORD NAVIGATION MODULE - NavigationModule
' Purpose: Functions for navigating between table records
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.6.2
' Date: 10.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

Public Sub NavigateToRecord(RowNumber As Long)
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo NavigationError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    ' Bounds check
    If tblData.ListRows.Count = 0 Then
        Call DataManager.ClearForm
        Exit Sub
    End If
    
    If RowNumber < 1 Then RowNumber = 1
    If RowNumber > tblData.ListRows.Count Then RowNumber = tblData.ListRows.Count
    
    ' Check for unsaved changes
    If DataManager.HasUnsavedChanges() And DataManager.CurrentRecordRow <> RowNumber And DataManager.CurrentRecordRow > 0 Then
        Dim response As VbMsgBoxResult
        response = MsgBox(LocalizationManager.GetText("Save changes to the current record before navigating?"), vbYesNo + vbQuestion, LocalizationManager.GetText("Unsaved Changes"))
        If response = vbYes Then
            Call DataManager.SaveCurrentRecord
        End If
    End If
    
    DataManager.CurrentRecordRow = RowNumber
    DataManager.IsNewRecord = False
    
    ' Call procedure from form
    Call UserFormVhIsh.LoadRecordToForm(RowNumber)
    Call UpdateNavigationButtons
    Call UpdateStatusBar
    
    DataManager.FormDataChanged = False
    
    Exit Sub
    
NavigationError:
    MsgBox LocalizationManager.GetText("Navigation error: ") & Err.description, vbCritical, LocalizationManager.GetText("Error")
End Sub

Public Sub NavigateToPrevious()
    Call NavigateToRecord(DataManager.CurrentRecordRow - 1)
End Sub

Public Sub NavigateToNext()
    Call NavigateToRecord(DataManager.CurrentRecordRow + 1)
End Sub

Public Sub NavigateToFirst()
    Call NavigateToRecord(1)
End Sub

Public Sub NavigateToLast()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo LastError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    Call NavigateToRecord(tblData.ListRows.Count)
    
    Exit Sub
    
LastError:
    Call NavigateToRecord(1)
End Sub

Public Sub NavigateToRecordByNumber(recordNumber As Long)
    ' Navigate to record by Sequence Number (Seq No)
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    
    On Error GoTo FindError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    ' Find record with specified Seq No
    For i = 1 To tblData.ListRows.Count
        If CLng(tblData.DataBodyRange.Cells(i, 1).value) = recordNumber Then
            Call NavigateToRecord(i)
            Exit Sub
        End If
    Next i
    
    MsgBox LocalizationManager.GetText("Record with Seq No ") & recordNumber & LocalizationManager.GetText(" not found!"), vbExclamation, LocalizationManager.GetText("Record Search")
    
    Exit Sub
    
FindError:
    MsgBox LocalizationManager.GetText("Error searching for record: ") & Err.description, vbExclamation, LocalizationManager.GetText("Error")
End Sub

Private Sub UpdateNavigationButtons()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo NavigationError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    With UserFormVhIsh
        .btnFirst.Enabled = (DataManager.CurrentRecordRow > 1)
        .btnPrevious.Enabled = (DataManager.CurrentRecordRow > 1)
        .btnNext.Enabled = (DataManager.CurrentRecordRow < tblData.ListRows.Count)
        .btnLast.Enabled = (DataManager.CurrentRecordRow < tblData.ListRows.Count)
        
        ' Update status bar with navigation info
        If tblData.ListRows.Count > 0 Then
            Dim navInfo As String
            navInfo = LocalizationManager.GetText("Navigation: ")
            
            If DataManager.CurrentRecordRow > 1 Then
                navInfo = navInfo & LocalizationManager.GetText("Prev.(") & (DataManager.CurrentRecordRow - 1) & ") "
            End If
            
            navInfo = navInfo & LocalizationManager.GetText("Curr.(") & DataManager.CurrentRecordRow & ") "
            
            If DataManager.CurrentRecordRow < tblData.ListRows.Count Then
                navInfo = navInfo & LocalizationManager.GetText("Next(") & (DataManager.CurrentRecordRow + 1) & ") "
            End If
            
            navInfo = navInfo & LocalizationManager.GetText("Last(") & tblData.ListRows.Count & ")"
            
            ' Temporarily show info in status bar
            .lblStatusBar.Caption = navInfo
        End If
    End With
    
    Exit Sub
    
NavigationError:
    ' Disable all navigation buttons on error
    With UserFormVhIsh
        .btnFirst.Enabled = False
        .btnPrevious.Enabled = False
        .btnNext.Enabled = False
        .btnLast.Enabled = False
    End With
End Sub

Public Sub UpdateStatusBar()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim statusText As String
    
    On Error GoTo StatusError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If DataManager.IsNewRecord Then
        statusText = LocalizationManager.GetText("New record")
    ElseIf tblData.ListRows.Count = 0 Then
        statusText = LocalizationManager.GetText("No records in table")
    Else
        statusText = LocalizationManager.GetText("Record ") & DataManager.CurrentRecordRow & LocalizationManager.GetText(" of ") & tblData.ListRows.Count
        
        ' Add business data info of current record
        If DataManager.CurrentRecordRow > 0 And DataManager.CurrentRecordRow <= tblData.ListRows.Count Then
            Dim serviceName As String
            Dim docNumber As String
            serviceName = CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 2).value)
            docNumber = CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 5).value)
            
            If Trim(serviceName) <> "" And Trim(docNumber) <> "" Then
                statusText = statusText & " | " & serviceName & " | " & LocalizationManager.GetText("Doc.No.") & docNumber
            End If
        End If
    End If
    
    If DataManager.FormDataChanged Then
        statusText = statusText & LocalizationManager.GetText(" (changed)")
    End If
    
    UserFormVhIsh.lblStatusBar.Caption = statusText
    
    Exit Sub
    
StatusError:
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Status update error")
End Sub

Public Function GetCurrentRecordInfo() As String
    ' Returns info about current record
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim recordInfo As String
    
    On Error GoTo InfoError
    
    If DataManager.IsNewRecord Then
        GetCurrentRecordInfo = LocalizationManager.GetText("New record")
        Exit Function
    End If
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If DataManager.CurrentRecordRow > 0 And DataManager.CurrentRecordRow <= tblData.ListRows.Count Then
        recordInfo = LocalizationManager.GetText("Record No.") & DataManager.CurrentRecordRow & ": " & _
                     CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 2).value) & " - " & _
                     CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 4).value) & " " & LocalizationManager.GetText("No.") & _
                     CStr(tblData.DataBodyRange.Cells(DataManager.CurrentRecordRow, 5).value)
    Else
        recordInfo = LocalizationManager.GetText("Invalid record")
    End If
    
    GetCurrentRecordInfo = recordInfo
    
    Exit Function
    
InfoError:
    GetCurrentRecordInfo = LocalizationManager.GetText("Error getting record information")
End Function

Public Sub RefreshNavigation()
    ' Update navigation after table changes
    Call UpdateNavigationButtons
    Call UpdateStatusBar
    
    ' Check that current record still exists
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo RefreshError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If DataManager.CurrentRecordRow > tblData.ListRows.Count Then
        ' If current record was deleted, go to last
        If tblData.ListRows.Count > 0 Then
            Call NavigateToRecord(tblData.ListRows.Count)
        Else
            Call DataManager.ClearForm
        End If
    End If
    
    Exit Sub
    
RefreshError:
    Call DataManager.ClearForm
End Sub

' Function FormatDateCell moved to CommonUtilities.bas

Public Function CanNavigate() As Boolean
    ' Check if navigation is possible
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo CannotNavigate
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    CanNavigate = (tblData.ListRows.Count > 0)
    
    Exit Function
    
CannotNavigate:
    CanNavigate = False
End Function

Public Sub ShowNavigationSummary()
    ' Shows navigation summary
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim summaryText As String
    
    On Error GoTo SummaryError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    summaryText = LocalizationManager.GetText("Records Summary:") & vbCrLf & _
                  LocalizationManager.GetText("Total records: ") & tblData.ListRows.Count & vbCrLf
    
    If tblData.ListRows.Count > 0 Then
        summaryText = summaryText & _
                      LocalizationManager.GetText("Current record: ") & DataManager.CurrentRecordRow & vbCrLf & _
                      LocalizationManager.GetText("First record: 1") & vbCrLf & _
                      LocalizationManager.GetText("Last record: ") & tblData.ListRows.Count & vbCrLf
        
        If DataManager.IsNewRecord Then
            summaryText = summaryText & LocalizationManager.GetText("Status: New record")
        Else
            summaryText = summaryText & LocalizationManager.GetText("Status: Existing record")
            If DataManager.FormDataChanged Then
                summaryText = summaryText & LocalizationManager.GetText(" (changed)")
            End If
        End If
    Else
        summaryText = summaryText & LocalizationManager.GetText("Status: Table is empty")
    End If
    
    MsgBox summaryText, vbInformation, LocalizationManager.GetText("Navigation Summary")
    
    Exit Sub
    
SummaryError:
    MsgBox LocalizationManager.GetText("Error getting navigation summary"), vbExclamation, LocalizationManager.GetText("Error")
End Sub

Public Sub ShowNavigationHelp()
    ' Shows navigation help
    Dim helpText As String
    
    helpText = LocalizationManager.GetText("Navigation Help:") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("|< First - jump to first record") & vbCrLf & _
               LocalizationManager.GetText("< Prev. - jump to previous record") & vbCrLf & _
               LocalizationManager.GetText("Next > - jump to next record") & vbCrLf & _
               LocalizationManager.GetText("Last >| - jump to last record") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Hotkeys:") & vbCrLf & _
               LocalizationManager.GetText("Ctrl+S - save record") & vbCrLf & _
               LocalizationManager.GetText("Ctrl+N - new record") & vbCrLf & _
               LocalizationManager.GetText("Ctrl+F - search") & vbCrLf & _
               LocalizationManager.GetText("F3 - next search result") & vbCrLf & _
               LocalizationManager.GetText("Esc - cancel changes")
    
    MsgBox helpText, vbInformation, LocalizationManager.GetText("Navigation Help")
End Sub

