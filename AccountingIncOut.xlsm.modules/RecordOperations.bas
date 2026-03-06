Attribute VB_Name = "RecordOperations"
'==============================================
' CORE RECORD OPERATIONS MODULE - RecordOperations
' Purpose: Unified module for Data Management, Navigation, and Search
' Date: 06.03.2026 (Refactored)
'==============================================

Option Explicit

' --- √ÀŒ¡¿À‹Õ€≈ œ≈–≈Ã≈ÕÕ€≈ ”œ–¿¬À≈Õ»ﬂ ƒ¿ÕÕ€Ã» ---
Public CurrentRecordRow As Long
Public IsNewRecord As Boolean
Public FormDataChanged As Boolean

' --- œ≈–≈Ã≈ÕÕ€≈ œŒ»— ¿ ---
Private LastSearchText As String
Private SearchResultsVisible As Boolean

' =====================================================================
' 1. ¡ÀŒ  ”œ–¿¬À≈Õ»ﬂ ƒ¿ÕÕ€Ã» (DATA MANAGEMENT)
' =====================================================================

Public Sub SaveCurrentRecord()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim newRow As ListRow
    
    If Not ValidateRequiredFields() Then Exit Sub
    If Not ValidateDates() Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If IsNewRecord Then
        Set newRow = tblData.ListRows.Add
        CurrentRecordRow = newRow.Index
        UserFormVhIsh.txtNomerPP.Text = CStr(tblData.ListRows.Count)
    End If
    
    Call WriteFormDataToTable(tblData, CurrentRecordRow)
    
    IsNewRecord = False
    FormDataChanged = False
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Record saved successfully")
    Exit Sub
    
ErrorHandler:
    MsgBox LocalizationManager.GetText("Error saving data: ") & Err.description, vbCritical, LocalizationManager.GetText("Save Error")
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Error saving data")
End Sub

Public Function ValidateRequiredFields() As Boolean
    ValidateRequiredFields = True
    With UserFormVhIsh
        If Trim(CStr(.cmbSlujba.value)) = "" Then
             MsgBox LocalizationManager.GetText("Field 'Service' is required!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .cmbSlujba.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        If Trim(CStr(.cmbVidDoc.value)) = "" Then
             MsgBox LocalizationManager.GetText("Field 'Document Type' is required!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .cmbVidDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        If Trim(CStr(.cmbVidDocumenta.value)) = "" Then
             MsgBox LocalizationManager.GetText("Field 'Document Group (Inc./Out.)' is required!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .cmbVidDocumenta.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        If Trim(.txtNomerDoc.Text) = "" Then
             MsgBox LocalizationManager.GetText("Field 'Document Number' is required!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .txtNomerDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        If Trim(.txtSummaDoc.Text) = "" Then
             MsgBox LocalizationManager.GetText("Field 'Document Amount' is required!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .txtSummaDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        If Not IsNumeric(.txtSummaDoc.Text) Then
             MsgBox LocalizationManager.GetText("Field 'Document Amount' must contain a numeric value!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .txtSummaDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        If Trim(.txtVhFRP.Text) = "" Then
             MsgBox LocalizationManager.GetText("Field 'Inc.FRP/Out.FRP' is required!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .txtVhFRP.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        If Trim(.txtDataVhFRP.Text) = "" Then
            MsgBox LocalizationManager.GetText("Field 'Date Inc.FRP/Out.FRP' is required!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .txtDataVhFRP.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
    End With
End Function

Public Function ValidateDates() As Boolean
    ValidateDates = True
    With UserFormVhIsh
        If Trim(.txtDataVhFRP.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataVhFRP.Text) Then
                MsgBox LocalizationManager.GetText("Enter a valid date in DD.MM.YY format in the 'Date Inc.FRP/Out.FRP' field"), vbExclamation, LocalizationManager.GetText("Data Validation")
                .txtDataVhFRP.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        If Trim(.txtDataPeredachi.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataPeredachi.Text) Then
                MsgBox LocalizationManager.GetText("Enter a valid date in DD.MM.YY format in the 'Date Transferred to Executor' field"), vbExclamation, LocalizationManager.GetText("Data Validation")
                .txtDataPeredachi.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        If Trim(.txtDataIshVSlujbu.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataIshVSlujbu.Text) Then
                MsgBox LocalizationManager.GetText("Enter a valid date in DD.MM.YY format in the 'Out. Date to Service' field"), vbExclamation, LocalizationManager.GetText("Data Validation")
                .txtDataIshVSlujbu.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        If Trim(.txtDataVozvrata.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataVozvrata.Text) Then
                MsgBox LocalizationManager.GetText("Enter a valid date in DD.MM.YY format in the 'Return Date from Service' field"), vbExclamation, LocalizationManager.GetText("Data Validation")
                .txtDataVozvrata.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        If Trim(.txtDataIshKonvert.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataIshKonvert.Text) Then
                MsgBox LocalizationManager.GetText("Enter a valid date in DD.MM.YY format in the 'Out. Envelope Date' field"), vbExclamation, LocalizationManager.GetText("Data Validation")
                .txtDataIshKonvert.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
    End With
End Function

Public Sub WriteFormDataToTable(tbl As ListObject, RowIndex As Long)
    On Error GoTo WriteError
    With UserFormVhIsh
        tbl.DataBodyRange.Cells(RowIndex, 1).value = RowIndex
        tbl.DataBodyRange.Cells(RowIndex, 2).value = .cmbSlujba.value
        tbl.DataBodyRange.Cells(RowIndex, 3).value = .cmbVidDocumenta.value
        tbl.DataBodyRange.Cells(RowIndex, 4).value = .cmbVidDoc.value
        tbl.DataBodyRange.Cells(RowIndex, 5).value = .txtNomerDoc.Text
        If IsNumeric(.txtSummaDoc.Text) Then
            tbl.DataBodyRange.Cells(RowIndex, 6).value = CDbl(.txtSummaDoc.Text)
        Else
            tbl.DataBodyRange.Cells(RowIndex, 6).value = 0
        End If
        tbl.DataBodyRange.Cells(RowIndex, 7).value = .txtVhFRP.Text
        Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 8), .txtDataVhFRP.Text)
        tbl.DataBodyRange.Cells(RowIndex, 9).value = .cmbOtKogoPostupil.value
        Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 10), .txtDataPeredachi.Text)
        tbl.DataBodyRange.Cells(RowIndex, 11).value = .cmbIspolnitel.value
        tbl.DataBodyRange.Cells(RowIndex, 12).value = .txtNomerIshVSlujbu.Text
        Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 13), .txtDataIshVSlujbu.Text)
        tbl.DataBodyRange.Cells(RowIndex, 14).value = .txtNomerVozvrata.Text
        Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 15), .txtDataVozvrata.Text)
        tbl.DataBodyRange.Cells(RowIndex, 16).value = .txtNomerIshKonvert.Text
        Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 17), .txtDataIshKonvert.Text)
        tbl.DataBodyRange.Cells(RowIndex, 18).value = .txtOtmetkaIspolnenie.Text
        tbl.DataBodyRange.Cells(RowIndex, 19).value = .cmbStatusPodtverjdenie.value
        tbl.DataBodyRange.Cells(RowIndex, 20).value = .txtNaryadInfo.Text
    End With
    Exit Sub
WriteError:
    MsgBox LocalizationManager.GetText("Error writing data to table: ") & Err.description, vbCritical, LocalizationManager.GetText("Error")
End Sub

Public Sub MarkFormAsChanged()
    FormDataChanged = True
    Call UpdateStatusBar
End Sub

Public Function HasUnsavedChanges() As Boolean
    HasUnsavedChanges = FormDataChanged
End Function

Public Sub CancelChanges()
    If IsNewRecord Then
        Call ClearForm
    Else
        Call NavigateToRecord(CurrentRecordRow)
    End If
    FormDataChanged = False
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Changes cancelled")
End Sub

Public Sub ClearForm()
    IsNewRecord = True
    CurrentRecordRow = 0
    FormDataChanged = False
    
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim nextNumber As Long
    
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If Not wsData Is Nothing And Not tblData Is Nothing Then
        nextNumber = tblData.ListRows.Count + 1
    Else
        nextNumber = 1
    End If
    On Error GoTo 0
    
    With UserFormVhIsh
        .txtNomerPP.Text = CStr(nextNumber)
        .cmbVidDocumenta.value = ""
        .txtNomerDoc.Text = ""
        .txtSummaDoc.Text = ""
        .txtVhFRP.Text = ""
        .txtDataVhFRP.Text = ""
        .cmbOtKogoPostupil.value = ""
        .txtDataPeredachi.Text = ""
        .txtNomerIshVSlujbu.Text = ""
        .txtDataIshVSlujbu.Text = ""
        .txtNomerVozvrata.Text = ""
        .txtDataVozvrata.Text = ""
        .txtNomerIshKonvert.Text = ""
        .txtDataIshKonvert.Text = ""
        .txtOtmetkaIspolnenie.Text = ""
        .cmbStatusPodtverjdenie.ListIndex = 0
        .cmbSlujba.value = ""
        .cmbVidDoc.value = ""
        .cmbIspolnitel.value = ""
        .txtNaryadInfo.Text = ""
        .txtSearch.Text = ""
        .lstSearchResults.Clear
        .lstSearchResults.Visible = False
    End With
    
    Call UpdateStatusBar
End Sub

Public Function DuplicateRecord(sourceRowNumber As Long) As Long
    Dim wsData As Worksheet, tblData As ListObject
    Dim newRow As ListRow, newRowIndex As Long, i As Long
    
    On Error GoTo DuplicateError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If sourceRowNumber < 1 Or sourceRowNumber > tblData.ListRows.Count Then
        MsgBox LocalizationManager.GetText("Invalid record number for duplication: ") & sourceRowNumber, vbExclamation, LocalizationManager.GetText("Error")
        DuplicateRecord = 0
        Exit Function
    End If
    
    Set newRow = tblData.ListRows.Add
    newRowIndex = newRow.Index
    
    For i = 1 To tblData.ListColumns.Count
        Select Case i
            Case 1: tblData.DataBodyRange.Cells(newRowIndex, i).value = newRowIndex
            Case 5, 20: tblData.DataBodyRange.Cells(newRowIndex, i).value = ""
            Case 6: tblData.DataBodyRange.Cells(newRowIndex, i).value = 0
            Case Else: tblData.DataBodyRange.Cells(newRowIndex, i).value = tblData.DataBodyRange.Cells(sourceRowNumber, i).value
        End Select
    Next i
    
    DuplicateRecord = newRowIndex
    Exit Function
    
DuplicateError:
    MsgBox LocalizationManager.GetText("Error duplicating record: ") & Err.description, vbCritical, LocalizationManager.GetText("Critical Error")
    DuplicateRecord = 0
    On Error Resume Next
    If Not newRow Is Nothing Then newRow.Delete
    On Error GoTo 0
End Function

Public Function GetRecordInfo(RowNumber As Long) As String
    Dim wsData As Worksheet, tblData As ListObject
    Dim recordInfo As String
    
    On Error GoTo InfoError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If RowNumber < 1 Or RowNumber > tblData.ListRows.Count Then
        GetRecordInfo = LocalizationManager.GetText("Invalid record number")
        Exit Function
    End If
    
    recordInfo = LocalizationManager.GetText("Record No.") & RowNumber & ": "
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 2).value) & " - "
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 3).value) & " "
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 4).value) & " "
    recordInfo = recordInfo & LocalizationManager.GetText("No.") & CStr(tblData.DataBodyRange.Cells(RowNumber, 5).value)
    
    GetRecordInfo = recordInfo
    Exit Function
InfoError:
    GetRecordInfo = LocalizationManager.GetText("Error getting record information")
End Function

Public Sub SetupGroupBoxes()
End Sub

Public Function CanDuplicateRecord(RowNumber As Long) As Boolean
    Dim wsData As Worksheet, tblData As ListObject
    On Error GoTo CannotDuplicate
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If tblData Is Nothing Then Exit Function
    If tblData.DataBodyRange Is Nothing Then Exit Function
    If RowNumber < 1 Or RowNumber > tblData.ListRows.Count Then Exit Function
    
    CanDuplicateRecord = True
    Exit Function
CannotDuplicate:
    CanDuplicateRecord = False
End Function

' =====================================================================
' 2. ¡ÀŒ  Õ¿¬»√¿÷»» (NAVIGATION)
' =====================================================================

Public Sub NavigateToRecord(RowNumber As Long)
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo NavigationError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If tblData.ListRows.Count = 0 Then
        Call ClearForm
        Exit Sub
    End If
    
    If RowNumber < 1 Then RowNumber = 1
    If RowNumber > tblData.ListRows.Count Then RowNumber = tblData.ListRows.Count
    
    If HasUnsavedChanges() And CurrentRecordRow <> RowNumber And CurrentRecordRow > 0 Then
        If MsgBox(LocalizationManager.GetText("Save changes to the current record before navigating?"), vbYesNo + vbQuestion, LocalizationManager.GetText("Unsaved Changes")) = vbYes Then
            Call SaveCurrentRecord
        End If
    End If
    
    CurrentRecordRow = RowNumber
    IsNewRecord = False
    
    Call UserFormVhIsh.LoadRecordToForm(RowNumber)
    Call UpdateNavigationButtons
    Call UpdateStatusBar
    
    FormDataChanged = False
    Exit Sub
    
NavigationError:
    MsgBox LocalizationManager.GetText("Navigation error: ") & Err.description, vbCritical, LocalizationManager.GetText("Error")
End Sub

Public Sub NavigateToPrevious()
    Call NavigateToRecord(CurrentRecordRow - 1)
End Sub
Public Sub NavigateToNext()
    Call NavigateToRecord(CurrentRecordRow + 1)
End Sub
Public Sub NavigateToFirst()
    Call NavigateToRecord(1)
End Sub

Public Sub NavigateToLast()
    Dim wsData As Worksheet, tblData As ListObject
    On Error GoTo LastError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    Call NavigateToRecord(tblData.ListRows.Count)
    Exit Sub
LastError:
    Call NavigateToRecord(1)
End Sub

Public Sub NavigateToRecordByNumber(recordNumber As Long)
    Dim wsData As Worksheet, tblData As ListObject, i As Long
    On Error GoTo FindError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
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
    Dim wsData As Worksheet, tblData As ListObject
    On Error GoTo NavigationError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    With UserFormVhIsh
        .btnFirst.Enabled = (CurrentRecordRow > 1)
        .btnPrevious.Enabled = (CurrentRecordRow > 1)
        .btnNext.Enabled = (CurrentRecordRow < tblData.ListRows.Count)
        .btnLast.Enabled = (CurrentRecordRow < tblData.ListRows.Count)
    End With
    Exit Sub
NavigationError:
    With UserFormVhIsh
        .btnFirst.Enabled = False: .btnPrevious.Enabled = False
        .btnNext.Enabled = False: .btnLast.Enabled = False
    End With
End Sub

Public Sub UpdateStatusBar()
    Dim wsData As Worksheet, tblData As ListObject, statusText As String
    On Error GoTo StatusError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If IsNewRecord Then
        statusText = LocalizationManager.GetText("New record")
    ElseIf tblData.ListRows.Count = 0 Then
        statusText = LocalizationManager.GetText("No records in table")
    Else
        statusText = LocalizationManager.GetText("Record ") & CurrentRecordRow & LocalizationManager.GetText(" of ") & tblData.ListRows.Count
        If CurrentRecordRow > 0 And CurrentRecordRow <= tblData.ListRows.Count Then
            Dim serviceName As String, docNumber As String
            serviceName = CStr(tblData.DataBodyRange.Cells(CurrentRecordRow, 2).value)
            docNumber = CStr(tblData.DataBodyRange.Cells(CurrentRecordRow, 5).value)
            If Trim(serviceName) <> "" And Trim(docNumber) <> "" Then
                statusText = statusText & " | " & serviceName & " | " & LocalizationManager.GetText("Doc.No.") & docNumber
            End If
        End If
    End If
    
    If FormDataChanged Then statusText = statusText & LocalizationManager.GetText(" (changed)")
    UserFormVhIsh.lblStatusBar.Caption = statusText
    Exit Sub
StatusError:
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Status update error")
End Sub

Public Function GetCurrentRecordInfo() As String
    Dim wsData As Worksheet, tblData As ListObject, recordInfo As String
    On Error GoTo InfoError
    
    If IsNewRecord Then
        GetCurrentRecordInfo = LocalizationManager.GetText("New record")
        Exit Function
    End If
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If CurrentRecordRow > 0 And CurrentRecordRow <= tblData.ListRows.Count Then
        recordInfo = LocalizationManager.GetText("Record No.") & CurrentRecordRow & ": " & _
                     CStr(tblData.DataBodyRange.Cells(CurrentRecordRow, 2).value) & " - " & _
                     CStr(tblData.DataBodyRange.Cells(CurrentRecordRow, 4).value) & " " & LocalizationManager.GetText("No.") & _
                     CStr(tblData.DataBodyRange.Cells(CurrentRecordRow, 5).value)
    Else
        recordInfo = LocalizationManager.GetText("Invalid record")
    End If
    GetCurrentRecordInfo = recordInfo
    Exit Function
InfoError:
    GetCurrentRecordInfo = LocalizationManager.GetText("Error getting record information")
End Function

Public Sub RefreshNavigation()
    Call UpdateNavigationButtons
    Call UpdateStatusBar
    Dim wsData As Worksheet, tblData As ListObject
    On Error GoTo RefreshError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If CurrentRecordRow > tblData.ListRows.Count Then
        If tblData.ListRows.Count > 0 Then
            Call NavigateToRecord(tblData.ListRows.Count)
        Else
            Call ClearForm
        End If
    End If
    Exit Sub
RefreshError:
    Call ClearForm
End Sub

Public Function CanNavigate() As Boolean
    Dim wsData As Worksheet, tblData As ListObject
    On Error GoTo CannotNavigate
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    CanNavigate = (tblData.ListRows.Count > 0)
    Exit Function
CannotNavigate:
    CanNavigate = False
End Function

Public Sub ShowNavigationSummary()
    Dim wsData As Worksheet, tblData As ListObject, summaryText As String
    On Error GoTo SummaryError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    summaryText = LocalizationManager.GetText("Records Summary:") & vbCrLf & _
                  LocalizationManager.GetText("Total records: ") & tblData.ListRows.Count & vbCrLf
    
    If tblData.ListRows.Count > 0 Then
        summaryText = summaryText & LocalizationManager.GetText("Current record: ") & CurrentRecordRow & vbCrLf & _
                      LocalizationManager.GetText("First record: 1") & vbCrLf & _
                      LocalizationManager.GetText("Last record: ") & tblData.ListRows.Count & vbCrLf
        If IsNewRecord Then
            summaryText = summaryText & LocalizationManager.GetText("Status: New record")
        Else
            summaryText = summaryText & LocalizationManager.GetText("Status: Existing record")
            If FormDataChanged Then summaryText = summaryText & LocalizationManager.GetText(" (changed)")
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


' =====================================================================
' 3. ¡ÀŒ  œŒ»— ¿ (SEARCH)
' =====================================================================

Public Sub PerformSearch()
    Dim SearchText As String, wsData As Worksheet, tblData As ListObject
    Dim i As Long, j As Long, cellValue As String, FoundCount As Integer
    
    SearchText = Trim(UCase(UserFormVhIsh.txtSearch.Text))
    LastSearchText = SearchText
    
    If Len(SearchText) = 0 Then
        With UserFormVhIsh
            .lstSearchResults.Clear
            .lstSearchResults.Visible = False
            .lblStatusBar.Caption = LocalizationManager.GetText("Enter text to search")
        End With
        SearchResultsVisible = False
        Exit Sub
    End If
    
    On Error GoTo SearchError
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    UserFormVhIsh.lstSearchResults.Clear
    
    If tblData.ListRows.Count > 0 Then
        For i = 1 To tblData.ListRows.Count
            Dim hasMatch As Boolean: hasMatch = False
            For j = 1 To tblData.ListColumns.Count
                cellValue = UCase(CStr(tblData.DataBodyRange.Cells(i, j).value))
                If InStr(cellValue, SearchText) > 0 Then
                    hasMatch = True
                    Exit For
                End If
            Next j
            
            If hasMatch Then
                Call AddImprovedSearchResult(i, tblData)
                FoundCount = FoundCount + 1
            End If
            If FoundCount >= 25 Then Exit For
        Next i
    End If
    
    Call DisplaySearchResults(FoundCount)
    Exit Sub
SearchError:
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Search error: ") & Err.description
    UserFormVhIsh.lstSearchResults.Visible = False
    SearchResultsVisible = False
End Sub

Private Sub AddImprovedSearchResult(rowNum As Long, tbl As ListObject)
    Dim ResultText As String, serviceText As String, docTypeInOut As String
    Dim docTypeText As String, docNumberText As String, amountText As String
    Dim frpText As String, DateText As String, fromWhomText As String
    Dim executorText As String, statusText As String
    
    On Error GoTo AddError
    
    serviceText = CStr(tbl.DataBodyRange.Cells(rowNum, 2).value)
    docTypeInOut = CStr(tbl.DataBodyRange.Cells(rowNum, 3).value)
    docTypeText = CStr(tbl.DataBodyRange.Cells(rowNum, 4).value)
    docNumberText = CStr(tbl.DataBodyRange.Cells(rowNum, 5).value)
    amountText = CStr(tbl.DataBodyRange.Cells(rowNum, 6).value)
    frpText = CStr(tbl.DataBodyRange.Cells(rowNum, 7).value)
    
    If IsDate(tbl.DataBodyRange.Cells(rowNum, 8).value) Then
        DateText = Format(tbl.DataBodyRange.Cells(rowNum, 8).value, "dd.mm.yy")
    Else
        DateText = CStr(tbl.DataBodyRange.Cells(rowNum, 8).value)
    End If
    
    fromWhomText = CStr(tbl.DataBodyRange.Cells(rowNum, 9).value)
    executorText = CStr(tbl.DataBodyRange.Cells(rowNum, 11).value)
    statusText = CStr(tbl.DataBodyRange.Cells(rowNum, 19).value)
    
    ResultText = ">" & rowNum & ": "
    If Trim(serviceText) <> "" Then ResultText = ResultText & "[" & serviceText & "] "
    If Trim(docTypeInOut) <> "" Then ResultText = ResultText & docTypeInOut & " "
    If Trim(docTypeText) <> "" Then ResultText = ResultText & docTypeText & " "
    If Trim(docNumberText) <> "" Then ResultText = ResultText & LocalizationManager.GetText("No.") & docNumberText & " "
    If Trim(amountText) <> "" And amountText <> "0" Then ResultText = ResultText & "(" & amountText & LocalizationManager.GetText(" rub.") & ") "
    If Trim(frpText) <> "" Then ResultText = ResultText & LocalizationManager.GetText("FRP:") & frpText & " "
    If Trim(DateText) <> "" Then ResultText = ResultText & LocalizationManager.GetText("from ") & DateText & " "
    If Trim(fromWhomText) <> "" Then ResultText = ResultText & "| " & LocalizationManager.GetText("From: ") & fromWhomText & " "
    If Trim(executorText) <> "" Then ResultText = ResultText & "| " & LocalizationManager.GetText("Exec.: ") & executorText & " "
    If Trim(statusText) <> "" Then ResultText = ResultText & "| " & LocalizationManager.GetText("Status: ") & statusText & " "
    
    With UserFormVhIsh.lstSearchResults
        .AddItem ResultText
        If .ListCount > 0 Then
           .List(.ListCount - 1, 0) = ResultText & Chr(1) & CStr(rowNum)
        End If
    End With
AddError:
End Sub

Private Sub DisplaySearchResults(FoundCount As Integer)
    With UserFormVhIsh
        If .lstSearchResults.ListCount > 0 Then
            .lstSearchResults.Visible = True
            SearchResultsVisible = True
            Call AutoResizeSearchResultsWidth
            .lstSearchResults.Height = 120
            .lblStatusBar.Caption = LocalizationManager.GetText("Found: ") & FoundCount & LocalizationManager.GetText(" records | ") & _
                                    LocalizationManager.GetText("Navigation: ^v or click to jump | ") & _
                                    LocalizationManager.GetText("Search remains active")
        Else
            .lstSearchResults.Visible = False
            SearchResultsVisible = False
            .lblStatusBar.Caption = LocalizationManager.GetText("For query '") & .txtSearch.Text & LocalizationManager.GetText("' nothing found")
        End If
    End With
End Sub

Private Sub AutoResizeSearchResultsWidth()
    Dim maxWidth As Single, textWidth As Single, i As Long
    Dim testText As String, minWidth As Single, maxAllowedWidth As Single
    
    On Error GoTo ResizeError
    minWidth = 420
    maxAllowedWidth = 800
    maxWidth = minWidth
    
    With UserFormVhIsh.lstSearchResults
        For i = 0 To .ListCount - 1
            testText = .List(i, 0)
            If InStr(testText, Chr(1)) > 0 Then testText = Left(testText, InStr(testText, Chr(1)) - 1)
            textWidth = CalculateTextWidth(testText, .Font.Name, .Font.Size)
            If textWidth + 20 > maxWidth Then maxWidth = textWidth + 20
        Next i
        If maxWidth < minWidth Then maxWidth = minWidth
        If maxWidth > maxAllowedWidth Then maxWidth = maxAllowedWidth
        .Width = maxWidth
        .ColumnWidths = CStr(maxWidth - 10)
    End With
    Exit Sub
ResizeError:
    UserFormVhIsh.lstSearchResults.Width = minWidth
End Sub

Private Function CalculateTextWidth(Text As String, fontName As String, fontSize As Single) As Single
    On Error GoTo WidthError
    CalculateTextWidth = Len(Text) * (fontSize * 0.6)
    Exit Function
WidthError:
    CalculateTextWidth = 420
End Function

Public Sub SelectSearchResult()
    Dim selectedRow As Long, resultData As String, parts As Variant
    On Error GoTo SelectError
    With UserFormVhIsh.lstSearchResults
        If .ListIndex >= 0 Then
            resultData = .List(.ListIndex, 0)
            parts = Split(resultData, Chr(1))
            If UBound(parts) >= 1 Then
                selectedRow = CLng(parts(1))
                Call NavigateToRecord(selectedRow)
                UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Jump to record No.") & selectedRow & " | " & _
                                                     LocalizationManager.GetText("Found: ") & .ListCount & LocalizationManager.GetText(" records | ") & _
                                                     LocalizationManager.GetText("Search active: """) & UserFormVhIsh.txtSearch.Text & """"
                .BackColor = RGB(240, 248, 255)
            End If
        End If
    End With
    Exit Sub
SelectError:
    MsgBox LocalizationManager.GetText("Error selecting search result: ") & Err.description, vbExclamation, LocalizationManager.GetText("Error")
End Sub

Public Sub ClearSearch()
    With UserFormVhIsh
        .txtSearch.Text = ""
        .lstSearchResults.Clear
        .lstSearchResults.Visible = False
        .lblStatusBar.Caption = LocalizationManager.GetText("Search cleared")
    End With
    LastSearchText = ""
    SearchResultsVisible = False
End Sub

Public Sub RestoreLastSearch()
    If LastSearchText <> "" Then
        UserFormVhIsh.txtSearch.Text = LastSearchText
        Call PerformSearch
    End If
End Sub

Public Sub NavigateSearchResults(direction As String)
    With UserFormVhIsh.lstSearchResults
        If .Visible And .ListCount > 0 Then
            Select Case UCase(direction)
                Case "UP"
                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1 Else .ListIndex = .ListCount - 1
                Case "DOWN"
                    If .ListIndex < .ListCount - 1 Then .ListIndex = .ListIndex + 1 Else .ListIndex = 0
                Case "FIRST": .ListIndex = 0
                Case "LAST": .ListIndex = .ListCount - 1
            End Select
            Call SelectSearchResult
        End If
    End With
End Sub

Public Function IsSearchActive() As Boolean
    IsSearchActive = SearchResultsVisible And (UserFormVhIsh.lstSearchResults.ListCount > 0)
End Function

Public Function GetSearchInfo() As String
    If IsSearchActive() Then
        With UserFormVhIsh.lstSearchResults
            GetSearchInfo = LocalizationManager.GetText("Active search: """) & UserFormVhIsh.txtSearch.Text & """ | " & _
                            LocalizationManager.GetText("Found: ") & .ListCount & LocalizationManager.GetText(" records | ") & _
                            LocalizationManager.GetText("Selected: ") & (.ListIndex + 1) & LocalizationManager.GetText(" of ") & .ListCount
        End With
    Else
        GetSearchInfo = LocalizationManager.GetText("Search inactive")
    End If
End Function

Public Sub HighlightSearchTermInForm()
    Dim searchTerm As String
    searchTerm = UserFormVhIsh.txtSearch.Text
End Sub

