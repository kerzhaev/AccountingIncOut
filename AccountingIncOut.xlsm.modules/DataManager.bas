Attribute VB_Name = "DataManager"
'==============================================
' TABLE DATA MANAGEMENT MODULE - DataManager
' Purpose: Functions for working with the TableIncOut table
' State: ADDED RECORD DUPLICATION FUNCTION WITH FIELD EXCLUSIONS
' Version: 2.0.0
' Date: 09.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

Public CurrentRecordRow As Long
Public IsNewRecord As Boolean
Public FormDataChanged As Boolean

Public Sub SaveCurrentRecord()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim newRow As ListRow
    
    ' Validate required fields
    If Not ValidateRequiredFields() Then
        Exit Sub
    End If
    
    ' Validate dates
    If Not ValidateDates() Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' CRITICAL FIX: Correct sheet name
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If IsNewRecord Then
        ' Add new record
        Set newRow = tblData.ListRows.Add
        CurrentRecordRow = newRow.Index
        
        ' Auto-fill sequence number
        UserFormVhIsh.txtNomerPP.Text = CStr(tblData.ListRows.Count)
    End If
    
    ' Save data to table
    Call WriteFormDataToTable(tblData, CurrentRecordRow)
    
    ' Update status
    IsNewRecord = False
    FormDataChanged = False
    UserFormVhIsh.lblStatusBar.Caption = "Record saved successfully"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving data: " & Err.description, vbCritical, "Save Error"
    UserFormVhIsh.lblStatusBar.Caption = "Error saving data"
End Sub

Public Function ValidateRequiredFields() As Boolean
    ValidateRequiredFields = True
    
    ' Check required fields
    With UserFormVhIsh
        ' Use Value for ComboBox
        ' Service
        If Trim(CStr(.cmbSlujba.value)) = "" Then
            MsgBox "Field 'Service' is required!", vbExclamation, "Data Validation"
            .cmbSlujba.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Document Type (ComboBox) - 4th column
        If Trim(CStr(.cmbVidDoc.value)) = "" Then
            MsgBox "Field 'Document Type' is required!", vbExclamation, "Data Validation"
            .cmbVidDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Document Group (Inc./Out.) - 3rd column - cmbVidDocumenta
        If Trim(CStr(.cmbVidDocumenta.value)) = "" Then
            MsgBox "Field 'Document Group (Inc./Out.)' is required!", vbExclamation, "Data Validation"
            .cmbVidDocumenta.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Document Number
        If Trim(.txtNomerDoc.Text) = "" Then
            MsgBox "Field 'Document Number' is required!", vbExclamation, "Data Validation"
            .txtNomerDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Document Amount
        If Trim(.txtSummaDoc.Text) = "" Then
            MsgBox "Field 'Document Amount' is required!", vbExclamation, "Data Validation"
            .txtSummaDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Check that amount is numeric
        If Not IsNumeric(.txtSummaDoc.Text) Then
            MsgBox "Field 'Document Amount' must contain a numeric value!", vbExclamation, "Data Validation"
            .txtSummaDoc.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Inc.FRP/Out.FRP
        If Trim(.txtVhFRP.Text) = "" Then
            MsgBox "Field 'Inc.FRP/Out.FRP' is required!", vbExclamation, "Data Validation"
            .txtVhFRP.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
        
        ' Date Inc.FRP/Out.FRP
        If Trim(.txtDataVhFRP.Text) = "" Then
            MsgBox "Field 'Date Inc.FRP/Out.FRP' is required!", vbExclamation, "Data Validation"
            .txtDataVhFRP.SetFocus
            ValidateRequiredFields = False
            Exit Function
        End If
    End With
End Function

Public Function ValidateDates() As Boolean
    ValidateDates = True
    
    With UserFormVhIsh
        ' Check all date fields
        If Trim(.txtDataVhFRP.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataVhFRP.Text) Then
                MsgBox "Enter a valid date in DD.MM.YY format in the 'Date Inc.FRP/Out.FRP' field", vbExclamation, "Data Validation"
                .txtDataVhFRP.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDataPeredachi.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataPeredachi.Text) Then
                MsgBox "Enter a valid date in DD.MM.YY format in the 'Date Transferred to Executor' field", vbExclamation, "Data Validation"
                .txtDataPeredachi.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDataIshVSlujbu.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataIshVSlujbu.Text) Then
                MsgBox "Enter a valid date in DD.MM.YY format in the 'Out. Date to Service' field", vbExclamation, "Data Validation"
                .txtDataIshVSlujbu.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDataVozvrata.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataVozvrata.Text) Then
                MsgBox "Enter a valid date in DD.MM.YY format in the 'Return Date from Service' field", vbExclamation, "Data Validation"
                .txtDataVozvrata.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
        
        If Trim(.txtDataIshKonvert.Text) <> "" Then
            If Not CommonUtilities.IsValidDateFormat(.txtDataIshKonvert.Text) Then
                MsgBox "Enter a valid date in DD.MM.YY format in the 'Out. Envelope Date' field", vbExclamation, "Data Validation"
                .txtDataIshKonvert.SetFocus
                ValidateDates = False
                Exit Function
            End If
        End If
    End With
End Function

' Function IsValidDateFormat moved to CommonUtilities.bas

Public Sub WriteFormDataToTable(tbl As ListObject, RowIndex As Long)
    On Error GoTo WriteError
    
    With UserFormVhIsh
        tbl.DataBodyRange.Cells(RowIndex, 1).value = RowIndex ' Record number = row number
        tbl.DataBodyRange.Cells(RowIndex, 2).value = .cmbSlujba.value
        ' cmbVidDocumenta goes to 3rd column
        tbl.DataBodyRange.Cells(RowIndex, 3).value = .cmbVidDocumenta.value
        ' cmbVidDoc goes to 4th column
        tbl.DataBodyRange.Cells(RowIndex, 4).value = .cmbVidDoc.value
        tbl.DataBodyRange.Cells(RowIndex, 5).value = .txtNomerDoc.Text
        
        ' Safe amount conversion
        If IsNumeric(.txtSummaDoc.Text) Then
            tbl.DataBodyRange.Cells(RowIndex, 6).value = CDbl(.txtSummaDoc.Text)
        Else
            tbl.DataBodyRange.Cells(RowIndex, 6).value = 0
        End If
        
        tbl.DataBodyRange.Cells(RowIndex, 7).value = .txtVhFRP.Text
        
        ' Safe date writing
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
        
        ' Order Info - 20th column
        tbl.DataBodyRange.Cells(RowIndex, 20).value = .txtNaryadInfo.Text
    End With
    
    Exit Sub
    
WriteError:
    MsgBox "Error writing data to table: " & Err.description, vbCritical, "Error"
End Sub

' Function WriteDateToCell moved to CommonUtilities.bas

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
    UserFormVhIsh.lblStatusBar.Caption = "Changes cancelled"
End Sub

Public Sub ClearForm()
    IsNewRecord = True
    CurrentRecordRow = 0
    FormDataChanged = False
    
    ' Set next sequence number for new record
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim nextNumber As Long
    
    On Error Resume Next
    ' CRITICAL FIX: Correct sheet name
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If Not wsData Is Nothing And Not tblData Is Nothing Then
        nextNumber = tblData.ListRows.Count + 1
    Else
        nextNumber = 1
    End If
    On Error GoTo 0
    
    With UserFormVhIsh
        ' Clear all fields
        .txtNomerPP.Text = CStr(nextNumber) ' Set next number
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
        .cmbStatusPodtverjdenie.ListIndex = 0 ' Set empty value
        
        ' Clear other comboboxes
        .cmbSlujba.value = ""
        .cmbVidDoc.value = ""
        .cmbIspolnitel.value = ""
        
        ' Clear order info field
        .txtNaryadInfo.Text = ""
        
        ' Clear search
        .txtSearch.Text = ""
        .lstSearchResults.Clear
        .lstSearchResults.Visible = False
    End With
    
    Call UpdateStatusBar
End Sub

' NEW FUNCTION: Duplicate record excluding certain fields
Public Function DuplicateRecord(sourceRowNumber As Long) As Long
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim newRow As ListRow
    Dim newRowIndex As Long
    Dim i As Long
    
    On Error GoTo DuplicateError
    
    ' Get references to table
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    ' Check source row correctness
    If sourceRowNumber < 1 Or sourceRowNumber > tblData.ListRows.Count Then
        MsgBox "Invalid record number for duplication: " & sourceRowNumber, vbExclamation, "Error"
        DuplicateRecord = 0
        Exit Function
    End If
    
    ' Add new row
    Set newRow = tblData.ListRows.Add
    newRowIndex = newRow.Index
    
    ' Copy data from source row, excluding specific fields
    For i = 1 To tblData.ListColumns.Count
        Select Case i
            Case 1  ' Seq No - set new number
                tblData.DataBodyRange.Cells(newRowIndex, i).value = newRowIndex
                
            Case 5  ' txtNomerDoc - DO NOT copy (document number)
                tblData.DataBodyRange.Cells(newRowIndex, i).value = ""
                
            Case 6  ' txtSummaDoc - DO NOT copy (document amount)
                tblData.DataBodyRange.Cells(newRowIndex, i).value = 0
                
            Case 20 ' txtNaryadInfo - DO NOT copy (order info)
                tblData.DataBodyRange.Cells(newRowIndex, i).value = ""
                
            Case Else ' All other fields - copy
                tblData.DataBodyRange.Cells(newRowIndex, i).value = _
                    tblData.DataBodyRange.Cells(sourceRowNumber, i).value
        End Select
    Next i
    
    ' Return new record number
    DuplicateRecord = newRowIndex
    
    Exit Function
    
DuplicateError:
    MsgBox "Error duplicating record: " & Err.description, vbCritical, "Critical Error"
    DuplicateRecord = 0
    
    ' Try to delete created row on error
    On Error Resume Next
    If Not newRow Is Nothing Then
        newRow.Delete
    End If
    On Error GoTo 0
End Function

' NEW FUNCTION: Get record information for duplication
Public Function GetRecordInfo(RowNumber As Long) As String
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim recordInfo As String
    
    On Error GoTo InfoError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If RowNumber < 1 Or RowNumber > tblData.ListRows.Count Then
        GetRecordInfo = "Invalid record number"
        Exit Function
    End If
    
    ' Build information string
    recordInfo = "Record No." & RowNumber & ": "
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 2).value) & " - " ' Service
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 3).value) & " " ' Document Group
    recordInfo = recordInfo & CStr(tblData.DataBodyRange.Cells(RowNumber, 4).value) & " " ' Document Type
    recordInfo = recordInfo & "No." & CStr(tblData.DataBodyRange.Cells(RowNumber, 5).value) ' Document Number
    
    GetRecordInfo = recordInfo
    
    Exit Function
    
InfoError:
    GetRecordInfo = "Error getting record information"
End Function

Public Sub SetupGroupBoxes()
    ' Dummy function for grouping setup
    ' Will be implemented upon form UI creation
End Sub

' NEW FUNCTION: Check if record can be duplicated
Public Function CanDuplicateRecord(RowNumber As Long) As Boolean
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo CannotDuplicate
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    ' Check if table and record exist
    If tblData Is Nothing Then
        CanDuplicateRecord = False
        Exit Function
    End If
    
    If tblData.DataBodyRange Is Nothing Then
        CanDuplicateRecord = False
        Exit Function
    End If
    
    If RowNumber < 1 Or RowNumber > tblData.ListRows.Count Then
        CanDuplicateRecord = False
        Exit Function
    End If
    
    CanDuplicateRecord = True
    Exit Function
    
CannotDuplicate:
    CanDuplicateRecord = False
End Function

