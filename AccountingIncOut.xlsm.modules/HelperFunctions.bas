Attribute VB_Name = "HelperFunctions"
'==============================================
' HELPER FUNCTIONS - HelperFunctions
' Purpose: General functions for form operations
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.4.2
' Date: 10.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' DUPLICATE FUNCTIONS CancelChanges() and MarkFormAsChanged() REMOVED
' These functions are now only in the DataManager module

Public Sub SetupFormAppearance()
    Call SetupRequiredFieldsHighlight
    Call SetupFieldAlignment
    Call SetupNavigationButtons
    Call SetupStatusBar
End Sub

Private Sub SetupRequiredFieldsHighlight()
    ' Highlight required fields with light yellow color
    Dim requiredColor As Long
    requiredColor = RGB(255, 255, 224)
    
    With UserFormVhIsh
        .cmbSlujba.BackColor = requiredColor
        .cmbVidDoc.BackColor = requiredColor
        .cmbVidDocumenta.BackColor = requiredColor
        .txtNomerDoc.BackColor = requiredColor
        .txtSummaDoc.BackColor = requiredColor
        .txtVhFRP.BackColor = requiredColor
        .txtDataVhFRP.BackColor = requiredColor
        
        ' Sequence number field remains gray (not editable)
        .txtNomerPP.BackColor = RGB(240, 240, 240)
    End With
End Sub

Private Sub SetupFieldAlignment()
    With UserFormVhIsh
        ' Align numeric fields
        .txtSummaDoc.TextAlign = fmTextAlignRight
        .txtNomerPP.TextAlign = fmTextAlignCenter
        .txtNomerDoc.TextAlign = fmTextAlignCenter
        .txtVhFRP.TextAlign = fmTextAlignCenter
        
        ' Length limits for date fields
        .txtDataVhFRP.MaxLength = 8
        .txtDataPeredachi.MaxLength = 8
        .txtDataIshVSlujbu.MaxLength = 8
        .txtDataVozvrata.MaxLength = 8
        .txtDataIshKonvert.MaxLength = 8
    End With
End Sub

Private Sub SetupNavigationButtons()
    With UserFormVhIsh
        ' Improve navigation button captions
        .btnFirst.Caption = LocalizationManager.GetText("|< First")
        .btnPrevious.Caption = LocalizationManager.GetText("< Prev")
        .btnNext.Caption = LocalizationManager.GetText("Next >")
        .btnLast.Caption = LocalizationManager.GetText("Last >|")
        
        ' Make buttons slightly wider
        .btnFirst.Width = 70
        .btnPrevious.Width = 70
        .btnNext.Width = 70
        .btnLast.Width = 70
    End With
End Sub

Private Sub SetupStatusBar()
    With UserFormVhIsh.lblStatusBar
        .BackColor = RGB(245, 245, 245)
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 9
        .Font.Name = "Segoe UI"
    End With
End Sub

Public Sub ShowFieldTooltip(fieldName As String)
    ' Tooltips for fields
    Dim tooltip As String
    
    Select Case fieldName
        Case "txtSummaDoc"
            tooltip = LocalizationManager.GetText("Enter document amount in rubles (numbers only)")
        Case "txtDataVhFRP"
            tooltip = LocalizationManager.GetText("Date format: DD.MM.YY (e.g., 15.07.25)")
        Case "cmbSlujba"
            tooltip = LocalizationManager.GetText("Select a service from the list or enter a new one")
        Case "txtVhFRP"
            tooltip = LocalizationManager.GetText("Incoming/outgoing FRP document number")
        Case "cmbOtKogoPostupil"
            tooltip = LocalizationManager.GetText("From whom the document was received or where it was sent")
        Case "cmbIspolnitel"
            tooltip = LocalizationManager.GetText("Select an executor from the list or enter a new one")
        Case Else
            tooltip = LocalizationManager.GetText("Data entry field")
    End Select
    
    UserFormVhIsh.lblStatusBar.Caption = tooltip
End Sub

Public Sub SetupGroupBoxes()
    ' Function for setting up field grouping
    ' Grouping fields into logical blocks to improve UX
    
    ' Core document info (fields 1-6)
    ' FRP data (fields 7-8)
    ' Document movement (fields 9-11)
    ' Service operations (fields 12-15)
    ' Envelope conversion (fields 16-17)
    ' Statuses (fields 18-19)
    
    ' This function can be expanded if grouping form elements is needed
End Sub

Public Sub ValidateFormData()
    ' Additional data validation before saving
    Dim IsValid As Boolean
    IsValid = True
    
    With UserFormVhIsh
        ' Logical data integrity check
        
        ' If transfer date to executor is specified, executor must be selected
        If Trim(.txtDataPeredachi.Text) <> "" And Trim(.cmbIspolnitel.Text) = "" Then
            MsgBox LocalizationManager.GetText("If transfer date to executor is specified, executor must be selected!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .cmbIspolnitel.SetFocus
            IsValid = False
        End If
        
        ' If outgoing number to service is specified, date must be entered
        If Trim(.txtNomerIshVSlujbu.Text) <> "" And Trim(.txtDataIshVSlujbu.Text) = "" Then
            MsgBox LocalizationManager.GetText("If outgoing number to service is specified, date must be entered!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .txtDataIshVSlujbu.SetFocus
            IsValid = False
        End If
        
        ' If return number is specified, return date must be entered
        If Trim(.txtNomerVozvrata.Text) <> "" And Trim(.txtDataVozvrata.Text) = "" Then
            MsgBox LocalizationManager.GetText("If return number is specified, return date must be entered!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .txtDataVozvrata.SetFocus
            IsValid = False
        End If
        
        ' If outgoing envelope number is specified, date must be entered
        If Trim(.txtNomerIshKonvert.Text) <> "" And Trim(.txtDataIshKonvert.Text) = "" Then
            MsgBox LocalizationManager.GetText("If outgoing envelope number is specified, date must be entered!"), vbExclamation, LocalizationManager.GetText("Data Validation")
            .txtDataIshKonvert.SetFocus
            IsValid = False
        End If
    End With
    
    If Not IsValid Then
        DataManager.FormDataChanged = True
    End If
End Sub

Public Sub FormatCurrencyField()
    ' Format amount field as currency
    With UserFormVhIsh.txtSummaDoc
        If IsNumeric(.Text) And Trim(.Text) <> "" Then
            Dim amount As Double
            amount = CDbl(.Text)
            .Text = Format(amount, "#,##0.00")
        End If
    End With
End Sub

Public Sub ClearSearchResults()
    ' Clear search results
    With UserFormVhIsh
        .txtSearch.Text = ""
        .lstSearchResults.Clear
        .lstSearchResults.Visible = False
        .lblStatusBar.Caption = LocalizationManager.GetText("Search cleared")
    End With
End Sub

Public Function GetNextRecordNumber() As Long
    ' Get next record sequence number
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo ErrorHandler
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    GetNextRecordNumber = tblData.ListRows.Count + 1
    
    Exit Function
    
ErrorHandler:
    GetNextRecordNumber = 1
End Function

Public Sub ResetFormToDefaults()
    ' Reset form to default settings
    DataManager.CurrentRecordRow = 0
    DataManager.IsNewRecord = True
    DataManager.FormDataChanged = False
    
    Call DataManager.ClearForm
    Call SetupFormAppearance
    
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Form reset to default settings")
End Sub
