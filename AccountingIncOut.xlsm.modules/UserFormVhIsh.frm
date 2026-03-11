VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormVhIsh 
   Caption         =   "Система учёта входящих и исходящих"
   ClientHeight    =   12045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17940
   OleObjectBlob   =   "UserFormVhIsh.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormVhIsh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'==============================================
' FORM MANAGEMENT MODULE "IncOut" - UserFormVhIsh
' Purpose: Fully functional form for adding, editing, and searching records
' State: FIXED "Expected array" ERROR IN IsNaryadOnlyMode FUNCTION
' Version: 3.2.1
' Date: 09.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' PATCH_MARK_2026_03_11_2318_AUTOCOMPLETE_IMPORT_READY

' API declarations for working with screen resolution
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

' Constants for GetSystemMetrics
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

' Global variables for the form module
Private CurrentRecordRow As Long
Private IsNewRecord As Boolean
Private FormDataChanged As Boolean
Private IsComboFilterBusy As Boolean
Private ServiceItems As Variant
Private DocTypeItems As Variant
Private ExecutorItems As Variant

Public Sub SetFormRecordState(ByVal rowNumber As Long, ByVal newRecord As Boolean, Optional ByVal changed As Boolean = False)
    CurrentRecordRow = rowNumber
    IsNewRecord = newRecord
    FormDataChanged = changed
End Sub


' Button to clear all execution marks
Private Sub btnClearMarks_Click()
    Call ProvodkaIntegrationModule.ClearAllProvodkaMarks
    
    ' Update current record in form
    If RecordOperations.CurrentRecordRow > 0 Then
        Call Me.LoadRecordToForm(RecordOperations.CurrentRecordRow)
    End If
End Sub

Private Sub btnFindProvodka_Click()
    ' Button to search posting for current record
    Call ProvodkaIntegrationModule.FindProvodkaForCurrentRecord
End Sub

' Button for mass check of all records with 1C export
Private Sub btnMassCheck_Click()
    Dim response As VbMsgBoxResult
    Dim StartTime As Double
    
    On Error GoTo MassCheckError
    
    ' User warning
    response = MsgBox("[INFO] MASS RECORD CHECK" & vbCrLf & vbCrLf & _
                     "This will check ALL records in the TableIncOut table" & vbCrLf & _
                     "against the 1C export." & vbCrLf & vbCrLf & _
                     "This may take some time." & vbCrLf & _
                     "Continue?", _
                     vbYesNo + vbQuestion, "Mass Check Confirmation")
    
    If response = vbNo Then Exit Sub
    
    ' Update status in form
    Me.lblStatusBar.Caption = "Performing mass check..."
    Me.btnMassCheck.Enabled = False
    
    StartTime = Timer
    
    ' Start mass processing
    Call ProvodkaIntegrationModule.MassProcessWithFileSelection
    
    ' Restore UI
    Me.btnMassCheck.Enabled = True
    
    ' If current record is open, update its display
    If RecordOperations.CurrentRecordRow > 0 Then
        Call Me.LoadRecordToForm(RecordOperations.CurrentRecordRow)
    End If
    
    ' Update status
    Me.lblStatusBar.Caption = "Mass check completed in " & _
                             Format(Timer - StartTime, "0.0") & " sec."
    
    Exit Sub
    
MassCheckError:
    Me.btnMassCheck.Enabled = True
    Me.lblStatusBar.Caption = "Mass check error: " & Err.description
    MsgBox "Error performing mass check:" & vbCrLf & Err.description, vbCritical, "Error"
End Sub

' Button to show matching statistics
Private Sub btnShowStats_Click()
    Call ProvodkaIntegrationModule.ShowMatchingStatistics
End Sub

Private Sub CommandButton1_Click()

End Sub

' ===============================================
' FORM EVENTS
' ===============================================

Private Sub UserForm_Initialize()
    ' Масштабирование и центрирование
    Call ResizeAndCenterForm
    
    ' АВТОМАТИЧЕСКИЙ ПЕРЕВОД ВСЕГО ИНТЕРФЕЙСА ФОРМЫ
    Call LocalizationManager.TranslateForm(Me)
    
    Call InitializeForm
    Call LoadSettings
    Call RecordOperations.UpdateStatusBar
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If HasUnsavedChanges() Then
        Dim response As VbMsgBoxResult
        response = MsgBox("You have unsaved changes. Save before closing?", _
                         vbYesNoCancel + vbQuestion, "Unsaved Changes")
        
        Select Case response
            Case vbYes
                Call SaveCurrentRecord
                Call SaveSettings
            Case vbCancel
                Cancel = True
                Exit Sub
        End Select
    End If
    
    Call SaveSettings
    
    ' Return focus to table
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    wsData.Activate
    Application.StatusBar = "Form closed. Interactive table system active."
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Hotkeys
    If Shift = 2 Then ' Ctrl pressed
        Select Case KeyCode
            Case vbKeyS ' Ctrl+S = Save
                Call SaveCurrentRecord
            Case vbKeyN ' Ctrl+N = New record
                Call RecordOperations.ClearForm
            Case vbKeyF ' Ctrl+F = Focus on search
                Me.txtSearch.SetFocus
        End Select
    ElseIf KeyCode = vbKeyF3 Then ' F3 = Next search result
        If Me.lstSearchResults.Visible And Me.lstSearchResults.ListCount > 0 Then
            With Me.lstSearchResults
                If .ListIndex < .ListCount - 1 Then
                    .ListIndex = .ListIndex + 1
                Else
                    .ListIndex = 0
                End If
            End With
        End If
    End If
    
    ' Check Escape key press (code 27)
    If KeyCode = vbKeyEscape Then
        ' Close form
        Me.Hide
        ' Or for full unload use: Unload Me
    End If
End Sub

Private Sub InitializeForm()
    CurrentRecordRow = 0
    IsNewRecord = True
    FormDataChanged = False
    
    Call LoadComboBoxData
    Call SetupNewComboBoxes
    Call RecordOperations.ClearForm
    
    Me.txtSearch.Text = ""
    Me.lstSearchResults.Clear
    Me.lstSearchResults.Visible = False
    Call SetupSearchResultsList
    
    Me.txtNomerPP.Enabled = False
    Me.txtNomerPP.BackColor = RGB(240, 240, 240)
    Me.txtNomerPP.Locked = True
    
    Call SetupNaryadField
    
    Me.btnDelete.Enabled = False
    Call SetupFormAppearance
    
    Me.txtSearch.Width = 300
    Me.txtSearch.Height = 24
    
    ' Динамический перевод статуса
    Me.lblStatusBar.Caption = LocalizationManager.GetText("Ready to work. Hotkeys: Ctrl+S (save), Ctrl+N (new), Ctrl+F (search)")
End Sub

' Setup order field (now optional)
Private Sub SetupNaryadField()
    With Me.txtNaryadInfo
        .MaxLength = 100 ' Length limit
        .BackColor = RGB(255, 255, 255) ' White color - optional field
        .Text = ""
    End With
End Sub

' Setup new ComboBoxes
Private Sub SetupNewComboBoxes()
    ' Setup cmbVidDocumenta (Inc./Out.)
    With Me.cmbVidDocumenta
        .Clear
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        .AddItem "Inc."
        .AddItem "Out."
        .ListIndex = -1
    End With
    
    ' Setup cmbStatusPodtverjdenie
    With Me.cmbStatusPodtverjdenie
        .Clear
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        .AddItem ""
        .AddItem "Confirmed"
        .AddItem "Sent Out."
        .ListIndex = 0
    End With
End Sub

' Setup search results list
Private Sub SetupSearchResultsList()
    With Me.lstSearchResults
        .ColumnCount = 1
        .ColumnWidths = "550"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Width = 420
        .Height = 120
        .Visible = False
        .MultiSelect = fmMultiSelectSingle
    End With
End Sub

Private Sub LoadComboBoxData()
    Dim wsSettings As Worksheet
    
    ' Check if "Dictionaries" sheet exists
    On Error GoTo SettingsError
    Set wsSettings = ThisWorkbook.Worksheets("Dictionaries")
    On Error GoTo 0
    
    ' Load services
    Call LoadComboData(Me.cmbSlujba, wsSettings, "A:A", "Services", "Службы")
    
    ' Load document types to cmbVidDoc (4th column)
    Call LoadComboData(Me.cmbVidDoc, wsSettings, "C:C", "Document Types", "Виды документов")
    
    ' Load executors
    Call LoadComboData(Me.cmbIspolnitel, wsSettings, "E:E", "Executors", "Исполнители")
    Call CacheComboSources
    
    ' Load data for cmbOtKogoPostupil autocomplete
    Call LoadAutoCompleteData
    
    ' Load data for txtNaryadInfo autocomplete
    Call LoadNaryadAutoCompleteData
    
    Exit Sub
    
SettingsError:
    MsgBox "Sheet 'Dictionaries' not found. Create a sheet for dictionaries.", vbExclamation, "Error"
End Sub

' Load data for order field autocomplete
Private Sub LoadNaryadAutoCompleteData()
    ' In the future, autocomplete for orders can be added
    ' from a separate table or already entered data
End Sub

' Load data for autocomplete
Private Sub LoadAutoCompleteData()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    Dim cellValue As String
    Dim uniqueValues As Collection
    Dim CurrentValue As String
    
    On Error GoTo LoadAutoError
    
    CurrentValue = CStr(Me.cmbOtKogoPostupil.value)
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    Set uniqueValues = New Collection
    
    ' Collect unique values from the 9th column
    If tblData.ListRows.Count > 0 Then
        For i = 1 To tblData.ListRows.Count
            cellValue = Trim(CStr(tblData.DataBodyRange.Cells(i, 9).value))
            If cellValue <> "" Then
                On Error Resume Next
                uniqueValues.Add cellValue, cellValue
                On Error GoTo LoadAutoError
            End If
        Next i
    End If
    
    ' Setup autocomplete for cmbOtKogoPostupil
    With Me.cmbOtKogoPostupil
        .Clear
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        .MatchEntry = fmMatchEntryComplete
        
        Dim item As Variant
        For Each item In uniqueValues
            .AddItem item
        Next item
        
        If Trim(CurrentValue) <> "" Then
            .value = CurrentValue
        Else
            .ListIndex = -1
        End If
    End With
    
    Exit Sub
    
LoadAutoError:
    If Trim(CurrentValue) <> "" Then
        Me.cmbOtKogoPostupil.value = CurrentValue
    End If
End Sub

Private Sub CacheComboSources()
    ServiceItems = GetComboItems(Me.cmbSlujba)
    DocTypeItems = GetComboItems(Me.cmbVidDoc)
    ExecutorItems = GetComboItems(Me.cmbIspolnitel)
End Sub

Private Function GetComboItems(TargetCombo As MSForms.ComboBox) As Variant
    Dim Result() As String
    Dim i As Long

    If TargetCombo.ListCount = 0 Then
        GetComboItems = Empty
        Exit Function
    End If

    ReDim Result(0 To TargetCombo.ListCount - 1)

    For i = 0 To TargetCombo.ListCount - 1
        Result(i) = CStr(TargetCombo.List(i))
    Next i

    GetComboItems = Result
End Function

Private Sub ResetComboList(TargetCombo As MSForms.ComboBox, SourceItems As Variant)
    Dim i As Long
    Dim CurrentText As String

    If IsEmpty(SourceItems) Then Exit Sub

    CurrentText = TargetCombo.Text

    IsComboFilterBusy = True
    TargetCombo.Clear

    For i = LBound(SourceItems) To UBound(SourceItems)
        TargetCombo.AddItem CStr(SourceItems(i))
    Next i

    TargetCombo.Text = CurrentText
    TargetCombo.SelStart = Len(CurrentText)
    TargetCombo.SelLength = 0
    IsComboFilterBusy = False
End Sub

Private Sub FilterComboByText(TargetCombo As MSForms.ComboBox, SourceItems As Variant)
    Dim SearchText As String
    Dim FirstPrefixMatch As String
    Dim ItemText As String
    Dim i As Long

    If IsComboFilterBusy Then Exit Sub
    If IsEmpty(SourceItems) Then Exit Sub

    SearchText = TargetCombo.Text

    IsComboFilterBusy = True
    TargetCombo.Clear

    For i = LBound(SourceItems) To UBound(SourceItems)
        ItemText = CStr(SourceItems(i))

        If SearchText = "" Or InStr(1, ItemText, SearchText, vbTextCompare) > 0 Then
            TargetCombo.AddItem ItemText

            If FirstPrefixMatch = "" Then
                If LCase$(Left$(ItemText, Len(SearchText))) = LCase$(SearchText) Then
                    FirstPrefixMatch = ItemText
                End If
            End If
        End If
    Next i

    If SearchText = "" Then
        TargetCombo.Text = ""
        TargetCombo.SelStart = 0
        TargetCombo.SelLength = 0
    ElseIf FirstPrefixMatch <> "" Then
        TargetCombo.Text = FirstPrefixMatch
        TargetCombo.SelStart = Len(SearchText)
        TargetCombo.SelLength = Len(FirstPrefixMatch) - Len(SearchText)
    Else
        TargetCombo.Text = SearchText
        TargetCombo.SelStart = Len(SearchText)
        TargetCombo.SelLength = 0
    End If

    If TargetCombo.ListCount > 0 Then
        TargetCombo.DropDown
    End If

    IsComboFilterBusy = False
End Sub

Private Sub HandleComboKeyUp(TargetCombo As MSForms.ComboBox, SourceItems As Variant, ByVal KeyCode As MSForms.ReturnInteger)
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyReturn, vbKeyEscape, vbKeyTab
            Exit Sub
        Case vbKeyDelete, vbKeyBack
            Call FilterComboByText(TargetCombo, SourceItems)
        Case Else
            Call FilterComboByText(TargetCombo, SourceItems)
    End Select
End Sub

Private Function FindHeaderCell(SourceSheet As Worksheet, SourceColumn As String, PrimaryHeader As String, Optional SecondaryHeader As String = "") As Range
    Set FindHeaderCell = SourceSheet.Range(SourceColumn).Find(PrimaryHeader, LookIn:=xlValues, LookAt:=xlWhole)

    If FindHeaderCell Is Nothing And SecondaryHeader <> "" Then
        Set FindHeaderCell = SourceSheet.Range(SourceColumn).Find(SecondaryHeader, LookIn:=xlValues, LookAt:=xlWhole)
    End If
End Function

Private Sub LoadComboData(TargetCombo As MSForms.ComboBox, SourceSheet As Worksheet, SourceColumn As String, PrimaryHeader As String, Optional SecondaryHeader As String = "")
    Dim HeaderCell As Range
    Dim CurrentCell As Range
    Dim StartRow As Long

    On Error GoTo LoadError

    With TargetCombo
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        .Clear
    End With

    Set HeaderCell = FindHeaderCell(SourceSheet, SourceColumn, PrimaryHeader, SecondaryHeader)
    If HeaderCell Is Nothing Then
        Exit Sub
    End If

    StartRow = HeaderCell.Row + 1
    Set CurrentCell = SourceSheet.Cells(StartRow, HeaderCell.Column)

    Do While Trim(CStr(CurrentCell.Value)) <> ""
        TargetCombo.AddItem CStr(CurrentCell.Value)
        Set CurrentCell = CurrentCell.Offset(1, 0)
    Loop

    TargetCombo.ListIndex = -1
    Exit Sub

LoadError:
    TargetCombo.Clear
End Sub

' ===============================================
' NAVIGATION BUTTON HANDLERS
' ===============================================

Private Sub btnFirst_Click()
    Call NavigateToRecord(1)
End Sub

Private Sub btnPrevious_Click()
    Call NavigateToPrevious
End Sub

Private Sub btnNext_Click()
    Call NavigateToNext
End Sub

Private Sub btnLast_Click()
    Call NavigateToLast
End Sub

' ===============================================
' MAIN BUTTON HANDLERS
' ===============================================

Private Sub btnSave_Click()
    Call SaveCurrentRecord
End Sub

Private Sub btnCancel_Click()
    Call CancelChanges
End Sub

Private Sub btnClear_Click()
    Call RecordOperations.ClearForm
End Sub

Private Sub btnNew_Click()
    If HasUnsavedChanges() Then
        Dim response As VbMsgBoxResult
        response = MsgBox("You have unsaved changes. Save before creating a new record?", _
                         vbYesNoCancel + vbQuestion, "Unsaved Changes")
        
        Select Case response
            Case vbYes
                Call SaveCurrentRecord
            Case vbCancel
                Exit Sub
        End Select
    End If
    
    Call RecordOperations.ClearForm
    Me.lblStatusBar.Caption = "Creating new record"
End Sub

Private Sub btnDelete_Click()
    MsgBox "Delete function temporarily disabled", vbInformation, "Information"
End Sub

' ===============================================
' REAL-TIME SEARCH
' ===============================================

Private Sub txtSearch_Change()
    Call PerformSearch
End Sub

Private Sub lstSearchResults_Click()
    Call SelectSearchResult
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call ClearSearch
    ElseIf KeyCode = vbKeyReturn Then
        If Me.lstSearchResults.Visible And Me.lstSearchResults.ListCount > 0 Then
            Me.lstSearchResults.ListIndex = 0
            Call SelectSearchResult
        End If
    ElseIf KeyCode = vbKeyDown Then
        If Me.lstSearchResults.Visible And Me.lstSearchResults.ListCount > 0 Then
            Call NavigateSearchResults("DOWN")
        End If
    ElseIf KeyCode = vbKeyUp Then
        If Me.lstSearchResults.Visible And Me.lstSearchResults.ListCount > 0 Then
            Call NavigateSearchResults("UP")
        End If
    End If
End Sub

Private Sub lstSearchResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call SelectSearchResult
        Case vbKeyEscape
            Me.txtSearch.SetFocus
        Case vbKeyDelete
            Call ClearSearch
            Me.txtSearch.SetFocus
        Case vbKeyF3
            Call NavigateSearchResults("DOWN")
        Case vbKeyF3 + 256
            Call NavigateSearchResults("UP")
    End Select
End Sub

Private Sub lstSearchResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call SelectSearchResult
    Call HighlightSearchTermInForm
End Sub

' ===============================================
' FIELD CHANGE HANDLERS
' ===============================================

Private Sub cmbVidDocumenta_Change()
    Call MarkFormAsChanged
    Call UpdateFieldAppearanceBasedOnNaryad ' Update field highlights
End Sub

Private Sub txtNomerDoc_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtSummaDoc_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtVhFRP_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataVhFRP_Change()
    Call MarkFormAsChanged
End Sub

Private Sub cmbOtKogoPostupil_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataPeredachi_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtNomerIshVSlujbu_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataIshVSlujbu_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtNomerVozvrata_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataVozvrata_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtNomerIshKonvert_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtDataIshKonvert_Change()
    Call MarkFormAsChanged
End Sub

Private Sub txtOtmetkaIspolnenie_Change()
    Call MarkFormAsChanged
End Sub

Private Sub cmbStatusPodtverjdenie_Change()
    Call MarkFormAsChanged
End Sub

Private Sub cmbSlujba_Change()
    If IsComboFilterBusy Then Exit Sub
    Call MarkFormAsChanged
End Sub

Private Sub cmbVidDoc_Change()
    If IsComboFilterBusy Then Exit Sub
    Call MarkFormAsChanged
End Sub

Private Sub cmbIspolnitel_Change()
    If IsComboFilterBusy Then Exit Sub
    Call MarkFormAsChanged
End Sub

Private Sub cmbSlujba_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call HandleComboKeyUp(Me.cmbSlujba, ServiceItems, KeyCode)
End Sub

Private Sub cmbVidDoc_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call HandleComboKeyUp(Me.cmbVidDoc, DocTypeItems, KeyCode)
End Sub

Private Sub cmbIspolnitel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call HandleComboKeyUp(Me.cmbIspolnitel, ExecutorItems, KeyCode)
End Sub

Private Sub cmbSlujba_DropButtonClick()
    Call ResetComboList(Me.cmbSlujba, ServiceItems)
End Sub

Private Sub cmbVidDoc_DropButtonClick()
    Call ResetComboList(Me.cmbVidDoc, DocTypeItems)
End Sub

Private Sub cmbIspolnitel_DropButtonClick()
    Call ResetComboList(Me.cmbIspolnitel, ExecutorItems)
End Sub

' Change order field with highlight update
Private Sub txtNaryadInfo_Change()
    Call MarkFormAsChanged
    Call UpdateFieldAppearanceBasedOnNaryad ' Update highlights when order changes
End Sub

' Update field appearance based on order mode
Private Sub UpdateFieldAppearanceBasedOnNaryad()
    On Error Resume Next
    Dim isNaryadOnly As Boolean
    Dim requiredColor As Long
    Dim optionalColor As Long
    
    requiredColor = RGB(255, 255, 224) ' Light yellow for required
    optionalColor = RGB(255, 255, 255) ' White for optional
    
    ' IMPORTANT: get function result into ANOTHER variable
    isNaryadOnly = IsNaryadOnlyMode()
    
    If isNaryadOnly Then
        ' "Order only" mode - make main fields optional
        Me.cmbVidDoc.BackColor = optionalColor
        Me.txtNomerDoc.BackColor = optionalColor
        Me.txtSummaDoc.BackColor = optionalColor
        Me.txtVhFRP.BackColor = optionalColor
        Me.txtDataVhFRP.BackColor = optionalColor
        
        Me.lblStatusBar.Caption = "ORDER MODE: Main document fields are optional"
    Else
        ' Normal mode - main fields are required
        Me.cmbVidDoc.BackColor = requiredColor
        Me.txtNomerDoc.BackColor = requiredColor
        Me.txtSummaDoc.BackColor = requiredColor
        Me.txtVhFRP.BackColor = requiredColor
        Me.txtDataVhFRP.BackColor = requiredColor
        
        If Me.lblStatusBar.Caption = "ORDER MODE: Main document fields are optional" Then
            Me.lblStatusBar.Caption = "Normal mode: all required fields must be filled"
        End If
    End If
End Sub

' CRITICALLY FIXED FUNCTION: Check "order only" mode
Private Function IsNaryadOnlyMode() As Boolean
    ' "Order only" mode is active if:
    ' 1. Order field is filled
    ' 2. Service is filled
    ' 3. Main fields are NOT filled: doc type, number, amount, IncFRP, date IncFRP
    
    On Error GoTo ModeError
    
    Dim naryadText As String
    Dim slujbaText As String
    Dim vidDocText As String
    Dim nomerDocText As String
    Dim summaDocText As String
    Dim vhFrpText As String
    Dim dataVhFrpText As String

    ' Safe read of form element values into local variables
    naryadText = ""
    If Not Me.txtNaryadInfo Is Nothing Then naryadText = Trim(CStr(Me.txtNaryadInfo.Text))

    slujbaText = ""
    If Not Me.cmbSlujba Is Nothing Then slujbaText = Trim(CStr(Me.cmbSlujba.Text))

    vidDocText = ""
    If Not Me.cmbVidDoc Is Nothing Then vidDocText = Trim(CStr(Me.cmbVidDoc.Text))

    nomerDocText = ""
    If Not Me.txtNomerDoc Is Nothing Then nomerDocText = Trim(CStr(Me.txtNomerDoc.Text))

    summaDocText = ""
    If Not Me.txtSummaDoc Is Nothing Then summaDocText = Trim(CStr(Me.txtSummaDoc.Text))

    vhFrpText = ""
    If Not Me.txtVhFRP Is Nothing Then vhFrpText = Trim(CStr(Me.txtVhFRP.Text))

    dataVhFrpText = ""
    If Not Me.txtDataVhFRP Is Nothing Then dataVhFrpText = Trim(CStr(Me.txtDataVhFRP.Text))

    ' 1) Order must be filled
    If naryadText = "" Then
        IsNaryadOnlyMode = False
        Exit Function
    End If

    ' 2) Service must be filled
    If slujbaText = "" Then
        IsNaryadOnlyMode = False
        Exit Function
    End If

    ' 3) Main fields must be empty
    If vidDocText <> "" Then
        IsNaryadOnlyMode = False
        Exit Function
    End If

    If nomerDocText <> "" Then
        IsNaryadOnlyMode = False
        Exit Function
    End If

    If summaDocText <> "" Then
        IsNaryadOnlyMode = False
        Exit Function
    End If

    If vhFrpText <> "" Then
        IsNaryadOnlyMode = False
        Exit Function
    End If

    If dataVhFrpText <> "" Then
        IsNaryadOnlyMode = False
        Exit Function
    End If

    ' All conditions met - "order only" mode is active
    IsNaryadOnlyMode = True
    Exit Function

ModeError:
    ' Return False on any error
    IsNaryadOnlyMode = False
End Function

' ===============================================
' FIELD VALIDATION ON EXIT
' ===============================================

Private Sub txtSummaDoc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Me.txtSummaDoc.Text) > 0 And Not IsNumeric(Me.txtSummaDoc.Text) Then
        MsgBox "Amount must be a number!", vbExclamation, "Data Validation"
        Cancel = True
        Me.txtSummaDoc.BackColor = RGB(255, 200, 200)
    Else
        ' Restore correct color based on mode
        Call UpdateFieldAppearanceBasedOnNaryad
    End If
End Sub

Private Sub txtDataVhFRP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataVhFRP, Cancel)
End Sub

Private Sub txtDataPeredachi_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataPeredachi, Cancel)
End Sub

Private Sub txtDataIshVSlujbu_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataIshVSlujbu, Cancel)
End Sub

Private Sub txtDataVozvrata_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataVozvrata, Cancel)
End Sub

Private Sub txtDataIshKonvert_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ValidateAndFormatDate(Me.txtDataIshKonvert, Cancel)
End Sub

' Validate order field on exit
Private Sub txtNaryadInfo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Order format validation (optional field)
    If Len(Trim(Me.txtNaryadInfo.Text)) > 0 Then
        Me.txtNaryadInfo.BackColor = RGB(255, 255, 255) ' White color for optional field
    End If
    
    ' Update other fields highlight when exiting order field
    Call UpdateFieldAppearanceBasedOnNaryad
End Sub

Private Sub ValidateAndFormatDate(dateField As MSForms.TextBox, Cancel As MSForms.ReturnBoolean)
    If Len(dateField.Text) > 0 Then
        ' Automatically add dots when entering date
        Dim DateText As String
        DateText = Replace(dateField.Text, ".", "")
        
        If Len(DateText) = 6 And IsNumeric(DateText) Then
            dateField.Text = Left(DateText, 2) & "." & Mid(DateText, 3, 2) & "." & Right(DateText, 2)
        End If
        
        ' Format validation
        If Not CommonUtilities.IsValidDateFormat(dateField.Text) Then
            MsgBox "Enter date in DD.MM.YY format", vbExclamation, "Data Validation"
            Cancel = True
            dateField.BackColor = RGB(255, 200, 200)
        Else
            ' Restore correct color based on mode
            Call UpdateFieldAppearanceBasedOnNaryad
        End If
    End If
End Sub

' ===============================================
' DATA MANAGEMENT FUNCTIONS
' ===============================================

Private Sub SaveCurrentRecord()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim newRow As ListRow
    Dim targetRowIndex As Long

    If Not ValidateRequiredFieldsConditional() Then
        Exit Sub
    End If

    If Not RecordOperations.ValidateDates() Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")

    If IsNewRecord Then
        Set newRow = tblData.ListRows.Add
        targetRowIndex = newRow.Index
        CurrentRecordRow = targetRowIndex

        Me.txtNomerPP.Text = CStr(targetRowIndex)
        Me.lblStatusBar.Caption = "Creating new record No. " & targetRowIndex
    Else
        If CurrentRecordRow < 1 Or CurrentRecordRow > tblData.ListRows.Count Then
            MsgBox "Error: invalid record number to update.", vbCritical, "Save Error"
            Exit Sub
        End If

        targetRowIndex = CurrentRecordRow
        Me.lblStatusBar.Caption = "Updating record No. " & targetRowIndex
    End If

    Call WriteFormDataToTable(tblData, targetRowIndex)

    IsNewRecord = False
    FormDataChanged = False

    RecordOperations.CurrentRecordRow = targetRowIndex
    RecordOperations.IsNewRecord = False
    RecordOperations.FormDataChanged = False

    Me.lblStatusBar.Caption = "Record No. " & targetRowIndex & " saved successfully"

    If targetRowIndex = tblData.ListRows.Count Then
        Call LoadAutoCompleteData
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error saving data: " & Err.description, vbCritical, "Save Error"
    Me.lblStatusBar.Caption = "Data save error"
End Sub

' Conditional required fields validation
Private Function ValidateRequiredFieldsConditional() As Boolean
    ValidateRequiredFieldsConditional = True
    
    ' Service is always required
    If Trim(CStr(Me.cmbSlujba.value)) = "" Then
        MsgBox "Field 'Service' is required!", vbExclamation, "Data Validation"
        Me.cmbSlujba.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    ' Check "order only" mode
    If IsNaryadOnlyMode() Then
        ' In "order only" mode, main fields are optional
        ' But at least the order must be filled
        If Trim(Me.txtNaryadInfo.Text) = "" Then
            MsgBox "In order mode, 'Order Info' field must be filled!", vbExclamation, "Data Validation"
            Me.txtNaryadInfo.SetFocus
            ValidateRequiredFieldsConditional = False
            Exit Function
        End If
        
        ' Successful validation for "order only" mode
        Exit Function
    End If
    
    ' Normal mode - check all required fields
    If Trim(CStr(Me.cmbVidDoc.value)) = "" Then
        MsgBox "Field 'Document Type' is required!", vbExclamation, "Data Validation"
        Me.cmbVidDoc.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(CStr(Me.cmbVidDocumenta.value)) = "" Then
        MsgBox "Field 'Document Group (Inc./Out.)' is required!", vbExclamation, "Data Validation"
        Me.cmbVidDocumenta.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(Me.txtNomerDoc.Text) = "" Then
        MsgBox "Field 'Document Number' is required!", vbExclamation, "Data Validation"
        Me.txtNomerDoc.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(Me.txtSummaDoc.Text) = "" Then
        MsgBox "Field 'Document Amount' is required!", vbExclamation, "Data Validation"
        Me.txtSummaDoc.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Not IsNumeric(Me.txtSummaDoc.Text) Then
        MsgBox "Field 'Document Amount' must contain a numeric value!", vbExclamation, "Data Validation"
        Me.txtSummaDoc.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(Me.txtVhFRP.Text) = "" Then
        MsgBox "Field 'Inc.FRP/Out.FRP' is required!", vbExclamation, "Data Validation"
        Me.txtVhFRP.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
    
    If Trim(Me.txtDataVhFRP.Text) = "" Then
        MsgBox "Field 'Date Inc.FRP/Out.FRP' is required!", vbExclamation, "Data Validation"
        Me.txtDataVhFRP.SetFocus
        ValidateRequiredFieldsConditional = False
        Exit Function
    End If
End Function

' Write form data to table
Private Sub WriteFormDataToTable(tbl As ListObject, RowIndex As Long)
    On Error GoTo WriteError
    
    tbl.DataBodyRange.Cells(RowIndex, 1).value = RowIndex
    tbl.DataBodyRange.Cells(RowIndex, 2).value = Me.cmbSlujba.value
    tbl.DataBodyRange.Cells(RowIndex, 3).value = Me.cmbVidDocumenta.value
    tbl.DataBodyRange.Cells(RowIndex, 4).value = Me.cmbVidDoc.value
    tbl.DataBodyRange.Cells(RowIndex, 5).value = Me.txtNomerDoc.Text
    
    If IsNumeric(Me.txtSummaDoc.Text) Then
        tbl.DataBodyRange.Cells(RowIndex, 6).value = CDbl(Me.txtSummaDoc.Text)
    Else
        tbl.DataBodyRange.Cells(RowIndex, 6).value = 0
    End If
    
    tbl.DataBodyRange.Cells(RowIndex, 7).value = Me.txtVhFRP.Text
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 8), Me.txtDataVhFRP.Text)
    tbl.DataBodyRange.Cells(RowIndex, 9).value = Me.cmbOtKogoPostupil.value
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 10), Me.txtDataPeredachi.Text)
    tbl.DataBodyRange.Cells(RowIndex, 11).value = Me.cmbIspolnitel.value
    tbl.DataBodyRange.Cells(RowIndex, 12).value = Me.txtNomerIshVSlujbu.Text
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 13), Me.txtDataIshVSlujbu.Text)
    tbl.DataBodyRange.Cells(RowIndex, 14).value = Me.txtNomerVozvrata.Text
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 15), Me.txtDataVozvrata.Text)
    tbl.DataBodyRange.Cells(RowIndex, 16).value = Me.txtNomerIshKonvert.Text
    Call CommonUtilities.WriteDateToCell(tbl.DataBodyRange.Cells(RowIndex, 17), Me.txtDataIshKonvert.Text)
    tbl.DataBodyRange.Cells(RowIndex, 18).value = Me.txtOtmetkaIspolnenie.Text
    tbl.DataBodyRange.Cells(RowIndex, 19).value = Me.cmbStatusPodtverjdenie.value
    tbl.DataBodyRange.Cells(RowIndex, 20).value = Me.txtNaryadInfo.Text
    
    Exit Sub
    
WriteError:
    MsgBox "Error writing data to table: " & Err.description, vbCritical, "Error"
End Sub

' ===============================================
' LOAD RECORD FROM TABLE TO FORM
' ===============================================

Private Sub UpdateProvodkaIndicator()
    Dim otmetka As String
    otmetka = Trim(Me.txtOtmetkaIspolnenie.Text)
    
    If otmetka <> "" Then
        Me.txtOtmetkaIspolnenie.BackColor = RGB(200, 255, 200) ' Light green
        Me.lblStatusBar.Caption = Me.lblStatusBar.Caption & " | + Posting: " & otmetka
    Else
        Me.txtOtmetkaIspolnenie.BackColor = RGB(255, 255, 255) ' White
    End If
End Sub

Public Sub LoadRecordToForm(RowIndex As Long)
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo LoadError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    If RowIndex < 1 Or RowIndex > tblData.ListRows.Count Then
        Exit Sub
    End If
    
    IsNewRecord = False
    CurrentRecordRow = RowIndex
    
    Me.txtNomerPP.Text = CStr(RowIndex)
    Me.cmbSlujba.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 2).value)
    Me.cmbVidDocumenta.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 3).value)
    Me.cmbVidDoc.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 4).value)
    Me.txtNomerDoc.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 5).value)
    Me.txtSummaDoc.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 6).value)
    Me.txtVhFRP.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 7).value)
    Me.txtDataVhFRP.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 8))
    Me.cmbOtKogoPostupil.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 9).value)
    Me.txtDataPeredachi.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 10))
    Me.cmbIspolnitel.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 11).value)
    Me.txtNomerIshVSlujbu.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 12).value)
    Me.txtDataIshVSlujbu.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 13))
    Me.txtNomerVozvrata.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 14).value)
    Me.txtDataVozvrata.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 15))
    Me.txtNomerIshKonvert.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 16).value)
    Me.txtDataIshKonvert.Text = CommonUtilities.FormatDateCell(tblData.DataBodyRange.Cells(RowIndex, 17))
    Me.txtOtmetkaIspolnenie.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 18).value)
    Me.cmbStatusPodtverjdenie.value = CStr(tblData.DataBodyRange.Cells(RowIndex, 19).value)
    
    If tblData.DataBodyRange.Columns.Count >= 20 Then
        Me.txtNaryadInfo.Text = CStr(tblData.DataBodyRange.Cells(RowIndex, 20).value)
    Else
        Me.txtNaryadInfo.Text = ""
    End If
    
    FormDataChanged = False
    
    ' Update fields highlight after loading record
    Call UpdateFieldAppearanceBasedOnNaryad
    
    Exit Sub
    
LoadError:
    MsgBox "Error loading record: " & Err.description, vbExclamation, "Error"
    Call RecordOperations.ClearForm
End Sub

' ===============================================
' HELPER FUNCTIONS
' ===============================================

Private Function HasUnsavedChanges() As Boolean
    HasUnsavedChanges = FormDataChanged
End Function

Private Sub CancelChanges()
    If IsNewRecord Then
        Call RecordOperations.ClearForm
    Else
        Call NavigateToRecord(CurrentRecordRow)
    End If
    
    FormDataChanged = False
    Me.lblStatusBar.Caption = "Changes cancelled"
End Sub

' Methods for working with form from table
Public Sub ShowFromTable(rowNumber As Long, fieldName As String)
    Call Me.LoadRecordToForm(rowNumber)
    
    CurrentRecordRow = rowNumber
    IsNewRecord = False
    FormDataChanged = False
    
    Me.Show vbModeless
    
    Application.OnTime Now + TimeValue("00:00:01"), "'SetFieldFocus """ & fieldName & """'"
    
    Me.lblStatusBar.Caption = "Opened from table | Record No." & rowNumber & " | Field: " & TableEventHandler.GetFieldDisplayName(fieldName)
End Sub

Public Sub SetFieldFocus(fieldName As String)
    Call TableEventHandler.ActivateFormField(fieldName)
End Sub

Private Sub UpdateNavigationButtons()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    On Error GoTo NavigationError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    Me.btnFirst.Enabled = (CurrentRecordRow > 1)
    Me.btnPrevious.Enabled = (CurrentRecordRow > 1)
    Me.btnNext.Enabled = (CurrentRecordRow < tblData.ListRows.Count)
    Me.btnLast.Enabled = (CurrentRecordRow < tblData.ListRows.Count)
    
    Exit Sub
    
NavigationError:
    Me.btnFirst.Enabled = False
    Me.btnPrevious.Enabled = False
    Me.btnNext.Enabled = False
    Me.btnLast.Enabled = False
End Sub

Private Sub MarkFormAsChanged()
    FormDataChanged = True
    Call RecordOperations.UpdateStatusBar
End Sub

' ===============================================
' SETTINGS FUNCTIONS
' ===============================================

Private Sub SaveSettings()
    Dim wsSettings As Worksheet
    
    On Error GoTo SaveError
    
    Set wsSettings = GetSettingsWorksheet()
    
    With wsSettings
        .Cells(1, 1).value = "CurrentRecordRow"
        .Cells(1, 2).value = CurrentRecordRow
        
        .Cells(2, 1).value = "FormTop"
        .Cells(2, 2).value = Me.Top
        
        .Cells(3, 1).value = "FormLeft"
        .Cells(3, 2).value = Me.Left
        
        .Cells(4, 1).value = "LastSaved"
        .Cells(4, 2).value = Now()
    End With
    
    Exit Sub
    
SaveError:
    ' Ignore settings save errors
End Sub

Private Sub LoadSettings()
    Dim wsSettings As Worksheet
    
    On Error GoTo DefaultSettings
    
    Set wsSettings = GetSettingsWorksheet()
    
    With wsSettings
        If .Cells(1, 2).value <> "" Then
            CurrentRecordRow = .Cells(1, 2).value
        End If
        
        ' DO NOT load saved positions - form is always centered automatically
        ' Dimensions and position are set in ResizeAndCenterForm
    End With
    
    If CurrentRecordRow > 0 Then
        Call NavigateToRecord(CurrentRecordRow)
    Else
        Call RecordOperations.ClearForm
    End If
    
    Exit Sub
    
DefaultSettings:
    CurrentRecordRow = 0
    IsNewRecord = True
    FormDataChanged = False
End Sub

Private Function GetSettingsWorksheet() As Worksheet
    Dim Ws As Worksheet
    
    On Error GoTo CreateSheet
    
    Set Ws = ThisWorkbook.Worksheets("SystemSettings")
    Set GetSettingsWorksheet = Ws
    Exit Function
    
CreateSheet:
    Set Ws = ThisWorkbook.Worksheets.Add
    Ws.Name = "SystemSettings"
    Ws.Visible = xlSheetHidden
    Set GetSettingsWorksheet = Ws
End Function

' ===============================================
' UI APPEARANCE FUNCTIONS
' ===============================================

Private Sub SetupFormAppearance()
    Call SetupRequiredFieldsHighlight
    Call SetupFieldAlignment
    Call SetupNavigationButtons
    Call SetupStatusBar
End Sub

' Setup highlight based on order mode
Private Sub SetupRequiredFieldsHighlight()
    Dim requiredColor As Long
    requiredColor = RGB(255, 255, 224)
    
    With Me
        ' Fields that are always required
        .cmbSlujba.BackColor = requiredColor
        .cmbVidDocumenta.BackColor = requiredColor
        
        ' Order field is always optional (white)
        .txtNaryadInfo.BackColor = RGB(255, 255, 255)
        
        ' Seq no field is disabled
        .txtNomerPP.BackColor = RGB(240, 240, 240)
    End With
    
    ' Conditionally required fields - update based on mode
    Call UpdateFieldAppearanceBasedOnNaryad
End Sub

Private Sub SetupFieldAlignment()
    With Me
        .txtSummaDoc.TextAlign = fmTextAlignRight
        .txtNomerPP.TextAlign = fmTextAlignCenter
        .txtNomerDoc.TextAlign = fmTextAlignCenter
        .txtVhFRP.TextAlign = fmTextAlignCenter
        
        .txtDataVhFRP.MaxLength = 8
        .txtDataPeredachi.MaxLength = 8
        .txtDataIshVSlujbu.MaxLength = 8
        .txtDataVozvrata.MaxLength = 8
        .txtDataIshKonvert.MaxLength = 8
        
        .txtNaryadInfo.MaxLength = 100
        .txtNaryadInfo.TextAlign = fmTextAlignLeft
    End With
End Sub

Private Sub SetupNavigationButtons()
    With Me
        .btnFirst.Caption = "|< First"
        .btnPrevious.Caption = "< Prev"
        .btnNext.Caption = "Next >"
        .btnLast.Caption = "Last >|"
        
        .btnFirst.Width = 70
        .btnPrevious.Width = 70
        .btnNext.Width = 70
        .btnLast.Width = 70
    End With
End Sub

Private Sub SetupStatusBar()
    With Me.lblStatusBar
        .BackColor = RGB(245, 245, 245)
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = 9
        .Font.Name = "Segoe UI"
    End With
End Sub

' ===============================================
' SCREEN RESOLUTION AND SCALING FUNCTIONS
' ===============================================

Private Function GetScreenResolution() As String
    On Error GoTo ErrorHandler
    
    Dim screenWidth As Long
    Dim screenHeight As Long
    
    screenWidth = GetSystemMetrics(SM_CXSCREEN)
    screenHeight = GetSystemMetrics(SM_CYSCREEN)
    
    GetScreenResolution = screenWidth & "x" & screenHeight
    
    Exit Function
    
ErrorHandler:
    GetScreenResolution = "1920x1080"
End Function

Private Function GetScreenWidth() As Long
    On Error GoTo ErrorHandler
    
    GetScreenWidth = GetSystemMetrics(SM_CXSCREEN)
    Exit Function
    
ErrorHandler:
    GetScreenWidth = 1920 ' Default
End Function

Private Function GetScreenHeight() As Long
    On Error GoTo ErrorHandler
    
    GetScreenHeight = GetSystemMetrics(SM_CYSCREEN)
    Exit Function
    
ErrorHandler:
    GetScreenHeight = 1080 ' Default
End Function

Private Sub ResizeAndCenterForm()
    On Error GoTo ErrorHandler
    
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim formWidth As Single
    Dim formHeight As Single
    Dim scaleFactor As Double
    Dim minScaleFactor As Double
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim newLeft As Single
    Dim newTop As Single
    Dim usableScreenWidth As Single
    Dim usableScreenHeight As Single
    
    originalWidth = Me.Width
    originalHeight = Me.Height
    
    screenWidth = GetScreenWidth()
    screenHeight = GetScreenHeight()
    
    Dim screenWidthPoints As Double
    Dim screenHeightPoints As Double
    screenWidthPoints = screenWidth * (72 / 96)
    screenHeightPoints = screenHeight * (72 / 96)
    
    usableScreenWidth = screenWidthPoints * 0.9
    usableScreenHeight = screenHeightPoints * 0.9
    
    If originalWidth <= usableScreenWidth And originalHeight <= usableScreenHeight Then
        formWidth = originalWidth
        formHeight = originalHeight
        scaleFactor = 1#
    Else
        Dim scaleX As Double
        Dim scaleY As Double
        scaleX = usableScreenWidth / originalWidth
        scaleY = usableScreenHeight / originalHeight
        
        scaleFactor = Application.WorksheetFunction.Min(scaleX, scaleY)
        
        minScaleFactor = 0.75
        If scaleFactor < minScaleFactor Then scaleFactor = minScaleFactor
        
        formWidth = originalWidth * scaleFactor
        formHeight = originalHeight * scaleFactor
        
        If formWidth > usableScreenWidth Then
            formWidth = usableScreenWidth
            scaleFactor = formWidth / originalWidth
            formHeight = originalHeight * scaleFactor
        End If
        If formHeight > usableScreenHeight Then
            formHeight = usableScreenHeight
            scaleFactor = formHeight / originalHeight
            formWidth = originalWidth * scaleFactor
        End If
    End If
    
    Me.Width = formWidth
    Me.Height = formHeight
    
    newLeft = (screenWidthPoints - formWidth) / 2
    newTop = (screenHeightPoints - formHeight) / 2
    
    If newLeft < 0 Then newLeft = 0
    If newTop < 0 Then newTop = 0
    If newLeft + formWidth > screenWidthPoints Then newLeft = screenWidthPoints - formWidth
    If newTop + formHeight > screenHeightPoints Then newTop = screenHeightPoints - formHeight
    
    Me.Left = newLeft
    Me.Top = newTop
    
    Exit Sub
    
ErrorHandler:
    Me.StartUpPosition = 1 ' CenterOwner
End Sub

