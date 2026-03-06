Attribute VB_Name = "TableEventHandler"
'==============================================
' TABLE EVENT HANDLER MODULE - TableEventHandler
' Purpose: Handling table clicks to open form with activation of specific field
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.7.0
'==============================================

Option Explicit

Private wsData As Worksheet
Private tblVhIsh As ListObject
Private isSystemActive As Boolean
Private lastSelectedRow As Long
Private lastSelectedColumn As Long
Private currentRightClickRow As Long
Private contextMenuInitialized As Boolean
Private AppEventsHandler As AppEventHandler

Public Sub InitializeTableEvents()
    On Error GoTo InitError
    isSystemActive = False
    contextMenuInitialized = False
    Dim wsExists As Boolean: wsExists = False
    Dim Ws As Worksheet
    For Each Ws In ThisWorkbook.Worksheets
        If Ws.Name = "IncOut" Then
            wsExists = True
            Set wsData = Ws
            Exit For
        End If
    Next Ws
    If Not wsExists Then
        MsgBox LocalizationManager.GetText("Sheet 'IncOut' not found in the workbook!"), vbCritical, LocalizationManager.GetText("Initialization Error")
        Exit Sub
    End If
    Dim tblExists As Boolean: tblExists = False
    Dim tbl As ListObject
    For Each tbl In wsData.ListObjects
        If tbl.Name = "TableIncOut" Then
            tblExists = True
            Set tblVhIsh = tbl
            Exit For
        End If
    Next tbl
    If Not tblExists Then
        MsgBox LocalizationManager.GetText("Table 'TableIncOut' not found on sheet 'IncOut'!") & vbCrLf & _
               LocalizationManager.GetText("Make sure the table exists and has the correct name."), vbCritical, LocalizationManager.GetText("Initialization Error")
        Exit Sub
    End If
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = LocalizationManager.GetText("Warning: table is empty. Add data to work.")
    End If
    lastSelectedRow = 0
    lastSelectedColumn = 0
    currentRightClickRow = 0
    Set AppEventsHandler = New AppEventHandler
    Call AppEventsHandler.InitializeAppEvents
    isSystemActive = True
    contextMenuInitialized = True
    Application.StatusBar = LocalizationManager.GetText("Interactive forms system is active (Excel 2021). Use Ctrl+Shift+M for menu.")
    Exit Sub
InitError:
    isSystemActive = False
    contextMenuInitialized = False
    MsgBox LocalizationManager.GetText("Critical error initializing table events system:") & vbCrLf & Err.Number & " - " & Err.description, vbCritical, LocalizationManager.GetText("Critical Error")
End Sub

Public Sub ShowActionMenuExcel2021()
    Dim userChoice As String, currentCell As Range, intersectRange As Range, RowNumber As Long
    On Error GoTo MenuError
    If ActiveSheet.Name <> "IncOut" Then
        Application.StatusBar = LocalizationManager.GetText("Switch to sheet 'IncOut' to use the action menu.")
        Exit Sub
    End If
    Set currentCell = Selection
    If Not tblVhIsh Is Nothing Then
        If Not tblVhIsh.DataBodyRange Is Nothing Then
            Set intersectRange = Intersect(currentCell, tblVhIsh.DataBodyRange)
            If Not intersectRange Is Nothing And currentCell.Cells.Count = 1 Then
                RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
                userChoice = InputBox( _
                    LocalizationManager.GetText("ACTION MENU FOR RECORD No.") & RowNumber & ":" & vbCrLf & vbCrLf & _
                    LocalizationManager.GetText("Available commands:") & vbCrLf & _
                    "1 - " & LocalizationManager.GetText("Edit in form") & vbCrLf & _
                    "2 - " & LocalizationManager.GetText("Duplicate record") & vbCrLf & _
                    "3 - " & LocalizationManager.GetText("Show record information") & vbCrLf & _
                    "0 - " & LocalizationManager.GetText("Cancel") & vbCrLf & vbCrLf & _
                    LocalizationManager.GetText("Enter command number:"), _
                    LocalizationManager.GetText("Excel 2021 Action Menu"), "1")
                Select Case userChoice
                    Case "1": Call OpenFormForCurrentSelection
                    Case "2": Call DuplicateCurrentRecord
                    Case "3": Call ShowRecordInfo(RowNumber)
                    Case "0", "": ' Cancel
                    Case Else: Application.StatusBar = LocalizationManager.GetText("Invalid choice. Use numbers 0-3.")
                End Select
            Else
                Application.StatusBar = LocalizationManager.GetText("Select a cell inside the data table to use the menu.")
            End If
        Else
            Application.StatusBar = LocalizationManager.GetText("Table is empty!")
        End If
    Else
        Application.StatusBar = LocalizationManager.GetText("Table not found!")
    End If
    Exit Sub
MenuError:
    Application.StatusBar = LocalizationManager.GetText("Error displaying menu: ") & Err.description
End Sub

Private Sub ShowRecordInfo(RowNumber As Long)
    Dim info As String
    On Error GoTo InfoError
    If tblVhIsh Is Nothing Or tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    If RowNumber < 1 Or RowNumber > tblVhIsh.ListRows.Count Then Exit Sub
    info = LocalizationManager.GetText("RECORD No.") & RowNumber & LocalizationManager.GetText(" INFO:") & vbCrLf & vbCrLf
    info = info & LocalizationManager.GetText("Service: ") & CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 2).value) & vbCrLf
    info = info & LocalizationManager.GetText("Document Type: ") & CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 4).value) & vbCrLf
    info = info & LocalizationManager.GetText("Document Number: ") & CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 5).value) & vbCrLf
    info = info & LocalizationManager.GetText("Amount: ") & CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 6).value) & LocalizationManager.GetText(" rub.") & vbCrLf & vbCrLf
    info = info & LocalizationManager.GetText("Press OK to continue.")
    MsgBox info, vbInformation, LocalizationManager.GetText("Record Information")
    Exit Sub
InfoError:
End Sub

Public Sub AddContextMenuButton()
    On Error GoTo MenuError
    Call RemoveContextMenuButton
    Dim contextMenu As CommandBar
    Set contextMenu = Application.CommandBars("Cell")
    With contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
        .Caption = LocalizationManager.GetText("Edit in form")
        .OnAction = "TableEventHandler.OpenFormForCurrentSelection"
        .FaceId = 162
        .BeginGroup = True
        .Tag = "CustomEdit_VhIsh_2021"
    End With
    With contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
        .Caption = LocalizationManager.GetText("Duplicate record")
        .OnAction = "TableEventHandler.DuplicateCurrentRecord"
        .FaceId = 19
        .Tag = "CustomDuplicate_VhIsh_2021"
    End With
    With contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
        .Caption = LocalizationManager.GetText("Action menu (Ctrl+Shift+M)")
        .OnAction = "TableEventHandler.ShowActionMenuExcel2021"
        .FaceId = 923
        .Tag = "CustomAltMenu_VhIsh_2021"
    End With
    With contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
        .Caption = LocalizationManager.GetText("Find posting in 1C")
        .OnAction = "ProvodkaIntegrationModule.FindProvodkaForCurrentRecord"
        .FaceId = 1219
        .Tag = "CustomProvodka_VhIsh_2025"
    End With
    contextMenu.Reset
    Exit Sub
MenuError:
End Sub

Public Sub RemoveContextMenuButton()
    On Error GoTo RemoveMenuError
    Dim ctrl As CommandBarControl
    For Each ctrl In Application.CommandBars("Cell").Controls
        If InStr(ctrl.Tag, "Custom") > 0 And InStr(ctrl.Tag, "VhIsh") > 0 Then ctrl.Delete
    Next ctrl
    Exit Sub
RemoveMenuError:
End Sub

Public Sub DeactivateTableEvents()
    On Error Resume Next
    isSystemActive = False
    contextMenuInitialized = False
    If Not AppEventsHandler Is Nothing Then
        Call AppEventsHandler.DeactivateAppEvents
        Set AppEventsHandler = Nothing
    End If
    Set wsData = Nothing
    Set tblVhIsh = Nothing
    Application.StatusBar = LocalizationManager.GetText("Interactive forms system deactivated")
    On Error GoTo 0
End Sub

Public Function GetFieldNameByColumn(columnIndex As Long) As String
    On Error GoTo FieldNameError
    Select Case columnIndex
        Case 1: GetFieldNameByColumn = "txtNomerPP"
        Case 2: GetFieldNameByColumn = "cmbSlujba"
        Case 3: GetFieldNameByColumn = "cmbVidDocumenta"
        Case 4: GetFieldNameByColumn = "cmbVidDoc"
        Case 5: GetFieldNameByColumn = "txtNomerDoc"
        Case 6: GetFieldNameByColumn = "txtSummaDoc"
        Case 7: GetFieldNameByColumn = "txtVhFRP"
        Case 8: GetFieldNameByColumn = "txtDataVhFRP"
        Case 9: GetFieldNameByColumn = "cmbOtKogoPostupil"
        Case 10: GetFieldNameByColumn = "txtDataPeredachi"
        Case 11: GetFieldNameByColumn = "cmbIspolnitel"
        Case 12: GetFieldNameByColumn = "txtNomerIshVSlujbu"
        Case 13: GetFieldNameByColumn = "txtDataIshVSlujbu"
        Case 14: GetFieldNameByColumn = "txtNomerVozvrata"
        Case 15: GetFieldNameByColumn = "txtDataVozvrata"
        Case 16: GetFieldNameByColumn = "txtNomerIshKonvert"
        Case 17: GetFieldNameByColumn = "txtDataIshKonvert"
        Case 18: GetFieldNameByColumn = "txtOtmetkaIspolnenie"
        Case 19: GetFieldNameByColumn = "cmbStatusPodtverjdenie"
        Case 20: GetFieldNameByColumn = "txtNaryadInfo"
        Case Else: GetFieldNameByColumn = "txtNomerDoc"
    End Select
    Exit Function
FieldNameError:
    GetFieldNameByColumn = "txtNomerDoc"
End Function

Public Function GetFieldDisplayName(fieldName As String) As String
    On Error GoTo DisplayNameError
    Select Case fieldName
        Case "txtNomerPP": GetFieldDisplayName = LocalizationManager.GetText("Seq No")
        Case "cmbSlujba": GetFieldDisplayName = LocalizationManager.GetText("Service")
        Case "cmbVidDocumenta": GetFieldDisplayName = LocalizationManager.GetText("Document Group")
        Case "cmbVidDoc": GetFieldDisplayName = LocalizationManager.GetText("Document Type")
        Case "txtNomerDoc": GetFieldDisplayName = LocalizationManager.GetText("Document Number")
        Case "txtSummaDoc": GetFieldDisplayName = LocalizationManager.GetText("Document Amount")
        Case "txtVhFRP": GetFieldDisplayName = LocalizationManager.GetText("Inc.FRP/Out.FRP")
        Case "txtDataVhFRP": GetFieldDisplayName = LocalizationManager.GetText("Date Inc.FRP/Out.FRP")
        Case "cmbOtKogoPostupil": GetFieldDisplayName = LocalizationManager.GetText("Received From")
        Case "txtDataPeredachi": GetFieldDisplayName = LocalizationManager.GetText("Date Transferred to Executor")
        Case "cmbIspolnitel": GetFieldDisplayName = LocalizationManager.GetText("Executor")
        Case "txtNomerIshVSlujbu": GetFieldDisplayName = LocalizationManager.GetText("Out. No. to Service")
        Case "txtDataIshVSlujbu": GetFieldDisplayName = LocalizationManager.GetText("Out. Date to Service")
        Case "txtNomerVozvrata": GetFieldDisplayName = LocalizationManager.GetText("Return No.")
        Case "txtDataVozvrata": GetFieldDisplayName = LocalizationManager.GetText("Return Date")
        Case "txtNomerIshKonvert": GetFieldDisplayName = LocalizationManager.GetText("Out. Envelope No.")
        Case "txtDataIshKonvert": GetFieldDisplayName = LocalizationManager.GetText("Out. Envelope Date")
        Case "txtOtmetkaIspolnenie": GetFieldDisplayName = LocalizationManager.GetText("Execution Mark")
        Case "cmbStatusPodtverjdenie": GetFieldDisplayName = LocalizationManager.GetText("Confirmation Status")
        Case "txtNaryadInfo": GetFieldDisplayName = LocalizationManager.GetText("Order Info")
        Case Else: GetFieldDisplayName = LocalizationManager.GetText("Data Field")
    End Select
    Exit Function
DisplayNameError:
    GetFieldDisplayName = LocalizationManager.GetText("Data Field")
End Function

Public Sub OpenFormForCurrentSelection()
    Dim intersectRange As Range, RowNumber As Long, columnNumber As Long
    On Error GoTo OpenError
    If Not isSystemActive Then
        Application.StatusBar = LocalizationManager.GetText("Interactive forms system is not active!")
        Exit Sub
    End If
    If wsData Is Nothing Or tblVhIsh Is Nothing Then
        Application.StatusBar = LocalizationManager.GetText("Error: system objects not initialized!")
        Exit Sub
    End If
    If ActiveSheet.Name <> "IncOut" Then
        Application.StatusBar = LocalizationManager.GetText("Switch to sheet 'IncOut' to work with the table.")
        Exit Sub
    End If
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = LocalizationManager.GetText("Table is empty! Add data to work with the form.")
        Exit Sub
    End If
    If Selection.Cells.Count = 1 Then
        Set intersectRange = Intersect(Selection, tblVhIsh.DataBodyRange)
        If Not intersectRange Is Nothing Then
            RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
            columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
            Call OpenFormWithRecord(RowNumber, GetFieldNameByColumn(columnNumber), columnNumber)
        Else
            Application.StatusBar = LocalizationManager.GetText("Select a cell inside the data table.")
        End If
    Else
        Application.StatusBar = LocalizationManager.GetText("Select a single cell in the table for editing.")
    End If
    Exit Sub
OpenError:
    Application.StatusBar = LocalizationManager.GetText("Error opening form: ") & Err.description
End Sub

Private Sub OpenFormWithRecord(RowNumber As Long, fieldName As String, columnNumber As Long)
    On Error GoTo OpenError
    If Not IsFormOpen("UserFormVhIsh") Then
        Load UserFormVhIsh
        UserFormVhIsh.Show vbModeless
    End If
    
    Call UserFormVhIsh.LoadRecordToForm(RowNumber)
    RecordOperations.CurrentRecordRow = RowNumber
    RecordOperations.IsNewRecord = False
    RecordOperations.FormDataChanged = False
    
    Application.OnTime Now + TimeValue("00:00:01"), "'TableEventHandler.ActivateFormField """ & fieldName & """'"
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Loaded record No.") & RowNumber & LocalizationManager.GetText(" | Active field: ") & GetFieldDisplayName(fieldName)
    Call HighlightActiveField(fieldName, columnNumber)
    Exit Sub
OpenError:
End Sub

Public Function IsFormOpen(formName As String) As Boolean
    Dim frm As Object
    IsFormOpen = False
    On Error GoTo FormNotOpen
    Set frm = VBA.UserForms(formName)
    IsFormOpen = True
FormNotOpen:
    On Error GoTo 0
End Function

Public Sub ActivateFormField(fieldName As String)
    On Error Resume Next
    If IsFormOpen("UserFormVhIsh") Then
        CallByName UserFormVhIsh, fieldName, VbGet
        UserFormVhIsh.Controls(fieldName).SetFocus
    End If
End Sub

Private Sub HighlightActiveField(fieldName As String, columnIndex As Long)
    On Error Resume Next
    If tblVhIsh Is Nothing Or tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    tblVhIsh.ListColumns(columnIndex).DataBodyRange.Interior.Color = RGB(255, 255, 200)
    Application.OnTime Now + TimeValue("00:00:02"), "'TableEventHandler.RemoveColumnHighlight " & columnIndex & "'"
End Sub

Public Sub RemoveColumnHighlight(columnIndex As Long)
    On Error Resume Next
    tblVhIsh.ListColumns(columnIndex).DataBodyRange.Interior.ColorIndex = xlNone
End Sub

Public Sub DuplicateCurrentRecord()
    Dim intersectRange As Range, RowNumber As Long
    On Error GoTo DuplicateError
    If Not isSystemActive Or tblVhIsh Is Nothing Or tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    If Selection.Cells.Count = 1 Then
        Set intersectRange = Intersect(Selection, tblVhIsh.DataBodyRange)
        If Not intersectRange Is Nothing Then
            RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
            If MsgBox(LocalizationManager.GetText("Duplicate record No.") & RowNumber & "?" & vbCrLf & vbCrLf & _
                      LocalizationManager.GetText("All data will be copied except:") & vbCrLf & _
                      LocalizationManager.GetText("- Document Number") & vbCrLf & _
                      LocalizationManager.GetText("- Document Amount") & vbCrLf & _
                      LocalizationManager.GetText("- Order Info"), _
                      vbYesNo + vbQuestion, LocalizationManager.GetText("Duplication Confirmation")) = vbYes Then
                Call PerformDuplication(RowNumber)
            End If
        End If
    End If
    Exit Sub
DuplicateError:
End Sub

Private Sub PerformDuplication(sourceRowNumber As Long)
    Dim newRow As ListRow, i As Long
    On Error GoTo PerformDuplicationError
    Set newRow = tblVhIsh.ListRows.Add
    For i = 1 To tblVhIsh.ListColumns.Count
        Select Case i
            Case 1: tblVhIsh.DataBodyRange.Cells(newRow.Index, i).value = newRow.Index
            Case 5, 20: tblVhIsh.DataBodyRange.Cells(newRow.Index, i).value = ""
            Case 6: tblVhIsh.DataBodyRange.Cells(newRow.Index, i).value = 0
            Case Else: tblVhIsh.DataBodyRange.Cells(newRow.Index, i).value = tblVhIsh.DataBodyRange.Cells(sourceRowNumber, i).value
        End Select
    Next i
    MsgBox LocalizationManager.GetText("Record No.") & sourceRowNumber & LocalizationManager.GetText(" successfully duplicated!") & vbCrLf & _
           LocalizationManager.GetText("New record: No.") & newRow.Index & vbCrLf & vbCrLf & _
           LocalizationManager.GetText("Please note:") & vbCrLf & _
           LocalizationManager.GetText("- Document number not copied (requires filling)") & vbCrLf & _
           LocalizationManager.GetText("- Document amount reset to zero (requires filling)") & vbCrLf & _
           LocalizationManager.GetText("- Order info not copied"), _
           vbInformation, LocalizationManager.GetText("Duplication Completed")
    tblVhIsh.DataBodyRange.Cells(newRow.Index, 1).Select
    Application.StatusBar = LocalizationManager.GetText("Created duplicate of record No.") & newRow.Index & LocalizationManager.GetText(" based on record No.") & sourceRowNumber
    Exit Sub
PerformDuplicationError:
    MsgBox LocalizationManager.GetText("Error performing duplication: ") & Err.description, vbCritical, LocalizationManager.GetText("Critical Error")
End Sub

Public Sub ProcessCellClick(Target As Range)
    Dim intersectRange As Range, RowNumber As Long, columnNumber As Long
    On Error Resume Next
    If Not isSystemActive Or tblVhIsh Is Nothing Or tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    Set intersectRange = Intersect(Target, tblVhIsh.DataBodyRange)
    If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
        RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
        columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
        lastSelectedRow = RowNumber
        lastSelectedColumn = columnNumber
        Application.StatusBar = LocalizationManager.GetText("Row ") & RowNumber & LocalizationManager.GetText(", Col ") & columnNumber & _
                               LocalizationManager.GetText(" | Ctrl+Shift+M: menu | Ctrl+E: edit | Ctrl+D: duplicate | ") & _
                               GetFieldDisplayName(GetFieldNameByColumn(columnNumber))
    Else
        If lastSelectedRow <> 0 Then
            lastSelectedRow = 0
            lastSelectedColumn = 0
            Application.StatusBar = LocalizationManager.GetText("Interactive forms system is active (Excel 2021). Use hotkeys.")
        End If
    End If
End Sub

Public Sub OpenFormForSpecificCell(Target As Range)
    Dim intersectRange As Range
    On Error Resume Next
    If isSystemActive And Not tblVhIsh Is Nothing And Not tblVhIsh.DataBodyRange Is Nothing Then
        Set intersectRange = Intersect(Target, tblVhIsh.DataBodyRange)
        If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
            Call OpenFormWithRecord(intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1, _
                                    GetFieldNameByColumn(intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1), _
                                    intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1)
        End If
    End If
End Sub

