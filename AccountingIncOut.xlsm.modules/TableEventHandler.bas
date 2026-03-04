Attribute VB_Name = "TableEventHandler"
'==============================================
' TABLE EVENT HANDLER MODULE - TableEventHandler
' Purpose: Handling table clicks to open form with activation of specific field
' State: ALL REFERENCES TO GLOBAL VARIABLES FIXED
' Version: 1.6.0
' Date: 10.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' Global variables for table operations
Private wsData As Worksheet
Private tblVhIsh As ListObject
Private isSystemActive As Boolean

' Variables to track the last selection
Private lastSelectedRow As Long
Private lastSelectedColumn As Long

' Variables for context menu
Private currentRightClickRow As Long
Private contextMenuInitialized As Boolean

' Application event handler via class
Private AppEventsHandler As AppEventHandler

' POPUP WINDOWS DISABLED: Event handler initialization
Public Sub InitializeTableEvents()
    On Error GoTo InitError
    
    ' Deactivate system before initialization
    isSystemActive = False
    contextMenuInitialized = False
    
    ' Check if sheet exists
    Dim wsExists As Boolean
    wsExists = False
    
    Dim Ws As Worksheet
    For Each Ws In ThisWorkbook.Worksheets
        If Ws.Name = "IncOut" Then
            wsExists = True
            Set wsData = Ws
            Exit For
        End If
    Next Ws
    
    If Not wsExists Then
        ' KEPT: Critical errors are shown
        MsgBox "Sheet 'IncOut' not found in the workbook!", vbCritical, "Initialization Error"
        Exit Sub
    End If
    
    ' Check if table exists
    Dim tblExists As Boolean
    tblExists = False
    
    Dim tbl As ListObject
    For Each tbl In wsData.ListObjects
        If tbl.Name = "TableIncOut" Then
            tblExists = True
            Set tblVhIsh = tbl
            Exit For
        End If
    Next tbl
    
    If Not tblExists Then
        ' KEPT: Critical errors are shown
        MsgBox "Table 'TableIncOut' not found on sheet 'IncOut'!" & vbCrLf & _
               "Make sure the table exists and has the correct name.", vbCritical, "Initialization Error"
        Exit Sub
    End If
    
    ' Check table content (Debug only, no MsgBox)
    If tblVhIsh.DataBodyRange Is Nothing Then
        Debug.Print "WARNING: Table 'TableIncOut' is empty!"
        Application.StatusBar = "Warning: table is empty. Add data to work."
    End If
    
    ' Initialize tracking variables
    lastSelectedRow = 0
    lastSelectedColumn = 0
    currentRightClickRow = 0
    
    ' Initialize Application event handler via class
    Set AppEventsHandler = New AppEventHandler
    Call AppEventsHandler.InitializeAppEvents
    
    ' Activate system
    isSystemActive = True
    contextMenuInitialized = True
    
    ' POPUP WINDOWS DISABLED: Status bar and Debug only
    Application.StatusBar = "Interactive forms system is active (Excel 2021). Use Ctrl+Shift+M for menu."
    Debug.Print "Interactive forms system activated for Excel 2021!"
    Debug.Print "Found table: " & tblVhIsh.Name
    Debug.Print "Data rows: " & IIf(tblVhIsh.DataBodyRange Is Nothing, 0, tblVhIsh.ListRows.Count)
    Debug.Print "WORK METHODS: Double click, Ctrl+E, Ctrl+Shift+M, Ctrl+D"
    
    Exit Sub
    
InitError:
    isSystemActive = False
    contextMenuInitialized = False
    ' KEPT: Critical errors are shown
    MsgBox "Critical error initializing table events system:" & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.description, vbCritical, "Critical Error"
End Sub

' POPUP WINDOWS DISABLED: Show action menu for Excel 2021
Public Sub ShowActionMenuExcel2021()
    Dim userChoice As String
    Dim currentCell As Range
    Dim intersectRange As Range
    Dim RowNumber As Long
    
    On Error GoTo MenuError
    
    ' Check if we are in the correct sheet
    If ActiveSheet.Name <> "IncOut" Then
        Application.StatusBar = "Switch to sheet 'IncOut' to use the action menu."
        Exit Sub
    End If
    
    Set currentCell = Selection
    
    If Not tblVhIsh Is Nothing Then
        If Not tblVhIsh.DataBodyRange Is Nothing Then
            Set intersectRange = Intersect(currentCell, tblVhIsh.DataBodyRange)
            
            If Not intersectRange Is Nothing And currentCell.Cells.Count = 1 Then
                RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
                
                ' KEPT: Menu via InputBox (core functionality)
                userChoice = InputBox( _
                    "ACTION MENU FOR RECORD No." & RowNumber & ":" & vbCrLf & vbCrLf & _
                    "Available commands:" & vbCrLf & _
                    "1 - Edit in form" & vbCrLf & _
                    "2 - Duplicate record" & vbCrLf & _
                    "3 - Show record information" & vbCrLf & _
                    "0 - Cancel" & vbCrLf & vbCrLf & _
                    "Enter command number:", _
                    "Excel 2021 Action Menu", "1")
                
                ' Process user choice
                Select Case userChoice
                    Case "1"
                        Call OpenFormForCurrentSelection
                    Case "2"
                        Call DuplicateCurrentRecord
                    Case "3"
                        Call ShowRecordInfo(RowNumber)
                    Case "0", ""
                        ' Cancel - do nothing
                    Case Else
                        Application.StatusBar = "Invalid choice. Use numbers 0-3."
                End Select
            Else
                Application.StatusBar = "Select a cell inside the data table to use the menu."
            End If
        Else
            Application.StatusBar = "Table is empty!"
        End If
    Else
        Application.StatusBar = "Table not found!"
    End If
    
    Exit Sub
    
MenuError:
    Application.StatusBar = "Error displaying menu: " & Err.description
End Sub

' POPUP WINDOWS DISABLED: Show record information
Private Sub ShowRecordInfo(RowNumber As Long)
    Dim info As String
    Dim serviceText As String
    Dim docTypeText As String
    Dim docNumberText As String
    Dim amountText As String
    
    On Error GoTo InfoError
    
    If tblVhIsh Is Nothing Or tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    If RowNumber < 1 Or RowNumber > tblVhIsh.ListRows.Count Then Exit Sub
    
    ' Get basic record info
    serviceText = CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 2).value)     ' Service
    docTypeText = CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 4).value)     ' Document Type
    docNumberText = CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 5).value)   ' Document Number
    amountText = CStr(tblVhIsh.DataBodyRange.Cells(RowNumber, 6).value)      ' Amount
    
    info = "RECORD No." & RowNumber & " INFO:" & vbCrLf & vbCrLf
    info = info & "Service: " & serviceText & vbCrLf
    info = info & "Document Type: " & docTypeText & vbCrLf
    info = info & "Document Number: " & docNumberText & vbCrLf
    info = info & "Amount: " & amountText & " rub." & vbCrLf & vbCrLf
    info = info & "Press OK to continue."
    
    ' KEPT: Information windows (core functionality)
    MsgBox info, vbInformation, "Record Information"
    
    Exit Sub
    
InfoError:
    Application.StatusBar = "Error getting record information: " & Err.description
End Sub

' Add context menu
Public Sub AddContextMenuButton()
    On Error GoTo MenuError
    
    ' Try standard approach for compatibility
    Call RemoveContextMenuButton
    
    Dim contextMenu As CommandBar
    Set contextMenu = Application.CommandBars("Cell")
    
    ' Edit button
    Dim editButton As CommandBarButton
    Set editButton = contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
    
    With editButton
        .Caption = "Edit in form"
        .OnAction = "TableEventHandler.OpenFormForCurrentSelection"
        .FaceId = 162
        .BeginGroup = True
        .Tag = "CustomEdit_VhIsh_2021"
        .Visible = True
        .Enabled = True
    End With
    
    ' Duplicate button
    Dim duplicateButton As CommandBarButton
    Set duplicateButton = contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
    
    With duplicateButton
        .Caption = "Duplicate record"
        .OnAction = "TableEventHandler.DuplicateCurrentRecord"
        .FaceId = 19
        .BeginGroup = False
        .Tag = "CustomDuplicate_VhIsh_2021"
        .Visible = True
        .Enabled = True
    End With
    
    ' Alternative menu for Excel 2021
    Dim altMenuButton As CommandBarButton
    Set altMenuButton = contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
    
    With altMenuButton
        .Caption = "Action menu (Ctrl+Shift+M)"
        .OnAction = "TableEventHandler.ShowActionMenuExcel2021"
        .FaceId = 923
        .BeginGroup = False
        .Tag = "CustomAltMenu_VhIsh_2021"
        .Visible = True
        .Enabled = True
    End With
    
    ' Find posting in 1C button
    Dim provodkaButton As CommandBarButton
    Set provodkaButton = contextMenu.Controls.Add(Type:=msoControlButton, temporary:=True)

    With provodkaButton
        .Caption = "Find posting in 1C"
        .OnAction = "ProvodkaIntegrationModule.FindProvodkaForCurrentRecord"
        .FaceId = 1219
        .BeginGroup = False
        .Tag = "CustomProvodka_VhIsh_2025"
        .Visible = True
        .Enabled = True
    End With
    
    ' Force update for Excel 2021
    contextMenu.Reset
    
    Exit Sub
    
MenuError:
    Debug.Print "Error creating context menu: " & Err.description
End Sub

' Remove context menu
Public Sub RemoveContextMenuButton()
    On Error GoTo RemoveMenuError
    
    Dim contextMenu As CommandBar
    Dim ctrl As CommandBarControl
    
    Set contextMenu = Application.CommandBars("Cell")
    
    ' Remove all our buttons by tags
    For Each ctrl In contextMenu.Controls
        If InStr(ctrl.Tag, "Custom") > 0 And InStr(ctrl.Tag, "VhIsh") > 0 Then
            ctrl.Delete
        End If
    Next ctrl
    
    ' Additional cleanup by captions
    For Each ctrl In contextMenu.Controls
        If InStr(ctrl.Caption, "Edit in form") > 0 Or _
           InStr(ctrl.Caption, "Duplicate record") > 0 Or _
           InStr(ctrl.Caption, "Action menu") > 0 Or _
           InStr(ctrl.Caption, "Find posting") > 0 Then
            ctrl.Delete
        End If
    Next ctrl
    
    Exit Sub
    
RemoveMenuError:
    Debug.Print "Error removing context menu: " & Err.description
End Sub

' Deactivate event handler
Public Sub DeactivateTableEvents()
    On Error Resume Next
    
    isSystemActive = False
    contextMenuInitialized = False
    
    ' Deactivate Application event handler via class
    If Not AppEventsHandler Is Nothing Then
        Call AppEventsHandler.DeactivateAppEvents
        Set AppEventsHandler = Nothing
    End If
    
    ' Safely clear objects
    Set wsData = Nothing
    Set tblVhIsh = Nothing
    
    lastSelectedRow = 0
    lastSelectedColumn = 0
    currentRightClickRow = 0
    
    Application.StatusBar = "Interactive forms system deactivated"
    
    On Error GoTo 0
End Sub

' Function to get field name by column number
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

' Function to get field display name
Public Function GetFieldDisplayName(fieldName As String) As String
    On Error GoTo DisplayNameError
    
    Select Case fieldName
        Case "txtNomerPP": GetFieldDisplayName = "Seq No"
        Case "cmbSlujba": GetFieldDisplayName = "Service"
        Case "cmbVidDocumenta": GetFieldDisplayName = "Document Group"
        Case "cmbVidDoc": GetFieldDisplayName = "Document Type"
        Case "txtNomerDoc": GetFieldDisplayName = "Document Number"
        Case "txtSummaDoc": GetFieldDisplayName = "Document Amount"
        Case "txtVhFRP": GetFieldDisplayName = "Inc.FRP/Out.FRP"
        Case "txtDataVhFRP": GetFieldDisplayName = "Date Inc.FRP/Out.FRP"
        Case "cmbOtKogoPostupil": GetFieldDisplayName = "Received From"
        Case "txtDataPeredachi": GetFieldDisplayName = "Date Transferred to Executor"
        Case "cmbIspolnitel": GetFieldDisplayName = "Executor"
        Case "txtNomerIshVSlujbu": GetFieldDisplayName = "Out. No. to Service"
        Case "txtDataIshVSlujbu": GetFieldDisplayName = "Out. Date to Service"
        Case "txtNomerVozvrata": GetFieldDisplayName = "Return No."
        Case "txtDataVozvrata": GetFieldDisplayName = "Return Date"
        Case "txtNomerIshKonvert": GetFieldDisplayName = "Out. Envelope No."
        Case "txtDataIshKonvert": GetFieldDisplayName = "Out. Envelope Date"
        Case "txtOtmetkaIspolnenie": GetFieldDisplayName = "Execution Mark"
        Case "cmbStatusPodtverjdenie": GetFieldDisplayName = "Confirmation Status"
        Case "txtNaryadInfo": GetFieldDisplayName = "Order Info"
        Case Else: GetFieldDisplayName = "Data Field"
    End Select
    
    Exit Function
    
DisplayNameError:
    GetFieldDisplayName = "Data Field"
End Function

' POPUP WINDOWS DISABLED: Open form by command
Public Sub OpenFormForCurrentSelection()
    Dim currentCell As Range
    Dim intersectRange As Range
    Dim RowNumber As Long
    Dim columnNumber As Long
    Dim fieldToActivate As String
    
    On Error GoTo OpenError
    
    ' Check system state (status bar only)
    If Not isSystemActive Then
        Application.StatusBar = "Interactive forms system is not active!"
        Exit Sub
    End If
    
    If wsData Is Nothing Or tblVhIsh Is Nothing Then
        Application.StatusBar = "Error: system objects not initialized!"
        Exit Sub
    End If
    
    ' Check active sheet
    If ActiveSheet.Name <> "IncOut" Then
        Application.StatusBar = "Switch to sheet 'IncOut' to work with the table."
        Exit Sub
    End If
    
    ' Check: Table is not empty
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = "Table is empty! Add data to work with the form."
        Exit Sub
    End If
    
    Set currentCell = Selection
    
    ' Check that one cell is selected in the table area
    If currentCell.Cells.Count = 1 Then
        Set intersectRange = Intersect(currentCell, tblVhIsh.DataBodyRange)
        
        If Not intersectRange Is Nothing Then
            ' Get coordinates in table
            RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
            columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
            fieldToActivate = GetFieldNameByColumn(columnNumber)
            
            ' Open form
            Call OpenFormWithRecord(RowNumber, fieldToActivate, columnNumber)
        Else
            Application.StatusBar = "Select a cell inside the data table."
        End If
    Else
        Application.StatusBar = "Select a single cell in the table for editing."
    End If
    
    Exit Sub
    
OpenError:
    Application.StatusBar = "Error opening form: " & Err.description
End Sub

' Procedure to open form with activation of specific field
Private Sub OpenFormWithRecord(RowNumber As Long, fieldName As String, columnNumber As Long)
    On Error GoTo OpenError
    
    ' Check if form is already open
    Dim formAlreadyOpen As Boolean
    formAlreadyOpen = IsFormOpen("UserFormVhIsh")
    
    If Not formAlreadyOpen Then
        ' Load form
        Load UserFormVhIsh
        ' Show form
        UserFormVhIsh.Show vbModeless ' Show modeless for working with table
    End If
    
    ' Load specific record
    Call UserFormVhIsh.LoadRecordToForm(RowNumber)
    
    ' FIXED: Correct calls to global variables of DataManager
    DataManager.CurrentRecordRow = RowNumber
    DataManager.IsNewRecord = False
    DataManager.FormDataChanged = False
    
    ' Activate specific field with a slight delay
    Application.OnTime Now + TimeValue("00:00:01"), "'TableEventHandler.ActivateFormField """ & fieldName & """'"
    
    ' Update form status bar
    UserFormVhIsh.lblStatusBar.Caption = "Loaded record No." & RowNumber & " | Active field: " & GetFieldDisplayName(fieldName)
    
    ' Highlight active field
    Call HighlightActiveField(fieldName, columnNumber)
    
    Exit Sub
    
OpenError:
    Application.StatusBar = "Error opening form: " & Err.description
End Sub

' Function to check if form is open
Public Function IsFormOpen(formName As String) As Boolean
    Dim frm As Object
    IsFormOpen = False
    
    On Error GoTo FormNotOpen
    Set frm = VBA.UserForms(formName)
    IsFormOpen = True
    
FormNotOpen:
    On Error GoTo 0
End Function

' Procedure to activate form field (called via OnTime)
Public Sub ActivateFormField(fieldName As String)
    On Error GoTo ActivateError
    
    ' Check that form is still open
    If IsFormOpen("UserFormVhIsh") Then
        ' Set focus to specific field
        Select Case fieldName
            Case "txtNomerPP": UserFormVhIsh.txtNomerPP.SetFocus
            Case "cmbSlujba": UserFormVhIsh.cmbSlujba.SetFocus
            Case "cmbVidDocumenta": UserFormVhIsh.cmbVidDocumenta.SetFocus
            Case "cmbVidDoc": UserFormVhIsh.cmbVidDoc.SetFocus
            Case "txtNomerDoc": UserFormVhIsh.txtNomerDoc.SetFocus
            Case "txtSummaDoc": UserFormVhIsh.txtSummaDoc.SetFocus
            Case "txtVhFRP": UserFormVhIsh.txtVhFRP.SetFocus
            Case "txtDataVhFRP": UserFormVhIsh.txtDataVhFRP.SetFocus
            Case "cmbOtKogoPostupil": UserFormVhIsh.cmbOtKogoPostupil.SetFocus
            Case "txtDataPeredachi": UserFormVhIsh.txtDataPeredachi.SetFocus
            Case "cmbIspolnitel": UserFormVhIsh.cmbIspolnitel.SetFocus
            Case "txtNomerIshVSlujbu": UserFormVhIsh.txtNomerIshVSlujbu.SetFocus
            Case "txtDataIshVSlujbu": UserFormVhIsh.txtDataIshVSlujbu.SetFocus
            Case "txtNomerVozvrata": UserFormVhIsh.txtNomerVozvrata.SetFocus
            Case "txtDataVozvrata": UserFormVhIsh.txtDataVozvrata.SetFocus
            Case "txtNomerIshKonvert": UserFormVhIsh.txtNomerIshKonvert.SetFocus
            Case "txtDataIshKonvert": UserFormVhIsh.txtDataIshKonvert.SetFocus
            Case "txtOtmetkaIspolnenie": UserFormVhIsh.txtOtmetkaIspolnenie.SetFocus
            Case "cmbStatusPodtverjdenie": UserFormVhIsh.cmbStatusPodtverjdenie.SetFocus
            Case "txtNaryadInfo": UserFormVhIsh.txtNaryadInfo.SetFocus
            Case Else: UserFormVhIsh.txtNomerDoc.SetFocus ' Default
        End Select
    End If
    
    Exit Sub
    
ActivateError:
    ' Ignore field activation errors
End Sub

' Procedure to highlight active field
Private Sub HighlightActiveField(fieldName As String, columnIndex As Long)
    On Error GoTo HighlightError
    
    ' Check existence of table and column
    If tblVhIsh Is Nothing Then Exit Sub
    If columnIndex < 1 Or columnIndex > tblVhIsh.ListColumns.Count Then Exit Sub
    If tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    
    ' Temporarily highlight corresponding column in table
    Dim columnRange As Range
    Set columnRange = tblVhIsh.ListColumns(columnIndex).DataBodyRange
    
    ' Apply temporary highlight
    columnRange.Interior.Color = RGB(255, 255, 200) ' Light yellow
    
    ' Remove highlight after 2 seconds
    Application.OnTime Now + TimeValue("00:00:02"), "'TableEventHandler.RemoveColumnHighlight " & columnIndex & "'"
    
    Exit Sub
    
HighlightError:
    ' Ignore highlight errors
End Sub

' Procedure to remove column highlight
Public Sub RemoveColumnHighlight(columnIndex As Long)
    On Error GoTo RemoveError
    
    ' Check existence of table and column
    If tblVhIsh Is Nothing Then Exit Sub
    If columnIndex < 1 Or columnIndex > tblVhIsh.ListColumns.Count Then Exit Sub
    If tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    
    Dim columnRange As Range
    Set columnRange = tblVhIsh.ListColumns(columnIndex).DataBodyRange
    columnRange.Interior.ColorIndex = xlNone ' Remove highlight
    
    Exit Sub
    
RemoveError:
    ' Ignore remove highlight errors
End Sub

' CONFIRMATION KEPT: Duplicate current record
Public Sub DuplicateCurrentRecord()
    Dim currentCell As Range
    Dim intersectRange As Range
    Dim RowNumber As Long
    
    On Error GoTo DuplicateError
    
    ' Check system state (status bar)
    If Not isSystemActive Then
        Application.StatusBar = "Interactive forms system is not active!"
        Exit Sub
    End If
    
    If wsData Is Nothing Or tblVhIsh Is Nothing Then
        Application.StatusBar = "Error: system objects not initialized!"
        Exit Sub
    End If
    
    ' Check active sheet
    If ActiveSheet.Name <> "IncOut" Then
        Application.StatusBar = "Switch to sheet 'IncOut' to work with the table."
        Exit Sub
    End If
    
    ' Check that table is not empty
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = "Table is empty! No records to duplicate."
        Exit Sub
    End If
    
    Set currentCell = Selection
    
    ' Check that one cell is selected in the table area
    If currentCell.Cells.Count = 1 Then
        Set intersectRange = Intersect(currentCell, tblVhIsh.DataBodyRange)
        
        If Not intersectRange Is Nothing Then
            ' Get row number in table
            RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
            
            ' KEPT: Duplication confirmation (core functionality)
            Dim response As VbMsgBoxResult
            response = MsgBox("Duplicate record No." & RowNumber & "?" & vbCrLf & vbCrLf & _
                             "All data will be copied except:" & vbCrLf & _
                             "- Document Number" & vbCrLf & _
                             "- Document Amount" & vbCrLf & _
                             "- Order Info", _
                             vbYesNo + vbQuestion, "Duplication Confirmation")
            
            If response = vbYes Then
                Call PerformDuplication(RowNumber)
            End If
        Else
            Application.StatusBar = "Select a cell inside the data table to duplicate."
        End If
    Else
        Application.StatusBar = "Select a single cell in the table to duplicate the record."
    End If
    
    Exit Sub
    
DuplicateError:
    Application.StatusBar = "Error duplicating record: " & Err.description
End Sub

' NOTIFICATION KEPT: Perform record duplication
Private Sub PerformDuplication(sourceRowNumber As Long)
    Dim newRow As ListRow
    Dim sourceRowIndex As Long
    Dim newRowIndex As Long
    Dim i As Long
    
    On Error GoTo PerformDuplicationError
    
    ' Add new row to table
    Set newRow = tblVhIsh.ListRows.Add
    newRowIndex = newRow.Index
    sourceRowIndex = sourceRowNumber
    
    ' Copy data from source row, excluding specific fields
    For i = 1 To tblVhIsh.ListColumns.Count
        Select Case i
            Case 1  ' Seq No - set new number
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = newRowIndex
                
            Case 5  ' txtNomerDoc - DO NOT copy (document number)
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = ""
                
            Case 6  ' txtSummaDoc - DO NOT copy (document amount)
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = 0
                
            Case 20 ' txtNaryadInfo - DO NOT copy (order info)
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = ""
                
            Case Else ' All other fields - copy
                tblVhIsh.DataBodyRange.Cells(newRowIndex, i).value = _
                    tblVhIsh.DataBodyRange.Cells(sourceRowIndex, i).value
        End Select
    Next i
    
    ' KEPT: Successful duplication notification (core functionality)
    MsgBox "Record No." & sourceRowNumber & " successfully duplicated!" & vbCrLf & _
           "New record: No." & newRowIndex & vbCrLf & vbCrLf & _
           "Please note:" & vbCrLf & _
           "- Document number not copied (requires filling)" & vbCrLf & _
           "- Document amount reset to zero (requires filling)" & vbCrLf & _
           "- Order info not copied", _
           vbInformation, "Duplication Completed"
    
    ' Go to new record
    Dim newCellRange As Range
    Set newCellRange = tblVhIsh.DataBodyRange.Cells(newRowIndex, 1)
    newCellRange.Select
    
    ' Update status bar
    Application.StatusBar = "Created duplicate of record No." & newRowIndex & " based on record No." & sourceRowNumber
    
    Exit Sub
    
PerformDuplicationError:
    ' KEPT: Critical errors are shown
    MsgBox "Error performing duplication: " & Err.description, vbCritical, "Critical Error"
    
    ' Try to delete created row on error
    On Error Resume Next
    If Not newRow Is Nothing Then
        newRow.Delete
    End If
    On Error GoTo 0
End Sub

' Handle cell click from worksheet event
Public Sub ProcessCellClick(Target As Range)
    Dim intersectRange As Range
    Dim RowNumber As Long
    Dim columnNumber As Long
    Dim fieldToActivate As String
    
    On Error GoTo ProcessError
    
    ' Check system activity
    If Not isSystemActive Then Exit Sub
    
    ' Check object existence
    If wsData Is Nothing Or tblVhIsh Is Nothing Then Exit Sub
    
    ' Check that table is not empty
    If tblVhIsh.DataBodyRange Is Nothing Then Exit Sub
    
    ' Check that click is in table data area
    Set intersectRange = Intersect(Target, tblVhIsh.DataBodyRange)
    
    If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
        ' Get coordinates in table
        RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
        columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
        fieldToActivate = GetFieldNameByColumn(columnNumber)
        
        ' Update selection tracking
        lastSelectedRow = RowNumber
        lastSelectedColumn = columnNumber
        
        ' Update status bar with hotkeys info
        Application.StatusBar = "Row " & RowNumber & ", Col " & columnNumber & _
                               " | Ctrl+Shift+M: menu | Ctrl+E: edit | Ctrl+D: duplicate | " & _
                               GetFieldDisplayName(fieldToActivate)
    Else
        ' Reset if we exited table
        If lastSelectedRow <> 0 Then
            lastSelectedRow = 0
            lastSelectedColumn = 0
            Application.StatusBar = "Interactive forms system is active (Excel 2021). Use hotkeys."
        End If
    End If
    
    Exit Sub
    
ProcessError:
    ' Log error but do not interrupt
    Debug.Print "Error processing cell click: " & Err.description
End Sub

' Open form for specific cell (from double click)
Public Sub OpenFormForSpecificCell(Target As Range)
    Dim intersectRange As Range
    Dim RowNumber As Long
    Dim columnNumber As Long
    Dim fieldToActivate As String
    
    On Error GoTo OpenSpecificError
    
    ' Check system activity (status bar)
    If Not isSystemActive Then
        Application.StatusBar = "Interactive forms system is not active!"
        Exit Sub
    End If
    
    ' Check object existence
    If wsData Is Nothing Or tblVhIsh Is Nothing Then
        Application.StatusBar = "Error: system objects not initialized!"
        Exit Sub
    End If
    
    ' Check that table is not empty
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = "Table is empty! Add data to work with the form."
        Exit Sub
    End If
    
    ' Check that click is in table data area
    Set intersectRange = Intersect(Target, tblVhIsh.DataBodyRange)
    
    If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
        ' Get coordinates in table
        RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
        columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
        fieldToActivate = GetFieldNameByColumn(columnNumber)
        
        ' Open form
        Call OpenFormWithRecord(RowNumber, fieldToActivate, columnNumber)
    Else
        Application.StatusBar = "Select a cell inside the data table."
    End If
    
    Exit Sub
    
OpenSpecificError:
    Application.StatusBar = "Error opening form: " & Err.description
End Sub

