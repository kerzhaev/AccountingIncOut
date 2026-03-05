VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================
' WORKSHEET CODE FOR "IncOut" SHEET - click event handling
' Purpose: Handling cell click events for interactive form work
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 2.1.1
' Date: 05.03.2026
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' Variables for event tracking
Private lastClickTime As Double
Private lastClickCell As String
Private rightClickRow As Long
Private rightClickColumn As Long

' Configuration constants
Private Const DOUBLE_CLICK_THRESHOLD As Double = 0.5 ' Double click threshold (seconds)

' FIXED: Worksheet cell click event
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo SelectionError
    
    ' FIXED: Direct module procedure call without object check
    Call TableEventHandler.ProcessCellClick(Target)
    
    Exit Sub
    
SelectionError:
    ' Ignore selection event errors
    Debug.Print "Error in Worksheet_SelectionChange: " & Err.description
End Sub

' FIXED: Cell double click event
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo DoubleClickError
    
    ' Cancel standard Excel behavior (entering cell edit mode)
    Cancel = True
    
    ' FIXED: Direct module procedure call
    Call TableEventHandler.OpenFormForSpecificCell(Target)
    
    ' Update status bar
    Application.StatusBar = LocalizationManager.GetText("Double click processed - opening edit form")
    
    Exit Sub
    
DoubleClickError:
    Cancel = False ' Restore standard behavior on error
    Debug.Print "Error in Worksheet_BeforeDoubleClick: " & Err.description
End Sub

' FIXED: Right click event
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Dim intersectRange As Range
    Dim tblData As ListObject
    
    On Error GoTo RightClickError
    
    ' Check that sheet contains our table
    Set tblData = Me.ListObjects("TableIncOut")
    
    If Not tblData Is Nothing Then
        If Not tblData.DataBodyRange Is Nothing Then
            ' Check that right click occurred in table data area
            Set intersectRange = Intersect(Target, tblData.DataBodyRange)
            
            If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
                ' Right click in table area - show our context menu
                Cancel = True ' Cancel standard Excel context menu
                
                ' Save click coordinates for context menu usage
                rightClickRow = intersectRange.Row - tblData.DataBodyRange.Row + 1
                rightClickColumn = intersectRange.Column - tblData.DataBodyRange.Column + 1
                
                ' FIXED: Direct call to add context menu
                Call TableEventHandler.AddContextMenuButton
                
                ' Update status bar
                Application.StatusBar = LocalizationManager.GetText("Right click on row ") & rightClickRow & LocalizationManager.GetText(", col ") & rightClickColumn & LocalizationManager.GetText(" - context menu shown")
                
                ' Do not cancel - allow menu to show
                Cancel = False
            Else
                ' Right click outside table - use standard Excel menu
                Cancel = False
                rightClickRow = 0
                rightClickColumn = 0
            End If
        Else
            ' Table is empty - standard menu
            Cancel = False
            rightClickRow = 0
            rightClickColumn = 0
        End If
    Else
        ' Table not found - standard menu
        Cancel = False
        rightClickRow = 0
        rightClickColumn = 0
    End If
    
    Exit Sub
    
RightClickError:
    ' Restore standard Excel context menu on error
    Cancel = False
    rightClickRow = 0
    rightClickColumn = 0
    Debug.Print "Error in Worksheet_BeforeRightClick: " & Err.description
End Sub

' FIXED: Sheet activation event
Private Sub Worksheet_Activate()
    On Error GoTo ActivateError
    
    ' FIXED: Direct call to add context menu
    Call TableEventHandler.AddContextMenuButton
    
    ' Update status bar
    Application.StatusBar = LocalizationManager.GetText("IncOut sheet activated - interactive forms system ready to work")
    
    Exit Sub
    
ActivateError:
    Debug.Print "Error in Worksheet_Activate: " & Err.description
End Sub

' FIXED: Sheet deactivation event
Private Sub Worksheet_Deactivate()
    On Error GoTo DeactivateError
    
    ' FIXED: Direct call to remove context menu
    Call TableEventHandler.RemoveContextMenuButton
    
    ' Reset click variables
    rightClickRow = 0
    rightClickColumn = 0
    
    ' Update status bar
    Application.StatusBar = LocalizationManager.GetText("IncOut sheet deactivated")
    
    Exit Sub
    
DeactivateError:
    Debug.Print "Error in Worksheet_Deactivate: " & Err.description
End Sub

' Cell value change event
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim intersectRange As Range
    Dim tblData As ListObject
    
    On Error GoTo ChangeError
    
    ' Check if change occurred in our table
    Set tblData = Me.ListObjects("TableIncOut")
    
    If Not tblData Is Nothing Then
        If Not tblData.DataBodyRange Is Nothing Then
            Set intersectRange = Intersect(Target, tblData.DataBodyRange)
            
            If Not intersectRange Is Nothing Then
                ' Table data changed - can add additional logic here
                ' E.g., auto-updating related fields
                
                ' Update status bar
                Application.StatusBar = LocalizationManager.GetText("Table data changed - changes saved automatically")
            End If
        End If
    End If
    
    Exit Sub
    
ChangeError:
    Debug.Print "Error in Worksheet_Change: " & Err.description
End Sub

' Function to get info about last right click
Public Function GetLastRightClickInfo() As String
    If rightClickRow > 0 And rightClickColumn > 0 Then
        GetLastRightClickInfo = LocalizationManager.GetText("Row: ") & rightClickRow & LocalizationManager.GetText(", Col: ") & rightClickColumn
    Else
        GetLastRightClickInfo = LocalizationManager.GetText("No right click information")
    End If
End Function

' Function to check if last right click was in table
Public Function IsLastRightClickInTable() As Boolean
    IsLastRightClickInTable = (rightClickRow > 0 And rightClickColumn > 0)
End Function

' Procedure to reset click info
Public Sub ResetClickInfo()
    rightClickRow = 0
    rightClickColumn = 0
    lastClickTime = 0
    lastClickCell = ""
End Sub

' Procedure for programmatic table cell selection
Public Sub SelectTableCell(RowNumber As Long, columnNumber As Long)
    Dim tblData As ListObject
    Dim targetCell As Range
    
    On Error GoTo SelectError
    
    Set tblData = Me.ListObjects("TableIncOut")
    
    If Not tblData Is Nothing Then
        If Not tblData.DataBodyRange Is Nothing Then
            If RowNumber >= 1 And RowNumber <= tblData.ListRows.Count Then
                If columnNumber >= 1 And columnNumber <= tblData.ListColumns.Count Then
                    Set targetCell = tblData.DataBodyRange.Cells(RowNumber, columnNumber)
                    targetCell.Select
                    
                    Application.StatusBar = LocalizationManager.GetText("Cell selected: row ") & RowNumber & LocalizationManager.GetText(", col ") & columnNumber
                End If
            End If
        End If
    End If
    
    Exit Sub
    
SelectError:
    Debug.Print "Error in SelectTableCell: " & Err.description
End Sub

' Procedure to get coordinates of current table selection
Public Sub GetCurrentTableSelection(ByRef RowNumber As Long, ByRef columnNumber As Long)
    Dim intersectRange As Range
    Dim tblData As ListObject
    Dim currentCell As Range
    
    On Error GoTo GetSelectionError
    
    RowNumber = 0
    columnNumber = 0
    
    Set tblData = Me.ListObjects("TableIncOut")
    Set currentCell = ActiveCell
    
    If Not tblData Is Nothing And Not currentCell Is Nothing Then
        If Not tblData.DataBodyRange Is Nothing Then
            Set intersectRange = Intersect(currentCell, tblData.DataBodyRange)
            
            If Not intersectRange Is Nothing Then
                RowNumber = intersectRange.Row - tblData.DataBodyRange.Row + 1
                columnNumber = intersectRange.Column - tblData.DataBodyRange.Column + 1
            End If
        End If
    End If
    
    Exit Sub
    
GetSelectionError:
    RowNumber = 0
    columnNumber = 0
    Debug.Print "Error in GetCurrentTableSelection: " & Err.description
End Sub

' Table structure validation function
Public Function ValidateTableStructure() As Boolean
    Dim tblData As ListObject
    Dim expectedColumns As Long
    
    On Error GoTo ValidateError
    
    expectedColumns = 20 ' Expected number of columns including new order info field
    
    Set tblData = Me.ListObjects("TableIncOut")
    
    If tblData Is Nothing Then
        ValidateTableStructure = False
        Exit Function
    End If
    
    ' Check columns count
    If tblData.ListColumns.Count <> expectedColumns Then
        MsgBox LocalizationManager.GetText("Invalid table structure!") & vbCrLf & _
               LocalizationManager.GetText("Expected columns: ") & expectedColumns & vbCrLf & _
               LocalizationManager.GetText("Found columns: ") & tblData.ListColumns.Count, _
               vbExclamation, LocalizationManager.GetText("Table Structure Error")
        ValidateTableStructure = False
        Exit Function
    End If
    
    ValidateTableStructure = True
    Exit Function
    
ValidateError:
    ValidateTableStructure = False
    Debug.Print "Table structure validation error: " & Err.description
End Function

' Procedure for worksheet events debug info
Public Sub ShowWorksheetEventInfo()
    Dim info As String
    Dim tblData As ListObject
    
    On Error GoTo InfoError
    
    info = LocalizationManager.GetText("EVENTS INFO FOR 'IncOut' SHEET:") & vbCrLf & vbCrLf
    
    ' Table info
    Set tblData = Me.ListObjects("TableIncOut")
    If Not tblData Is Nothing Then
        info = info & LocalizationManager.GetText("Table found: ") & tblData.Name & vbCrLf
        info = info & LocalizationManager.GetText("Columns: ") & tblData.ListColumns.Count & vbCrLf
        info = info & LocalizationManager.GetText("Data rows: ") & IIf(tblData.DataBodyRange Is Nothing, 0, tblData.ListRows.Count) & vbCrLf
    Else
        info = info & LocalizationManager.GetText("Table NOT found!") & vbCrLf
    End If
    
    info = info & vbCrLf
    
    ' Last right click info
    info = info & LocalizationManager.GetText("Last right click: ") & GetLastRightClickInfo() & vbCrLf
    info = info & LocalizationManager.GetText("Click in table: ") & IIf(IsLastRightClickInTable(), LocalizationManager.GetText("YES"), LocalizationManager.GetText("NO")) & vbCrLf
    
    info = info & vbCrLf
    
    ' Current selection info
    Dim CurrentRow As Long, currentCol As Long
    Call GetCurrentTableSelection(CurrentRow, currentCol)
    info = info & LocalizationManager.GetText("Current table selection: row ") & CurrentRow & LocalizationManager.GetText(", col ") & currentCol & vbCrLf
    
    info = info & vbCrLf & LocalizationManager.GetText("Table validation status: ") & IIf(ValidateTableStructure(), LocalizationManager.GetText("VALID"), LocalizationManager.GetText("ERROR"))
    
    MsgBox info, vbInformation, LocalizationManager.GetText("Worksheet Events Diagnostics")
    
    Exit Sub
    
InfoError:
    MsgBox LocalizationManager.GetText("Error getting events info: ") & Err.description, vbExclamation, LocalizationManager.GetText("Error")
End Sub

