VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================
' WORKSHEET EVENT HANDLER CLASS - Class Module WorksheetEventHandler
' Purpose: Class for handling worksheet events with a table
' Type: Class Module
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.3.1
' Date: 05.03.2026
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' Declaration of variable with events (only in class!)
Private WithEvents wsData As Worksheet
Attribute wsData.VB_VarHelpID = -1
Private tblVhIsh As ListObject

' Property to set worksheet
Public Property Set Worksheet(Ws As Worksheet)
    Set wsData = Ws
    On Error Resume Next
    Set tblVhIsh = Ws.ListObjects("TableIncOut")
    On Error GoTo 0
End Property

' Selection change event handler
Private Sub wsData_SelectionChange(ByVal Target As Range)
    Dim intersectRange As Range
    Dim RowNumber As Long
    Dim columnNumber As Long
    Dim fieldToActivate As String
    
    On Error GoTo EventError
    
    ' Check if tblVhIsh is initialized
    If tblVhIsh Is Nothing Then Exit Sub
    
    ' Check if DataBodyRange exists
    If tblVhIsh.DataBodyRange Is Nothing Then
        Application.StatusBar = LocalizationManager.GetText("Table is empty. Add data to work with the system.")
        Exit Sub
    End If
    
    ' Check if click occurred in table data area (excluding headers)
    Set intersectRange = Intersect(Target, tblVhIsh.DataBodyRange)
    
    If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
        ' Get row number in table
        RowNumber = intersectRange.Row - tblVhIsh.DataBodyRange.Row + 1
        
        ' Get column number in table
        columnNumber = intersectRange.Column - tblVhIsh.DataBodyRange.Column + 1
        
        ' FIXED: Use function from TableEventHandler
        fieldToActivate = TableEventHandler.GetFieldNameByColumn(columnNumber)
        
        ' Show info in status bar
        Application.StatusBar = LocalizationManager.GetText("Row ") & RowNumber & LocalizationManager.GetText(", Col ") & columnNumber & _
                               LocalizationManager.GetText(" | Ctrl+E to edit | ") & _
                               TableEventHandler.GetFieldDisplayName(fieldToActivate)
    Else
        ' Reset if clicked outside table
        Application.StatusBar = LocalizationManager.GetText("Interactive forms system is active. Select a table cell.")
    End If
    
    Exit Sub
    
EventError:
    ' Ignore event handling errors
    Debug.Print "Error in WorksheetEventHandler.wsData_SelectionChange: " & Err.description
End Sub

' FIXED: Public method to open form via TableEventHandler
Public Sub OpenFormForSelection()
    On Error GoTo OpenError
    
    If wsData Is Nothing Or tblVhIsh Is Nothing Then Exit Sub
    
    ' Check if DataBodyRange exists
    If tblVhIsh.DataBodyRange Is Nothing Then
        MsgBox LocalizationManager.GetText("Table is empty! Add data to work with the form."), vbExclamation, LocalizationManager.GetText("Empty Table")
        Exit Sub
    End If
    
    ' FIXED: Use method from TableEventHandler instead of duplicating
    Call TableEventHandler.OpenFormForCurrentSelection
    
    Exit Sub
    
OpenError:
    MsgBox LocalizationManager.GetText("Error opening form: ") & Err.description, vbExclamation, LocalizationManager.GetText("Error")
End Sub

' Public method to get class info
Public Function GetClassInfo() As String
    GetClassInfo = "WorksheetEventHandler | " & _
                  LocalizationManager.GetText("Sheet: ") & IIf(wsData Is Nothing, LocalizationManager.GetText("NOT SET"), wsData.Name) & " | " & _
                  LocalizationManager.GetText("Table: ") & IIf(tblVhIsh Is Nothing, LocalizationManager.GetText("NOT FOUND"), tblVhIsh.Name) & " | " & _
                  LocalizationManager.GetText("Rows: ") & IIf(tblVhIsh Is Nothing Or tblVhIsh.DataBodyRange Is Nothing, 0, tblVhIsh.ListRows.Count)
End Function

' Method for safe object cleanup
Public Sub Cleanup()
    On Error Resume Next
    Set wsData = Nothing
    Set tblVhIsh = Nothing
    On Error GoTo 0
End Sub

' Public method to programmatically trigger SelectionChange event
Public Sub TriggerSelectionChange(Target As Range)
    Call wsData_SelectionChange(Target)
End Sub

' Method to check class activity
Public Function IsActive() As Boolean
    IsActive = (Not wsData Is Nothing) And (Not tblVhIsh Is Nothing)
End Function

' Method to get current worksheet
Public Property Get CurrentWorksheet() As Worksheet
    Set CurrentWorksheet = wsData
End Property

' Method to get current table
Public Property Get CurrentTable() As ListObject
    Set CurrentTable = tblVhIsh
End Property

