VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================
' WORKBOOK MODULE - ThisWorkbook
' Purpose: Handling workbook events
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 2.0.1
'==============================================

Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    
    ' Automatically activate system on workbook open
    ' (works only if modules are already loaded)
    Call AutoStart.InitializeSystem
    
    ' NEW: Stage 1 infrastructure initialization
    Call SystemSettings.InitializeSettings
    Call SystemLogger.InitializeLog
    
    ' Log system start
    Call SystemLogger.LogOperation("SystemStart", LocalizationManager.GetText("IncOut system started"), "SUCCESS", 0)
    
    ' Open form modelessly
    Load UserFormVhIsh
    UserFormVhIsh.Show vbModeless
    
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    On Error Resume Next
    
    ' Export modules to folder on save
    ' Uncomment the following line if you want automatic export:
    ' Call ExportModules(ThisWorkbook.Name & ".modules")
    
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' Deactivate system on workbook close
    Call AutoStart.ShutdownSystem
    
    ' MODULE REMOVAL DISABLED - do it manually if needed
    ' Call RemoveModules(ShowMsgBox:=True)
    
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    
    If Sh.Name = "IncOut" Then
        Application.StatusBar = LocalizationManager.GetText("Table sheet is active. Select a cell and press Ctrl+E to edit.")
    End If
    
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    
    If Sh.Name = "IncOut" Then
        Call UpdateStatusForSelection(Target)
    End If
    
    On Error GoTo 0
End Sub

Private Sub UpdateStatusForSelection(Target As Range)
    On Error Resume Next
    
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim intersectRange As Range
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    Set intersectRange = Intersect(Target, tblData.DataBodyRange)
    
    If Not intersectRange Is Nothing And Target.Cells.Count = 1 Then
        Dim RowNumber As Long
        Dim columnNumber As Long
        RowNumber = intersectRange.Row - tblData.DataBodyRange.Row + 1
        columnNumber = intersectRange.Column - tblData.DataBodyRange.Column + 1
        
        Dim fieldName As String
        Dim displayName As String
        fieldName = TableEventHandler.GetFieldNameByColumn(columnNumber)
        displayName = TableEventHandler.GetFieldDisplayName(fieldName)
        Application.StatusBar = LocalizationManager.GetText("Row ") & RowNumber & LocalizationManager.GetText(", Col ") & columnNumber & _
                               LocalizationManager.GetText(" | Ctrl+E to edit | ") & displayName
    Else
        Application.StatusBar = LocalizationManager.GetText("Interactive forms system is active. Select a table cell.")
    End If
    
    On Error GoTo 0
End Sub

