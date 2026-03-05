VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================
' APPLICATION EVENT HANDLER CLASS MODULE - AppEventHandler
' Purpose: Handling Application events for Excel 2021
' State: CREATED TO FIX "Only valid in object module" ERROR
' Version: 1.0.1
' Date: 05.03.2026
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' CORRECT USAGE OF WithEvents - only in a class module
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' Initialize application event handler
Public Sub InitializeAppEvents()
    Set App = Application
End Sub

' Deactivate application event handler
Public Sub DeactivateAppEvents()
    Set App = Nothing
End Sub

' Window activation event
Private Sub App_WindowActivate(ByVal wb As Workbook, ByVal Wn As Window)
    On Error Resume Next
    
    ' Activate hotkeys on Excel window activation
    Application.OnKey "^+m", "TableEventHandler.ShowActionMenuExcel2021"  ' Ctrl+Shift+M
    Application.OnKey "^d", "TableEventHandler.DuplicateCurrentRecord"     ' Ctrl+D
    Application.OnKey "^e", "TableEventHandler.OpenFormForCurrentSelection" ' Ctrl+E
    
    On Error GoTo 0
End Sub

' Window deactivation event
Private Sub App_WindowDeactivate(ByVal wb As Workbook, ByVal Wn As Window)
    On Error Resume Next
    
    ' Reset hotkeys on deactivation
    Application.OnKey "^+m"  ' Reset Ctrl+Shift+M
    Application.OnKey "^d"   ' Reset Ctrl+D
    Application.OnKey "^e"   ' Reset Ctrl+E
    
    On Error GoTo 0
End Sub

' Workbook open event
Private Sub App_WorkbookOpen(ByVal wb As Workbook)
    ' Additional logic can be added on workbook open
    If wb Is ThisWorkbook Then
        ' Activate system when our workbook is opened
        Call TableEventHandler.InitializeTableEvents
    End If
End Sub
