Attribute VB_Name = "AutoStart"
'==============================================
' SYSTEM AUTOSTART MODULE - AutoStart
' Purpose: Automatic system initialization upon workbook opening
' State: POPUP WINDOWS DISABLED - STATUS BAR AND DEBUG ONLY
' Version: 2.3.0
' Date: 10.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' Global variable to track system state
Public SystemInitialized As Boolean

' POPUP WINDOWS DISABLED: Main system initialization
Public Sub InitializeSystem()
    On Error GoTo InitError
    
    ' Reset initialization flag
    SystemInitialized = False
    
    ' Inform user via status bar (no MsgBox)
    Application.StatusBar = "Initializing interactive forms system..."
    
    ' 1. Initialize table event handler
    Call TableEventHandler.InitializeTableEvents
    
    ' 2. Force context menu activation
    Call ForceActivateContextMenu
    
    ' 3. Check system state (Debug only)
    Call DiagnoseSystemState
    
    ' 4. Activate sheet with the table
    On Error Resume Next
    ThisWorkbook.Worksheets("IncOut").Activate
    On Error GoTo InitError
    
    ' 5. Set successful initialization flag
    SystemInitialized = True
    
    ' POPUP WINDOW DISABLED: Status bar only
    Application.StatusBar = "Interactive forms system is active. Ready to work."
    
    ' DISABLED: ShowUserInstructions - removed popup instructions
    ' Debug information for developer
    Debug.Print "Interactive forms system successfully initialized"
    
    Exit Sub
    
InitError:
    SystemInitialized = False
    ' KEPT: Critical errors are shown via MsgBox
    MsgBox "CRITICAL ERROR initializing system!" & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.description, vbCritical, "System Error"
    
    Application.StatusBar = "ERROR initializing system."
End Sub

' Force activation of the context menu
Public Sub ForceActivateContextMenu()
    On Error GoTo MenuError
    
    ' Remove existing buttons (if any)
    Call TableEventHandler.RemoveContextMenuButton
    
    ' Short pause
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Add context menu
    Call TableEventHandler.AddContextMenuButton
    
    ' Check if successfully added
    If Not CheckContextMenuExists() Then
        ' Retry
        Application.Wait Now + TimeValue("00:00:01")
        Call TableEventHandler.AddContextMenuButton
    End If
    
    Exit Sub
    
MenuError:
    Debug.Print "Error forcing context menu activation: " & Err.description
End Sub

' Check if context menu exists
Public Function CheckContextMenuExists() As Boolean
    Dim contextMenu As CommandBar
    Dim ctrl As CommandBarControl
    
    On Error GoTo CheckError
    
    Set contextMenu = Application.CommandBars("Cell")
    
    For Each ctrl In contextMenu.Controls
        If ctrl.Caption = "Duplicate record" Or InStr(ctrl.Caption, "Duplicate") > 0 Then
            CheckContextMenuExists = True
            Exit Function
        End If
    Next ctrl
    
    CheckContextMenuExists = False
    Exit Function
    
CheckError:
    CheckContextMenuExists = False
End Function

' System state diagnostics (Debug only, no MsgBox)
Public Sub DiagnoseSystemState()
    Dim diagResult As String
    
    diagResult = "SYSTEM DIAGNOSTICS:" & vbCrLf & vbCrLf
    
    ' Check sheet
    On Error Resume Next
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("IncOut")
    If Ws Is Nothing Then
        diagResult = diagResult & "[X] Sheet 'IncOut' NOT found!" & vbCrLf
    Else
        diagResult = diagResult & "[OK] Sheet 'IncOut' found" & vbCrLf
        
        ' Check table
        Dim tbl As ListObject
        Set tbl = Ws.ListObjects("TableIncOut")
        If tbl Is Nothing Then
            diagResult = diagResult & "[X] Table 'TableIncOut' NOT found!" & vbCrLf
        Else
            diagResult = diagResult & "[OK] Table found (" & tbl.ListRows.Count & " rows)" & vbCrLf
        End If
    End If
    On Error GoTo 0
    
    ' Check context menu
    If CheckContextMenuExists() Then
        diagResult = diagResult & "[OK] Context menu is active" & vbCrLf
    Else
        diagResult = diagResult & "[X] Context menu is NOT active!" & vbCrLf
    End If
    
    ' POPUP WINDOW DISABLED: Write result to Debug only
    Debug.Print diagResult
End Sub

' POPUP WINDOW DISABLED: Show instructions to user
' Public Sub ShowUserInstructions()
'     ' This procedure is disabled to prevent popups
'     Debug.Print "Instructions disabled. System ready to work."
' End Sub

' Manual system initialization (no popup windows)
Public Sub ManualInitializeSystem()
    Application.StatusBar = "Manual system initialization..."
    Call InitializeSystem
    Application.StatusBar = "Manual initialization completed."
End Sub

' POPUP WINDOWS DISABLED: Force context menu reload
Public Sub ReloadContextMenu()
    On Error GoTo ReloadError
    
    Application.StatusBar = "Reloading context menu..."
    
    ' Remove
    Call TableEventHandler.RemoveContextMenuButton
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Add
    Call TableEventHandler.AddContextMenuButton
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Check (Status bar only)
    If CheckContextMenuExists() Then
        Application.StatusBar = "[OK] Context menu successfully reloaded!"
        Debug.Print "Context menu successfully reloaded"
    Else
        Application.StatusBar = "[X] Error reloading context menu!"
        Debug.Print "Error reloading context menu"
    End If
    
    Exit Sub
    
ReloadError:
    Application.StatusBar = "Reload error: " & Err.description
    Debug.Print "Reload error: " & Err.description
End Sub

' POPUP WINDOWS DISABLED: Context menu test
Public Sub TestContextMenu()
    Dim result As String
    
    result = "CONTEXT MENU TEST:" & vbCrLf & vbCrLf
    
    ' Activate sheet
    On Error Resume Next
    ThisWorkbook.Worksheets("IncOut").Activate
    On Error GoTo 0
    
    ' Check if menu exists
    If CheckContextMenuExists() Then
        result = result & "[OK] Context menu found in the system" & vbCrLf
        result = result & "Right-click a cell INSIDE the table"
        Application.StatusBar = "Context menu found - ready to use"
    Else
        result = result & "[X] Context menu NOT found!" & vbCrLf
        result = result & "Forcing activation..."
        
        Call ForceActivateContextMenu
        
        If CheckContextMenuExists() Then
            result = result & vbCrLf & "[OK] Activation successful!"
            Application.StatusBar = "Context menu activated successfully"
        Else
            result = result & vbCrLf & "[X] Activation failed!"
            Application.StatusBar = "Error activating context menu"
        End If
    End If
    
    ' POPUP WINDOW DISABLED: Debug and status bar only
    Debug.Print result
End Sub

' System deactivation (renamed from DeactivateSystem)
Public Sub ShutdownSystem()
    On Error Resume Next
    
    ' Deactivate event handler
    Call TableEventHandler.DeactivateTableEvents
    
    ' Remove context menu
    Call TableEventHandler.RemoveContextMenuButton
    
    ' Reset system flag
    SystemInitialized = False
    
    ' Update status
    Application.StatusBar = "Interactive forms system shut down"
    Debug.Print "System correctly shut down"
    
    On Error GoTo 0
End Sub

' Alternative name for compatibility
Public Sub DeactivateSystem()
    ' Call main shutdown procedure
    Call ShutdownSystem
End Sub

' Function to check system state
Public Function IsSystemReady() As Boolean
    IsSystemReady = SystemInitialized And CheckContextMenuExists()
End Function

' POPUP WINDOWS DISABLED: Full diagnostics for user
Public Sub FullDiagnostic()
    Dim Report As String
    
    Report = "=== FULL SYSTEM DIAGNOSTICS ===" & vbCrLf & vbCrLf
    
    ' 1. Check initialization
    Report = Report & "System initialized: " & IIf(SystemInitialized, "[OK] YES", "[X] NO") & vbCrLf
    
    ' 2. Check sheet and table
    On Error Resume Next
    Dim Ws As Worksheet
    Dim tbl As ListObject
    Set Ws = ThisWorkbook.Worksheets("IncOut")
    
    If Not Ws Is Nothing Then
        Report = Report & "Sheet 'IncOut': [OK] Found" & vbCrLf
        Set tbl = Ws.ListObjects("TableIncOut")
        If Not tbl Is Nothing Then
            Report = Report & "Table: [OK] Found (" & tbl.ListRows.Count & " rows)" & vbCrLf
        Else
            Report = Report & "Table: [X] NOT found!" & vbCrLf
        End If
    Else
        Report = Report & "Sheet 'IncOut': [X] NOT found!" & vbCrLf
    End If
    On Error GoTo 0
    
    ' 3. Check context menu
    Report = Report & "Context menu: " & IIf(CheckContextMenuExists(), "[OK] Active", "[X] NOT active") & vbCrLf
    
    ' 4. Check active sheet
    Report = Report & "Active sheet: " & ActiveSheet.Name & vbCrLf
    
    ' 5. Office version
    Report = Report & "Office version: " & Application.Version & vbCrLf
    
    ' POPUP WINDOW DISABLED: Debug and status bar only
    Debug.Print Report
    Application.StatusBar = "Diagnostics completed. See results in Debug (Ctrl+G)"
End Sub

' Emergency system shutdown (no popup windows)
Public Sub EmergencyShutdown()
    On Error Resume Next
    
    ' Force clear all resources
    Call TableEventHandler.RemoveContextMenuButton
    Call TableEventHandler.DeactivateTableEvents
    
    ' Reset all state variables
    SystemInitialized = False
    
    ' Clear status bar
    Application.StatusBar = "Emergency system shutdown completed"
    Debug.Print "Emergency system shutdown completed. All resources freed."
    
    On Error GoTo 0
End Sub

' Check system integrity
Public Function SystemIntegrityCheck() As Boolean
    Dim allGood As Boolean
    
    On Error GoTo IntegrityError
    
    allGood = True
    
    ' Check 1: Sheet exists
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("IncOut")
    If Ws Is Nothing Then allGood = False
    
    ' Check 2: Table exists
    If Not Ws Is Nothing Then
        Dim tbl As ListObject
        Set tbl = Ws.ListObjects("TableIncOut")
        If tbl Is Nothing Then allGood = False
    End If
    
    ' Check 3: Modules exist
    ' (If VBA reached this point, modules exist)
    
    SystemIntegrityCheck = allGood
    Exit Function
    
IntegrityError:
    SystemIntegrityCheck = False
End Function

' POPUP WINDOWS DISABLED: Quick system restart
Public Sub RestartSystem()
    Application.StatusBar = "Restarting interactive forms system..."
    Debug.Print "System restart initiated"
    
    ' Shutdown current system
    Call ShutdownSystem
    
    ' Pause
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Start again
    Call InitializeSystem
    
    Application.StatusBar = "System successfully restarted!"
    Debug.Print "System successfully restarted!"
End Sub

