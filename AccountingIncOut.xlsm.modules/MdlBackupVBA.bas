Attribute VB_Name = "MdlBackupVBA"
'==============================================
' MdlBackup module for creating VBA project backups
' Version: 1.0.0
' Date: 12.07.2025
' Description: Exports all VBA modules to separate files to create a project snapshot
' Functionality: Creates a backup of all modules, forms, and classes in a specified folder
'==============================================

Option Explicit

Sub CreateVBASnapshot()
    Dim vbComp As Object
    Dim exportPath As String
    Dim FileName As String
    Dim timeStamp As String
    Dim fso As Object
    
    ' Create timestamp for uniqueness
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    
    ' Create folder for snapshot
    exportPath = ThisWorkbook.Path & "\VBA_Snapshots\Snapshot_" & timeStamp & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder Left(exportPath, Len(exportPath) - 1)
    End If
    
    ' Export all VBA project components
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule - standard modules
                FileName = vbComp.Name & ".bas"
            Case 2 ' vbext_ct_ClassModule - class modules
                FileName = vbComp.Name & ".cls"
            Case 3 ' vbext_ct_MSForm - user forms
                FileName = vbComp.Name & ".frm"
            Case 100 ' vbext_ct_Document - sheet/workbook modules
                FileName = vbComp.Name & ".cls"
        End Select
        
        ' Export component
        vbComp.Export exportPath & FileName
        Debug.Print "Exported: " & FileName
    Next vbComp
    
    ' Create file with snapshot info
    Dim infoFile As String
    infoFile = exportPath & "SnapshotInfo.txt"
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open infoFile For Output As FileNum
    Print #FileNum, "VBA Project Snapshot"
    Print #FileNum, "Creation date: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
    Print #FileNum, "Project file: " & ThisWorkbook.Name
    Print #FileNum, "Path: " & ThisWorkbook.FullName
    Print #FileNum, "Number of modules: " & ThisWorkbook.VBProject.VBComponents.Count
    Print #FileNum, ""
    Print #FileNum, "Module list:"
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Print #FileNum, "- " & vbComp.Name & " (Type: " & GetComponentTypeName(vbComp.Type) & ")"
    Next vbComp
    Close FileNum
    
    MsgBox "VBA project snapshot created!" & vbCrLf & _
           "Path: " & exportPath & vbCrLf & _
           "Modules exported: " & ThisWorkbook.VBProject.VBComponents.Count, _
           vbInformation, "Snapshot Created"
    
    ' Open snapshot folder
    Shell "explorer.exe " & exportPath, vbNormalFocus
End Sub

' Helper function to get component type name
Function GetComponentTypeName(componentType As Integer) As String
    Select Case componentType
        Case 1: GetComponentTypeName = "Standard module"
        Case 2: GetComponentTypeName = "Class module"
        Case 3: GetComponentTypeName = "User form"
        Case 100: GetComponentTypeName = "Document module"
        Case Else: GetComponentTypeName = "Unknown type"
    End Select
End Function

' Module for restoring VBA project from a snapshot
' Version: 1.0.0
' Date: 12.07.2025
' Description: Imports VBA modules from snapshot folder to restore previous state
' Functionality: Deletes current modules and loads modules from selected backup

Sub RestoreFromSnapshot()
    Dim importPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim vbComp As Object
    Dim componentName As String
    
    ' Select snapshot folder
    importPath = SelectSnapshotFolder()
    If importPath = "" Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(importPath)
    
    ' User warning
    Dim response As VbMsgBoxResult
    response = MsgBox("WARNING!" & vbCrLf & _
                      "This action will delete all current VBA modules and replace them with modules from the snapshot." & vbCrLf & _
                      "Are you sure you want to continue?", _
                      vbYesNo + vbExclamation, "Restore Confirmation")
    
    If response = vbNo Then Exit Sub
    
    ' Delete all user modules (except document modules)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then ' Standard modules, classes, forms
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
    Next vbComp
    
    ' Import modules from snapshot
    For Each file In folder.Files
        If Right(LCase(file.Name), 4) = ".bas" Or _
           Right(LCase(file.Name), 4) = ".cls" Or _
           Right(LCase(file.Name), 4) = ".frm" Then
            
            ThisWorkbook.VBProject.VBComponents.Import file.Path
            Debug.Print "Imported: " & file.Name
        End If
    Next file
    
    MsgBox "Snapshot successfully restored!" & vbCrLf & _
           "Snapshot folder: " & importPath, _
           vbInformation, "Restore Completed"
End Sub

' Function to select snapshot folder
Function SelectSnapshotFolder() As String
    Dim snapshotsPath As String
    Dim selectedPath As String
    
    snapshotsPath = ThisWorkbook.Path & "\VBA_Snapshots\"
    
    ' Check if snapshots folder exists
    If dir(snapshotsPath, vbDirectory) = "" Then
        MsgBox "Snapshots folder not found: " & snapshotsPath, vbExclamation
        SelectSnapshotFolder = ""
        Exit Function
    End If
    
    ' Simple InputBox for folder selection
    selectedPath = InputBox("Enter the snapshot folder name:" & vbCrLf & _
                            "Available snapshots are located in: " & snapshotsPath, _
                            "Select Snapshot", "Snapshot_")
    
    If selectedPath <> "" Then
        SelectSnapshotFolder = snapshotsPath & selectedPath & "\"
    Else
        SelectSnapshotFolder = ""
    End If
End Function

' Module for version control via comments
' Version: 1.0.0
' Date: 12.07.2025
' Description: Adds version tags to code for tracking changes
' Functionality: Inserts comments with versions and dates at the beginning of each module

Sub AddVersionTagsToAllModules()
    Dim vbComp As Object
    Dim codeModule As Object
    Dim versionTag As String
    Dim CurrentDate As String
    
    CurrentDate = Format(Now, "dd.mm.yyyy hh:mm")
    versionTag = "' === VERSION SNAPSHOT === " & CurrentDate & " ==="
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Then ' Standard modules only
            Set codeModule = vbComp.codeModule
            
            ' Add version tag to the beginning of the module
            codeModule.InsertLines 1, versionTag
            codeModule.InsertLines 2, "' Working version saved: " & CurrentDate
            codeModule.InsertLines 3, ""
            
            Debug.Print "Version tag added to module: " & vbComp.Name
        End If
    Next vbComp
    
    MsgBox "Version tags added to all modules!", vbInformation
End Sub

' Module for creating named Excel file copies
' Version: 1.0.0
' Date: 12.07.2025
' Description: Creates copies of current Excel file with timestamps for quick rollback
' Functionality: Saves a file copy with descriptive name and timestamp

Sub CreateWorkbookSnapshot()
    Dim originalPath As String
    Dim snapshotPath As String
    Dim timeStamp As String
    Dim description As String
    Dim FileName As String
    
    ' Get description from user
    description = InputBox("Enter a brief description for this snapshot:", _
                          "Snapshot Description", "Working_version")
    
    If description = "" Then Exit Sub
    
    ' Create timestamp
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm")
    
    ' Format file name
    FileName = Replace(ThisWorkbook.Name, ".xlsm", "") & "_" & description & "_" & timeStamp & ".xlsm"
    
    ' Define save path
    originalPath = ThisWorkbook.Path
    snapshotPath = originalPath & "\Snapshots\"
    
    ' Create folder if it doesn't exist
    If dir(snapshotPath, vbDirectory) = "" Then
        MkDir snapshotPath
    End If
    
    ' Save copy
    ThisWorkbook.SaveCopyAs snapshotPath & FileName
    
    MsgBox "Snapshot created!" & vbCrLf & _
           "File: " & FileName & vbCrLf & _
           "Path: " & snapshotPath, _
           vbInformation, "Snapshot Saved"
    
    ' Open snapshots folder
    Shell "explorer.exe " & snapshotPath, vbNormalFocus
End Sub

' Quick command to create a snapshot
Sub QuickSnapshot()
    Call AddVersionTagsToAllModules
    Call CreateVBASnapshot
    Call CreateWorkbookSnapshot
End Sub

' Add to any module for quick start
Public Sub QuickStart()
    UserFormVhIsh.Show
End Sub

