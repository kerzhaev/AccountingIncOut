Attribute VB_Name = "VbaModuleManager"
Option Explicit
Option Private Module

#Const MANAGING_WORD = 0
#Const MANAGING_EXCEL = 1

Private Const MY_NAME = "VbaModuleManager"
Private Const ERR_SUPPORTED_APPS = MY_NAME & " currently only supports Microsoft Word and Excel."

Dim allComponents As VBComponents
Dim fileSys As New FileSystemObject
Dim alreadySaved As Boolean

Public Sub ImportModules(FromDirectory As String, Optional ShowMsgBox As Boolean = True)
    On Error GoTo ErrorHandler
    
    'Cache some references
    'If the given directory does not exist then show an error dialog and exit
    Dim fromPath As String: fromPath = FromDirectory
    Dim basePath As String: basePath = getFilePath()
    
    ' Если книга не сохранена, basePath будет пустой
    If basePath = "" Then
        MsgBox "Ошибка: Файл должен быть сохранён перед импортом модулей!", vbCritical, "Ошибка импорта"
        Exit Sub
    End If
    
    ' Проверяем относительный путь относительно папки с файлом
    If Not fileSys.FolderExists(fromPath) Then
        fromPath = basePath & "\" & FromDirectory
        If Not fileSys.FolderExists(fromPath) Then
            Dim errMsg As String
            errMsg = "Не удалось найти директорию импорта: " & FromDirectory & vbCrLf & vbCrLf
            errMsg = errMsg & "Проверенные пути:" & vbCrLf
            errMsg = errMsg & "1. " & FromDirectory & vbCrLf
            errMsg = errMsg & "2. " & fromPath
            MsgBox errMsg, vbCritical, "Ошибка импорта"
            Exit Sub
        End If
    End If
    
    Dim dir As folder: Set dir = fileSys.GetFolder(fromPath)
                
    'Import all VB code files from the given directory if any)
    Dim f As file
    Dim imports As New Scripting.Dictionary 'Must be qualified to distinguish it from an MS Word Dictionary
    Dim numFiles As Long: numFiles = 0
    For Each f In dir.Files
        Dim dotIndex As Long: dotIndex = InStrRev(f.Name, ".")
        If dotIndex > 0 Then
            Dim ext As String: ext = UCase(Right(f.Name, Len(f.Name) - dotIndex))
            Dim correctType As Boolean: correctType = (ext = "BAS" Or ext = "CLS" Or ext = "FRM")
            Dim allowedName As Boolean: allowedName = Left(f.Name, dotIndex - 1) <> MY_NAME
            If correctType And allowedName Then
                numFiles = numFiles + 1
                Dim replaced As Boolean: replaced = doImport(f)
                imports.Add f.Name, replaced
            End If
        End If
    Next f
    
    'Show a success message box, if requested
    If ShowMsgBox Then
        If numFiles = 0 Then
            MsgBox "Модули не найдены в директории: " & fromPath, vbInformation, "Импорт модулей"
        Else
            Dim i As Long
            Dim msg As String: msg = numFiles & " модулей импортировано:" & vbCrLf & vbCrLf
            For i = 0 To imports.Count - 1
                msg = msg & "    " & imports.Keys()(i) & IIf(imports.Items()(i), " (заменён)", " (новый)") & vbCrLf
            Next i
            Dim result As VbMsgBoxResult: result = MsgBox(msg, vbOKOnly, "Импорт модулей")
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при импорте модулей: " & Err.description & vbCrLf & "Код ошибки: " & Err.Number, vbCritical, "Ошибка импорта"
End Sub

Public Sub ExportModules(ToDirectory As String)
    On Error GoTo ErrorHandler
    
    'Cache some references
    'If the given directory does not exist then show an error dialog and exit
    Dim toPath As String: toPath = ToDirectory
    Dim basePath As String: basePath = getFilePath()
    
    If basePath = "" Then
        MsgBox "Ошибка: Файл должен быть сохранён перед экспортом модулей!", vbCritical, "Ошибка экспорта"
        Exit Sub
    End If
    
    If Not fileSys.FolderExists(toPath) Then
        toPath = basePath & "\" & ToDirectory
        If Not fileSys.FolderExists(toPath) Then
            fileSys.CreateFolder (toPath)
        End If
    End If
    Dim dir As folder: Set dir = fileSys.GetFolder(toPath)
    
    'Export all modules from this file (except default MS Office modules)
    Dim vbc As VBComponent
    Dim allComponents As VBComponents: Set allComponents = getAllComponents()
    For Each vbc In allComponents
        Dim correctType As Boolean: correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType And vbc.Name <> MY_NAME Then
            Call doExport(vbc, dir.Path)
        End If
    Next vbc
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при экспорте модулей: " & Err.description & vbCrLf & "Код ошибки: " & Err.Number, vbCritical, "Ошибка экспорта"
End Sub

Public Sub RemoveModules(Optional ShowMsgBox As Boolean = True)
    On Error GoTo ErrorHandler
    
    'Check the saved flag to prevent a save event loop
    If alreadySaved Then
        alreadySaved = False
        Exit Sub
    End If

    'Remove all modules from this file (except default MS Office modules obviously)
    Dim removals As New Collection
    Dim vbc As VBComponent
    Dim numModules As Long: numModules = 0
    Dim allComponents As VBComponents: Set allComponents = getAllComponents()
    For Each vbc In allComponents
        Dim correctType As Boolean: correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType And vbc.Name <> MY_NAME Then
            numModules = numModules + 1
            removals.Add vbc.Name
            allComponents.Remove vbc
        End If
    Next vbc
    
    'Set the saved flag to prevent a save event loop
    'Save file again now that all modules have been removed
    alreadySaved = True
    Call saveFile
        
    'Show a success message box
    If ShowMsgBox Then
        Dim item As Variant
        Dim msg As String: msg = numModules & " модулей успешно удалено:" & vbCrLf & vbCrLf
        For Each item In removals
            msg = msg & "    " & item & vbCrLf
        Next item
        msg = msg & vbCrLf & "Не забудьте удалить пустые строки после строк Attribute в .frm файлах..." _
                  & vbCrLf & "VbaModuleManager никогда не будет переимпортирован или экспортирован. Вы должны сделать это вручную при необходимости." _
                  & vbCrLf & "НИКОГДА не редактируйте код в VBE и отдельном редакторе одновременно!"
        Dim result As VbMsgBoxResult: result = MsgBox(msg, vbOKOnly, "Удаление модулей")
    End If
    
    Exit Sub
    
ErrorHandler:
    alreadySaved = False
    MsgBox "Ошибка при удалении модулей: " & Err.description & vbCrLf & "Код ошибки: " & Err.Number, vbCritical, "Ошибка удаления"
End Sub

Private Function getFilePath() As String
    #If MANAGING_WORD Then
        getFilePath = ThisDocument.Path
    #ElseIf MANAGING_EXCEL Then
        getFilePath = ThisWorkbook.Path
    #Else
        Call raiseUnsupportedAppError
    #End If
End Function

Private Function getAllComponents() As VBComponents
    #If MANAGING_WORD Then
        Set getAllComponents = ThisDocument.VBProject.VBComponents
    #ElseIf MANAGING_EXCEL Then
        Set getAllComponents = ThisWorkbook.VBProject.VBComponents
    #Else
        Call raiseUnsupportedAppError
    #End If
End Function

Private Sub saveFile()
    #If MANAGING_WORD Then
        ThisDocument.Save
    #ElseIf MANAGING_EXCEL Then
        ThisWorkbook.Save
    #Else
        Call raiseUnsupportedAppError
    #End If
End Sub

Private Sub raiseUnsupportedAppError()
    Err.Raise Number:=vbObjectError + 1, description:=ERR_SUPPORTED_APPS
End Sub

Private Function doImport(ByRef codeFile As file) As Boolean
    On Error GoTo ErrorHandler
    
    'Determine whether a module with this name already exists
    Dim Name As String: Name = Left(codeFile.Name, Len(codeFile.Name) - 4)
    On Error Resume Next
    Dim allComponents As VBComponents: Set allComponents = getAllComponents()
    Dim m As VBComponent: Set m = allComponents.item(Name)
    If Err.Number <> 0 Then
        Set m = Nothing
    End If
    On Error GoTo ErrorHandler
        
    'If so, remove it
    Dim alreadyExists As Boolean: alreadyExists = Not (m Is Nothing)
    If alreadyExists Then
        allComponents.Remove m
    End If
    
    'Then import the new module
    allComponents.Import (codeFile.Path)
    doImport = alreadyExists
    
    Exit Function
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "Ошибка при импорте файла " & codeFile.Name & ":" & vbCrLf & Err.description
    MsgBox errMsg, vbCritical, "Ошибка импорта"
    doImport = False
End Function

Private Function doExport(ByRef module As VBComponent, dirPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    'Determine whether a file with this component's name already exists
    Dim ext As String
    Select Case module.Type
        Case vbext_ct_MSForm
            ext = "frm"
        Case vbext_ct_ClassModule
            ext = "cls"
        Case vbext_ct_StdModule
            ext = "bas"
    End Select
    Dim filePath As String: filePath = dirPath & "\" & module.Name & "." & ext
    Dim alreadyExists As Boolean: alreadyExists = fileSys.FileExists(filePath)
        
    'If so, remove it (even if its ReadOnly)
    If alreadyExists Then
        Dim f As file: Set f = fileSys.GetFile(filePath)
        If (f.Attributes And 1) Then
            f.Attributes = f.Attributes - 1 'The bitmask for ReadOnly file attribute
        End If
        fileSys.DeleteFile (filePath)
    End If
    
    'Then export the module
    'Remove it also, so that the file stays small (and unchanged according to version control)
    module.Export (filePath)
    doExport = alreadyExists
    
    Exit Function
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "Ошибка при экспорте модуля " & module.Name & ":" & vbCrLf & Err.description
    MsgBox errMsg, vbCritical, "Ошибка экспорта"
    doExport = False
End Function


' =============================================
' ПУБЛИЧНЫЕ ПРОЦЕДУРЫ ДЛЯ ВЫЗОВА ИЗ МАКРОСОВ EXCEL
' =============================================

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Импортирует все модули из папки .modules
' =============================================
Public Sub ИмпортМодулей()
    On Error GoTo ErrorHandler
    
    Dim folderPath As String
    folderPath = ThisWorkbook.Name & ".modules"
    
    Call ImportModules(folderPath, ShowMsgBox:=True)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при импорте модулей: " & Err.description, vbCritical, "Ошибка"
End Sub

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Экспортирует все модули в папку .modules
' =============================================
Public Sub ЭкспортМодулей()
    On Error GoTo ErrorHandler
    
    Dim folderPath As String
    folderPath = ThisWorkbook.Name & ".modules"
    
    Call ExportModules(folderPath)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при экспорте модулей: " & Err.description, vbCritical, "Ошибка"
End Sub

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Удаляет все модули из проекта (кроме VbaModuleManager)
' =============================================
Public Sub УдалениеМодулей()
    On Error GoTo ErrorHandler
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Вы уверены, что хотите удалить все модули?" & vbCrLf & _
                     "VbaModuleManager будет сохранен.", _
                     vbYesNo + vbQuestion + vbDefaultButton2, _
                     "Подтверждение удаления")
    
    If response = vbYes Then
        Call RemoveModules(ShowMsgBox:=True)
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при удалении модулей: " & Err.description, vbCritical, "Ошибка"
End Sub
