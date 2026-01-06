Attribute VB_Name = "MdlBackupVBA"

' Модуль MdlBackup для создания резервных копий VBA проекта
' Версия: 1.0.0
' Дата: 12.07.2025
' Описание: Экспортирует все модули VBA в отдельные файлы для создания снапшота проекта
' Функциональность: Создает резервную копию всех модулей, форм и классов в указанную папку

Sub CreateVBASnapshot()
    Dim vbComp As Object
    Dim exportPath As String
    Dim FileName As String
    Dim timeStamp As String
    Dim fso As Object
    
    ' Создаем временную метку для уникальности
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    
    ' Создаем папку для снапшота
    exportPath = ThisWorkbook.Path & "\VBA_Snapshots\Snapshot_" & timeStamp & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder Left(exportPath, Len(exportPath) - 1)
    End If
    
    ' Экспортируем все компоненты VBA проекта
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule - стандартные модули
                FileName = vbComp.Name & ".bas"
            Case 2 ' vbext_ct_ClassModule - модули классов
                FileName = vbComp.Name & ".cls"
            Case 3 ' vbext_ct_MSForm - пользовательские формы
                FileName = vbComp.Name & ".frm"
            Case 100 ' vbext_ct_Document - модули листов/книги
                FileName = vbComp.Name & ".cls"
        End Select
        
        ' Экспортируем компонент
        vbComp.Export exportPath & FileName
        Debug.Print "Экспортирован: " & FileName
    Next vbComp
    
    ' Создаем файл с информацией о снапшоте
    Dim infoFile As String
    infoFile = exportPath & "SnapshotInfo.txt"
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open infoFile For Output As FileNum
    Print #FileNum, "VBA Project Snapshot"
    Print #FileNum, "Дата создания: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
    Print #FileNum, "Файл проекта: " & ThisWorkbook.Name
    Print #FileNum, "Путь: " & ThisWorkbook.FullName
    Print #FileNum, "Количество модулей: " & ThisWorkbook.VBProject.VBComponents.Count
    Print #FileNum, ""
    Print #FileNum, "Список модулей:"
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Print #FileNum, "- " & vbComp.Name & " (Тип: " & GetComponentTypeName(vbComp.Type) & ")"
    Next vbComp
    Close FileNum
    
    MsgBox "Снапшот VBA проекта создан!" & vbCrLf & _
           "Путь: " & exportPath & vbCrLf & _
           "Экспортировано модулей: " & ThisWorkbook.VBProject.VBComponents.Count, _
           vbInformation, "Снапшот создан"
    
    ' Открываем папку со снапшотом
    Shell "explorer.exe " & exportPath, vbNormalFocus
End Sub

' Вспомогательная функция для получения названия типа компонента
Function GetComponentTypeName(componentType As Integer) As String
    Select Case componentType
        Case 1: GetComponentTypeName = "Стандартный модуль"
        Case 2: GetComponentTypeName = "Модуль класса"
        Case 3: GetComponentTypeName = "Пользовательская форма"
        Case 100: GetComponentTypeName = "Модуль документа"
        Case Else: GetComponentTypeName = "Неизвестный тип"
    End Select
End Function


' Модуль для восстановления VBA проекта из снапшота
' Версия: 1.0.0
' Дата: 12.07.2025
' Описание: Импортирует модули VBA из папки снапшота для восстановления предыдущего состояния
' Функциональность: Удаляет текущие модули и загружает модули из выбранной резервной копии

Sub RestoreFromSnapshot()
    Dim importPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim vbComp As Object
    Dim componentName As String
    
    ' Выбираем папку со снапшотом
    importPath = SelectSnapshotFolder()
    If importPath = "" Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(importPath)
    
    ' Предупреждение пользователя
    Dim response As VbMsgBoxResult
    response = MsgBox("ВНИМАНИЕ!" & vbCrLf & _
                     "Это действие удалит все текущие модули VBA и заменит их модулями из снапшота." & vbCrLf & _
                     "Вы уверены, что хотите продолжить?", _
                     vbYesNo + vbExclamation, "Подтверждение восстановления")
    
    If response = vbNo Then Exit Sub
    
    ' Удаляем все пользовательские модули (кроме модулей листов)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then ' Стандартные модули, классы, формы
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
    Next vbComp
    
    ' Импортируем модули из снапшота
    For Each file In folder.Files
        If Right(LCase(file.Name), 4) = ".bas" Or _
           Right(LCase(file.Name), 4) = ".cls" Or _
           Right(LCase(file.Name), 4) = ".frm" Then
            
            ThisWorkbook.VBProject.VBComponents.Import file.Path
            Debug.Print "Импортирован: " & file.Name
        End If
    Next file
    
    MsgBox "Снапшот успешно восстановлен!" & vbCrLf & _
           "Папка снапшота: " & importPath, _
           vbInformation, "Восстановление завершено"
End Sub

' Функция для выбора папки со снапшотом
Function SelectSnapshotFolder() As String
    Dim snapshotsPath As String
    Dim selectedPath As String
    
    snapshotsPath = ThisWorkbook.Path & "\VBA_Snapshots\"
    
    ' Проверяем существование папки со снапшотами
    If dir(snapshotsPath, vbDirectory) = "" Then
        MsgBox "Папка со снапшотами не найдена: " & snapshotsPath, vbExclamation
        SelectSnapshotFolder = ""
        Exit Function
    End If
    
    ' Здесь можно добавить диалог выбора папки
    ' Для простоты используем InputBox
    selectedPath = InputBox("Введите имя папки снапшота:" & vbCrLf & _
                           "Доступные снапшоты находятся в: " & snapshotsPath, _
                           "Выбор снапшота", "Snapshot_")
    
    If selectedPath <> "" Then
        SelectSnapshotFolder = snapshotsPath & selectedPath & "\"
    Else
        SelectSnapshotFolder = ""
    End If
End Function


' Модуль для управления версиями кода через комментарии
' Версия: 1.0.0
' Дата: 12.07.2025
' Описание: Добавляет метки версий в код для отслеживания изменений
' Функциональность: Вставляет комментарии с версиями и датами в начало каждого модуля

Sub AddVersionTagsToAllModules()
    Dim vbComp As Object
    Dim codeModule As Object
    Dim versionTag As String
    Dim CurrentDate As String
    
    CurrentDate = Format(Now, "dd.mm.yyyy hh:mm")
    versionTag = "' === СНАПШОТ ВЕРСИИ === " & CurrentDate & " ==="
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Then ' Только стандартные модули
            Set codeModule = vbComp.codeModule
            
            ' Добавляем метку версии в начало модуля
            codeModule.InsertLines 1, versionTag
            codeModule.InsertLines 2, "' Рабочая версия сохранена: " & CurrentDate
            codeModule.InsertLines 3, ""
            
            Debug.Print "Метка версии добавлена в модуль: " & vbComp.Name
        End If
    Next vbComp
    
    MsgBox "Метки версий добавлены во все модули!", vbInformation
End Sub


' Модуль для создания именованных копий Excel файла
' Версия: 1.0.0
' Дата: 12.07.2025
' Описание: Создает копии текущего Excel файла с временными метками для быстрого отката
' Функциональность: Сохраняет копию файла с описательным именем и временной меткой

Sub CreateWorkbookSnapshot()
    Dim originalPath As String
    Dim snapshotPath As String
    Dim timeStamp As String
    Dim description As String
    Dim FileName As String
    
    ' Получаем описание от пользователя
    description = InputBox("Введите краткое описание этого снапшота:", _
                          "Описание снапшота", "Рабочая_версия")
    
    If description = "" Then Exit Sub
    
    ' Создаем временную метку
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm")
    
    ' Формируем имя файла
    FileName = Replace(ThisWorkbook.Name, ".xlsm", "") & "_" & description & "_" & timeStamp & ".xlsm"
    
    ' Определяем путь для сохранения
    originalPath = ThisWorkbook.Path
    snapshotPath = originalPath & "\Snapshots\"
    
    ' Создаем папку, если её нет
    If dir(snapshotPath, vbDirectory) = "" Then
        MkDir snapshotPath
    End If
    
    ' Сохраняем копию
    ThisWorkbook.SaveCopyAs snapshotPath & FileName
    
    MsgBox "Снапшот создан!" & vbCrLf & _
           "Файл: " & FileName & vbCrLf & _
           "Путь: " & snapshotPath, _
           vbInformation, "Снапшот сохранен"
    
    ' Открываем папку со снапшотами
    Shell "explorer.exe " & snapshotPath, vbNormalFocus
End Sub


' Быстрая команда для создания снапшота
Sub QuickSnapshot()
    Call AddVersionTagsToAllModules
    Call CreateVBASnapshot
    Call CreateWorkbookSnapshot
End Sub



' Добавить в любой модуль для быстрого запуска
Public Sub QuickStart()
    UserFormVhIsh.Show
End Sub
