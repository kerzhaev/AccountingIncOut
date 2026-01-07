VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================
' МОДУЛЬ КНИГИ - ThisWorkbook
' Назначение: Обработка событий книги
' Состояние: УПРОЩЁН - ручное управление модулями
' Версия: 2.0.0
'==============================================

Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    
    ' Автоматически активируем систему при открытии книги
    ' (работает только если модули уже загружены)
    Call AutoStart.InitializeSystem
    
    ' НОВОЕ: Инициализация инфраструктуры этапа 1
    Call SystemSettings.InitializeSettings
    Call SystemLogger.InitializeLog
    
    ' Логируем старт системы
    Call SystemLogger.LogOperation("SystemStart", "Система УчётВхИсх запущена", "SUCCESS", 0)
    
    ' Открытие формы немодально
    Load UserFormVhIsh
    UserFormVhIsh.Show vbModeless
    
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    On Error Resume Next
    
    ' Экспортируем модули в папку при сохранении
    ' Раскомментируйте следующую строку, если хотите автоматический экспорт:
    ' Call ExportModules(ThisWorkbook.Name & ".modules")
    
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' Деактивируем систему при закрытии книги
    Call AutoStart.ShutdownSystem
    
    ' УДАЛЕНИЕ МОДУЛЕЙ ОТКЛЮЧЕНО - делайте вручную если нужно
    ' Call RemoveModules(ShowMsgBox:=True)
    
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    
    If Sh.Name = "ВхИсх" Then
        Application.StatusBar = "Лист с таблицей активен. Выберите ячейку и нажмите Ctrl+E для редактирования."
    End If
    
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    
    If Sh.Name = "ВхИсх" Then
        Call UpdateStatusForSelection(Target)
    End If
    
    On Error GoTo 0
End Sub

Private Sub UpdateStatusForSelection(Target As Range)
    On Error Resume Next
    
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim intersectRange As Range
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
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
        Application.StatusBar = "Строка " & RowNumber & ", столбец " & columnNumber & _
                               " | Ctrl+E для редактирования | " & displayName
    Else
        Application.StatusBar = "Система интерактивных форм активна. Выберите ячейку таблицы."
    End If
    
    On Error GoTo 0
End Sub

