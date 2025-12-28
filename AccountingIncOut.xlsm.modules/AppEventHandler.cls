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
' КЛАСС МОДУЛЬ ДЛЯ ОБРАБОТКИ СОБЫТИЙ ПРИЛОЖЕНИЯ - AppEventHandler
' Назначение: Обработка событий Application для Excel 2021
' Состояние: СОЗДАН ДЛЯ ИСПРАВЛЕНИЯ ОШИБКИ "Only valid in object module"
' Версия: 1.0.0
' Дата: 09.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' ПРАВИЛЬНОЕ ИСПОЛЬЗОВАНИЕ WithEvents - только в модуле класса
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' Инициализация обработчика событий приложения
Public Sub InitializeAppEvents()
    Set App = Application
End Sub

' Деактивация обработчика событий приложения
Public Sub DeactivateAppEvents()
    Set App = Nothing
End Sub

' Событие активации окна
Private Sub App_WindowActivate(ByVal wb As Workbook, ByVal Wn As Window)
    On Error Resume Next
    
    ' Активируем горячие клавиши при активации окна Excel
    Application.OnKey "^+m", "TableEventHandler.ShowActionMenuExcel2021"  ' Ctrl+Shift+M
    Application.OnKey "^d", "TableEventHandler.DuplicateCurrentRecord"     ' Ctrl+D
    Application.OnKey "^e", "TableEventHandler.OpenFormForCurrentSelection" ' Ctrl+E
    
    On Error GoTo 0
End Sub

' Событие деактивации окна
Private Sub App_WindowDeactivate(ByVal wb As Workbook, ByVal Wn As Window)
    On Error Resume Next
    
    ' Сбрасываем горячие клавиши при деактивации
    Application.OnKey "^+m"  ' Сбрасываем Ctrl+Shift+M
    Application.OnKey "^d"   ' Сбрасываем Ctrl+D
    Application.OnKey "^e"   ' Сбрасываем Ctrl+E
    
    On Error GoTo 0
End Sub

' Событие открытия книги
Private Sub App_WorkbookOpen(ByVal wb As Workbook)
    ' Можно добавить дополнительную логику при открытии книги
    If wb Is ThisWorkbook Then
        ' Активируем систему при открытии нашей книги
        Call TableEventHandler.InitializeTableEvents
    End If
End Sub

