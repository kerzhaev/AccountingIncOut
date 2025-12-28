Attribute VB_Name = "ProgressManager"
' ИСПРАВЛЕННЫЙ модуль ProgressManager
'==============================================
' МОДУЛЬ УПРАВЛЕНИЯ ПРОГРЕССОМ - ProgressManager
' Назначение: Отображение прогресса длительных операций
' Состояние: ИСПРАВЛЕНО - переименован UpdateStatusBar
' Версия: 1.0.1
' Дата: 23.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

Private ProgressForm As UserFormProgress
Private operationCancelled As Boolean
Private StartTime As Double

' Показать прогресс-бар
Public Sub ShowProgress(operationTitle As String, Optional allowCancel As Boolean = True)
    Set ProgressForm = New UserFormProgress
    
    With ProgressForm
        .Caption = "Выполнение операции"
        .lblOperation.Caption = operationTitle
        .ProgressBar.value = 0
        .btnCancel.Visible = allowCancel
        .Show vbModeless
    End With
    
    operationCancelled = False
    StartTime = Timer
    
    ' Обновляем статус Excel
    Application.StatusBar = operationTitle & " - 0%"
End Sub

' Обновить прогресс
Public Sub UpdateProgress(CurrentValue As Long, MaxValue As Long, Optional statusText As String = "")
    Dim percentage As Double
    
    If ProgressForm Is Nothing Then Exit Sub
    
    percentage = (CurrentValue / MaxValue) * 100
    
    With ProgressForm
        .ProgressBar.value = percentage
        .lblStatus.Caption = "Выполнено: " & CurrentValue & " из " & MaxValue & " (" & Format(percentage, "0.0") & "%)"
        
        If statusText <> "" Then
            .lblDetails.Caption = statusText
        End If
        
        ' Обновляем время выполнения
        .lblTime.Caption = "Время: " & Format(Timer - StartTime, "0.0") & " сек."
    End With
    
    ' Обновляем статус Excel
    Application.StatusBar = ProgressForm.lblOperation.Caption & " - " & Format(percentage, "0.0") & "%"
    
    ' Обрабатываем события Windows
    DoEvents
End Sub

' Скрыть прогресс-бар
Public Sub HideProgress()
    If Not ProgressForm Is Nothing Then
        Unload ProgressForm
        Set ProgressForm = Nothing
    End If
    
    Application.StatusBar = False
End Sub

' Проверка отмены операции
Public Function IsOperationCancelled() As Boolean
    IsOperationCancelled = operationCancelled
End Function

' Отмена операции (вызывается из формы)
Public Sub CancelOperation()
    operationCancelled = True
    Call HideProgress
    Application.StatusBar = "Операция отменена пользователем"
End Sub

' ИСПРАВЛЕНО: Переименованная функция быстрого обновления статус-бара
Public Sub UpdateProgressStatusBar(Message As String, Optional percentage As Double = -1)
    If percentage >= 0 Then
        Application.StatusBar = Message & " - " & Format(percentage, "0.0") & "%"
    Else
        Application.StatusBar = Message
    End If
End Sub

