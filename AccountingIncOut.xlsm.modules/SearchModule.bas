Attribute VB_Name = "SearchModule"
'==============================================
' МОДУЛЬ ПОИСКА ПО ТАБЛИЦЕ SearchModule
' Назначение: Реализация мгновенного поиска по всем столбцам таблицы
' Состояние: ДОБАВЛЕНО АВТОМАТИЧЕСКОЕ РАСШИРЕНИЕ ШИРИНЫ ПОЛЯ ПО СОДЕРЖИМОМУ
' Версия: 1.7.2
' Дата: 09.08.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' Глобальные переменные для сохранения состояния поиска
Private LastSearchText As String
Private SearchResultsVisible As Boolean

Public Sub PerformSearch()
    Dim SearchText As String
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long, j As Long
    Dim cellValue As String
    Dim FoundCount As Integer
    
    SearchText = Trim(UCase(UserFormVhIsh.txtSearch.Text))
    
    ' Сохраняем текст поиска для возможности восстановления
    LastSearchText = SearchText
    
    ' Очистка результатов если поиск пустой
    If Len(SearchText) = 0 Then
        With UserFormVhIsh
            .lstSearchResults.Clear
            .lstSearchResults.Visible = False
            .lblStatusBar.Caption = "Введите текст для поиска"
        End With
        SearchResultsVisible = False
        Exit Sub
    End If
    
    On Error GoTo SearchError
    
    Set wsData = ThisWorkbook.Worksheets("ВхИсх")
    Set tblData = wsData.ListObjects("ВходящиеИсходящие")
    
    UserFormVhIsh.lstSearchResults.Clear
    
    ' Поиск по всем строкам и столбцам таблицы
    If tblData.ListRows.Count > 0 Then
        For i = 1 To tblData.ListRows.Count
            Dim hasMatch As Boolean
            hasMatch = False
            
            For j = 1 To tblData.ListColumns.Count
                cellValue = UCase(CStr(tblData.DataBodyRange.Cells(i, j).value))
                
                If InStr(cellValue, SearchText) > 0 Then
                    hasMatch = True
                    Exit For
                End If
            Next j
            
            If hasMatch Then
                Call AddImprovedSearchResult(i, tblData)
                FoundCount = FoundCount + 1
            End If
            
            ' Ограничение количества результатов
            If FoundCount >= 25 Then Exit For
        Next i
    End If
    
    ' Отображение результатов
    Call DisplaySearchResults(FoundCount)
    
    Exit Sub
    
SearchError:
    UserFormVhIsh.lblStatusBar.Caption = "Ошибка поиска: " & Err.description
    UserFormVhIsh.lstSearchResults.Visible = False
    SearchResultsVisible = False
End Sub

' УЛУЧШЕННАЯ ПРОЦЕДУРА: Добавление результата поиска с нумерацией
' Добавлены поля: От кого поступил (9), Исполнитель (11), Статус подтверждения (19)
' Убрано дублирование номера строки в конце записи (скрытый ключ хранится отдельно)
Private Sub AddImprovedSearchResult(rowNum As Long, tbl As ListObject)
    Dim ResultText As String
    Dim serviceText As String
    Dim docTypeInOut As String
    Dim docTypeText As String
    Dim docNumberText As String
    Dim amountText As String
    Dim frpText As String
    Dim DateText As String
    Dim fromWhomText As String
    Dim executorText As String
    Dim statusText As String
    
    On Error GoTo AddError
    
    ' Основные данные
    serviceText = CStr(tbl.DataBodyRange.Cells(rowNum, 2).value)     ' Служба
    docTypeInOut = CStr(tbl.DataBodyRange.Cells(rowNum, 3).value)    ' Вид документа (Вх./Исх.)
    docTypeText = CStr(tbl.DataBodyRange.Cells(rowNum, 4).value)     ' Тип документа
    docNumberText = CStr(tbl.DataBodyRange.Cells(rowNum, 5).value)   ' Номер документа
    amountText = CStr(tbl.DataBodyRange.Cells(rowNum, 6).value)      ' Сумма
    frpText = CStr(tbl.DataBodyRange.Cells(rowNum, 7).value)         ' Номер ФРП
    
    ' Дата (8)
    If IsDate(tbl.DataBodyRange.Cells(rowNum, 8).value) Then
        DateText = Format(tbl.DataBodyRange.Cells(rowNum, 8).value, "dd.mm.yy")
    Else
        DateText = CStr(tbl.DataBodyRange.Cells(rowNum, 8).value)
    End If
    
    ' Новые поля
    fromWhomText = CStr(tbl.DataBodyRange.Cells(rowNum, 9).value)    ' От кого поступил
    executorText = CStr(tbl.DataBodyRange.Cells(rowNum, 11).value)   ' Исполнитель
    statusText = CStr(tbl.DataBodyRange.Cells(rowNum, 19).value)     ' Статус подтверждения
    
    ' Формирование строки результата
    ' Старт с ">" и номером строки
    ResultText = ">" & rowNum & ": "
    
    If Trim(serviceText) <> "" Then
        ResultText = ResultText & "[" & serviceText & "] "
    End If
    
    If Trim(docTypeInOut) <> "" Then
        ResultText = ResultText & docTypeInOut & " "
    End If
    
    If Trim(docTypeText) <> "" Then
        ResultText = ResultText & docTypeText & " "
    End If
    
    If Trim(docNumberText) <> "" Then
        ResultText = ResultText & "№" & docNumberText & " "
    End If
    
    If Trim(amountText) <> "" And amountText <> "0" Then
        ResultText = ResultText & "(" & amountText & "р.) "
    End If
    
    If Trim(frpText) <> "" Then
        ResultText = ResultText & "ФРП:" & frpText & " "
    End If
    
    If Trim(DateText) <> "" Then
        ResultText = ResultText & "от " & DateText & " "
    End If
    
    ' Добавлены поля в хвост результата:
    If Trim(fromWhomText) <> "" Then
        ResultText = ResultText & "| От кого: " & fromWhomText & " "
    End If
    
    If Trim(executorText) <> "" Then
        ResultText = ResultText & "| Исп.: " & executorText & " "
    End If
    
    If Trim(statusText) <> "" Then
        ResultText = ResultText & "| Статус: " & statusText & " "
    End If
    
    ' ИЗМЕНЕНО: Не обрезаем строку для корректного расчета ширины
    ' If Len(resultText) > 120 Then
    '     resultText = Left(resultText, 117) & "..."
    ' End If
    
    ' Добавление в список
    With UserFormVhIsh.lstSearchResults
        .AddItem ResultText
        ' Сохраняем скрытый идентификатор записи через разделитель,
        ' но НЕ добавляем его визуально в текст (нет дублирования номера)
        If .ListCount > 0 Then
            .List(.ListCount - 1, 0) = ResultText & Chr(1) & CStr(rowNum)
        End If
    End With
    
    Exit Sub
    
AddError:
    ' Пропускаем ошибки добавления результатов
End Sub

' ОБНОВЛЕНА ПРОЦЕДУРА: Отображение результатов с автоматическим расширением ширины
Private Sub DisplaySearchResults(FoundCount As Integer)
    With UserFormVhIsh
        If .lstSearchResults.ListCount > 0 Then
            .lstSearchResults.Visible = True
            SearchResultsVisible = True
            
            ' НОВОЕ: Автоматическое вычисление ширины по содержимому
            Call AutoResizeSearchResultsWidth
            
            ' Фиксированная высота остается
            .lstSearchResults.Height = 120
            
            ' Статус-бар
            .lblStatusBar.Caption = "Найдено: " & FoundCount & " записей | " & _
                                   "Навигация: ^v или щелчок для перехода | " & _
                                   "Поиск остается активным"
        Else
            .lstSearchResults.Visible = False
            SearchResultsVisible = False
            .lblStatusBar.Caption = "По запросу '" & .txtSearch.Text & "' ничего не найдено"
        End If
    End With
End Sub

' НОВАЯ ПРОЦЕДУРА: Автоматическое изменение размера поля по содержимому
Private Sub AutoResizeSearchResultsWidth()
    Dim maxWidth As Single
    Dim textWidth As Single
    Dim i As Long
    Dim testText As String
    Dim minWidth As Single
    Dim maxAllowedWidth As Single
    
    On Error GoTo ResizeError
    
    ' Минимальная и максимальная ширина
    minWidth = 420 ' Минимальная ширина
    maxAllowedWidth = 800 ' Максимальная ширина (чтобы не выходить за границы формы)
    
    maxWidth = minWidth
    
    With UserFormVhIsh.lstSearchResults
        ' Проходим по всем элементам списка для поиска самого длинного
        For i = 0 To .ListCount - 1
            ' Получаем текст без скрытой части (до разделителя Chr(1))
            testText = .List(i, 0)
            If InStr(testText, Chr(1)) > 0 Then
                testText = Left(testText, InStr(testText, Chr(1)) - 1)
            End If
            
            ' НОВОЕ: Вычисляем ширину текста с учетом шрифта
            textWidth = CalculateTextWidth(testText, .Font.Name, .Font.Size)
            
            ' Обновляем максимальную ширину с небольшим отступом
            If textWidth + 20 > maxWidth Then
                maxWidth = textWidth + 20
            End If
        Next i
        
        ' Ограничиваем ширину разумными пределами
        If maxWidth < minWidth Then
            maxWidth = minWidth
        ElseIf maxWidth > maxAllowedWidth Then
            maxWidth = maxAllowedWidth
        End If
        
        ' Применяем новую ширину
        .Width = maxWidth
        
        ' НОВОЕ: Обновляем ColumnWidths для корректного отображения
        .columnWidths = CStr(maxWidth - 10)
    End With
    
    Exit Sub
    
ResizeError:
    ' При ошибке используем стандартную ширину
    UserFormVhIsh.lstSearchResults.Width = minWidth
End Sub

' НОВАЯ ФУНКЦИЯ: Вычисление ширины текста
Private Function CalculateTextWidth(Text As String, fontName As String, fontSize As Single) As Single
    Dim textLength As Long
    Dim avgCharWidth As Single
    
    On Error GoTo WidthError
    
    textLength = Len(Text)
    
    ' Приблизительная ширина символа в пикселях в зависимости от размера шрифта
    ' Для Segoe UI средняя ширина символа составляет примерно 0.6 от высоты шрифта
    avgCharWidth = fontSize * 0.6
    
    ' Вычисляем общую ширину с небольшим запасом
    CalculateTextWidth = textLength * avgCharWidth
    
    Exit Function
    
WidthError:
    ' При ошибке возвращаем базовую ширину
    CalculateTextWidth = 420
End Function

' КАРДИНАЛЬНО ПЕРЕРАБОТАННАЯ ПРОЦЕДУРА: SelectSearchResult
Public Sub SelectSearchResult()
    Dim selectedRow As Long
    Dim resultData As String
    Dim parts As Variant
    
    On Error GoTo SelectError
    
    With UserFormVhIsh.lstSearchResults
        If .listIndex >= 0 Then
            ' Извлекаем номер строки из скрытых данных элемента
            resultData = .List(.listIndex, 0)
            parts = Split(resultData, Chr(1))
            
            If UBound(parts) >= 1 Then
                selectedRow = CLng(parts(1))
                
                ' Переходим к найденной записи
                Call NavigateToRecord(selectedRow)
                
                ' КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: НЕ очищаем поиск и НЕ скрываем результаты!
                ' Вместо этого подсвечиваем выбранный элемент и обновляем статус
                
                ' Обновляем статус-бар с информацией о выбранной записи
                UserFormVhIsh.lblStatusBar.Caption = "Переход к записи №" & selectedRow & " | " & _
                                                    "Найдено: " & .ListCount & " записей | " & _
                                                    "Поиск активен: """ & UserFormVhIsh.txtSearch.Text & """"
                
                ' Визуально выделяем выбранный элемент (остается выделенным)
                .BackColor = RGB(240, 248, 255) ' Светло-голубой фон для активного поиска
            End If
        End If
    End With
    
    Exit Sub
    
SelectError:
    MsgBox "Ошибка выбора результата поиска: " & Err.description, vbExclamation, "Ошибка"
End Sub

' НОВАЯ ПРОЦЕДУРА: Очистка поиска по требованию пользователя
Public Sub ClearSearch()
    With UserFormVhIsh
        .txtSearch.Text = ""
        .lstSearchResults.Clear
        .lstSearchResults.Visible = False
        .lblStatusBar.Caption = "Поиск очищен"
    End With
    LastSearchText = ""
    SearchResultsVisible = False
End Sub

' НОВАЯ ПРОЦЕДУРА: Восстановление последнего поиска
Public Sub RestoreLastSearch()
    If LastSearchText <> "" Then
        UserFormVhIsh.txtSearch.Text = LastSearchText
        Call PerformSearch
    End If
End Sub

' НОВАЯ ПРОЦЕДУРА: Навигация по результатам поиска с клавиатуры
Public Sub NavigateSearchResults(direction As String)
    With UserFormVhIsh.lstSearchResults
        If .Visible And .ListCount > 0 Then
            Select Case UCase(direction)
                Case "UP"
                    If .listIndex > 0 Then
                        .listIndex = .listIndex - 1
                    Else
                        .listIndex = .ListCount - 1 ' Переход к последнему элементу
                    End If
                Case "DOWN"
                    If .listIndex < .ListCount - 1 Then
                        .listIndex = .listIndex + 1
                    Else
                        .listIndex = 0 ' Переход к первому элементу
                    End If
                Case "FIRST"
                    .listIndex = 0
                Case "LAST"
                    .listIndex = .ListCount - 1
            End Select
            
            ' Автоматически переходим к выбранной записи
            Call SelectSearchResult
        End If
    End With
End Sub

' НОВАЯ ПРОЦЕДУРА: Проверка активности поиска
Public Function IsSearchActive() As Boolean
    IsSearchActive = SearchResultsVisible And (UserFormVhIsh.lstSearchResults.ListCount > 0)
End Function

' НОВАЯ ПРОЦЕДУРА: Получение информации о текущем поиске
Public Function GetSearchInfo() As String
    If IsSearchActive() Then
        With UserFormVhIsh.lstSearchResults
            GetSearchInfo = "Активный поиск: """ & UserFormVhIsh.txtSearch.Text & """ | " & _
                           "Найдено: " & .ListCount & " записей | " & _
                           "Выбрана: " & (.listIndex + 1) & " из " & .ListCount
        End With
    Else
        GetSearchInfo = "Поиск неактивен"
    End If
End Function

' НОВАЯ ПРОЦЕДУРА: Подсветка найденного текста в форме
Public Sub HighlightSearchTermInForm()
    ' Эта процедура может быть расширена для подсветки найденного текста в полях формы
    ' Пока оставляем как заглушку для будущего развития
    Dim searchTerm As String
    searchTerm = UserFormVhIsh.txtSearch.Text
    
    If Len(searchTerm) > 0 Then
        ' Здесь можно добавить логику подсветки найденного текста в полях формы
        ' Например, изменение цвета фона полей, содержащих искомый текст
    End If
End Sub


