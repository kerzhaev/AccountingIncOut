Attribute VB_Name = "DoverennostMatchingModule"
'==============================================
' DoverennostMatchingModule - ТРЕХПРОХОДНАЯ СИСТЕМА
' Версия: 1.2.2
' Дата: 07.09.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================

Option Explicit

' =============================================
' СТРУКТУРЫ ДАННЫХ
' =============================================

Public Type ParsedNaryad
    OriginalText As String
    DocumentType As String
    DocumentNumber As String
    DocumentDate As String
    IsValid As Boolean
    MatchDetails As String
End Type

Public Type MatchResult
    FoundMatch As Boolean
    MatchScore As Double
    OperationRow As Long
    OperationNumber As String
    OperationDate As Date
    MatchDetails As String
    PassNumber As Long
End Type

' =============================================
' СТРУКТУРЫ ДЛЯ УМНОГО ПОИСКА
' =============================================

Public Type ComponentMatchResult
    IsMatch As Boolean
    Confidence As Double
    MatchDetails As String
    FoundNumber As String
    FoundDate As String
    NumberPosition As Long
    DatePosition As Long
End Type

Public Type NumberCandidate
    Text As String
    Position As Long
    Length As Long
    ConfidenceScore As Double
    NumberType As String ' "SIMPLE", "SLASH", "DASH", "MILITARY", "COMPLEX"
End Type

Public Type DateCandidate
    Text As String
    Position As Long
    Length As Long
    ParsedDate As Date
    DateFormat As String ' "FULL", "SHORT", "PARTIAL"
End Type
' =============================================
' ОСНОВНАЯ ПРОЦЕДУРА
' =============================================

Public Sub ProcessDoverennostMatching()
    Dim FileDoverennosti As String
    Dim FileOperations As String
    Dim WbDoverennosti As Workbook
    Dim WbOperations As Workbook
    
    On Error GoTo ProcessError
    
    Debug.Print "=== НАЧАЛО ТРЕХПРОХОДНОЙ СИСТЕМЫ С ИСПРАВЛЕННЫМ REGEX ==="
    
    FileDoverennosti = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        "Выберите файл журнала доверенностей")
    If FileDoverennosti = "False" Then Exit Sub
    
    FileOperations = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        "Выберите файл журнала операций 1С")
    If FileOperations = "False" Then Exit Sub
    
    Set WbDoverennosti = Workbooks.Open(FileDoverennosti, ReadOnly:=False)
    Set WbOperations = Workbooks.Open(FileOperations, ReadOnly:=True)
    
    Call PerformThreePassMatching(WbDoverennosti, WbOperations)
    
    WbOperations.Close False
    WbDoverennosti.Save
    WbDoverennosti.Close
    
    Debug.Print "=== ТРЕХПРОХОДНАЯ СИСТЕМА С ИСПРАВЛЕННЫМ REGEX ЗАВЕРШЕНА ==="
    
    Exit Sub
    
ProcessError:
    On Error Resume Next
    If Not WbOperations Is Nothing Then WbOperations.Close False
    If Not WbDoverennosti Is Nothing Then WbDoverennosti.Close False
    
    Debug.Print "Ошибка: " & Err.description
    MsgBox "Ошибка: " & Err.description, vbCritical
End Sub

' =============================================
' ТРЕХПРОХОДНАЯ ФУНКЦИЯ СОПОСТАВЛЕНИЯ
' =============================================

Private Sub PerformThreePassMatching(WbDover As Workbook, WbOper As Workbook)
    Dim WsDover As Worksheet, WsOper As Worksheet
    Dim LastRowDover As Long, LastRowOper As Long
    Dim i As Long
    Dim ProcessedCount As Long, SkippedEmpty As Long, SkippedFound As Long

    
    ' Статистика по проходам
    Dim Pass1Found As Long, Pass2Found As Long, Pass3Found As Long
    Dim Pass3ByType(4) As Long
    
    Dim StartTime As Double
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    StartTime = Timer
    
    On Error GoTo MatchingError
    
    Set WsDover = WbDover.Worksheets(1)
    Set WsOper = WbOper.Worksheets(1)
    
    LastRowDover = WsDover.Cells(WsDover.Rows.Count, 1).End(xlUp).Row
    LastRowOper = WsOper.Cells(WsOper.Rows.Count, 1).End(xlUp).Row
    
    Dim DoverData As Variant
    DoverData = WsDover.Range("A2:I" & LastRowDover).Value2
    
    Dim OperData As Variant
    OperData = WsOper.Range("A2:I" & LastRowOper).Value2
    
    ' Фильтруем операции
    Dim FilteredOperData As Variant
    Dim FilteredCount As Long
    FilteredOperData = FilterOperationsByCorrespondent(OperData, FilteredCount, SkippedEmpty)
    
    Call AddThreePassResultColumns(WsDover)
    
    ' Определяем номер столбца с результатами для проверки уже найденных
    Dim ResultColumn As Long
    ResultColumn = FindResultColumn(WsDover)
    
    Debug.Print "ТРЕХПРОХОДНАЯ СИСТЕМА С ИСПРАВЛЕННЫМ REGEX: " & (LastRowDover - 1) & " доверенностей против " & FilteredCount & " операций"

    Debug.Print "Пропущено " & SkippedEmpty & " операций с пустым корреспондентом"
    
    ' Массивы для хранения не найденных записей
    Dim Pass1NotFound() As Long, Pass1NotFoundData() As ParsedNaryad
    Dim Pass2NotFound() As Long, Pass2NotFoundData() As ParsedNaryad
    Dim Pass1Count As Long, Pass2Count As Long
    
    Pass1Count = 0
    Pass2Count = 0
    
    Dim DoverComment As String
    Dim ParsedDover As ParsedNaryad
    Dim MatchResult As MatchResult
    Dim CurrentRow As Long
    
    ' ОСНОВНОЙ ЦИКЛ ОБРАБОТКИ ДОВЕРЕННОСТЕЙ
    For i = 1 To UBound(DoverData, 1)
        CurrentRow = i + 1
        DoverComment = Trim(CStr(DoverData(i, 9)))
        
        If i Mod 25 = 0 Then
            Application.StatusBar = "3-проходная система с исправленным REGEX: " & i & " из " & UBound(DoverData, 1)
        End If
        
        ' Проверяем, была ли строка уже обработана с результатом "НАЙДЕНО"
        If IsAlreadyProcessed(WsDover, CurrentRow, ResultColumn) Then
            SkippedFound = SkippedFound + 1
            Debug.Print "Пропуск строки " & CurrentRow & " - уже найдено"
            GoTo NextIteration
        End If
        
        If DoverComment <> "" Then

            
            ' === 1-Й ПРОХОД: ТОЧНЫЕ НАРЯДЫ ===
            ParsedDover = ParseNaryadForSubstring(DoverComment)
            
            If ParsedDover.IsValid And UCase(ParsedDover.DocumentType) = "НАРЯД" Then
                MatchResult = FindBySubstringEnhanced(ParsedDover, FilteredOperData)
                MatchResult.PassNumber = 1
                
                If MatchResult.FoundMatch Then
                    Pass1Found = Pass1Found + 1
                    MatchResult.MatchDetails = "1-й проход: " & MatchResult.MatchDetails
                Else
                    ' Сохраняем для 2-го прохода
                    Pass1Count = Pass1Count + 1
                    ReDim Preserve Pass1NotFound(Pass1Count - 1)
                    ReDim Preserve Pass1NotFoundData(Pass1Count - 1)
                    Pass1NotFound(Pass1Count - 1) = CurrentRow
                    Pass1NotFoundData(Pass1Count - 1) = ParsedDover
                    MatchResult.MatchDetails = "Ждет 2-й проход (исправленный REGEX)"
                End If
            Else
                ' Если не наряд или не распарсился - сохраняем для 3-го прохода
                Pass2Count = Pass2Count + 1
                ReDim Preserve Pass2NotFound(Pass2Count - 1)
                ReDim Preserve Pass2NotFoundData(Pass2Count - 1)
                Pass2NotFound(Pass2Count - 1) = CurrentRow
                Pass2NotFoundData(Pass2Count - 1).OriginalText = DoverComment
                Pass2NotFoundData(Pass2Count - 1).IsValid = False
                
                MatchResult = CreateEmptyMatchResult("Ждет 3-й проход")
                MatchResult.PassNumber = 0
            End If
            
            ProcessedCount = ProcessedCount + 1
        Else
            MatchResult = CreateEmptyMatchResult("Пустой комментарий")
            MatchResult.PassNumber = 0
        End If
        
        Call WriteThreePassResult(WsDover, CurrentRow, MatchResult)
        
NextIteration:
        Next i

    
    ' === 2-Й ПРОХОД: ИСПРАВЛЕННЫЙ АДАПТИВНЫЙ REGEX ===
    If Pass1Count > 0 Then
        Debug.Print "=== 2-Й ПРОХОД: ИСПРАВЛЕННЫЙ АДАПТИВНЫЙ REGEX ==="
'Pass2Found = ProcessSecondPassWithSmartComponent(WsDover, FilteredOperData, Pass1NotFound, Pass1NotFoundData, Pass1Count, Pass2NotFound, Pass2NotFoundData, Pass2Count)

Pass2Found = ProcessSecondPassWithSupplierService(WsDover, FilteredOperData, Pass1NotFound, Pass1NotFoundData, Pass1Count, Pass2NotFound, Pass2NotFoundData, Pass2Count)

    End If
    
    ' === 3-Й ПРОХОД: УНИВЕРСАЛЬНЫЕ ДОКУМЕНТЫ ===
    If Pass2Count > 0 Then
        Debug.Print "=== 3-Й ПРОХОД: УНИВЕРСАЛЬНЫЕ ДОКУМЕНТЫ ==="
        Pass3Found = ProcessThirdPass(WsDover, FilteredOperData, Pass2NotFound, Pass2NotFoundData, Pass2Count, Pass3ByType)
        Debug.Print "3-й проход нашел: " & Pass3Found
        Debug.Print "  Разнарядки: " & Pass3ByType(0)
        Debug.Print "  Телеграммы: " & Pass3ByType(1)
        Debug.Print "  Аттестаты: " & Pass3ByType(2)
        Debug.Print "  Заявки: " & Pass3ByType(3)
    End If
    
    ' Восстанавливаем настройки
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    Dim TotalTime As Double
    TotalTime = Timer - StartTime
    
    Debug.Print "=== РЕЗУЛЬТАТЫ ТРЕХПРОХОДНОЙ СИСТЕМЫ С ИСПРАВЛЕННЫМ REGEX ==="
    Debug.Print "Время выполнения: " & Format(TotalTime, "0.0") & " сек"
    Debug.Print "Обработано: " & ProcessedCount
    Debug.Print "Пропущено уже найденных: " & SkippedFound
    Debug.Print "1-й проход (точные наряды): " & Pass1Found

    Debug.Print "2-й проход (исправленный REGEX наряды): " & Pass2Found
    Debug.Print "3-й проход (другие документы): " & Pass3Found
    Debug.Print "Всего найдено: " & (Pass1Found + Pass2Found + Pass3Found) & " (" & Format((Pass1Found + Pass2Found + Pass3Found) / ProcessedCount * 100, "0.0") & "%)"
    Debug.Print "Пропущено операций с пустым корреспондентом: " & SkippedEmpty
    
    Call ShowThreePassStatistics(ProcessedCount, Pass1Found, Pass2Found, Pass3Found, Pass3ByType, SkippedEmpty, SkippedFound)

    
    Application.Calculate
    WsDover.Calculate
    
    Exit Sub
    
MatchingError:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    Debug.Print "Ошибка трехпроходной системы: " & Err.description
    MsgBox "Ошибка: " & Err.description, vbCritical
End Sub

' =============================================
' 2-Й ПРОХОД: ИСПРАВЛЕННЫЙ АДАПТИВНЫЙ REGEX
' =============================================

' ИСПРАВЛЕННАЯ ФУНКЦИЯ: 2-й проход с улучшенным regex
' НОВЫЙ 2-Й ПРОХОД: УМНЫЙ ПОИСК ПО КОМПОНЕНТАМ
Private Function ProcessSecondPassWithSmartComponent(WsDover As Worksheet, OperData As Variant, NotFoundRows() As Long, NotFoundData() As ParsedNaryad, NotFoundCount As Long, ByRef Pass2NotFound() As Long, ByRef Pass2NotFoundData() As ParsedNaryad, ByRef Pass2Count As Long) As Long
    Dim i As Long
    Dim SecondPassFound As Long
    Dim MatchResult As MatchResult
    
    SecondPassFound = 0
    
    Debug.Print "=== 2-Й ПРОХОД: УМНЫЙ ПОИСК ПО КОМПОНЕНТАМ ==="
    
    For i = 0 To NotFoundCount - 1
        Debug.Print "2-й проход УМНЫЙ для: " & NotFoundData(i).DocumentNumber & " от " & NotFoundData(i).DocumentDate
        
        ' ИСПОЛЬЗУЕМ УМНЫЙ ПОИСК ПО КОМПОНЕНТАМ!
        MatchResult = FindBySmartComponents(NotFoundData(i), OperData)
        MatchResult.PassNumber = 2
        
        If MatchResult.FoundMatch Then
            SecondPassFound = SecondPassFound + 1
            MatchResult.MatchDetails = "2-й проход умный (" & Format(MatchResult.MatchScore, "0") & "%): " & MatchResult.MatchDetails
            
            Call WriteThreePassResult(WsDover, NotFoundRows(i), MatchResult)
            
            Debug.Print "  ? Найдено 2-м проходом УМНЫЙ! Уверенность: " & MatchResult.MatchScore & "%"
        Else
            ' Сохраняем для 3-го прохода
            Pass2Count = Pass2Count + 1
            ReDim Preserve Pass2NotFound(Pass2Count - 1)
            ReDim Preserve Pass2NotFoundData(Pass2Count - 1)
            Pass2NotFound(Pass2Count - 1) = NotFoundRows(i)
            Pass2NotFoundData(Pass2Count - 1) = NotFoundData(i)
            
            Debug.Print "  ? Не найдено умным поиском, переходит в 3-й"
        End If
    Next i
    
    ProcessSecondPassWithSmartComponent = SecondPassFound
End Function


' ОСНОВНАЯ ИСПРАВЛЕННАЯ ФУНКЦИЯ АДАПТИВНОГО REGEX ПОИСКА
' ОСНОВНАЯ ФУНКЦИЯ УМНОГО ПОИСКА ПО КОМПОНЕНТАМ
Private Function FindBySmartComponents(DoverData As ParsedNaryad, OperData As Variant) As MatchResult
    Dim Result As MatchResult
    Dim i As Long
    Dim OperComment As String
    
    Result.FoundMatch = False
    Result.MatchScore = 0
    
'    Debug.Print ">>> УМНЫЙ ПОИСК КОМПОНЕНТОВ ДЛЯ: № '" & DoverData.DocumentNumber & "' от '" & DoverData.DocumentDate & "'"
    
    ' Проверяем каждую операцию
    For i = 1 To UBound(OperData, 1)
        OperComment = Trim(CStr(OperData(i, 9)))
        
        If OperComment <> "" Then
'            Debug.Print "    Анализируем операцию " & (i + 1) & ": " & Left(OperComment, 120) & "..."
            
            ' КЛЮЧЕВАЯ ФУНКЦИЯ: Умный анализ компонентов
            Dim ComponentMatch As ComponentMatchResult
            ComponentMatch = AnalyzeComponents(OperComment, DoverData.DocumentNumber, DoverData.DocumentDate)
            
            ' НОВАЯ ПРОВЕРКА: Добавляем проверку корреспондента!
            If ComponentMatch.IsMatch And ComponentMatch.Confidence >= 25 Then ' Снижен порог с 40 до 25
                Dim CorrespondentBonus As Double
                CorrespondentBonus = CheckCorrespondentMatch(OperData(i, 6), DoverData.OriginalText, i + 1)
                ComponentMatch.Confidence = ComponentMatch.Confidence + CorrespondentBonus
                
                ' Обновляем детали с учетом корреспондента
                If CorrespondentBonus > 0 Then
                    ComponentMatch.MatchDetails = ComponentMatch.MatchDetails & " корр:+" & Format(CorrespondentBonus, "0") & "%"
                End If
                
                ' Финальная проверка с учетом корреспондента
                If ComponentMatch.Confidence >= 40 Then
                    Result.FoundMatch = True
                    Result.MatchScore = ComponentMatch.Confidence
                    Result.OperationRow = i + 1
                    Result.OperationNumber = Trim(CStr(OperData(i, 3)))
                    
                    ' НОВОЕ: Добавляем комментарий из 1С в детали
                    ComponentMatch.MatchDetails = ComponentMatch.MatchDetails & " | 1С: " & OperComment

                    Result.MatchDetails = ComponentMatch.MatchDetails
                    
                    On Error Resume Next
                    Result.OperationDate = CDate(OperData(i, 2))
                    On Error GoTo 0
                    
'                    Debug.Print ">>> ? НАЙДЕНО УМНЫМ ПОИСКОМ! Строка " & (i + 1) & " Уверенность: " & ComponentMatch.Confidence & "%"
'                    Debug.Print "    Детали: " & ComponentMatch.MatchDetails
                    Exit For
                End If
            End If
        End If
    Next i
    
    If Not Result.FoundMatch Then
        Result.MatchDetails = "умный поиск не нашел подходящих компонентов"
'        Debug.Print ">>> ? УМНЫЙ ПОИСК НЕ НАШЕЛ ПОДХОДЯЩИХ КОМПОНЕНТОВ"
    End If
    
    FindBySmartComponents = Result
End Function



' НОВАЯ ФУНКЦИЯ: Создание гибких паттернов для поиска внутри текста

Private Sub CreateFlexibleTextPatterns(DocumentNumber As String, DocumentDate As String, ByRef Patterns() As String)
    Dim PatternCount As Long
    PatternCount = 0
    
    ReDim Patterns(15) ' Уменьшаем до 15 более точных паттернов
    
    Debug.Print "    Создание УПРОЩЕННЫХ паттернов для: № '" & DocumentNumber & "' дата '" & DocumentDate & "'"
    
    ' Подготавливаем компоненты без агрессивного экранирования
    Dim SimpleNumber As String
    Dim SimpleDatePatterns() As String
    
    SimpleNumber = EscapeMinimal(DocumentNumber)
    Call CreateSimpleDatePatterns(DocumentDate, SimpleDatePatterns)
    
    Debug.Print "    Простой номер: " & SimpleNumber
    Debug.Print "    Создано вариантов дат: " & (UBound(SimpleDatePatterns) + 1)
    
    ' ========================================
    ' ГРУППА 1: ОСНОВНЫЕ ПАТТЕРНЫ (УПРОЩЕННЫЕ)
    ' ========================================
    
    Dim DateVar As Long
    For DateVar = 0 To UBound(SimpleDatePatterns)
        If PatternCount >= 15 Then Exit For
        
        ' Основной паттерн - очень простой
        Patterns(PatternCount) = "наряд\s*№?\s*" & SimpleNumber & "\s*от\s*" & SimpleDatePatterns(DateVar)
        PatternCount = PatternCount + 1
    Next DateVar
    
    ' ========================================
    ' ГРУППА 2: СПЕЦИАЛЬНЫЕ ПАТТЕРНЫ ДЛЯ КОНКРЕТНЫХ СЛУЧАЕВ
    ' ========================================
    
    ' Паттерн для номеров со слешами (901/243450, 36/1, 565/4/24/2213)
    If InStr(DocumentNumber, "/") > 0 And PatternCount < 14 Then
        Patterns(PatternCount) = CreateSlashPattern(DocumentNumber, DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    ' Паттерн для номеров с дефисами (561/22/176-а-25, ПС-4, 1202-П, 5/ЮВО/16-С)
    If InStr(DocumentNumber, "-") > 0 And PatternCount < 14 Then
        Patterns(PatternCount) = CreateDashPattern(DocumentNumber, DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    ' Паттерн для военных номеров (60/ЮВО, 46/3/32Б, 46/3/919БТ, 197/ЮВО565/ЗЧ-556)
    If ContainsMilitaryCode(DocumentNumber) And PatternCount < 14 Then
        Patterns(PatternCount) = CreateMilitaryPattern(DocumentNumber, DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    ' Паттерн для цифровых номеров (2007, 15030018)
    If IsNumericOnly(DocumentNumber) And PatternCount < 14 Then
        Patterns(PatternCount) = CreateNumericPattern(DocumentNumber, DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    ' ========================================
    ' ГРУППА 3: РЕЗЕРВНЫЕ ПАТТЕРНЫ
    ' ========================================
    
    If PatternCount < 15 Then
        ' Очень гибкий поиск - игнорируем регистр полностью
        Patterns(PatternCount) = "(?i)наряд.*?" & EscapeMinimal(DocumentNumber) & ".*?от.*?" & EscapeMinimal(DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    ' Обрезаем массив до реального размера
    ReDim Preserve Patterns(PatternCount - 1)
    
    Debug.Print "    Итого создано УПРОЩЕННЫХ паттернов: " & PatternCount
End Sub





' НОВАЯ ФУНКЦИЯ: Поиск паттерна в тексте с возвратом найденного фрагмента
Private Function FindPatternInText(TextToSearch As String, Patterns() As String, ByRef MatchedPattern As String, ByRef MatchedText As String) As Boolean
    Dim RegEx As Object
    Dim Matches As Object
    Dim Match As Object
    
    Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        .IgnoreCase = True
        .Global = True ' Ищем все вхождения
        .Multiline = True
    End With
    
    Dim i As Long
    
    Debug.Print "      Проверяем " & (UBound(Patterns) + 1) & " паттернов в тексте:"
    Debug.Print "      Текст для поиска: " & Left(TextToSearch, 150) & "..."
    
    ' Пробуем каждый паттерн
    For i = 0 To UBound(Patterns)
        RegEx.Pattern = Patterns(i)
        
        Debug.Print "        Паттерн " & (i + 1) & ": " & Left(Patterns(i), 70) & "..."
        
        On Error Resume Next
        Set Matches = RegEx.Execute(TextToSearch)
        
        If Matches.Count > 0 Then
            Set Match = Matches(0) ' Берем первое совпадение
            MatchedPattern = Patterns(i)
            MatchedText = Match.value
            
            Debug.Print "        ? НАЙДЕНО СОВПАДЕНИЕ: '" & MatchedText & "'"
            FindPatternInText = True
            Exit Function
        End If
        On Error GoTo 0
        
        Debug.Print "        ? Не найдено"
    Next i
    
    Debug.Print "        > НИ ОДИН ПАТТЕРН НЕ СРАБОТАЛ"
    FindPatternInText = False
End Function

' НОВАЯ ФУНКЦИЯ: Создание всех вариаций дат
Private Sub CreateSimpleDatePatterns(OriginalDate As String, ByRef DatePatterns() As String)
    Dim PatternCount As Long
    PatternCount = 0
    
    ReDim DatePatterns(5) ' Максимум 6 вариантов
    
    ' Парсим дату
    Dim DateParts() As String
    DateParts = Split(OriginalDate, ".")
    
    If UBound(DateParts) >= 2 Then
        Dim Day As String, Month As String, Year As String
        Day = DateParts(0)
        Month = DateParts(1)
        Year = DateParts(2)
        
        ' Вариант 1: Точная дата
        DatePatterns(PatternCount) = Day & "\." & Month & "\." & Year & "(?:\s*г\.?)?"
        PatternCount = PatternCount + 1
        
        ' Вариант 2: Без ведущих нулей
        DatePatterns(PatternCount) = CLng(Day) & "\." & CLng(Month) & "\." & Year & "(?:\s*г\.?)?"
        PatternCount = PatternCount + 1
        
        ' Вариант 3: Короткий год
        If Len(Year) = 4 Then
            DatePatterns(PatternCount) = CLng(Day) & "\." & CLng(Month) & "\." & Right(Year, 2) & "(?:\s*г\.?)?"
            PatternCount = PatternCount + 1
        End If
        
        ' Вариант 4: Очень гибкий
        DatePatterns(PatternCount) = "\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2})(?:\s*г\.?)?"
        PatternCount = PatternCount + 1
        
    Else
        DatePatterns(PatternCount) = EscapeMinimal(OriginalDate)
        PatternCount = PatternCount + 1
    End If
    
    ReDim Preserve DatePatterns(PatternCount - 1)
End Sub


' УЛУЧШЕННАЯ ФУНКЦИЯ: Умное экранирование regex символов
Private Function EscapeMinimal(Text As String) As String
    Dim Result As String
    Result = Text
    
    ' Экранируем только самые критичные символы для regex
    Result = Replace(Result, ".", "\.")
    Result = Replace(Result, "(", "\(")
    Result = Replace(Result, ")", "\)")
    Result = Replace(Result, "[", "\[")
    Result = Replace(Result, "]", "\]")
    Result = Replace(Result, "*", "\*")
    Result = Replace(Result, "+", "\+")
    Result = Replace(Result, "?", "\?")
    Result = Replace(Result, "^", "\^")
    Result = Replace(Result, "$", "\$")
    Result = Replace(Result, "|", "\|")
    
    ' НЕ экранируем слеши и дефисы - пусть остаются как есть
    
    EscapeMinimal = Result
End Function


' ФУНКЦИЯ ДЛЯ ДОБАВЛЕНИЯ ПОЛЬЗОВАТЕЛЬСКИХ ПАТТЕРНОВ
Public Sub AddCustomPatternForText(ByRef Patterns() As String, ByRef PatternCount As Long, DocumentNumber As String, DocumentDate As String)
    ' ЗДЕСЬ МОЖНО ДОБАВЛЯТЬ СВОИ СПЕЦИАЛЬНЫЕ ПАТТЕРНЫ
    ' Пример:
    ' If PatternCount < UBound(Patterns) Then
    '     Patterns(PatternCount) = "ваш_специальный_паттерн_для_текста"
    '     PatternCount = PatternCount + 1
    ' End If
End Sub

' =============================================
' ПАРСИНГ ДОКУМЕНТОВ (СУЩЕСТВУЮЩИЕ ФУНКЦИИ)
' =============================================

' Парсинг наряда для 1-го прохода
Private Function ParseNaryadForSubstring(Text As String) As ParsedNaryad
    Dim Result As ParsedNaryad
    
    Result.OriginalText = Text
    Result.IsValid = False
    Result.DocumentType = "НАРЯД"
    
    On Error GoTo SubstringParseError
    
    Debug.Print "=== ПАРСИНГ НАРЯДА ДЛЯ ПОДСТРОК ==="
    Debug.Print "Текст: " & Text
    
    If InStr(1, UCase(Text), "НАРЯД") = 0 Then
        Debug.Print "? Слово 'наряд' не найдено"
        Result.MatchDetails = "Нет слова 'наряд'"
        ParseNaryadForSubstring = Result
        Exit Function
    End If
    
    Debug.Print "? Найдено слово 'наряд'"
    
    Result.DocumentNumber = ExtractSubstringNumber(Text, "НАРЯД")
    Debug.Print "Номер: '" & Result.DocumentNumber & "'"
    
    Result.DocumentDate = ExtractSubstringDate(Text)
    Debug.Print "Дата: '" & Result.DocumentDate & "'"
    
    If Result.DocumentNumber <> "" And Result.DocumentDate <> "" Then
        Result.IsValid = True
        Result.MatchDetails = "Подстроки: успешно"
        Debug.Print "? ПАРСИНГ НАРЯДА УСПЕШЕН"
    Else
        Result.MatchDetails = "Подстроки: нет номера или даты"
        Debug.Print "? ПАРСИНГ НАРЯДА НЕУСПЕШЕН"
    End If
    
    Debug.Print "========================="
    
    ParseNaryadForSubstring = Result
    Exit Function
    
SubstringParseError:
    Debug.Print "? Ошибка парсинга наряда: " & Err.description
    Result.IsValid = False
    Result.MatchDetails = "Ошибка парсинга: " & Err.description
    ParseNaryadForSubstring = Result
End Function

' Извлечение номера документа
Private Function ExtractSubstringNumber(Text As String, DocumentType As String) As String
    Dim UpperText As String
    Dim TypePos As Long
    Dim FromPos As Long
    Dim BetweenText As String
    
    UpperText = UCase(Text)
    
    Debug.Print "  === ИЗВЛЕЧЕНИЕ НОМЕРА ==="
    Debug.Print "  Тип документа: " & DocumentType
    
    TypePos = InStr(UpperText, UCase(DocumentType))
    If TypePos = 0 Then
        ExtractSubstringNumber = ""
        Exit Function
    End If
    
    FromPos = InStr(TypePos, UpperText, " ОТ ")
    If FromPos = 0 Then FromPos = InStr(TypePos, UpperText, "ОТ ")
    
    If FromPos > 0 Then
        BetweenText = Trim(Mid(Text, TypePos + Len(DocumentType), FromPos - TypePos - Len(DocumentType)))
        
        BetweenText = Replace(BetweenText, "№", "")
        BetweenText = Trim(BetweenText)
        
        If BetweenText <> "" Then
            ExtractSubstringNumber = BetweenText
            Debug.Print "  ? Номер: '" & BetweenText & "'"
            Exit Function
        End If
    End If
    
    Debug.Print "  ? Номер не найден"
    ExtractSubstringNumber = ""
End Function

' Извлечение даты документа
Private Function ExtractSubstringDate(Text As String) As String
    Dim UpperText As String
    Dim FromPos As Long
    Dim DateStart As Long
    Dim DateStr As String
    Dim i As Long, Char As String
    
    UpperText = UCase(Text)
    
    Debug.Print "  === ИЗВЛЕЧЕНИЕ ДАТЫ ==="
    
    FromPos = InStr(UpperText, " ОТ ")
    If FromPos = 0 Then FromPos = InStr(UpperText, "ОТ ")
    
    If FromPos > 0 Then
        DateStart = FromPos + 3
        If Mid(UpperText, FromPos, 4) = " ОТ " Then DateStart = FromPos + 4
        
        Do While DateStart <= Len(Text) And Mid(Text, DateStart, 1) = " "
            DateStart = DateStart + 1
        Loop
        
        For i = DateStart To Len(Text)
            Char = Mid(Text, i, 1)
            If (Char >= "0" And Char <= "9") Or Char = "." Then
                DateStr = DateStr & Char
            ElseIf DateStr <> "" And (Char = " " Or Char = "г" Or Char = "Г") Then
                Exit For
            End If
        Next i
        
        If Len(DateStr) >= 8 And InStr(DateStr, ".") > 0 Then
            DateStr = NormalizeDateForSubstring(DateStr)
            ExtractSubstringDate = DateStr
            Debug.Print "  ? Дата: '" & DateStr & "'"
            Exit Function
        End If
    End If
    
    Debug.Print "  ? Дата не найдена"
    ExtractSubstringDate = ""
End Function

' Нормализация даты для поиска подстрок
Private Function NormalizeDateForSubstring(DateStr As String) As String
    Dim DateParts() As String
    Dim Day As String, Month As String, Year As String
    
    On Error GoTo NormalizeError
    
    DateParts = Split(DateStr, ".")
    
    If UBound(DateParts) >= 2 Then
        Day = DateParts(0)
        Month = DateParts(1)
        Year = DateParts(2)
        
        If Len(Day) = 1 Then Day = "0" & Day
        If Len(Month) = 1 Then Month = "0" & Month
        
        If Len(Year) = 2 Then
            If CLng(Year) > 50 Then
                Year = "19" & Year
            Else
                Year = "20" & Year
            End If
        End If
        
        NormalizeDateForSubstring = Day & "." & Month & "." & Year
        Debug.Print "    Нормализованная дата: " & NormalizeDateForSubstring
    Else
        NormalizeDateForSubstring = DateStr
    End If
    
    Exit Function
    
NormalizeError:
    NormalizeDateForSubstring = DateStr
End Function

' =============================================
' ФУНКЦИИ ПОИСКА (СУЩЕСТВУЮЩИЕ)
' =============================================

' Поиск по подстрокам (1-й проход)
Private Function FindBySubstringEnhanced(DoverData As ParsedNaryad, OperArray As Variant) As MatchResult
    Dim Result As MatchResult
    Dim i As Long
    Dim OperComment As String
    
    Result.FoundMatch = False
    Result.MatchScore = 0
    
    For i = 1 To UBound(OperArray, 1)
        OperComment = Trim(CStr(OperArray(i, 9)))
        
        If OperComment <> "" Then
            Dim HasNumber As Boolean, HasDate As Boolean
            
            HasNumber = (InStr(1, OperComment, DoverData.DocumentNumber, vbTextCompare) > 0)
            HasDate = (InStr(1, OperComment, DoverData.DocumentDate, vbTextCompare) > 0)
            
            If Not HasDate Then
                Dim ShortDate As String
                ShortDate = ConvertToShortDate(DoverData.DocumentDate)
                If ShortDate <> "" Then
                    HasDate = (InStr(1, OperComment, ShortDate, vbTextCompare) > 0)
                End If
            End If
            
            If HasNumber And HasDate Then
                Result.FoundMatch = True
                Result.MatchScore = 100
                Result.OperationRow = i + 1
                Result.OperationNumber = Trim(CStr(OperArray(i, 3)))
                Result.MatchDetails = "найден номер И дата"
                
                On Error Resume Next
                Result.OperationDate = CDate(OperArray(i, 2))
                On Error GoTo 0
                
                Exit For
            End If
        End If
    Next i
    
    If Not Result.FoundMatch Then
        Result.MatchDetails = "номер или дата не найдены"
    End If
    
    FindBySubstringEnhanced = Result
End Function

' Преобразование в короткий формат даты
Private Function ConvertToShortDate(FullDate As String) As String
    Dim DateParts() As String
    
    On Error GoTo ConvertError
    
    DateParts = Split(FullDate, ".")
    
    If UBound(DateParts) >= 2 And Len(DateParts(2)) = 4 Then
        ConvertToShortDate = DateParts(0) & "." & DateParts(1) & "." & Right(DateParts(2), 2)
    Else
        ConvertToShortDate = ""
    End If
    
    Exit Function
    
ConvertError:
    ConvertToShortDate = ""
End Function

' =============================================
' 3-Й ПРОХОД: УНИВЕРСАЛЬНЫЕ ДОКУМЕНТЫ
' =============================================

' 3-й проход: универсальные документы
Private Function ProcessThirdPass(WsDover As Worksheet, OperData As Variant, NotFoundRows() As Long, NotFoundData() As ParsedNaryad, NotFoundCount As Long, ByRef Pass3ByType() As Long) As Long
    Dim i As Long
    Dim ThirdPassFound As Long
    Dim MatchResult As MatchResult
    Dim UniversalDoc As ParsedNaryad
    
    ThirdPassFound = 0
    
    For i = 0 To NotFoundCount - 1
        Debug.Print "3-й проход для: " & NotFoundData(i).OriginalText
        
        UniversalDoc = ParseUniversalDocument(NotFoundData(i).OriginalText)
        
        If UniversalDoc.IsValid Then
            MatchResult = FindByUniversalDocument(UniversalDoc, OperData)
            MatchResult.PassNumber = 3
            
            If MatchResult.FoundMatch Then
                ThirdPassFound = ThirdPassFound + 1
                MatchResult.MatchDetails = "3-й проход: " & UniversalDoc.DocumentType & " № " & UniversalDoc.DocumentNumber
                
                Select Case UCase(UniversalDoc.DocumentType)
                    Case "РАЗНАРЯДКА": Pass3ByType(0) = Pass3ByType(0) + 1
                    Case "ТЕЛЕГРАММА": Pass3ByType(1) = Pass3ByType(1) + 1
                    Case "АТТЕСТАТ": Pass3ByType(2) = Pass3ByType(2) + 1
                    Case "ЗАЯВКА": Pass3ByType(3) = Pass3ByType(3) + 1
                End Select
                
                Call WriteThreePassResult(WsDover, NotFoundRows(i), MatchResult)
                
                Debug.Print "  ? Найдено 3-м проходом: " & UniversalDoc.DocumentType
            Else
                Debug.Print "  ? Не найдено 3-м проходом"
            End If
        Else
            Debug.Print "  ? Не удалось распарсить для 3-го прохода"
        End If
    Next i
    
    ProcessThirdPass = ThirdPassFound
End Function

' Парсинг универсальных документов
Private Function ParseUniversalDocument(Text As String) As ParsedNaryad
    Dim Result As ParsedNaryad
    Dim DocumentTypes(3) As String
    Dim i As Long
    Dim FoundType As String
    
    Result.OriginalText = Text
    Result.IsValid = False
    
    DocumentTypes(0) = "РАЗНАРЯДКА"
    DocumentTypes(1) = "ТЕЛЕГРАММА"
    DocumentTypes(2) = "АТТЕСТАТ"
    DocumentTypes(3) = "ЗАЯВКА"
    
    Debug.Print "=== ПАРСИНГ УНИВЕРСАЛЬНОГО ДОКУМЕНТА ==="
    Debug.Print "Текст: " & Text
    
    For i = 0 To UBound(DocumentTypes)
        If InStr(1, UCase(Text), DocumentTypes(i)) > 0 Then
            FoundType = DocumentTypes(i)
            Debug.Print "? Найден тип документа: " & FoundType
            Exit For
        End If
    Next i
    
    If FoundType = "" Then
        Debug.Print "? Тип универсального документа не найден"
        Result.MatchDetails = "Тип документа не найден"
        ParseUniversalDocument = Result
        Exit Function
    End If
    
    Result.DocumentType = FoundType
    
    Result.DocumentNumber = ExtractSubstringNumber(Text, FoundType)
    Debug.Print "Номер: '" & Result.DocumentNumber & "'"
    
    Result.DocumentDate = ExtractSubstringDate(Text)
    Debug.Print "Дата: '" & Result.DocumentDate & "'"
    
    If Result.DocumentNumber <> "" And Result.DocumentDate <> "" Then
        Result.IsValid = True
        Result.MatchDetails = "3-й проход: успешно"
        Debug.Print "? УНИВЕРСАЛЬНЫЙ ДОКУМЕНТ РАСПАРСЕН"
    Else
        Result.MatchDetails = "3-й проход: нет номера или даты"
        Debug.Print "? УНИВЕРСАЛЬНЫЙ ДОКУМЕНТ НЕ РАСПАРСЕН"
    End If
    
    Debug.Print "========================="
    
    ParseUniversalDocument = Result
End Function

' Поиск по универсальному документу
Private Function FindByUniversalDocument(UniversalDoc As ParsedNaryad, OperData As Variant) As MatchResult
    Dim Result As MatchResult
    Dim i As Long
    Dim OperComment As String
    
    Result.FoundMatch = False
    Result.MatchScore = 0
    
    Debug.Print ">>> 3-Й ПРОХОД ПОИСКА ДЛЯ: " & UniversalDoc.DocumentType & " № '" & UniversalDoc.DocumentNumber & "' от '" & UniversalDoc.DocumentDate & "'"
    
    For i = 1 To UBound(OperData, 1)
        OperComment = Trim(CStr(OperData(i, 9)))
        
        If OperComment <> "" Then
            Dim HasNumber As Boolean, HasDate As Boolean
            
            HasNumber = (InStr(1, OperComment, UniversalDoc.DocumentNumber, vbTextCompare) > 0)
            HasDate = (InStr(1, OperComment, UniversalDoc.DocumentDate, vbTextCompare) > 0)
            
            If Not HasDate Then
                Dim ShortDate As String
                ShortDate = ConvertToShortDate(UniversalDoc.DocumentDate)
                If ShortDate <> "" Then
                    HasDate = (InStr(1, OperComment, ShortDate, vbTextCompare) > 0)
                End If
            End If
            
            If HasNumber And HasDate Then
                Result.FoundMatch = True
                Result.MatchScore = 100
                Result.OperationRow = i + 1
                Result.OperationNumber = Trim(CStr(OperData(i, 3)))
                Result.MatchDetails = "найден " & UniversalDoc.DocumentType
                
                On Error Resume Next
                Result.OperationDate = CDate(OperData(i, 2))
                On Error GoTo 0
                
                Debug.Print ">>> ? НАЙДЕНО 3-М ПРОХОДОМ в строке " & (i + 1)
                Exit For
            End If
        End If
    Next i
    
    If Not Result.FoundMatch Then
        Result.MatchDetails = "номер или дата не найдены"
        Debug.Print ">>> ? НЕ НАЙДЕНО 3-М ПРОХОДОМ"
    End If
    
    FindByUniversalDocument = Result
End Function

' =============================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' =============================================

' Фильтрация операций по корреспонденту
Private Function FilterOperationsByCorrespondent(OperData As Variant, ByRef FilteredCount As Long, ByRef SkippedCount As Long) As Variant
    Dim i As Long
    Dim TempArray() As Variant
    Dim FilteredRows As Long
    
    FilteredCount = 0
    SkippedCount = 0
    FilteredRows = 0
    
    For i = 1 To UBound(OperData, 1)
        If Trim(CStr(OperData(i, 6))) <> "" Then
            FilteredRows = FilteredRows + 1
        Else
            SkippedCount = SkippedCount + 1
        End If
    Next i
    
    If FilteredRows > 0 Then
        ReDim TempArray(1 To FilteredRows, 1 To UBound(OperData, 2))
        
        For i = 1 To UBound(OperData, 1)
            If Trim(CStr(OperData(i, 6))) <> "" Then
                FilteredCount = FilteredCount + 1
                Dim j As Long
                For j = 1 To UBound(OperData, 2)
                    TempArray(FilteredCount, j) = OperData(i, j)
                Next j
            End If
        Next i
    End If
    
    FilterOperationsByCorrespondent = TempArray
    FilteredCount = FilteredRows
End Function

' Добавление столбцов результатов
Private Sub AddThreePassResultColumns(Ws As Worksheet)
    Dim LastCol As Long
    Dim i As Long
    Dim AlreadyExists As Boolean
    
    For i = 1 To 25
        If Trim(CStr(Ws.Cells(1, i).value)) = "Трехпроходный поиск" Then
            AlreadyExists = True
            Exit For
        End If
    Next i
    
    If Not AlreadyExists Then
        LastCol = Ws.Cells(1, Ws.Columns.Count).End(xlToLeft).Column
        
        With Ws
            .Cells(1, LastCol + 1).value = "Трехпроходный поиск"
            .Cells(1, LastCol + 2).value = "Номер операции"
            .Cells(1, LastCol + 3).value = "Дата операции"
            .Cells(1, LastCol + 4).value = "Процент"
            .Cells(1, LastCol + 5).value = "Детали"
            .Cells(1, LastCol + 6).value = "Комментарий 1С" ' НОВЫЙ СТОЛБЕЦ!
            
            .Range(.Cells(1, LastCol + 1), .Cells(1, LastCol + 6)).Font.Bold = True
            .Range(.Cells(1, LastCol + 1), .Cells(1, LastCol + 6)).Interior.Color = RGB(220, 220, 255)
            .Range(.Cells(1, LastCol + 1), .Cells(1, LastCol + 6)).Borders.LineStyle = xlContinuous
        End With
    End If
End Sub


' Запись результата с цветовой маркировкой
Private Sub WriteThreePassResult(Ws As Worksheet, Row As Long, MatchResult As MatchResult)
    Dim ResultCol As Long
    Static CachedCol As Long
    
    If CachedCol = 0 Then
        For ResultCol = 10 To 30
            If Trim(CStr(Ws.Cells(1, ResultCol).value)) = "Трехпроходный поиск" Then
                CachedCol = ResultCol
                Exit For
            End If
        Next ResultCol
    End If
    
    If CachedCol > 0 Then
        If MatchResult.FoundMatch Then
            ' Определяем, нужно ли добавить "ВЕРОЯТНО"
            Dim ResultText As String
            If MatchResult.MatchScore >= 80 Then
                ResultText = "НАЙДЕНО"
            ElseIf MatchResult.MatchScore >= 60 Then
                ResultText = "ВЕРОЯТНО НАЙДЕНО"
            Else
                ResultText = "ВЕРОЯТНО"
            End If
            
            Ws.Cells(Row, CachedCol).value = ResultText
            Ws.Cells(Row, CachedCol + 1).value = MatchResult.OperationNumber
            Ws.Cells(Row, CachedCol + 2).value = MatchResult.OperationDate
            Ws.Cells(Row, CachedCol + 3).value = Format(MatchResult.MatchScore, "0") & "%"
            Ws.Cells(Row, CachedCol + 4).value = MatchResult.MatchDetails
            
            ' НОВОЕ: Извлекаем комментарий из 1С из MatchDetails
            ' ИСПРАВЛЕННОЕ: Извлекаем полный комментарий из 1С
Dim Comment1C As String
Dim CommentStart As Long
CommentStart = InStr(MatchResult.MatchDetails, "| 1С: ")

If CommentStart > 0 Then
    Comment1C = Mid(MatchResult.MatchDetails, CommentStart + 5)
    Comment1C = Trim(Comment1C)
    Call WriteFullComment(Ws, Row, CachedCol + 5, Comment1C)
End If


            
            ' ЦВЕТОВАЯ МАРКИРОВКА ПО ПРОХОДАМ И УВЕРЕННОСТИ
            Select Case MatchResult.PassNumber
                Case 1: ' 1-й проход - темно-зеленый
                    Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 5)).Interior.Color = RGB(0, 200, 0)
                Case 2: ' 2-й проход - зависит от уверенности
                    If MatchResult.MatchScore >= 80 Then
                        Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 5)).Interior.Color = RGB(150, 255, 150) ' Светло-зеленый
                    Else
                        Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 5)).Interior.Color = RGB(255, 255, 150) ' Желтый - требует проверки
                    End If
                Case 3: ' 3-й проход - голубой
                    Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 5)).Interior.Color = RGB(150, 200, 255)
            End Select
        Else
            Ws.Cells(Row, CachedCol).value = "НЕ НАЙДЕНО"
            Ws.Cells(Row, CachedCol + 3).value = "0%"
            Ws.Cells(Row, CachedCol + 4).value = MatchResult.MatchDetails
            
            If InStr(MatchResult.MatchDetails, "проход") > 0 Then
                Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 5)).Interior.Color = RGB(255, 255, 150) ' Желтый - ожидание
            Else
                Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 5)).Interior.Color = RGB(255, 200, 200) ' Красный - не найдено
            End If
        End If
    End If
End Sub


' Создание пустого результата
Private Function CreateEmptyMatchResult(Message As String) As MatchResult
    Dim Result As MatchResult
    Result.FoundMatch = False
    Result.MatchDetails = Message
    Result.MatchScore = 0
    Result.PassNumber = 0
    CreateEmptyMatchResult = Result
End Function

' Показ статистики
Private Sub ShowThreePassStatistics(Processed As Long, Pass1 As Long, Pass2 As Long, Pass3 As Long, Pass3ByType() As Long, SkippedEmpty As Long, SkippedFound As Long)

    Dim TotalFound As Long
    Dim Message As String
    
    TotalFound = Pass1 + Pass2 + Pass3
    
    Message = "?? РЕЗУЛЬТАТЫ ТРЕХПРОХОДНОЙ СИСТЕМЫ С ИСПРАВЛЕННЫМ REGEX:" & vbCrLf & vbCrLf & _
              "?? Всего обработано: " & Processed & vbCrLf & vbCrLf & _
              "? 1-й проход (точные наряды): " & Pass1 & " (" & Format(Pass1 / Processed * 100, "0.0") & "%)" & vbCrLf & _
              "?? 2-й проход (исправленный REGEX наряды): " & Pass2 & " (" & Format(Pass2 / Processed * 100, "0.0") & "%)" & vbCrLf & _
              "?? 3-й проход (другие документы): " & Pass3 & " (" & Format(Pass3 / Processed * 100, "0.0") & "%)" & vbCrLf & _
              "   • Разнарядки: " & Pass3ByType(0) & vbCrLf & _
              "   • Телеграммы: " & Pass3ByType(1) & vbCrLf & _
              "   • Аттестаты: " & Pass3ByType(2) & vbCrLf & _
              "   • Заявки: " & Pass3ByType(3) & vbCrLf & vbCrLf & _
              "?? Общий результат: " & TotalFound & " (" & Format(TotalFound / Processed * 100, "0.0") & "%)" & vbCrLf & _
              "? Не найдено: " & (Processed - TotalFound) & vbCrLf & vbCrLf & _
              "?? Пропущено операций с пустым корреспондентом: " & SkippedEmpty & vbCrLf & vbCrLf & _
              "?? ЦВЕТОВАЯ МАРКИРОВКА:" & vbCrLf & _
              "?? Темно-зеленый = 1-й проход (точные наряды)" & vbCrLf & _
              "?? Светло-зеленый = 2-й проход (исправленный REGEX наряды)" & vbCrLf & _
              "?? Голубой = 3-й проход (другие документы)" & vbCrLf & vbCrLf & _
              "?? Результаты в столбцах 'Трехпроходный поиск'"
    
    MsgBox Message, vbInformation, "Трехпроходная система с исправленным REGEX завершена"
End Sub

' Проверка на наличие военных кодов
Private Function ContainsMilitaryCode(Number As String) As Boolean
    Dim MilitaryCodes As Variant
    Dim i As Long
    Dim UpperNumber As String
    
    UpperNumber = UCase(Number)
    MilitaryCodes = Array("ЮВО", "АТ", "ЦВО", "ВВО", "СВО", "БТ", "ПВО", "ВДВ", "ВМФ", "ЗЧ", "И-", "А-", "Б-", "В-", "С-")
    
    For i = 0 To UBound(MilitaryCodes)
        If InStr(UpperNumber, CStr(MilitaryCodes(i))) > 0 Then
            ContainsMilitaryCode = True
            Exit Function
        End If
    Next i
    
    ContainsMilitaryCode = False
End Function

' Проверка на наличие русских букв
Private Function ContainsRussianLetters(Text As String) As Boolean
    Dim i As Long
    Dim Char As String
    
    For i = 1 To Len(Text)
        Char = Mid(Text, i, 1)
        If (Char >= "А" And Char <= "я") Or (Char >= "Ё" And Char <= "ё") Then
            ContainsRussianLetters = True
            Exit Function
        End If
    Next i
    
    ContainsRussianLetters = False
End Function

' Проверка на только цифры
Private Function IsNumericOnly(Text As String) As Boolean
    Dim i As Long
    Dim Char As String
    
    For i = 1 To Len(Text)
        Char = Mid(Text, i, 1)
        If Not (Char >= "0" And Char <= "9") Then
            IsNumericOnly = False
            Exit Function
        End If
    Next i
    
    IsNumericOnly = True
End Function

' НОВЫЕ ФУНКЦИИ - ДОБАВИТЬ В КОНЕЦ МОДУЛЯ

Private Function CreateSlashPattern(Number As String, DateStr As String) As String
    Dim FlexNumber As String
    FlexNumber = Replace(Number, "/", "\s*/\s*") ' Пробелы вокруг слешей опциональны
    CreateSlashPattern = "(?i)наряд\s*№?\s*" & FlexNumber & "\s*от\s*" & EscapeMinimal(DateStr) & "(?:\s*г\.?)?"
End Function

Private Function CreateDashPattern(Number As String, DateStr As String) As String
    Dim FlexNumber As String
    FlexNumber = Replace(Number, "-", "\s*-\s*") ' Пробелы вокруг дефисов опциональны
    CreateDashPattern = "(?i)наряд\s*№?\s*" & FlexNumber & "\s*от\s*" & EscapeMinimal(DateStr) & "(?:\s*г\.?)?"
End Function

Private Function CreateMilitaryPattern(Number As String, DateStr As String) As String
    Dim FlexNumber As String
    FlexNumber = Number
    
    ' Делаем военные коды регистронезависимыми
    FlexNumber = Replace(FlexNumber, "ЮВО", "(?:ЮВО|юво|Юво)")
    FlexNumber = Replace(FlexNumber, "БТ", "(?:БТ|бт|Бт)")
    FlexNumber = Replace(FlexNumber, "ЗЧ", "(?:ЗЧ|зч|Зч)")
    
    ' Гибкие разделители
    FlexNumber = Replace(FlexNumber, "/", "\s*/\s*")
    FlexNumber = Replace(FlexNumber, "-", "\s*-\s*")
    
    CreateMilitaryPattern = "(?i)наряд\s*№?\s*" & FlexNumber & "\s*от\s*" & EscapeMinimal(DateStr) & "(?:\s*г\.?)?"
End Function

Private Function CreateNumericPattern(Number As String, DateStr As String) As String
    CreateNumericPattern = "(?i)наряд\s*№?\s*" & Number & "\s*от\s*" & EscapeMinimal(DateStr) & "(?:\s*г\.?)?"
End Function



' =============================================
' ОСНОВНАЯ ФУНКЦИЯ АНАЛИЗА КОМПОНЕНТОВ
' =============================================

Private Function AnalyzeComponents(CommentText As String, TargetNumber As String, TargetDate As String) As ComponentMatchResult
    Dim Result As ComponentMatchResult
    
    Result.IsMatch = False
    Result.Confidence = 0
    
    Debug.Print "      >>> АНАЛИЗ КОМПОНЕНТОВ В: " & CommentText
    Debug.Print "      Ищем номер: '" & TargetNumber & "' и дату: '" & TargetDate & "'"
    
    ' ЭТАП 1: Найти все потенциальные номера
    Dim NumberCandidates() As NumberCandidate
    Dim NumberCount As Long
    Call ExtractNumberCandidates(CommentText, NumberCandidates, NumberCount)
    
    ' ЭТАП 2: Найти все потенциальные даты
    Dim DateCandidates() As DateCandidate
    Dim DateCount As Long
    Call ExtractDateCandidates(CommentText, DateCandidates, DateCount)
    
    Debug.Print "      Найдено номеров: " & NumberCount & ", дат: " & DateCount
    
    ' ЭТАП 3: Интеллектуальное сопоставление
    If NumberCount > 0 And DateCount > 0 Then
        Result = FindBestComponentMatch(CommentText, NumberCandidates, NumberCount, DateCandidates, DateCount, TargetNumber, TargetDate)
    End If
    
    AnalyzeComponents = Result
End Function

' =============================================
' ИЗВЛЕЧЕНИЕ НОМЕРОВ-КАНДИДАТОВ
' =============================================

Private Sub ExtractNumberCandidates(Text As String, ByRef NumberCandidates() As NumberCandidate, ByRef Count As Long)
    Dim i As Long, j As Long
    Dim Char As String
    Dim CurrentNumber As String
    Dim InNumber As Boolean
    Dim StartPos As Long
    
    Count = 0
    ReDim NumberCandidates(50) ' Максимум 50 кандидатов
    
    Debug.Print "        Извлечение номеров из: " & Text
    
    ' Поиск паттернов номеров
    For i = 1 To Len(Text) - 1
        Char = Mid(Text, i, 1)
        
        ' Начинаем номер если найден контекст
        If Not InNumber Then
            If IsNumberStart(Text, i) Then
                InNumber = True
                StartPos = i
                CurrentNumber = ""
                Debug.Print "        Начало номера в позиции " & i
            End If
        End If
        
        ' Собираем символы номера
        If InNumber Then
            If IsNumberChar(Char) Then
                CurrentNumber = CurrentNumber & Char
            ElseIf IsNumberSeparator(Char) Then
                CurrentNumber = CurrentNumber & Char
            Else
                ' Конец номера
                If Len(CurrentNumber) > 0 Then
                    Call AddNumberCandidate(NumberCandidates, Count, CurrentNumber, StartPos, Len(CurrentNumber))
                    Debug.Print "        Найден номер: '" & CurrentNumber & "' в позиции " & StartPos
                End If
                InNumber = False
                CurrentNumber = ""
            End If
        End If
    Next i
    
    ' Обрабатываем последний номер
    If InNumber And Len(CurrentNumber) > 0 Then
        Call AddNumberCandidate(NumberCandidates, Count, CurrentNumber, StartPos, Len(CurrentNumber))
        Debug.Print "        Найден номер (в конце): '" & CurrentNumber & "' в позиции " & StartPos
    End If
    
    Debug.Print "        Всего найдено номеров: " & Count
End Sub

' =============================================
' ИЗВЛЕЧЕНИЕ ДАТ-КАНДИДАТОВ
' =============================================

Private Sub ExtractDateCandidates(Text As String, ByRef DateCandidates() As DateCandidate, ByRef Count As Long)
    Dim RegEx As Object
    Dim Matches As Object
    Dim Match As Object
    Dim i As Long
    
    Set RegEx = CreateObject("VBScript.RegExp")
    
    Count = 0
    ReDim DateCandidates(20) ' Максимум 20 дат
    
    Debug.Print "        Извлечение дат из: " & Text
    
    With RegEx
        .IgnoreCase = True
        .Global = True
        .Pattern = "\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2})(?:\s*г\.?)?"
    End With
    
    On Error Resume Next
    Set Matches = RegEx.Execute(Text)
    
    If Matches.Count > 0 Then
        For i = 0 To Matches.Count - 1
            Set Match = Matches(i)
            Call AddDateCandidate(DateCandidates, Count, Match.value, Match.FirstIndex + 1, Match.Length)
            Debug.Print "        Найдена дата: '" & Match.value & "' в позиции " & (Match.FirstIndex + 1)
        Next i
    End If
    On Error GoTo 0
    
    Debug.Print "        Всего найдено дат: " & Count
End Sub

' =============================================
' ИНТЕЛЛЕКТУАЛЬНОЕ СОПОСТАВЛЕНИЕ
' =============================================

Private Function FindBestComponentMatch(Text As String, NumberCandidates() As NumberCandidate, NumberCount As Long, DateCandidates() As DateCandidate, DateCount As Long, TargetNumber As String, TargetDate As String) As ComponentMatchResult
    Dim Result As ComponentMatchResult
    Dim BestScore As Double
    Dim i As Long, j As Long
    Dim TempResult As ComponentMatchResult
    
    Result.IsMatch = False
    BestScore = 0
    
'    Debug.Print "        Сопоставление " & NumberCount & " номеров с " & DateCount & " датами"
    
    ' Проверяем все комбинации номер+дата
    For i = 0 To NumberCount - 1
        For j = 0 To DateCount - 1
'            Debug.Print "          Проверяем: '" & NumberCandidates(i).Text & "' + '" & DateCandidates(j).Text & "'"
            
            TempResult = EvaluateComponentPair(Text, NumberCandidates(i), DateCandidates(j), TargetNumber, TargetDate)
            
            If TempResult.Confidence > BestScore Then
                BestScore = TempResult.Confidence
                Result = TempResult
'                Debug.Print "          ? НОВЫЙ ЛУЧШИЙ РЕЗУЛЬТАТ: " & BestScore & "%"
            End If
        Next j
    Next i
    
    If BestScore >= 40 Then
        Result.IsMatch = True
'        Debug.Print "        ИТОГО: Лучшее совпадение " & BestScore & "% - " & Result.MatchDetails
    Else
'        Debug.Print "        ИТОГО: Лучший результат " & BestScore & "% недостаточен (нужно ?40%)"
    End If
    
    FindBestComponentMatch = Result
End Function

' =============================================
' ОЦЕНКА ПАРЫ НОМЕР+ДАТА
' =============================================

Private Function EvaluateComponentPair(Text As String, NumberCand As NumberCandidate, DateCand As DateCandidate, TargetNumber As String, TargetDate As String) As ComponentMatchResult
    Dim Result As ComponentMatchResult
    Dim Score As Double
    
    Score = 0
    Result.FoundNumber = NumberCand.Text
    Result.FoundDate = DateCand.Text
    Result.NumberPosition = NumberCand.Position
    Result.DatePosition = DateCand.Position
    
    ' КРИТЕРИЙ 1: Совпадение номера (0-50 баллов)
    Dim NumberScore As Double
    NumberScore = CompareNumbers(TargetNumber, NumberCand.Text)
    Score = Score + NumberScore
    
    ' КРИТЕРИЙ 2: Совпадение даты (0-40 баллов)
    Dim DateScore As Double
    DateScore = CompareDates(TargetDate, DateCand.Text)
    Score = Score + DateScore
    
    ' КРИТЕРИЙ 3: Позиционный анализ (0-10 баллов)
    Dim PositionScore As Double
    PositionScore = EvaluatePositions(Text, NumberCand.Position, DateCand.Position)
    Score = Score + PositionScore
    
    Result.Confidence = Score
    Result.MatchDetails = "номер:" & NumberCand.Text & "(" & Format(NumberScore, "0") & "%) дата:" & DateCand.Text & "(" & Format(DateScore, "0") & "%) поз:(" & Format(PositionScore, "0") & "%)"
    
    EvaluateComponentPair = Result
End Function

' =============================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' =============================================

Private Function CompareNumbers(Target As String, Found As String) As Double
    Dim Score As Double
    
    ' Точное совпадение = 50 баллов
    If UCase(Trim(Target)) = UCase(Trim(Found)) Then
        Score = 50
    ' Содержится в найденном = 35 баллов
    ElseIf InStr(1, UCase(Found), UCase(Target), vbTextCompare) > 0 Then
        Score = 35
    ' Найденный содержится в целевом = 25 баллов
    ElseIf InStr(1, UCase(Target), UCase(Found), vbTextCompare) > 0 Then
        Score = 25
    ' Частичное совпадение = 10-20 баллов
    Else
        Score = CalculatePartialNumberMatch(Target, Found)
    End If
    
    CompareNumbers = Score
End Function

Private Function CompareDates(Target As String, Found As String) As Double
    Dim Score As Double
    Dim TargetDate As Date, FoundDate As Date
    
    On Error Resume Next
    TargetDate = CDate(Target)
    FoundDate = CDate(NormalizeDateString(Found))
    On Error GoTo 0
    
    ' СТРОГИЕ КРИТЕРИИ ДЛЯ ДАТ!
    ' Точное совпадение = 40 баллов
    If TargetDate = FoundDate And TargetDate > 0 And FoundDate > 0 Then
        Score = 40
    ' Разница в 1 день = 25 баллов (снижено!)
    ElseIf Abs(TargetDate - FoundDate) = 1 And TargetDate > 0 And FoundDate > 0 Then
        Score = 25
    ' Остальные случаи = 0 баллов (было 15-35, теперь 0!)
    Else
        Score = 0
    End If
    
    CompareDates = Score
End Function


Private Function EvaluatePositions(Text As String, NumberPos As Long, DatePos As Long) As Double
    Dim Score As Double
    
    ' Номер должен идти перед датой
    If NumberPos < DatePos Then
        Score = Score + 5
        
        ' Расстояние между номером и датой
        Dim Distance As Long
        Distance = DatePos - NumberPos
        
        If Distance <= 20 Then
            Score = Score + 5 ' Близко друг к другу
        ElseIf Distance <= 50 Then
            Score = Score + 3 ' Средняя дистанция
        Else
            Score = Score + 1 ' Далеко, но в правильном порядке
        End If
    End If
    
    EvaluatePositions = Score
End Function

' =============================================
' СЛУЖЕБНЫЕ ФУНКЦИИ
' =============================================

Private Function IsNumberStart(Text As String, Position As Long) As Boolean
    Dim Context As String
    Dim StartCheck As Long
    
    StartCheck = Position - 10
    If StartCheck < 1 Then StartCheck = 1
    
    Context = UCase(Mid(Text, StartCheck, 15))
    
    IsNumberStart = (InStr(Context, "НАРЯД") > 0) Or _
                   (InStr(Context, "№") > 0) Or _
                   (InStr(Context, "#") > 0)
End Function

Private Function IsNumberChar(Char As String) As Boolean
    IsNumberChar = (Char >= "0" And Char <= "9") Or _
                  (Char >= "A" And Char <= "Z") Or _
                  (Char >= "А" And Char <= "Я") Or _
                  (Char >= "a" And Char <= "z") Or _
                  (Char >= "а" And Char <= "я")
End Function

Private Function IsNumberSeparator(Char As String) As Boolean
    IsNumberSeparator = (Char = "/") Or (Char = "-") Or (Char = " ")
End Function

Private Sub AddNumberCandidate(ByRef Candidates() As NumberCandidate, ByRef Count As Long, NumberText As String, Position As Long, Length As Long)
    If Count < UBound(Candidates) Then
        With Candidates(Count)
            .Text = Trim(NumberText)
            .Position = Position
            .Length = Length
            .ConfidenceScore = 0
            .NumberType = ClassifyNumberType(NumberText)
        End With
        Count = Count + 1
    End If
End Sub

Private Sub AddDateCandidate(ByRef Candidates() As DateCandidate, ByRef Count As Long, DateText As String, Position As Long, Length As Long)
    If Count < UBound(Candidates) Then
        With Candidates(Count)
            .Text = Trim(DateText)
            .Position = Position
            .Length = Length
            .DateFormat = "AUTO"
            On Error Resume Next
            .ParsedDate = CDate(NormalizeDateString(DateText))
            On Error GoTo 0
        End With
        Count = Count + 1
    End If
End Sub

Private Function ClassifyNumberType(NumberText As String) As String
    If IsNumericOnly(NumberText) Then
        ClassifyNumberType = "SIMPLE"
    ElseIf InStr(NumberText, "/") > 0 And InStr(NumberText, "-") > 0 Then
        ClassifyNumberType = "COMPLEX"
    ElseIf InStr(NumberText, "/") > 0 Then
        ClassifyNumberType = "SLASH"
    ElseIf InStr(NumberText, "-") > 0 Then
        ClassifyNumberType = "DASH"
    ElseIf ContainsMilitaryCode(NumberText) Then
        ClassifyNumberType = "MILITARY"
    Else
        ClassifyNumberType = "OTHER"
    End If
End Function

Private Function NormalizeDateString(DateStr As String) As String
    Dim CleanDate As String
    CleanDate = DateStr
    
    ' Убираем "г." и лишние пробелы
    CleanDate = Replace(CleanDate, "г.", "")
    CleanDate = Replace(CleanDate, "г", "")
    CleanDate = Trim(CleanDate)
    
    NormalizeDateString = CleanDate
End Function

Private Function CalculatePartialNumberMatch(Target As String, Found As String) As Double
    Dim CommonChars As Long
    Dim i As Long, j As Long
    Dim MaxLen As Long
    
    MaxLen = IIf(Len(Target) > Len(Found), Len(Target), Len(Found))
    
    ' Подсчитываем общие символы
    For i = 1 To Len(Target)
        For j = 1 To Len(Found)
            If UCase(Mid(Target, i, 1)) = UCase(Mid(Found, j, 1)) Then
                CommonChars = CommonChars + 1
                Exit For
            End If
        Next j
    Next i
    
    If MaxLen > 0 Then
        CalculatePartialNumberMatch = (CommonChars / MaxLen) * 20
    Else
        CalculatePartialNumberMatch = 0
    End If
End Function


' НОВАЯ ФУНКЦИЯ: Проверка корреспондента
Private Function CheckCorrespondentMatch(OperCorrespondent As Variant, DoverText As String, OperationRow As Long) As Double
    Dim OperCorr As String, DoverCorr As String
    Dim Score As Double
    
    Score = 0
    OperCorr = Trim(CStr(OperCorrespondent))
    
    Debug.Print "        Проверка корреспондента операции " & OperationRow
    Debug.Print "        Операция: '" & OperCorr & "'"
    
    If OperCorr = "" Then
        Debug.Print "        Корреспондент операции пустой"
        CheckCorrespondentMatch = 0
    Exit Function

    End If
    
    ' Извлекаем корреспондента из доверенности (предполагаем, что он может быть в комментарии)
    ' Пока используем упрощенную логику - можно расширить
    DoverCorr = ExtractCorrespondentFromComment(DoverText)
    
    Debug.Print "        Доверенность: '" & DoverCorr & "'"
    
    If DoverCorr <> "" Then
        ' Убираем части в скобках для сравнения
        Dim CleanOperCorr As String, CleanDoverCorr As String
        CleanOperCorr = CleanCorrespondentName(OperCorr)
        CleanDoverCorr = CleanCorrespondentName(DoverCorr)
        
        Debug.Print "        Очищенная операция: '" & CleanOperCorr & "'"
        Debug.Print "        Очищенная доверенность: '" & CleanDoverCorr & "'"
        
        ' Точное совпадение = +50 баллов (почти гарантия!)
        If UCase(CleanOperCorr) = UCase(CleanDoverCorr) Then
            Score = 50
            Debug.Print "        ? ТОЧНОЕ СОВПАДЕНИЕ КОРРЕСПОНДЕНТОВ! +50 баллов"
        ' Частичное совпадение = +25 баллов
        ElseIf InStr(1, UCase(CleanOperCorr), UCase(CleanDoverCorr), vbTextCompare) > 0 Or _
               InStr(1, UCase(CleanDoverCorr), UCase(CleanOperCorr), vbTextCompare) > 0 Then
            Score = 25
            Debug.Print "        ? Частичное совпадение корреспондентов! +25 баллов"
        Else
            Debug.Print "        ? Корреспонденты не совпадают"
        End If
    Else
        Debug.Print "        Не удалось извлечь корреспондента из доверенности"
    End If
    
    CheckCorrespondentMatch = Score
End Function

' Очистка названия корреспондента (убираем скобки)
Private Function CleanCorrespondentName(FullName As String) As String
    Dim Result As String
    Dim BracketPos As Long
    
    Result = Trim(FullName)
    
    ' Убираем все с первой открывающейся скобки
    BracketPos = InStr(Result, "(")
    If BracketPos > 0 Then
        Result = Trim(Left(Result, BracketPos - 1))
    End If
    
    CleanCorrespondentName = Result
End Function

' Извлечение корреспондента из комментария доверенности (упрощенная версия)
Private Function ExtractCorrespondentFromComment(Comment As String) As String
    ' Пока возвращаем пустую строку - можно расширить логику
    ' если в комментариях доверенностей есть информация о корреспондентах
    ExtractCorrespondentFromComment = ""
End Function


' ФУНКЦИЯ ДЛЯ ЗАПИСИ ПОЛНОГО КОММЕНТАРИЯ
Private Sub WriteFullComment(Ws As Worksheet, Row As Long, Col As Long, FullComment As String)
    With Ws.Cells(Row, Col)
        .value = FullComment
        .WrapText = True
    End With
    
    ' Автоподбор высоты строки
    Ws.Rows(Row).AutoFit
    
    ' Ограничиваем максимальную высоту строки (чтобы не было слишком высоких строк)
    If Ws.Rows(Row).RowHeight > 100 Then
        Ws.Rows(Row).RowHeight = 100
    End If
End Sub


' =============================================
' ФУНКЦИИ ДЛЯ ПРОВЕРКИ УЖЕ НАЙДЕННЫХ ЗАПИСЕЙ
' =============================================

' Проверяет, была ли строка уже обработана с результатом "НАЙДЕНО"
Private Function IsAlreadyProcessed(Ws As Worksheet, Row As Long, ResultColumn As Long) As Boolean
    Dim ResultValue As String
    
    IsAlreadyProcessed = False
    
    If ResultColumn > 0 Then
        ResultValue = Trim(UCase(CStr(Ws.Cells(Row, ResultColumn).value)))
        
        If ResultValue = "НАЙДЕНО" Then
            IsAlreadyProcessed = True
        End If
    End If
End Function

' Находит номер столбца с результатами поиска
Private Function FindResultColumn(Ws As Worksheet) As Long
    Dim i As Long
    
    FindResultColumn = 0
    
    For i = 1 To 30
        If Trim(UCase(CStr(Ws.Cells(1, i).value))) = "ТРЕХПРОХОДНЫЙ ПОИСК" Then
            FindResultColumn = i
            Exit For
        End If
    Next i
End Function


' =============================================
' 2-Й ПРОХОД: ПОИСК ПО НОМЕРУ + ПОСТАВЩИКУ + СЛУЖБЕ
' =============================================

' НОВАЯ ФУНКЦИЯ: 2-й проход по номеру, поставщику и службе
Private Function ProcessSecondPassWithSupplierService(WsDover As Worksheet, OperData As Variant, NotFoundRows() As Long, NotFoundData() As ParsedNaryad, NotFoundCount As Long, ByRef Pass2NotFound() As Long, ByRef Pass2NotFoundData() As ParsedNaryad, ByRef Pass2Count As Long) As Long
    Dim i As Long
    Dim SecondPassFound As Long
    Dim MatchResult As MatchResult
    
    SecondPassFound = 0
    
    Debug.Print "=== 2-Й ПРОХОД: ПОИСК ПО НОМЕРУ + ПОСТАВЩИКУ + СЛУЖБЕ ==="
    
    For i = 0 To NotFoundCount - 1
        Debug.Print "2-й проход для: " & NotFoundData(i).DocumentNumber
        
        ' ИСПОЛЬЗУЕМ НОВЫЙ ПОИСК ПО ПОСТАВЩИКУ И СЛУЖБЕ!
        MatchResult = FindBySupplierAndService(NotFoundData(i), OperData)
        MatchResult.PassNumber = 2
        
        If MatchResult.FoundMatch Then
            SecondPassFound = SecondPassFound + 1
            MatchResult.MatchDetails = "2-й проход (" & Format(MatchResult.MatchScore, "0") & "%): " & MatchResult.MatchDetails
            
            Call WriteThreePassResult(WsDover, NotFoundRows(i), MatchResult)
            
            Debug.Print "  ? Найдено 2-м проходом! Уверенность: " & MatchResult.MatchScore & "%"
        Else
            ' Сохраняем для 3-го прохода
            Pass2Count = Pass2Count + 1
            ReDim Preserve Pass2NotFound(Pass2Count - 1)
            ReDim Preserve Pass2NotFoundData(Pass2Count - 1)
            Pass2NotFound(Pass2Count - 1) = NotFoundRows(i)
            Pass2NotFoundData(Pass2Count - 1) = NotFoundData(i)
            
            Debug.Print "  ? Не найдено по поставщику+службе, переходит в 3-й"
        End If
    Next i
    
    ProcessSecondPassWithSupplierService = SecondPassFound
End Function

' ОСНОВНАЯ ФУНКЦИЯ: Поиск по поставщику и службе
Private Function FindBySupplierAndService(DoverData As ParsedNaryad, OperData As Variant) As MatchResult
    Dim Result As MatchResult
    Dim i As Long
    Dim OperComment As String, OperSupplier As String
    Dim DoverService As String
    
    Result.FoundMatch = False
    Result.MatchScore = 0
    
    ' Извлекаем службу из комментария доверенности
    DoverService = ExtractServiceFromComment(DoverData.OriginalText)
    
    Debug.Print ">>> ПОИСК ПО ПОСТАВЩИКУ+СЛУЖБЕ ДЛЯ: № '" & DoverData.DocumentNumber & "' служба: '" & DoverService & "'"
    
    ' Проверяем каждую операцию
    For i = 1 To UBound(OperData, 1)
        OperComment = Trim(CStr(OperData(i, 9)))
        OperSupplier = Trim(CStr(OperData(i, 6)))
        
        If OperComment <> "" And OperSupplier <> "" Then
            Debug.Print "    Проверяем операцию " & (i + 1) & ": " & Left(OperComment, 50) & "... Поставщик: " & OperSupplier
            
            ' КЛЮЧЕВАЯ ФУНКЦИЯ: Анализ совпадения номера, поставщика и службы
            Dim MatchScore As Double
            Dim MatchDetails As String
            MatchScore = AnalyzeSupplierServiceMatch(OperComment, OperSupplier, DoverData.DocumentNumber, DoverService, MatchDetails)
            
            ' Если найдено хорошее совпадение
            If MatchScore >= 60 Then ' Порог для второго прохода
                Result.FoundMatch = True
                Result.MatchScore = MatchScore
                Result.OperationRow = i + 1
                Result.OperationNumber = Trim(CStr(OperData(i, 3)))
                Result.MatchDetails = MatchDetails & " | 1С: " & OperComment
                
                On Error Resume Next
                Result.OperationDate = CDate(OperData(i, 2))
                On Error GoTo 0
                
                Debug.Print ">>> ? НАЙДЕНО ПОСТАВЩИКОМ+СЛУЖБОЙ! Строка " & (i + 1) & " Уверенность: " & MatchScore & "%"
                Debug.Print "    Детали: " & MatchDetails
                Exit For
            End If
        End If
    Next i
    
    If Not Result.FoundMatch Then
        Result.MatchDetails = "поставщик+служба не совпадают"
        Debug.Print ">>> ? ПОИСК ПО ПОСТАВЩИКУ+СЛУЖБЕ НЕ НАШЕЛ ПОДХОДЯЩИХ"
    End If
    
    FindBySupplierAndService = Result
End Function

' НОВАЯ ФУНКЦИЯ: Анализ совпадения номера, поставщика и службы
Private Function AnalyzeSupplierServiceMatch(OperComment As String, OperSupplier As String, DoverNumber As String, DoverService As String, ByRef MatchDetails As String) As Double
    Dim TotalScore As Double
    Dim NumberScore As Double, SupplierScore As Double, ServiceScore As Double
    
    TotalScore = 0
    MatchDetails = ""
    
    ' КРИТЕРИЙ 1: Совпадение номера в комментарии операции (0-40 баллов)
    NumberScore = CheckNumberInComment(OperComment, DoverNumber)
    TotalScore = TotalScore + NumberScore
    
    ' КРИТЕРИЙ 2: Совпадение поставщика (0-40 баллов) - ОСНОВНОЙ КРИТЕРИЙ!
    SupplierScore = CheckSupplierMatch(OperSupplier, DoverService)
    TotalScore = TotalScore + SupplierScore
    
    ' КРИТЕРИЙ 3: Совпадение службы в комментарии (0-20 баллов) - БОНУС!
    If DoverService <> "" Then
        ServiceScore = CheckServiceInComment(OperComment, DoverService)
        TotalScore = TotalScore + ServiceScore
    End If
    
    ' Формируем детали
    MatchDetails = "номер:" & Format(NumberScore, "0") & "% поставщик:" & Format(SupplierScore, "0") & "%"
    If DoverService <> "" Then
        MatchDetails = MatchDetails & " служба:" & Format(ServiceScore, "0") & "%"
    End If
    
    Debug.Print "        Оценка: номер=" & NumberScore & " поставщик=" & SupplierScore & " служба=" & ServiceScore & " ИТОГО=" & TotalScore
    
    AnalyzeSupplierServiceMatch = TotalScore
End Function

' НОВАЯ ФУНКЦИЯ: Извлечение службы из комментария доверенности
Private Function ExtractServiceFromComment(Comment As String) As String
    Dim Words() As String
    Dim FirstWord As String
    Dim KnownServices As Variant
    Dim i As Long
    
    ExtractServiceFromComment = ""
    
    If Trim(Comment) = "" Then Exit Function
    
    Words = Split(Trim(Comment), " ")
    If UBound(Words) >= 0 Then
        FirstWord = UCase(Trim(Words(0)))
        
        ' Список известных служб
        KnownServices = Array("МС", "ВС", "ПС", "ТС", "ИС", "СС", "АС", "ГСМ", "РТС", "ВВС", "ВМФ", "РВСН")
        
        For i = 0 To UBound(KnownServices)
            If FirstWord = CStr(KnownServices(i)) Then
                ExtractServiceFromComment = FirstWord
                Debug.Print "    Извлечена служба: '" & FirstWord & "'"
                Exit Function
            End If
        Next i
        
        Debug.Print "    Служба не определена из: '" & FirstWord & "'"
    End If
End Function

' НОВАЯ ФУНКЦИЯ: Проверка номера в комментарии операции
Private Function CheckNumberInComment(OperComment As String, DoverNumber As String) As Double
    Dim Score As Double
    
    Score = 0
    
    ' Точное вхождение номера
    If InStr(1, OperComment, DoverNumber, vbTextCompare) > 0 Then
        Score = 40
        Debug.Print "        Номер найден точно: +40"
    ' Проверяем частичное совпадение для сложных номеров
    ElseIf InStr(DoverNumber, "/") > 0 Or InStr(DoverNumber, "-") > 0 Then
        Score = CheckPartialNumberMatch(OperComment, DoverNumber)
        If Score > 0 Then
            Debug.Print "        Номер найден частично: +" & Score
        End If
    Else
        Debug.Print "        Номер не найден: +0"
    End If
    
    CheckNumberInComment = Score
End Function

' НОВАЯ ФУНКЦИЯ: Проверка совпадения поставщика
Private Function CheckSupplierMatch(OperSupplier As String, DoverService As String) As Double
    Dim Score As Double
    Dim CleanOperSupplier As String
    
    Score = 0
    CleanOperSupplier = CleanSupplierName(OperSupplier)
    
    Debug.Print "        Поставщик операции: '" & CleanOperSupplier & "'"
    
    ' Пока упрощенная логика - можно расширить
    ' Основная логика сопоставления поставщика будет зависеть от
    ' конкретных данных в ваших файлах
    
    ' Временно даем базовую оценку
    If Len(CleanOperSupplier) >= 3 Then
        Score = 30 ' Базовая оценка за непустого поставщика
        Debug.Print "        Поставщик присутствует: +30"
    End If
    
    CheckSupplierMatch = Score
End Function

' НОВАЯ ФУНКЦИЯ: Проверка службы в комментарии операции
Private Function CheckServiceInComment(OperComment As String, DoverService As String) As Double
    Dim Score As Double
    
    Score = 0
    
    If DoverService <> "" Then
        If InStr(1, UCase(OperComment), DoverService, vbTextCompare) > 0 Then
            Score = 20 ' Бонус за совпадение службы
            Debug.Print "        Служба найдена в комментарии: +20"
        End If
    End If
    
    CheckServiceInComment = Score
End Function

' ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: Очистка названия поставщика
Private Function CleanSupplierName(SupplierName As String) As String
    Dim Result As String
    
    Result = Trim(SupplierName)
    
    ' Убираем кавычки
    Result = Replace(Result, """", "")
    Result = Replace(Result, "'", "")
    
    ' Убираем лишние пробелы
    Do While InStr(Result, "  ") > 0
        Result = Replace(Result, "  ", " ")
    Loop
    
    CleanSupplierName = Result
End Function

' ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: Проверка частичного совпадения номера
Private Function CheckPartialNumberMatch(OperComment As String, DoverNumber As String) As Double
    Dim Score As Double
    Dim NumberParts() As String
    Dim i As Long
    Dim FoundParts As Long
    
    Score = 0
    FoundParts = 0
    
    ' Разбиваем номер на части
    If InStr(DoverNumber, "/") > 0 Then
        NumberParts = Split(DoverNumber, "/")
    ElseIf InStr(DoverNumber, "-") > 0 Then
        NumberParts = Split(DoverNumber, "-")
    Else
        CheckPartialNumberMatch = 0
        Exit Function
    End If
    
    ' Проверяем каждую часть
    For i = 0 To UBound(NumberParts)
        If Len(Trim(NumberParts(i))) >= 2 Then ' Игнорируем слишком короткие части
            If InStr(1, OperComment, Trim(NumberParts(i)), vbTextCompare) > 0 Then
                FoundParts = FoundParts + 1
            End If
        End If
    Next i
    
    ' Оцениваем результат
    If FoundParts > 0 Then
        Score = (FoundParts / (UBound(NumberParts) + 1)) * 25 ' Максимум 25 баллов за частичное совпадение
    End If
    
    CheckPartialNumberMatch = Score
End Function










' =============================================
' ТЕСТОВЫЕ ФУНКЦИИ
' =============================================

' ТЕСТОВЫЕ ФУНКЦИИ - ДОБАВИТЬ ДЛЯ ОТЛАДКИ
Public Sub TestProblematicNumbers()
    Debug.Print "=== ТЕСТ ПРОБЛЕМНЫХ НОМЕРОВ ==="
    Debug.Print ""
    
    Call QuickTest("901/243450", "25.10.2024", "ГСМ Наряд № 901/243450 от 25.10.2024")
    Call QuickTest("561/22/176-а-25", "25.10.2024", "Наряд № 561/22/176-а-25 от 25.10.2024")
    Call QuickTest("ПС-4", "25.10.2024", "Наряд № ПС-4 от 25.10.2024")
    Call QuickTest("60/ЮВО", "25.10.2024", "Наряд № 60/ЮВО от 25.10.2024")
    
    Debug.Print "=== КОНЕЦ ТЕСТА ==="
End Sub

Public Sub QuickTest(TestNumber As String, TestDate As String, TestComment As String)
    Debug.Print "ТЕСТ: " & TestNumber
    
    Dim Patterns() As String
    Call CreateFlexibleTextPatterns(TestNumber, TestDate, Patterns)
    
    Dim MatchedPattern As String, MatchedText As String
    
    If FindPatternInText(TestComment, Patterns, MatchedPattern, MatchedText) Then
        Debug.Print "? НАЙДЕНО!"
    Else
        Debug.Print "? НЕ НАЙДЕНО"
    End If
End Sub


