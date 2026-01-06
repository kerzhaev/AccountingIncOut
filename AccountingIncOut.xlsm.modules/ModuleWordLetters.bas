Attribute VB_Name = "ModuleWordLetters"
'==============================================
' МОДУЛЬ ФОРМИРОВАНИЯ ПИСЕМ ПО ДОВЕРЕННОСТЯМ
' Версия: 1.2.3
' Дата: 07.09.2025
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================
Option Explicit

' =============================================
' СТРУКТУРЫ ДАННЫХ
' =============================================
Public Type AddressInfo
    CorrespondentName As String
    RecipientName As String
    Street As String
    City As String
    District As String
    Region As String
    PostalCode As String
    FullAddress As String
End Type

Public Type DoverennostInfo
    RowNumber As Long
    FIO As String
    VoyskovayaChast As String
    DoverennostNumber As String
    DoverennostDate As Date
    Comment As String
    Correspondent As String
End Type

Public Type CorrespondentGroup
    CorrespondentName As String
    Address As AddressInfo
    Doverennosti() As DoverennostInfo
    Count As Long
    AddressFound As Boolean
End Type

Public Type FIOGroup
    FIOName As String
    Doverennosti() As DoverennostInfo
    Count As Long
End Type

Public Type LetterInfo
    Number As String          ' Номер письма (7/25)
    Date As String           ' Дата письма (25.08.2025)
    FormattedDate As String  ' Дата письма (25 августа 2025 г.)
End Type


' =============================================
' КОНСТАНТЫ СТОЛБЦОВ
' =============================================
Private Const COL_DATE As String = "ДАТА"
Private Const COL_NUMBER As String = "НОМЕР"
Private Const COL_FIO As String = "ПОДОТЧЕТНОЕ ЛИЦО"
Private Const COL_ORGANIZATION As String = "ОРГАНИЗАЦИЯ"
Private Const COL_CORRESPONDENT As String = "ПОСТАВЩИК"
Private Const COL_COMMENT As String = "КОММЕНТАРИЙ"
Private Const COL_SEARCH_RESULT As String = "ТРЕХПРОХОДНЫЙ ПОИСК"


' =============================================
' ПЕРЕМЕННЫЕ ДЛЯ УЧЕТА ИСХОДЯЩИХ ПИСЕМ v1.2.0
' =============================================
Private CurrentLetterNumber As Long
Private LetterPrefix As String
Private CurrentDate As Date
Private FormattedCurrentDate As String


' =============================================
' ОСНОВНАЯ ПРОЦЕДУРА v1.1.0
' =============================================
Public Sub GenerateWordLetters()
    Dim FileDovernnosti As String
    Dim FileAddresses As String
    Dim SaveFolder As String
    Dim WbDover As Workbook
    Dim WbAddr As Workbook
    
    On Error GoTo GenerateError
    
    Debug.Print "=== НАЧАЛО ФОРМИРОВАНИЯ ПИСЕМ ПО ДОВЕРЕННОСТЯМ v1.1.0 ==="
    
    ' Выбор файла доверенностей
    FileDovernnosti = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        "Выберите файл с обработанными доверенностями")
    If FileDovernnosti = "False" Then Exit Sub
    
    ' Выбор файла адресов
    FileAddresses = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,Excel Macro Files (*.xlsm),*.xlsm,CSV Files (*.csv),*.csv", , _
        "Выберите файл с адресами корреспондентов")
    If FileAddresses = "False" Then Exit Sub
    
    ' Выбор папки для сохранения
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Выберите папку для сохранения писем"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SaveFolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Проверка шаблона
    Dim TemplatePath As String
    TemplatePath = ThisWorkbook.Path & "\ШаблонПисьмаДоверенность.docx"
    If dir(TemplatePath) = "" Then
        MsgBox "ОШИБКА: Шаблон не найден: " & TemplatePath & vbCrLf & vbCrLf & _
               "Поместите файл шаблона в папку с макросом.", vbCritical, "Шаблон не найден"
        Exit Sub
    End If
    
    ' Открытие файлов
    Set WbDover = Workbooks.Open(FileDovernnosti, ReadOnly:=False)
    Debug.Print "? Файл доверенностей открыт для записи"

    Set WbAddr = Workbooks.Open(FileAddresses, ReadOnly:=True)
    
    ' Проверка структуры файлов
    If Not ValidateFileStructure(WbDover, WbAddr) Then
        WbAddr.Close False
        WbDover.Close False
        Exit Sub
    End If
    

    ' Инициализация системы исходящих писем
    If Not InitializeLetterNumbering() Then
        WbAddr.Close False
        WbDover.Close False
        Exit Sub
    End If

' Основная обработка
Call ProcessWordLetterGeneration(WbDover, WbAddr, SaveFolder)

    
    ' Закрытие файлов
    WbAddr.Close False
    WbDover.Close False
    
    Debug.Print "=== ФОРМИРОВАНИЕ ПИСЕМ ПО ДОВЕРЕННОСТЯМ ЗАВЕРШЕНО ==="
    
    Exit Sub
    
GenerateError:
    On Error Resume Next
    If Not WbAddr Is Nothing Then WbAddr.Close False
    If Not WbDover Is Nothing Then WbDover.Close False
    
    Debug.Print "Ошибка формирования писем: " & Err.description
    MsgBox "Ошибка: " & Err.description, vbCritical
End Sub

' =============================================
' ПРОВЕРКА СТРУКТУРЫ ФАЙЛОВ v1.1.0
' =============================================
Private Function ValidateFileStructure(WbDover As Workbook, WbAddr As Workbook) As Boolean
    Dim WsDover As Worksheet
    Dim WsAddr As Worksheet
    Dim MissingColumns As String
    Dim i As Long
    
    ValidateFileStructure = False
    Set WsDover = WbDover.Worksheets(1)
    
    ' Поиск листа "Адреса"
    Dim AddrSheetFound As Boolean
    AddrSheetFound = False
    For i = 1 To WbAddr.Worksheets.Count
        If UCase(WbAddr.Worksheets(i).Name) = "АДРЕСА" Then
            Set WsAddr = WbAddr.Worksheets(i)
            AddrSheetFound = True
            Exit For
        End If
    Next i
    
    If Not AddrSheetFound Then
        MsgBox "ОШИБКА: Лист 'Адреса' не найден в файле адресов!" & vbCrLf & vbCrLf & _
               "Убедитесь, что лист называется точно 'Адреса'", vbCritical, "Структура файла нарушена"
        Exit Function
    End If
    
    ' Проверка обязательных столбцов в файле доверенностей
    MissingColumns = ""
    If FindColumnByName(WsDover, COL_DATE) = 0 Then MissingColumns = MissingColumns & "• " & COL_DATE & vbCrLf
    If FindColumnByName(WsDover, COL_NUMBER) = 0 Then MissingColumns = MissingColumns & "• " & COL_NUMBER & vbCrLf
    If FindColumnByName(WsDover, COL_FIO) = 0 Then MissingColumns = MissingColumns & "• " & COL_FIO & vbCrLf
    If FindColumnByName(WsDover, COL_ORGANIZATION) = 0 Then MissingColumns = MissingColumns & "• " & COL_ORGANIZATION & vbCrLf
    If FindColumnByName(WsDover, COL_CORRESPONDENT) = 0 Then MissingColumns = MissingColumns & "• " & COL_CORRESPONDENT & vbCrLf
    If FindColumnByName(WsDover, COL_SEARCH_RESULT) = 0 Then MissingColumns = MissingColumns & "• " & COL_SEARCH_RESULT & vbCrLf
    
    If Len(MissingColumns) > 0 Then
        MsgBox "ОШИБКА: В файле доверенностей отсутствуют обязательные столбцы:" & vbCrLf & vbCrLf & _
               MissingColumns & vbCrLf & _
               "Обработка прекращена. Проверьте структуру файла.", vbCritical, "Отсутствуют обязательные столбцы"
        Exit Function
    End If
    
    ' Проверка столбцов в файле адресов
    If WsAddr.Cells(1, 1).value = "" Or WsAddr.Cells(1, 2).value = "" Then
        MsgBox "ОШИБКА: Неверная структура файла адресов!" & vbCrLf & vbCrLf & _
               "Ожидаемые столбцы:" & vbCrLf & _
               "1. Наименование адресата" & vbCrLf & _
               "2. Улица, дом, квартира" & vbCrLf & _
               "3. Населенный пункт" & vbCrLf & _
               "4. Район" & vbCrLf & _
               "5. Область/край/республика" & vbCrLf & _
               "6. Почтовый индекс", vbCritical, "Структура файла нарушена"
        Exit Function
    End If
    
    ValidateFileStructure = True
    Debug.Print "? Структура файлов проверена успешно"
End Function

' =============================================
' ОСНОВНАЯ ЛОГИКА ОБРАБОТКИ v1.1.0
' =============================================
Private Sub ProcessWordLetterGeneration(WbDover As Workbook, WbAddr As Workbook, SaveFolder As String)
    Dim WsDover As Worksheet, WsAddr As Worksheet
    Dim LastRowDover As Long, LastRowAddr As Long
    Dim DoverData As Variant, AddrData As Variant
    
    Dim UnmatchedDoverennosti() As DoverennostInfo
    Dim UnmatchedCount As Long
    Dim CorrespondentGroups() As CorrespondentGroup
    Dim GroupCount As Long
    Dim NotFoundCorrespondents As String
    
    Dim i As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set WsDover = WbDover.Worksheets(1)
    Set WsAddr = WbAddr.Worksheets("Адреса")
    
    LastRowDover = WsDover.Cells(WsDover.Rows.Count, 1).End(xlUp).Row
    LastRowAddr = WsAddr.Cells(WsAddr.Rows.Count, 1).End(xlUp).Row
    
    ' Загрузка данных
    DoverData = WsDover.Range("A2:Z" & LastRowDover).Value2
    AddrData = WsAddr.Range("A2:F" & LastRowAddr).Value2
    
    Debug.Print "Загружено доверенностей: " & (LastRowDover - 1)
    Debug.Print "Загружено адресов: " & (LastRowAddr - 1)
    
    ' Фильтрация несопоставленных доверенностей
    Application.StatusBar = "Фильтрация несопоставленных доверенностей..."
    Call FilterUnmatchedDoverennosti(WsDover, DoverData, UnmatchedDoverennosti, UnmatchedCount)
    
    Debug.Print "Найдено несопоставленных доверенностей: " & UnmatchedCount
    
    If UnmatchedCount = 0 Then
        MsgBox "Несопоставленные доверенности не найдены!", vbInformation
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' Группировка по корреспондентам
    Application.StatusBar = "Группировка по корреспондентам..."
    Call GroupByCorrespondents(UnmatchedDoverennosti, UnmatchedCount, CorrespondentGroups, GroupCount)
    
    Debug.Print "Создано групп корреспондентов: " & GroupCount
    
    ' Поиск адресов корреспондентов
    Application.StatusBar = "Поиск адресов корреспондентов..."
    Call FindCorrespondentAddresses(CorrespondentGroups, GroupCount, AddrData, NotFoundCorrespondents)
    
    ' Создание документов Word
    Application.StatusBar = "Создание документов Word..."
    Call CreateWordDocuments(CorrespondentGroups, GroupCount, SaveFolder, WbDover)

    
    ' Создание отчета о не найденных адресах
    If Len(NotFoundCorrespondents) > 0 Then
        Call CreateMissingAddressesReport(NotFoundCorrespondents, SaveFolder)
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Показ результатов
    Call ShowGenerationResults(GroupCount, NotFoundCorrespondents, SaveFolder)
End Sub

' =============================================
' ИСПРАВЛЕННАЯ ФИЛЬТРАЦИЯ v1.1.0
' =============================================
Private Sub FilterUnmatchedDoverennosti(WsDover As Worksheet, DoverData As Variant, ByRef UnmatchedList() As DoverennostInfo, ByRef Count As Long)
    Dim i As Long
    Dim SearchCol As Long
    Dim DateCol As Long, NumberCol As Long, FIOCol As Long
    Dim OrganizationCol As Long, CorrespondentCol As Long, CommentCol As Long
    Dim SearchValue As String
    
    Count = 0
    ReDim UnmatchedList(UBound(DoverData, 1))
    
    ' Определение столбцов
    SearchCol = FindColumnByName(WsDover, COL_SEARCH_RESULT)
    DateCol = FindColumnByName(WsDover, COL_DATE)
    NumberCol = FindColumnByName(WsDover, COL_NUMBER)
    FIOCol = FindColumnByName(WsDover, COL_FIO)
    OrganizationCol = FindColumnByName(WsDover, COL_ORGANIZATION)
    CorrespondentCol = FindColumnByName(WsDover, COL_CORRESPONDENT)
    CommentCol = FindColumnByName(WsDover, COL_COMMENT)
    
    Debug.Print "=== НАЧИНАЕМ ФИЛЬТРАЦИЮ v1.1.0 ==="
    Debug.Print "Столбец поиска: " & SearchCol
    
    ' Фильтрация записей
    For i = 1 To UBound(DoverData, 1)
        If i Mod 50 = 0 Then
            Application.StatusBar = "Фильтрация строки " & i & " из " & UBound(DoverData, 1)
        End If
        
        SearchValue = ""
        If SearchCol > 0 And SearchCol <= UBound(DoverData, 2) Then
            SearchValue = Trim(UCase(CStr(DoverData(i, SearchCol))))
        End If
        
        If SearchValue = "НЕ НАЙДЕНО" Then
    ' Проверка на пустые обязательные поля
    If FIOCol > 0 And OrganizationCol > 0 And NumberCol > 0 And CorrespondentCol > 0 Then
        If Trim(CStr(DoverData(i, FIOCol))) <> "" And _
           Trim(CStr(DoverData(i, OrganizationCol))) <> "" And _
           Trim(CStr(DoverData(i, NumberCol))) <> "" And _
           Trim(CStr(DoverData(i, CorrespondentCol))) <> "" Then
           
            ' НОВАЯ ПРОВЕРКА: Проверяем 3-месячный период для существующих писем
            If ShouldIncludeDoverennost(WsDover, i + 1) Then
                With UnmatchedList(Count)
                    .RowNumber = i + 1
                    .FIO = Trim(CStr(DoverData(i, FIOCol)))
                    .VoyskovayaChast = Trim(CStr(DoverData(i, OrganizationCol)))
                    .DoverennostNumber = Trim(CStr(DoverData(i, NumberCol)))
                    .Correspondent = Trim(CStr(DoverData(i, CorrespondentCol)))
                    
                    ' Комментарий
                    If CommentCol > 0 And CommentCol <= UBound(DoverData, 2) Then
                        .Comment = Trim(CStr(DoverData(i, CommentCol)))
                    Else
                        .Comment = ""
                    End If
                    
                    ' Парсинг даты
                    On Error Resume Next
                    If DateCol > 0 And DateCol <= UBound(DoverData, 2) Then
                        If IsDate(DoverData(i, DateCol)) Then
                            .DoverennostDate = CDate(DoverData(i, DateCol))
                        Else
                            .DoverennostDate = Date
                        End If
                    Else
                        .DoverennostDate = Date
                    End If
                    On Error GoTo 0
                End With
                
                Count = Count + 1
                Debug.Print "? Добавлена: " & UnmatchedList(Count - 1).FIO & " (" & UnmatchedList(Count - 1).DoverennostNumber & ")"
            Else
                Debug.Print "? Пропущена (недавнее письмо): " & Trim(CStr(DoverData(i, FIOCol)))
            End If
        End If
    End If
End If

    Next i
    
    If Count > 0 Then
        ReDim Preserve UnmatchedList(Count - 1)
    End If
    Debug.Print "=== ИТОГ ФИЛЬТРАЦИИ: " & Count & " записей ==="
End Sub

' =============================================
' ИСПРАВЛЕННОЕ СОЗДАНИЕ ДОКУМЕНТОВ v1.1.0
' =============================================
Private Sub CreateWordDocuments(Groups() As CorrespondentGroup, GroupCount As Long, SaveFolder As String, WbDover As Workbook)

    Dim i As Long
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim TemplatePath As String
    Dim SavePath As String
    Dim CurrentDate As String
    Dim ShortDate As String           ' ДОБАВИТЬ ЭТУ СТРОКУ
    Dim LetterInfo As LetterInfo      ' И ЭТУ ТОЖЕ
    Dim FileNameNumber As String      ' И ЭТУ
    
    CurrentDate = Format(Now, "dd.mm.yyyy")
    TemplatePath = ThisWorkbook.Path & "\ШаблонПисьмаДоверенность.docx"
    
    Debug.Print "=== СОЗДАНИЕ ДОКУМЕНТОВ WORD v1.1.0 ==="
    
    ' Проверка и тестирование шаблона
    On Error GoTo TemplateError
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = False
    
    ' Тест открытия шаблона
    Set WordDoc = WordApp.Documents.Open(TemplatePath, ReadOnly:=True)
    WordDoc.Close False
    Set WordDoc = Nothing
    
    Debug.Print "? Шаблон успешно протестирован"
    
    For i = 0 To GroupCount - 1
        If Groups(i).AddressFound Then
            Application.StatusBar = "Создание документа " & (i + 1) & " из " & GroupCount & ": " & Groups(i).CorrespondentName
            
            On Error GoTo DocumentError
            
            ' Открытие шаблона
            Set WordDoc = WordApp.Documents.Open(TemplatePath)
            
            ' ИСПРАВЛЕННАЯ ПОСЛЕДОВАТЕЛЬНОСТЬ: сначала адрес, потом таблица
            Call FillWordDocumentFixed(WordDoc, Groups(i))
            
            ' Сохранение документа
            
    ' Формируем имя файла с номером письма v1.2.0
    LetterInfo = GetCurrentLetterInfo()
    FileNameNumber = Replace(LetterInfo.Number, "/", "_") ' Заменяем / на
   ShortDate = FormatShortDate(CurrentDate)  ' Используем введенную дату




    
    SavePath = SaveFolder & "\" & CleanFileName(Groups(i).CorrespondentName) & "_" & FileNameNumber & " от " & ShortDate & ".docx"



    WordDoc.SaveAs2 SavePath
    WordDoc.Close
    
    Debug.Print "? Документ создан: " & SavePath
    
    ' ОБНОВЛЯЕМ ФАЙЛ ДОВЕРЕННОСТЕЙ ПОСЛЕ УСПЕШНОГО СОЗДАНИЯ v1.2.0
    Call UpdateDoverennostiFile(WbDover, Groups(i), LetterInfo)

    
    ' Увеличиваем номер письма для следующего документа
    CurrentLetterNumber = CurrentLetterNumber + 1
    Debug.Print "? Следующий номер письма: " & LetterPrefix & CurrentLetterNumber

        Else
            Debug.Print "? Пропущен корреспондент без адреса: " & Groups(i).CorrespondentName
        End If
    Next i
    
    WordApp.Quit
    Set WordApp = Nothing
    
    Debug.Print "? Создание документов Word завершено"
    Exit Sub
    
TemplateError:
    On Error Resume Next
    If Not WordApp Is Nothing Then WordApp.Quit
    MsgBox "ОШИБКА: Шаблон поврежден или недоступен!" & vbCrLf & vbCrLf & _
           "Проверьте файл: " & TemplatePath & vbCrLf & vbCrLf & _
           "Ошибка: " & Err.description, vbCritical, "Проблема с шаблоном"
    Exit Sub
    
DocumentError:
    On Error Resume Next
    If Not WordDoc Is Nothing Then WordDoc.Close False
    Debug.Print "? Ошибка создания документа для: " & Groups(i).CorrespondentName & " - " & Err.description
    Resume Next
End Sub

' =============================================
' ЗАПОЛНЕНИЕ ДОКУМЕНТА WORD v1.1.1
' =============================================
Private Sub FillWordDocumentFixed(WordDoc As Object, GroupData As CorrespondentGroup)
    Dim WordTable As Object
    Dim i As Long, j As Long
    Dim RowIndex As Long
    
    Debug.Print "=== ЗАПОЛНЕНИЕ ДОКУМЕНТА v1.1.1 ==="
    Debug.Print "Корреспондент: " & GroupData.CorrespondentName
    
    On Error GoTo FillError
    

    ' ШАГ 1: ЗАМЕНЯЕМ ЗАКЛАДКИ (НЕ ТЕКСТ!) v1.1.6
    With GroupData.Address
    Call ReplaceWordBookmarkCorrect(WordDoc, "НаименованиеПолучателя", .CorrespondentName)
    Call ReplaceWordBookmarkCorrect(WordDoc, "АдресПолучателя", .FullAddress)
    Call ReplaceWordBookmarkCorrect(WordDoc, "КОРРЕСПОНДЕНТ", LCase(.CorrespondentName))
    
    ' НОВЫЕ ПЛЕЙСХОЛДЕРЫ ДЛЯ ИСХОДЯЩИХ ПИСЕМ v1.2.0
    Dim CurrentLetterInfo As LetterInfo
    CurrentLetterInfo = GetCurrentLetterInfo()
    Call ReplaceWordBookmarkCorrect(WordDoc, "НомерИсходящего", CurrentLetterInfo.Number)
    Call ReplaceWordBookmarkCorrect(WordDoc, "ДатаИсходящего", CurrentLetterInfo.FormattedDate)
    
    Debug.Print "? Плейсхолдеры письма заменены:"
    Debug.Print "  НомерИсходящего: " & CurrentLetterInfo.Number
    Debug.Print "  ДатаИсходящего: " & CurrentLetterInfo.FormattedDate

    End With
    
    ' ШАГ 2: СОЗДАЕМ ГРУППЫ ПО ФИО
    Dim FIOGroups() As FIOGroup
    Dim FIOGroupCount As Long
    Call CreateFIOGroups(GroupData.Doverennosti, GroupData.Count, FIOGroups, FIOGroupCount)
    
    ' ШАГ 3: ПОДСЧИТЫВАЕМ СТРОКИ ДЛЯ ТАБЛИЦЫ
' ШАГ 3: ПОДСЧИТЫВАЕМ СТРОКИ ДЛЯ ТАБЛИЦЫ
    Dim TotalDataRows As Long
    For i = 0 To FIOGroupCount - 1
        TotalDataRows = TotalDataRows + FIOGroups(i).Count
    Next i

    
    ' ШАГ 4: ИЩЕМ МЕСТО ДЛЯ ТАБЛИЦЫ
    Dim InsertRange As Object
    Set InsertRange = FindInsertionPoint(WordDoc)
    
    If InsertRange Is Nothing Then
        Debug.Print "? Место для таблицы не найдено, добавляем в конец"
        Set InsertRange = WordDoc.Range
        InsertRange.Collapse 0 ' wdCollapseEnd
        InsertRange.InsertParagraphBefore
        InsertRange.InsertParagraphBefore
        InsertRange.Collapse 0 ' wdCollapseEnd
    End If
    
    ' ШАГ 5: СОЗДАНИЕ ТАБЛИЦЫ
    Debug.Print "Создаем таблицу " & (TotalDataRows + 1) & "x4"
    Set WordTable = WordDoc.Tables.Add(InsertRange, TotalDataRows + 1, 4)
    
    ' ШАГ 6: ФОРМАТИРОВАНИЕ ТАБЛИЦЫ С ВЕРТИКАЛЬНЫМ ЦЕНТРИРОВАНИЕМ v1.1.4
With WordTable
    .Borders.Enable = True
    .Range.Font.Size = 12
    .Range.Font.Name = "Times New Roman"
    .Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter - горизонтальное центрирование
    
    ' ВЕРТИКАЛЬНОЕ ЦЕНТРИРОВАНИЕ ВСЕХ ЯЧЕЕК
    Dim CellRow As Long, CellCol As Long
    For CellRow = 1 To .Rows.Count
        For CellCol = 1 To .Columns.Count
            .cell(CellRow, CellCol).VerticalAlignment = 1 ' wdCellAlignVerticalCenter
        Next CellCol
    Next CellRow
    
    ' Заголовки
    .cell(1, 1).Range.Text = "ФИО"
    .cell(1, 2).Range.Text = "Войсковая часть получатель"
    .cell(1, 3).Range.Text = "Номер доверенности и её дата"
    .cell(1, 4).Range.Text = "Основание"
    
    ' Форматирование заголовков
    .Rows(1).Range.Font.Bold = True
    .Rows(1).Shading.BackgroundPatternColor = RGB(230, 230, 230)
    
    ' Ширина столбцов
    .Columns(1).Width = WordDoc.Application.CentimetersToPoints(4)
    .Columns(2).Width = WordDoc.Application.CentimetersToPoints(5)
    .Columns(3).Width = WordDoc.Application.CentimetersToPoints(4)
    .Columns(4).Width = WordDoc.Application.CentimetersToPoints(3)
End With

    
' ШАГ 7: ЗАПОЛНЕНИЕ ДАННЫМИ С ОБЪЕДИНЕНИЕМ ЯЧЕЕК ФИО v1.1.3
RowIndex = 2
For i = 0 To FIOGroupCount - 1
    Dim StartRow As Long
    StartRow = RowIndex
    
    ' Заполняем все доверенности для одного ФИО
    For j = 0 To FIOGroups(i).Count - 1
        With FIOGroups(i).Doverennosti(j)
            ' ФИО заполняем только в первой строке группы
            If j = 0 Then
                WordTable.cell(RowIndex, 1).Range.Text = .FIO
            Else
                WordTable.cell(RowIndex, 1).Range.Text = ""  ' Пустая ячейка для остальных
            End If
            
            WordTable.cell(RowIndex, 2).Range.Text = .VoyskovayaChast
            
            Dim DateText As String
            If .DoverennostDate > DateSerial(1900, 1, 2) Then
                DateText = .DoverennostNumber & " от " & Format(.DoverennostDate, "dd.mm.yyyy")
            Else
                DateText = .DoverennostNumber
            End If
            
            WordTable.cell(RowIndex, 3).Range.Text = DateText
            WordTable.cell(RowIndex, 4).Range.Text = IIf(Len(.Comment) > 0, .Comment, "Получение материальных ценностей")
        End With
        RowIndex = RowIndex + 1
    Next j
    
    ' ОБЪЕДИНЯЕМ ЯЧЕЙКИ ФИО, если у человека несколько доверенностей
    If FIOGroups(i).Count > 1 Then
        Dim EndRow As Long
        EndRow = RowIndex - 1
        
        On Error Resume Next
        WordTable.cell(StartRow, 1).Merge WordTable.cell(EndRow, 1)
        WordTable.cell(StartRow, 1).VerticalAlignment = 1  ' wdCellAlignVerticalCenter
        On Error GoTo 0
        
        Debug.Print "? Объединены ячейки ФИО: " & FIOGroups(i).FIOName & " (строки " & StartRow & "-" & EndRow & ")"
    End If
    
    Debug.Print "? Заполнено ФИО: " & FIOGroups(i).FIOName & " (" & FIOGroups(i).Count & " доверенностей)"
Next i

    
    Debug.Print "? Документ заполнен успешно! Строк: " & TotalDataRows
    Exit Sub
    
FillError:
    Debug.Print "? Ошибка заполнения: " & Err.description
    MsgBox "Ошибка заполнения документа: " & Err.description, vbCritical
End Sub


' =============================================
' ПРАВИЛЬНАЯ ЗАМЕНА ЗАКЛАДОК v1.1.1
' =============================================
Private Sub ReplaceWordBookmarkCorrect(WordDoc As Object, BookmarkName As String, NewText As String)
    Dim Success As Boolean
    Success = False
    
    Debug.Print "=== ЗАМЕНА ЗАКЛАДКИ: '" & BookmarkName & "' -> '" & NewText & "'"
    
    ' СПОСОБ 1: Прямая работа с закладками
    On Error Resume Next
    If WordDoc.Bookmarks.Exists(BookmarkName) Then
        WordDoc.Bookmarks(BookmarkName).Select
        WordDoc.Application.Selection.TypeText NewText
        Success = True
        Debug.Print "? Заменено через закладку: " & BookmarkName
    End If
    On Error GoTo 0
    
    ' СПОСОБ 2: Работа через Range закладки
    If Not Success Then
        On Error Resume Next
        If WordDoc.Bookmarks.Exists(BookmarkName) Then
            Dim BookmarkRange As Object
            Set BookmarkRange = WordDoc.Bookmarks(BookmarkName).Range
            BookmarkRange.Text = NewText
            ' Восстанавливаем закладку после замены текста
            WordDoc.Bookmarks.Add BookmarkName, BookmarkRange
            Success = True
            Debug.Print "? Заменено через Range закладки: " & BookmarkName
        End If
        On Error GoTo 0
    End If
    
    ' СПОСОБ 3: Поиск и замена как резервный вариант
    If Not Success Then
        On Error Resume Next
        With WordDoc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = BookmarkName
            .Replacement.Text = NewText
            .Forward = True
            .Wrap = 1 ' wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = True
            .MatchWildcards = False
            
            If .Execute(Replace:=2) Then ' wdReplaceAll
                Success = True
                Debug.Print "? Заменено через поиск: " & BookmarkName
            End If
        End With
        On Error GoTo 0
    End If
    
    If Success Then
        Debug.Print "? УСПЕШНО заменено: " & BookmarkName
    Else
        Debug.Print "? НЕ УДАЛОСЬ заменить: " & BookmarkName
        ' Дополнительная диагностика
        Call DiagnoseBookmarkIssues(WordDoc, BookmarkName)
    End If
End Sub


' =============================================
' ПОИСК ТОЧКИ ВСТАВКИ ТАБЛИЦЫ v1.1.0
' =============================================
Private Function FindInsertionPoint(WordDoc As Object) As Object
    Dim SearchRange As Object
    
    Set SearchRange = WordDoc.Content
    
    ' Ищем текст "материальных ценностей:"
    With SearchRange.Find
        .ClearFormatting
        .Text = "материальных ценностей:"
        .Forward = True
        .Wrap = 0 ' wdFindStop
        
        If .Execute Then
            SearchRange.Collapse 0 ' wdCollapseEnd
            SearchRange.InsertAfter vbCrLf & vbCrLf
            SearchRange.Collapse 0 ' wdCollapseEnd
            Set FindInsertionPoint = SearchRange
            Debug.Print "? Найдено место для вставки таблицы"
            Exit Function
        End If
    End With
    
    ' Если не найдено, вставляем в конец
    Set SearchRange = WordDoc.Range
    SearchRange.Collapse 0 ' wdCollapseEnd
    SearchRange.InsertBefore vbCrLf & vbCrLf
    Set FindInsertionPoint = SearchRange
    Debug.Print "? Таблица будет вставлена в конец документа"
End Function

' =============================================
' ОСТАЛЬНЫЕ ФУНКЦИИ БЕЗ ИЗМЕНЕНИЙ
' =============================================
Private Sub GroupByCorrespondents(UnmatchedList() As DoverennostInfo, UnmatchedCount As Long, ByRef Groups() As CorrespondentGroup, ByRef GroupCount As Long)
    Dim i As Long, j As Long
    Dim CorrespondentName As String
    Dim GroupFound As Boolean
    
    GroupCount = 0
    ReDim Groups(UnmatchedCount)
    
    For i = 0 To UnmatchedCount - 1
        CorrespondentName = ExtractCorrespondentName(UnmatchedList(i).Correspondent)
        GroupFound = False
        
        For j = 0 To GroupCount - 1
            If UCase(Groups(j).CorrespondentName) = UCase(CorrespondentName) Then
                GroupFound = True
                Groups(j).Count = Groups(j).Count + 1
                ReDim Preserve Groups(j).Doverennosti(Groups(j).Count - 1)
                Groups(j).Doverennosti(Groups(j).Count - 1) = UnmatchedList(i)
                Exit For
            End If
        Next j
        
        If Not GroupFound Then
            With Groups(GroupCount)
                .CorrespondentName = CorrespondentName
                .Count = 1
                .AddressFound = False
                ReDim .Doverennosti(0)
                .Doverennosti(0) = UnmatchedList(i)
            End With
            GroupCount = GroupCount + 1
        End If
    Next i
    
    ReDim Preserve Groups(GroupCount - 1)
    
    ' Сортировка внутри групп
    For i = 0 To GroupCount - 1
        Call SortDoverennostiInGroup(Groups(i))
    Next i
End Sub

Private Sub FindCorrespondentAddresses(ByRef Groups() As CorrespondentGroup, GroupCount As Long, AddrData As Variant, ByRef NotFoundList As String)
    Dim i As Long, j As Long
    Dim CorrespondentName As String
    Dim AddressFound As Boolean
    
    NotFoundList = ""
    
    For i = 0 To GroupCount - 1
        CorrespondentName = Groups(i).CorrespondentName
        AddressFound = False
        
        For j = 1 To UBound(AddrData, 1)
            Dim AddressName As String
            AddressName = Trim(CStr(AddrData(j, 1)))
            
            If UCase(AddressName) = UCase(CorrespondentName) Then
                With Groups(i).Address
                    .CorrespondentName = CorrespondentName
                    .RecipientName = AddressName
                    .Street = Trim(CStr(AddrData(j, 2)))
                    .City = Trim(CStr(AddrData(j, 3)))
                    .District = Trim(CStr(AddrData(j, 4)))
                    .Region = Trim(CStr(AddrData(j, 5)))
                    .PostalCode = Trim(CStr(AddrData(j, 6)))
                    .FullAddress = BuildFullAddress(.PostalCode, .Region, .District, .City, .Street)
                End With
                
                Groups(i).AddressFound = True
                AddressFound = True
                Exit For
            End If
        Next j
        
        If Not AddressFound Then
            Groups(i).AddressFound = False
            If Len(NotFoundList) > 0 Then NotFoundList = NotFoundList & vbCrLf
            NotFoundList = NotFoundList & CorrespondentName
        End If
    Next i
End Sub

Private Function FindColumnByName(Ws As Worksheet, ColumnName As String) As Long
    Dim i As Long
    Dim HeaderValue As String
    
    Debug.Print "=== ПОИСК СТОЛБЦА: " & ColumnName & " ==="
    
    For i = 1 To 30
        HeaderValue = Trim(UCase(CStr(Ws.Cells(1, i).value)))
        Debug.Print "Столбец " & i & ": '" & HeaderValue & "'"
        
        If InStr(HeaderValue, UCase(ColumnName)) > 0 Then
            Debug.Print "? НАЙДЕН столбец " & ColumnName & " в позиции " & i
            FindColumnByName = i
            Exit Function
        End If
    Next i
    
    Debug.Print "? СТОЛБЕЦ " & ColumnName & " НЕ НАЙДЕН!"
    FindColumnByName = 0
End Function


Private Function ExtractCorrespondentName(FullCorrespondent As String) As String
    Dim result As String
    Dim BracketPos As Long
    Dim i As Long
    Dim InNumberSequence As Boolean
    Dim LastValidPos As Long
    Dim CurrentChar As String
    Dim NextChar As String
    
    result = Trim(FullCorrespondent)
    
    Debug.Print "=== ИЗВЛЕЧЕНИЕ КОРРЕСПОНДЕНТА v1.1.9 ==="
    Debug.Print "Исходный текст: '" & result & "'"
    
    ' ПРАВИЛО 1: Если есть скобка - берем текст до первой скобки
    BracketPos = InStr(result, "(")
    If BracketPos > 0 Then
        result = Trim(Left(result, BracketPos - 1))
        Debug.Print "? Найдена скобка в позиции " & BracketPos
        Debug.Print "? Результат до скобки: '" & result & "'"
        ' Дальше обрабатываем как обычно для извлечения номера
    End If
    
    ' ПРАВИЛО 2: Берем до окончания номера (цифры + дефисы + буквы)
    Debug.Print "Ищем окончание номера с дефисами и буквами..."
    
    InNumberSequence = False
    LastValidPos = 0
    
    ' Проходим по всем символам
    For i = 1 To Len(result)
        CurrentChar = Mid(result, i, 1)
        NextChar = ""
        If i < Len(result) Then NextChar = Mid(result, i + 1, 1)
        
        If CurrentChar >= "0" And CurrentChar <= "9" Then
            ' Цифра - начинаем или продолжаем номер
            InNumberSequence = True
            LastValidPos = i
            Debug.Print "  Найдена цифра '" & CurrentChar & "' в позиции " & i
            
        ElseIf CurrentChar = "-" And InNumberSequence Then
            ' Дефис в номере - проверяем что после него цифра или буква
            If NextChar >= "0" And NextChar <= "9" Then
                ' После дефиса цифра - продолжаем номер
                LastValidPos = i
                Debug.Print "  Найден дефис перед цифрой в позиции " & i
            ElseIf (NextChar >= "А" And NextChar <= "я") Or (NextChar >= "A" And NextChar <= "Z") Or (NextChar >= "a" And NextChar <= "z") Then
                ' После дефиса буква - включаем её в номер
                LastValidPos = i + 1  ' +1 чтобы включить букву
                Debug.Print "  Найден дефис перед буквой '" & NextChar & "' в позиции " & i
            Else
                ' После дефиса что-то другое - заканчиваем номер
                Exit For
            End If
            
ElseIf InNumberSequence And ((CurrentChar >= "А" And CurrentChar <= "я") Or (CurrentChar >= "A" And CurrentChar <= "Z") Or (CurrentChar >= "a" And CurrentChar <= "z")) Then
    ' Буква в номере (например Т или ТА в 47084-ТА)
    Dim PrevChar As String
    Dim AfterDash As Boolean
    AfterDash = False
    
    ' Проверяем, идет ли эта буква после дефиса (напрямую или через другие буквы)
    Dim j As Long
    For j = i - 1 To 1 Step -1
        Dim CheckChar As String
        CheckChar = Mid(result, j, 1)
        If CheckChar = "-" Then
            AfterDash = True
            Exit For
        ElseIf Not ((CheckChar >= "А" And CheckChar <= "я") Or (CheckChar >= "A" And CheckChar <= "Z") Or (CheckChar >= "a" And CheckChar <= "z")) Then
            ' Встретили не-букву и не-дефис - значит буквы не после дефиса
            Exit For
        End If
    Next j
    
    If AfterDash Then
        ' Буква после дефиса (напрямую или в цепочке букв) - включаем её
        LastValidPos = i
        Debug.Print "  Найдена буква '" & CurrentChar & "' в составе номера после дефиса в позиции " & i
        ' Продолжаем, если следующий символ тоже буква
        ' Останавливаемся только на пробеле или другом разделителе
        If NextChar <> "" And NextChar <> " " And Not ((NextChar >= "А" And NextChar <= "я") Or (NextChar >= "A" And NextChar <= "z")) Then
            Exit For
        End If
    Else
        ' Буква не связана с дефисом - заканчиваем номер
        Exit For
    End If

            
        Else
            ' Символ НЕ является частью номера
            If InNumberSequence Then
                ' Мы вышли из номера
                Debug.Print "  Окончание номера в позиции " & LastValidPos
                Exit For
            End If
        End If
    Next i
    
    ' Если найден номер, обрезаем до его окончания
    If LastValidPos > 0 Then
        result = Trim(Left(result, LastValidPos))
        Debug.Print "? Результат до окончания номера: '" & result & "'"
    Else
        Debug.Print "? Номер не найден, оставляем текст как есть: '" & result & "'"
    End If
    
    ExtractCorrespondentName = result
    Debug.Print "=== ИТОГОВЫЙ РЕЗУЛЬТАТ: '" & result & "' ==="
End Function


Private Sub SortDoverennostiInGroup(ByRef Group As CorrespondentGroup)
    Dim i As Long, j As Long
    Dim TempDover As DoverennostInfo
    
    For i = 0 To Group.Count - 2
        For j = i + 1 To Group.Count - 1
            If Val(Group.Doverennosti(i).DoverennostNumber) > Val(Group.Doverennosti(j).DoverennostNumber) Then
                TempDover = Group.Doverennosti(i)
                Group.Doverennosti(i) = Group.Doverennosti(j)
                Group.Doverennosti(j) = TempDover
            End If
        Next j
    Next i
End Sub

Private Function BuildFullAddress(PostalCode As String, Region As String, District As String, City As String, Street As String) As String
    Dim AddressParts As String
    
    ' ИСПРАВЛЕННЫЙ ПОРЯДОК адреса v1.1.5:
    ' 1. Улица, дом, квартира
    ' 2. Населенный пункт
    ' 3. Район
    ' 4. Область/край/республика
    ' 5. Почтовый индекс
    
    If Len(Trim(Street)) > 0 Then
        AddressParts = Trim(Street)
    End If
    
    If Len(Trim(City)) > 0 Then
        If Len(AddressParts) > 0 Then AddressParts = AddressParts & ", "
        AddressParts = AddressParts & Trim(City)
    End If
    
    If Len(Trim(District)) > 0 And Trim(District) <> "" Then
        If Len(AddressParts) > 0 Then AddressParts = AddressParts & ", "
        AddressParts = AddressParts & Trim(District)
    End If
    
    If Len(Trim(Region)) > 0 Then
        If Len(AddressParts) > 0 Then AddressParts = AddressParts & ", "
        AddressParts = AddressParts & Trim(Region)
    End If
    
    If Len(Trim(PostalCode)) > 0 Then
        If Len(AddressParts) > 0 Then AddressParts = AddressParts & ", "
        AddressParts = AddressParts & Trim(PostalCode)
    End If
    
    BuildFullAddress = AddressParts
    Debug.Print "Сформированный адрес (v1.1.5): " & AddressParts
End Function


Private Function CleanFileName(FileName As String) As String
    Dim result As String
    
    result = FileName
    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    
    CleanFileName = result
End Function

Private Sub CreateFIOGroups(Doverennosti() As DoverennostInfo, Count As Long, ByRef FIOGroups() As FIOGroup, ByRef FIOGroupCount As Long)
    Dim i As Long, j As Long
    Dim FIOName As String
    Dim GroupFound As Boolean
    
    FIOGroupCount = 0
    ReDim FIOGroups(Count)
    
    For i = 0 To Count - 1
        FIOName = Doverennosti(i).FIO
        GroupFound = False
        
        For j = 0 To FIOGroupCount - 1
            If UCase(FIOGroups(j).FIOName) = UCase(FIOName) Then
                GroupFound = True
                FIOGroups(j).Count = FIOGroups(j).Count + 1
                ReDim Preserve FIOGroups(j).Doverennosti(FIOGroups(j).Count - 1)
                FIOGroups(j).Doverennosti(FIOGroups(j).Count - 1) = Doverennosti(i)
                Exit For
            End If
        Next j
        
        If Not GroupFound Then
            With FIOGroups(FIOGroupCount)
                .FIOName = FIOName
                .Count = 1
                ReDim .Doverennosti(0)
                .Doverennosti(0) = Doverennosti(i)
            End With
            FIOGroupCount = FIOGroupCount + 1
        End If
    Next i
    
    ReDim Preserve FIOGroups(FIOGroupCount - 1)
End Sub

Private Sub CreateMissingAddressesReport(NotFoundList As String, SaveFolder As String)
    Dim ReportPath As String
    Dim FileNum As Integer


    ReportPath = SaveFolder & "\Отчет_НеНайденныеАдреса_" & Replace(CurrentDate, ".", "") & ".txt"
    
    FileNum = FreeFile
    Open ReportPath For Output As FileNum
    
    Print #FileNum, "ОТЧЕТ О НЕ НАЙДЕННЫХ АДРЕСАХ КОРРЕСПОНДЕНТОВ"
    Print #FileNum, "Дата создания: " & Format(Now, "dd.mm.yyyy HH:mm:ss")
    Print #FileNum, "Автор: Кержаев Евгений, ФКУ ""95 ФЭС"" МО РФ"
    Print #FileNum, ""
    Print #FileNum, "Следующие корреспонденты не найдены в файле адресов:"
    Print #FileNum, "======================================================"
    Print #FileNum, ""
    Print #FileNum, NotFoundList
    Print #FileNum, ""
    Print #FileNum, "======================================================"
    Print #FileNum, "Необходимо добавить адреса этих корреспондентов в файл адресов"
    Print #FileNum, "для возможности формирования писем."
    
    Close FileNum
End Sub

Private Sub ShowGenerationResults(GroupCount As Long, NotFoundList As String, SaveFolder As String)
    Dim Message As String
    Dim FoundCount As Long
    Dim NotFoundCount As Long
    
    If Len(NotFoundList) > 0 Then
        NotFoundCount = UBound(Split(NotFoundList, vbCrLf)) + 1
    Else
        NotFoundCount = 0
    End If
    
    FoundCount = GroupCount - NotFoundCount
    
    Message = "? РЕЗУЛЬТАТЫ ФОРМИРОВАНИЯ ПИСЕМ:" & vbCrLf & vbCrLf & _
              "?? Всего групп корреспондентов: " & GroupCount & vbCrLf & _
              "? Создано документов Word: " & FoundCount & vbCrLf & _
              "? Не найдено адресов: " & NotFoundCount & vbCrLf & vbCrLf & _
              "?? Файлы сохранены в: " & SaveFolder & vbCrLf & vbCrLf
    
    If NotFoundCount > 0 Then
        Message = Message & "?? Создан отчет о не найденных адресах." & vbCrLf & _
                           "Добавьте их в файл адресов и повторите процедуру." & vbCrLf & vbCrLf
    End If
    
    Message = Message & "Автор: Кержаев Евгений, ФКУ ""95 ФЭС"" МО РФ"
    
    MsgBox Message, vbInformation, "Формирование писем завершено v1.1.0"
End Sub


' =============================================
' ДИАГНОСТИКА ПРОБЛЕМ С ЗАКЛАДКАМИ v1.1.1
' =============================================
Private Sub DiagnoseBookmarkIssues(WordDoc As Object, BookmarkName As String)
    Debug.Print "=== ДИАГНОСТИКА ЗАКЛАДКИ: " & BookmarkName & " ==="
    
    On Error Resume Next
    
    ' Список всех закладок
    Debug.Print "Всего закладок в документе: " & WordDoc.Bookmarks.Count
    Dim i As Long
    For i = 1 To WordDoc.Bookmarks.Count
        Debug.Print "  Закладка " & i & ": '" & WordDoc.Bookmarks(i).Name & "'"
    Next i
    
    ' Проверка существования конкретной закладки
    If WordDoc.Bookmarks.Exists(BookmarkName) Then
        Debug.Print "? Закладка '" & BookmarkName & "' существует"
        Dim bmRange As Object
        Set bmRange = WordDoc.Bookmarks(BookmarkName).Range
        Debug.Print "  Позиция: " & bmRange.Start & "-" & bmRange.End
        Debug.Print "  Текущий текст: '" & bmRange.Text & "'"
    Else
        Debug.Print "? Закладка '" & BookmarkName & "' НЕ найдена"
        
        ' Поиск похожих закладок
        For i = 1 To WordDoc.Bookmarks.Count
            If InStr(UCase(WordDoc.Bookmarks(i).Name), UCase(BookmarkName)) > 0 Or _
               InStr(UCase(BookmarkName), UCase(WordDoc.Bookmarks(i).Name)) > 0 Then
                Debug.Print "  Похожая закладка: '" & WordDoc.Bookmarks(i).Name & "'"
            End If
        Next i
    End If
    
    On Error GoTo 0
    Debug.Print "=== КОНЕЦ ДИАГНОСТИКИ ==="
End Sub



' =============================================
' ИНИЦИАЛИЗАЦИЯ СИСТЕМЫ ИСХОДЯЩИХ ПИСЕМ v1.2.0
' =============================================
Private Function InitializeLetterNumbering() As Boolean
    Dim StartNumber As String
    Dim LetterDateInput As String
    Dim InputLetterDate As Date
    
    On Error GoTo InitError
    
    InitializeLetterNumbering = False
    
    ' Запрос начального номера у пользователя
    StartNumber = InputBox("Введите начальный номер письма:", "Номер исходящего письма", "1")
    
    If StartNumber = "" Then
        Debug.Print "Пользователь отменил ввод номера письма"
        Exit Function
    End If
    
    ' Запрос даты письма у пользователя
    LetterDateInput = InputBox("Введите дату письма (дд.мм.гггг):", "Дата исходящего письма", Format(Date, "dd.mm.yyyy"))
    
    If LetterDateInput = "" Then
        Debug.Print "Пользователь отменил ввод даты письма"
        Exit Function
    End If
    
    ' Проверка и преобразование даты
    On Error Resume Next
    InputLetterDate = CDate(LetterDateInput)
    If Err.Number <> 0 Then
        MsgBox "Неверный формат даты! Используется текущая дата.", vbExclamation, "Ошибка даты"
        InputLetterDate = Date
    End If
    On Error GoTo InitError
    
    ' Инициализация переменных
    CurrentLetterNumber = CLng(StartNumber)
    LetterPrefix = "7/"
    CurrentDate = InputLetterDate  ' Используем введенную дату
    FormattedCurrentDate = FormatLetterDate(CurrentDate)

    
    InitializeLetterNumbering = True
    Exit Function
    
InitError:
    MsgBox "Ошибка инициализации номеров писем: " & Err.description, vbCritical
    InitializeLetterNumbering = False
End Function


Private Function FormatLetterDate(ByVal InputDate As Date) As String
    Dim MonthNames As Variant
    MonthNames = Array("", "января", "февраля", "марта", "апреля", "мая", "июня", _
                       "июля", "августа", "сентября", "октября", "ноября", "декабря")
    
    FormatLetterDate = Day(InputDate) & " " & MonthNames(Month(InputDate)) & " " & Year(InputDate) & " г."
End Function


Private Function FormatShortDate(ByVal InputDate As Date) As String
    On Error Resume Next
    FormatShortDate = Format(InputDate, "dd.mm.yy")
    If Err.Number <> 0 Then
        FormatShortDate = Format(Now, "dd.mm.yy")
        Debug.Print "? Ошибка форматирования даты, использую текущую"
    End If
    Debug.Print "Форматированная дата: '" & FormatShortDate & "'"
    On Error GoTo 0
End Function





' =============================================
' ПРОВЕРКА 3-МЕСЯЧНОГО ПЕРИОДА ДЛЯ ПИСЕМ v1.2.0
' =============================================
Private Function ShouldIncludeDoverennost(WsDover As Worksheet, RowNumber As Long) As Boolean
    Dim OperationDateCol As Long
    Dim OperationNumberCol As Long
    Dim ExistingDate As Date
    Dim ExistingNumber As String
    Dim MonthsDiff As Long
    
    ShouldIncludeDoverennost = True
    
    ' Находим столбцы операций
    OperationDateCol = FindColumnByName(WsDover, "ДАТА ОПЕРАЦИИ")
    OperationNumberCol = FindColumnByName(WsDover, "НОМЕР ОПЕРАЦИИ")
    
    If OperationDateCol = 0 Or OperationNumberCol = 0 Then
        Debug.Print "? Столбцы операций не найдены для строки " & RowNumber
        Exit Function
    End If
    
    ' Проверяем существующие данные
    ExistingNumber = Trim(CStr(WsDover.Cells(RowNumber, OperationNumberCol).value))
    
    If Len(ExistingNumber) > 0 And InStr(ExistingNumber, "/") > 0 Then
        ' Есть номер письма, проверяем дату
        On Error Resume Next
        ExistingDate = CDate(WsDover.Cells(RowNumber, OperationDateCol).value)
        On Error GoTo 0
        
        If ExistingDate > DateSerial(1900, 1, 2) Then
            ' Считаем разность в месяцах
            MonthsDiff = DateDiff("m", ExistingDate, CurrentDate)


            
            If MonthsDiff < 3 Then
                Debug.Print "? Пропускаем строку " & RowNumber & " - письмо отправлено " & Format(ExistingDate, "dd.mm.yyyy") & " (" & MonthsDiff & " мес. назад)"
                ShouldIncludeDoverennost = False
            Else
                Debug.Print "? Включаем строку " & RowNumber & " - старое письмо от " & Format(ExistingDate, "dd.mm.yyyy") & " (" & MonthsDiff & " мес. назад)"
            End If
        End If
    End If
End Function

Private Function GetCurrentLetterInfo() As LetterInfo
    With GetCurrentLetterInfo
        .Number = LetterPrefix & CurrentLetterNumber
        .Date = Format(CurrentDate, "dd.mm.yyyy")
        .FormattedDate = FormattedCurrentDate
    End With
End Function


' =============================================
' ОБНОВЛЕНИЕ ФАЙЛА ДОВЕРЕННОСТЕЙ v1.2.1
' =============================================
Private Sub UpdateDoverennostiFile(WbDover As Workbook, GroupData As CorrespondentGroup, LetterInfo As LetterInfo)
    Dim WsDover As Worksheet
    Dim OperationDateCol As Long
    Dim OperationNumberCol As Long
    Dim i As Long
    
    On Error GoTo UpdateError
    
    Set WsDover = WbDover.Worksheets(1)
    
    ' Находим столбцы операций
    OperationDateCol = FindColumnByName(WsDover, "ДАТА ОПЕРАЦИИ")
    OperationNumberCol = FindColumnByName(WsDover, "НОМЕР ОПЕРАЦИИ")
    
    Debug.Print "=== ОБНОВЛЕНИЕ ФАЙЛА ДОВЕРЕННОСТЕЙ v1.2.1 ==="
    Debug.Print "Столбец 'Номер операции': " & OperationNumberCol
    Debug.Print "Столбец 'Дата операции': " & OperationDateCol
    Debug.Print "Корреспондент: " & GroupData.CorrespondentName
    Debug.Print "Номер письма: " & LetterInfo.Number
    Debug.Print "Дата письма: " & LetterInfo.Date
    
    If OperationDateCol = 0 Or OperationNumberCol = 0 Then
        Debug.Print "? КРИТИЧНО: Не удалось найти столбцы операций!"
        Debug.Print "Проверьте наличие столбцов 'Номер операции' и 'Дата операции'"
        MsgBox "ОШИБКА: Не найдены столбцы 'Номер операции' и/или 'Дата операции' в файле доверенностей!" & vbCrLf & vbCrLf & _
               "Проверьте названия столбцов.", vbCritical, "Ошибка обновления"
        Exit Sub
    End If
    
    ' Обновляем все доверенности этой группы
    For i = 0 To GroupData.Count - 1
        With GroupData.Doverennosti(i)
            Debug.Print "Обновляем строку " & .RowNumber & " для ФИО: " & .FIO
            
            ' Проверяем текущие значения в ячейках
            Dim ExistingNumber As String
            Dim ExistingDate As String
            ExistingNumber = Trim(CStr(WsDover.Cells(.RowNumber, OperationNumberCol).value))
            ExistingDate = Trim(CStr(WsDover.Cells(.RowNumber, OperationDateCol).value))
            
            Debug.Print "  Текущий номер: '" & ExistingNumber & "'"
            Debug.Print "  Текущая дата: '" & ExistingDate & "'"
            
            ' Записываем номер и дату операции
            ' Записываем номер и дату операции с правильным форматированием
            ' Устанавливаем формат ячейки как ТЕКСТ для номера операции
            WsDover.Cells(.RowNumber, OperationNumberCol).NumberFormat = "@"
            WsDover.Cells(.RowNumber, OperationNumberCol).value = LetterInfo.Number
            
            ' Для даты операции оставляем обычное значение
            WsDover.Cells(.RowNumber, OperationDateCol).value = LetterInfo.Date
            
            Debug.Print "  ? ЗАПИСАНО - Номер: " & LetterInfo.Number & " (как текст), Дата: " & LetterInfo.Date

        End With
    Next i
    
    ' Принудительно сохраняем файл
On Error Resume Next
WbDover.Save
If Err.Number = 0 Then
    Debug.Print "? Файл доверенностей сохранен успешно"
Else
    Debug.Print "? ПРЕДУПРЕЖДЕНИЕ: Не удалось сохранить файл доверенностей"
    Debug.Print "   Причина: " & Err.description
    
    ' Попробуем сохранить как новый файл
    Dim BackupPath As String
    BackupPath = Replace(WbDover.FullName, ".xlsx", "_updated_" & Format(Now, "ddmmyyyy_hhmmss") & ".xlsx")
    WbDover.SaveAs BackupPath
    
    MsgBox "ВНИМАНИЕ!" & vbCrLf & vbCrLf & _
           "Исходный файл заблокирован для записи." & vbCrLf & _
           "Данные сохранены в новый файл:" & vbCrLf & vbCrLf & _
           BackupPath, vbExclamation, "Файл сохранен как копия"
    
    Debug.Print "? Данные сохранены в резервный файл: " & BackupPath
End If
On Error GoTo 0

    
    Exit Sub
    
UpdateError:
    Debug.Print "? ОШИБКА обновления файла доверенностей: " & Err.description
    Debug.Print "   Номер ошибки: " & Err.Number
    Debug.Print "   Источник: " & Err.Source
    
    MsgBox "КРИТИЧЕСКАЯ ОШИБКА обновления файла доверенностей!" & vbCrLf & vbCrLf & _
           "Корреспондент: " & GroupData.CorrespondentName & vbCrLf & _
           "Ошибка: " & Err.description & vbCrLf & vbCrLf & _
           "Номер письма НЕ СОХРАНЕН в файле!", vbCritical, "Критическая ошибка"
End Sub


