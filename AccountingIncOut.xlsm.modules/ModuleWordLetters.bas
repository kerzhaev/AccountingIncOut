Attribute VB_Name = "ModuleWordLetters"
'==============================================
' WORD LETTERS GENERATION MODULE
' Version: 1.2.3
' Date: 07.09.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================
Option Explicit

' =============================================
' DATA STRUCTURES
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
    Number As String          ' Letter number (7/25)
    Date As String            ' Letter date (25.08.2025)
    FormattedDate As String   ' Letter date (25 August 2025)
End Type

' =============================================
' COLUMN CONSTANTS
' =============================================
Private Const COL_DATE As String = "DATE"
Private Const COL_NUMBER As String = "NUMBER"
Private Const COL_FIO As String = "ACCOUNTABLE PERSON"
Private Const COL_ORGANIZATION As String = "ORGANIZATION"
Private Const COL_CORRESPONDENT As String = "SUPPLIER"
Private Const COL_COMMENT As String = "COMMENT"
Private Const COL_SEARCH_RESULT As String = "THREE-PASS SEARCH"

' =============================================
' OUTGOING LETTERS TRACKING VARIABLES
' =============================================
Private CurrentLetterNumber As Long
Private LetterPrefix As String
Private CurrentDate As Date
Private FormattedCurrentDate As String

' =============================================
' MAIN PROCEDURE
' =============================================
Public Sub GenerateWordLetters()
    Dim FileDovernnosti As String
    Dim FileAddresses As String
    Dim SaveFolder As String
    Dim WbDover As Workbook
    Dim WbAddr As Workbook
    
    On Error GoTo GenerateError
    
    Debug.Print "=== START WORD LETTERS GENERATION ==="
    
    ' Select Doverennosti file
    FileDovernnosti = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        "Select file with processed Doverennosti")
    If FileDovernnosti = "False" Then Exit Sub
    
    ' Select Addresses file
    FileAddresses = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,Excel Macro Files (*.xlsm),*.xlsm,CSV Files (*.csv),*.csv", , _
        "Select file with correspondent addresses")
    If FileAddresses = "False" Then Exit Sub
    
    ' Select save folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder to save letters"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SaveFolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Check template
    Dim TemplatePath As String
    TemplatePath = ThisWorkbook.Path & "\LetterTemplate.docx"
    If dir(TemplatePath) = "" Then
        MsgBox "ERROR: Template not found: " & TemplatePath & vbCrLf & vbCrLf & _
               "Place the template file in the macro folder.", vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    ' Open files
    Set WbDover = Workbooks.Open(FileDovernnosti, ReadOnly:=False)
    Debug.Print "  [OK] Doverennosti file opened for writing"

    Set WbAddr = Workbooks.Open(FileAddresses, ReadOnly:=True)
    
    ' Validate file structure
    If Not ValidateFileStructure(WbDover, WbAddr) Then
        WbAddr.Close False
        WbDover.Close False
        Exit Sub
    End If
    
    ' Initialize outgoing letters system
    If Not InitializeLetterNumbering() Then
        WbAddr.Close False
        WbDover.Close False
        Exit Sub
    End If

    ' Main processing
    Call ProcessWordLetterGeneration(WbDover, WbAddr, SaveFolder)
    
    ' Close files
    WbAddr.Close False
    WbDover.Close False
    
    Debug.Print "=== WORD LETTERS GENERATION COMPLETED ==="
    
    Exit Sub
    
GenerateError:
    On Error Resume Next
    If Not WbAddr Is Nothing Then WbAddr.Close False
    If Not WbDover Is Nothing Then WbDover.Close False
    
    Debug.Print "Letter generation error: " & Err.description
    MsgBox "Error: " & Err.description, vbCritical
End Sub

' =============================================
' FILE STRUCTURE VALIDATION
' =============================================
Private Function ValidateFileStructure(WbDover As Workbook, WbAddr As Workbook) As Boolean
    Dim WsDover As Worksheet
    Dim WsAddr As Worksheet
    Dim MissingColumns As String
    Dim i As Long
    
    ValidateFileStructure = False
    Set WsDover = WbDover.Worksheets(1)
    
    ' Find "Addresses" sheet
    Dim AddrSheetFound As Boolean
    AddrSheetFound = False
    For i = 1 To WbAddr.Worksheets.Count
        If UCase(WbAddr.Worksheets(i).Name) = "ADDRESSES" Then
            Set WsAddr = WbAddr.Worksheets(i)
            AddrSheetFound = True
            Exit For
        End If
    Next i
    
    If Not AddrSheetFound Then
        MsgBox "ERROR: 'Addresses' sheet not found in the addresses file!" & vbCrLf & vbCrLf & _
               "Ensure the sheet is named exactly 'Addresses'", vbCritical, "File Structure Invalid"
        Exit Function
    End If
    
    ' Check required columns in Doverennosti file
    MissingColumns = ""
    If FindColumnByName(WsDover, COL_DATE) = 0 Then MissingColumns = MissingColumns & "- " & COL_DATE & vbCrLf
    If FindColumnByName(WsDover, COL_NUMBER) = 0 Then MissingColumns = MissingColumns & "- " & COL_NUMBER & vbCrLf
    If FindColumnByName(WsDover, COL_FIO) = 0 Then MissingColumns = MissingColumns & "- " & COL_FIO & vbCrLf
    If FindColumnByName(WsDover, COL_ORGANIZATION) = 0 Then MissingColumns = MissingColumns & "- " & COL_ORGANIZATION & vbCrLf
    If FindColumnByName(WsDover, COL_CORRESPONDENT) = 0 Then MissingColumns = MissingColumns & "- " & COL_CORRESPONDENT & vbCrLf
    If FindColumnByName(WsDover, COL_SEARCH_RESULT) = 0 Then MissingColumns = MissingColumns & "- " & COL_SEARCH_RESULT & vbCrLf
    
    If Len(MissingColumns) > 0 Then
        MsgBox "ERROR: Missing required columns in Doverennosti file:" & vbCrLf & vbCrLf & _
               MissingColumns & vbCrLf & _
               "Processing stopped. Check file structure.", vbCritical, "Missing Columns"
        Exit Function
    End If
    
    ' Check columns in Addresses file
    If WsAddr.Cells(1, 1).value = "" Or WsAddr.Cells(1, 2).value = "" Then
        MsgBox "ERROR: Invalid addresses file structure!" & vbCrLf & vbCrLf & _
               "Expected columns:" & vbCrLf & _
               "1. Recipient Name" & vbCrLf & _
               "2. Street, building, apartment" & vbCrLf & _
               "3. City/Settlement" & vbCrLf & _
               "4. District" & vbCrLf & _
               "5. Region/State" & vbCrLf & _
               "6. Postal Code", vbCritical, "File Structure Invalid"
        Exit Function
    End If
    
    ValidateFileStructure = True
    Debug.Print "  [OK] File structure validated successfully"
End Function

' =============================================
' MAIN PROCESSING LOGIC
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
    Set WsAddr = WbAddr.Worksheets("Addresses")
    
    LastRowDover = WsDover.Cells(WsDover.Rows.Count, 1).End(xlUp).Row
    LastRowAddr = WsAddr.Cells(WsAddr.Rows.Count, 1).End(xlUp).Row
    
    ' Load data
    DoverData = WsDover.Range("A2:Z" & LastRowDover).Value2
    AddrData = WsAddr.Range("A2:F" & LastRowAddr).Value2
    
    Debug.Print "Loaded Doverennosti: " & (LastRowDover - 1)
    Debug.Print "Loaded Addresses: " & (LastRowAddr - 1)
    
    ' Filter unmatched Doverennosti
    Application.StatusBar = "Filtering unmatched records..."
    Call FilterUnmatchedDoverennosti(WsDover, DoverData, UnmatchedDoverennosti, UnmatchedCount)
    
    Debug.Print "Found unmatched Doverennosti: " & UnmatchedCount
    
    If UnmatchedCount = 0 Then
        MsgBox "No unmatched records found!", vbInformation
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' Group by correspondents
    Application.StatusBar = "Grouping by correspondents..."
    Call GroupByCorrespondents(UnmatchedDoverennosti, UnmatchedCount, CorrespondentGroups, GroupCount)
    
    Debug.Print "Correspondent groups created: " & GroupCount
    
    ' Find addresses
    Application.StatusBar = "Searching for correspondent addresses..."
    Call FindCorrespondentAddresses(CorrespondentGroups, GroupCount, AddrData, NotFoundCorrespondents)
    
    ' Create Word documents
    Application.StatusBar = "Creating Word documents..."
    Call CreateWordDocuments(CorrespondentGroups, GroupCount, SaveFolder, WbDover)
    
    ' Create missing addresses report
    If Len(NotFoundCorrespondents) > 0 Then
        Call CreateMissingAddressesReport(NotFoundCorrespondents, SaveFolder)
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Show results
    Call ShowGenerationResults(GroupCount, NotFoundCorrespondents, SaveFolder)
End Sub

' =============================================
' FILTERING
' =============================================
Private Sub FilterUnmatchedDoverennosti(WsDover As Worksheet, DoverData As Variant, ByRef UnmatchedList() As DoverennostInfo, ByRef Count As Long)
    Dim i As Long
    Dim SearchCol As Long
    Dim DateCol As Long, NumberCol As Long, FIOCol As Long
    Dim OrganizationCol As Long, CorrespondentCol As Long, CommentCol As Long
    Dim SearchValue As String
    
    Count = 0
    ReDim UnmatchedList(UBound(DoverData, 1))
    
    ' Define columns
    SearchCol = FindColumnByName(WsDover, COL_SEARCH_RESULT)
    DateCol = FindColumnByName(WsDover, COL_DATE)
    NumberCol = FindColumnByName(WsDover, COL_NUMBER)
    FIOCol = FindColumnByName(WsDover, COL_FIO)
    OrganizationCol = FindColumnByName(WsDover, COL_ORGANIZATION)
    CorrespondentCol = FindColumnByName(WsDover, COL_CORRESPONDENT)
    CommentCol = FindColumnByName(WsDover, COL_COMMENT)
    
    Debug.Print "=== START FILTERING ==="
    Debug.Print "Search column: " & SearchCol
    
    ' Filter records
    For i = 1 To UBound(DoverData, 1)
        If i Mod 50 = 0 Then
            Application.StatusBar = "Filtering row " & i & " of " & UBound(DoverData, 1)
        End If
        
        SearchValue = ""
        If SearchCol > 0 And SearchCol <= UBound(DoverData, 2) Then
            SearchValue = Trim(UCase(CStr(DoverData(i, SearchCol))))
        End If
        
        If SearchValue = "NOT FOUND" Then
            ' Check empty required fields
            If FIOCol > 0 And OrganizationCol > 0 And NumberCol > 0 And CorrespondentCol > 0 Then
                If Trim(CStr(DoverData(i, FIOCol))) <> "" And _
                   Trim(CStr(DoverData(i, OrganizationCol))) <> "" And _
                   Trim(CStr(DoverData(i, NumberCol))) <> "" And _
                   Trim(CStr(DoverData(i, CorrespondentCol))) <> "" Then
                    
                    ' Check 3-month period
                    If ShouldIncludeDoverennost(WsDover, i + 1) Then
                        With UnmatchedList(Count)
                            .RowNumber = i + 1
                            .FIO = Trim(CStr(DoverData(i, FIOCol)))
                            .VoyskovayaChast = Trim(CStr(DoverData(i, OrganizationCol)))
                            .DoverennostNumber = Trim(CStr(DoverData(i, NumberCol)))
                            .Correspondent = Trim(CStr(DoverData(i, CorrespondentCol)))
                            
                            If CommentCol > 0 And CommentCol <= UBound(DoverData, 2) Then
                                .Comment = Trim(CStr(DoverData(i, CommentCol)))
                            Else
                                .Comment = ""
                            End If
                            
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
                        Debug.Print "  [OK] Added: " & UnmatchedList(Count - 1).FIO & " (" & UnmatchedList(Count - 1).DoverennostNumber & ")"
                    Else
                        Debug.Print "  [Skip] Recent letter sent: " & Trim(CStr(DoverData(i, FIOCol)))
                    End If
                End If
            End If
        End If
    Next i
    
    If Count > 0 Then
        ReDim Preserve UnmatchedList(Count - 1)
    End If
    Debug.Print "=== FILTER RESULT: " & Count & " records ==="
End Sub

' =============================================
' DOCUMENT CREATION
' =============================================
Private Sub CreateWordDocuments(Groups() As CorrespondentGroup, GroupCount As Long, SaveFolder As String, WbDover As Workbook)
    Dim i As Long
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim TemplatePath As String
    Dim SavePath As String
    Dim CurrentDateStr As String
    Dim ShortDate As String
    Dim LetterInfo As LetterInfo
    Dim FileNameNumber As String
    
    CurrentDateStr = Format(Now, "dd.mm.yyyy")
    TemplatePath = ThisWorkbook.Path & "\LetterTemplate.docx"
    
    Debug.Print "=== CREATING WORD DOCUMENTS ==="
    
    On Error GoTo TemplateError
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = False
    
    Set WordDoc = WordApp.Documents.Open(TemplatePath, ReadOnly:=True)
    WordDoc.Close False
    Set WordDoc = Nothing
    
    Debug.Print "  [OK] Template successfully tested"
    
    For i = 0 To GroupCount - 1
        If Groups(i).AddressFound Then
            Application.StatusBar = "Creating document " & (i + 1) & " of " & GroupCount & ": " & Groups(i).CorrespondentName
            
            On Error GoTo DocumentError
            
            Set WordDoc = WordApp.Documents.Open(TemplatePath)
            
            Call FillWordDocumentFixed(WordDoc, Groups(i))
            
            LetterInfo = GetCurrentLetterInfo()
            FileNameNumber = Replace(LetterInfo.Number, "/", "_")
            ShortDate = FormatShortDate(CurrentDate)
            
            SavePath = SaveFolder & "\" & CleanFileName(Groups(i).CorrespondentName) & "_" & FileNameNumber & " from " & ShortDate & ".docx"
            
            WordDoc.SaveAs2 SavePath
            WordDoc.Close
            
            Debug.Print "  [OK] Document created: " & SavePath
            
            Call UpdateDoverennostiFile(WbDover, Groups(i), LetterInfo)
            
            CurrentLetterNumber = CurrentLetterNumber + 1
            Debug.Print "  [Info] Next letter number: " & LetterPrefix & CurrentLetterNumber
        Else
            Debug.Print "  [Skip] No address found for: " & Groups(i).CorrespondentName
        End If
    Next i
    
    WordApp.Quit
    Set WordApp = Nothing
    
    Debug.Print "  [OK] Document creation finished"
    Exit Sub
    
TemplateError:
    On Error Resume Next
    If Not WordApp Is Nothing Then WordApp.Quit
    MsgBox "ERROR: Template corrupted or unavailable!" & vbCrLf & vbCrLf & _
           "Check file: " & TemplatePath & vbCrLf & vbCrLf & _
           "Error: " & Err.description, vbCritical, "Template Issue"
    Exit Sub
    
DocumentError:
    On Error Resume Next
    If Not WordDoc Is Nothing Then WordDoc.Close False
    Debug.Print "  [X] Document creation error for: " & Groups(i).CorrespondentName & " - " & Err.description
    Resume Next
End Sub

' =============================================
' FILL WORD DOCUMENT
' =============================================
Private Sub FillWordDocumentFixed(WordDoc As Object, GroupData As CorrespondentGroup)
    Dim WordTable As Object
    Dim i As Long, j As Long
    Dim RowIndex As Long
    
    Debug.Print "=== FILLING DOCUMENT ==="
    Debug.Print "Correspondent: " & GroupData.CorrespondentName
    
    On Error GoTo FillError
    
    ' STEP 1: REPLACE BOOKMARKS
    With GroupData.Address
        Call ReplaceWordBookmarkCorrect(WordDoc, "RecipientName", .CorrespondentName)
        Call ReplaceWordBookmarkCorrect(WordDoc, "RecipientAddress", .FullAddress)
        Call ReplaceWordBookmarkCorrect(WordDoc, "CORRESPONDENT", LCase(.CorrespondentName))
        
        Dim CurrentLetterInfo As LetterInfo
        CurrentLetterInfo = GetCurrentLetterInfo()
        Call ReplaceWordBookmarkCorrect(WordDoc, "OutgoingNumber", CurrentLetterInfo.Number)
        Call ReplaceWordBookmarkCorrect(WordDoc, "OutgoingDate", CurrentLetterInfo.FormattedDate)
        
        Debug.Print "  [OK] Placeholders replaced"
    End With
    
    ' STEP 2: CREATE FIO GROUPS
    Dim FIOGroups() As FIOGroup
    Dim FIOGroupCount As Long
    Call CreateFIOGroups(GroupData.Doverennosti, GroupData.Count, FIOGroups, FIOGroupCount)
    
    ' STEP 3: COUNT ROWS
    Dim TotalDataRows As Long
    For i = 0 To FIOGroupCount - 1
        TotalDataRows = TotalDataRows + FIOGroups(i).Count
    Next i
    
    ' STEP 4: FIND INSERTION POINT
    Dim InsertRange As Object
    Set InsertRange = FindInsertionPoint(WordDoc)
    
    If InsertRange Is Nothing Then
        Debug.Print "  [Warn] Insertion point not found, adding to the end"
        Set InsertRange = WordDoc.Range
        InsertRange.Collapse 0 ' wdCollapseEnd
        InsertRange.InsertParagraphBefore
        InsertRange.InsertParagraphBefore
        InsertRange.Collapse 0 ' wdCollapseEnd
    End If
    
    ' STEP 5: CREATE TABLE
    Debug.Print "Creating table " & (TotalDataRows + 1) & "x4"
    Set WordTable = WordDoc.Tables.Add(InsertRange, TotalDataRows + 1, 4)
    
    ' STEP 6: FORMAT TABLE
    With WordTable
        .Borders.Enable = True
        .Range.Font.Size = 12
        .Range.Font.Name = "Times New Roman"
        .Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
        
        Dim CellRow As Long, CellCol As Long
        For CellRow = 1 To .Rows.Count
            For CellCol = 1 To .Columns.Count
                .cell(CellRow, CellCol).VerticalAlignment = 1 ' wdCellAlignVerticalCenter
            Next CellCol
        Next CellRow
        
        .cell(1, 1).Range.Text = "Full Name"
        .cell(1, 2).Range.Text = "Recipient Military Unit"
        .cell(1, 3).Range.Text = "Power of Attorney No. & Date"
        .cell(1, 4).Range.Text = "Basis"
        
        .Rows(1).Range.Font.Bold = True
        .Rows(1).Shading.BackgroundPatternColor = RGB(230, 230, 230)
        
        .Columns(1).Width = WordDoc.Application.CentimetersToPoints(4)
        .Columns(2).Width = WordDoc.Application.CentimetersToPoints(5)
        .Columns(3).Width = WordDoc.Application.CentimetersToPoints(4)
        .Columns(4).Width = WordDoc.Application.CentimetersToPoints(3)
    End With
    
    ' STEP 7: FILL DATA WITH MERGING
    RowIndex = 2
    For i = 0 To FIOGroupCount - 1
        Dim StartRow As Long
        StartRow = RowIndex
        
        For j = 0 To FIOGroups(i).Count - 1
            With FIOGroups(i).Doverennosti(j)
                If j = 0 Then
                    WordTable.cell(RowIndex, 1).Range.Text = .FIO
                Else
                    WordTable.cell(RowIndex, 1).Range.Text = ""
                End If
                
                WordTable.cell(RowIndex, 2).Range.Text = .VoyskovayaChast
                
                Dim DateText As String
                If .DoverennostDate > DateSerial(1900, 1, 2) Then
                    DateText = .DoverennostNumber & " from " & Format(.DoverennostDate, "dd.mm.yyyy")
                Else
                    DateText = .DoverennostNumber
                End If
                
                WordTable.cell(RowIndex, 3).Range.Text = DateText
                WordTable.cell(RowIndex, 4).Range.Text = IIf(Len(.Comment) > 0, .Comment, "Receiving material values")
            End With
            RowIndex = RowIndex + 1
        Next j
        
        If FIOGroups(i).Count > 1 Then
            Dim EndRow As Long
            EndRow = RowIndex - 1
            
            On Error Resume Next
            WordTable.cell(StartRow, 1).Merge WordTable.cell(EndRow, 1)
            WordTable.cell(StartRow, 1).VerticalAlignment = 1
            On Error GoTo 0
        End If
    Next i
    
    Debug.Print "  [OK] Document filled successfully"
    Exit Sub
    
FillError:
    Debug.Print "  [X] Fill error: " & Err.description
    MsgBox "Document filling error: " & Err.description, vbCritical
End Sub

' =============================================
' BOOKMARK REPLACEMENT
' =============================================
Private Sub ReplaceWordBookmarkCorrect(WordDoc As Object, BookmarkName As String, NewText As String)
    Dim Success As Boolean
    Success = False
    
    On Error Resume Next
    If WordDoc.Bookmarks.Exists(BookmarkName) Then
        WordDoc.Bookmarks(BookmarkName).Select
        WordDoc.Application.Selection.TypeText NewText
        Success = True
    End If
    On Error GoTo 0
    
    If Not Success Then
        On Error Resume Next
        If WordDoc.Bookmarks.Exists(BookmarkName) Then
            Dim BookmarkRange As Object
            Set BookmarkRange = WordDoc.Bookmarks(BookmarkName).Range
            BookmarkRange.Text = NewText
            WordDoc.Bookmarks.Add BookmarkName, BookmarkRange
            Success = True
        End If
        On Error GoTo 0
    End If
    
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
            End If
        End With
        On Error GoTo 0
    End If
    
    If Not Success Then
        Call DiagnoseBookmarkIssues(WordDoc, BookmarkName)
    End If
End Sub

' =============================================
' FIND TABLE INSERTION POINT
' =============================================
Private Function FindInsertionPoint(WordDoc As Object) As Object
    Dim SearchRange As Object
    Set SearchRange = WordDoc.Content
    
    With SearchRange.Find
        .ClearFormatting
        .Text = "material values:"
        .Forward = True
        .Wrap = 0 ' wdFindStop
        
        If .Execute Then
            SearchRange.Collapse 0 ' wdCollapseEnd
            SearchRange.InsertAfter vbCrLf & vbCrLf
            SearchRange.Collapse 0 ' wdCollapseEnd
            Set FindInsertionPoint = SearchRange
            Exit Function
        End If
    End With
    
    Set SearchRange = WordDoc.Range
    SearchRange.Collapse 0 ' wdCollapseEnd
    SearchRange.InsertBefore vbCrLf & vbCrLf
    Set FindInsertionPoint = SearchRange
End Function

' =============================================
' UTILITY FUNCTIONS
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
    
    For i = 1 To 30
        HeaderValue = Trim(UCase(CStr(Ws.Cells(1, i).value)))
        If InStr(HeaderValue, UCase(ColumnName)) > 0 Then
            FindColumnByName = i
            Exit Function
        End If
    Next i
    
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
    
    BracketPos = InStr(result, "(")
    If BracketPos > 0 Then
        result = Trim(Left(result, BracketPos - 1))
    End If
    
    InNumberSequence = False
    LastValidPos = 0
    
    For i = 1 To Len(result)
        CurrentChar = Mid(result, i, 1)
        NextChar = ""
        If i < Len(result) Then NextChar = Mid(result, i + 1, 1)
        
        If CurrentChar >= "0" And CurrentChar <= "9" Then
            InNumberSequence = True
            LastValidPos = i
        ElseIf CurrentChar = "-" And InNumberSequence Then
            If NextChar >= "0" And NextChar <= "9" Then
                LastValidPos = i
            ElseIf (NextChar >= "A" And NextChar <= "Z") Or (NextChar >= "a" And NextChar <= "z") Then
                LastValidPos = i + 1
            Else
                Exit For
            End If
        ElseIf InNumberSequence And ((CurrentChar >= "A" And CurrentChar <= "Z") Or (CurrentChar >= "a" And CurrentChar <= "z")) Then
            Dim j As Long
            Dim AfterDash As Boolean
            AfterDash = False
            
            For j = i - 1 To 1 Step -1
                Dim CheckChar As String
                CheckChar = Mid(result, j, 1)
                If CheckChar = "-" Then
                    AfterDash = True
                    Exit For
                ElseIf Not ((CheckChar >= "A" And CheckChar <= "Z") Or (CheckChar >= "a" And CheckChar <= "z")) Then
                    Exit For
                End If
            Next j
            
            If AfterDash Then
                LastValidPos = i
                If NextChar <> "" And NextChar <> " " And Not ((NextChar >= "A" And NextChar <= "Z") Or (NextChar >= "a" And NextChar <= "z")) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Else
            If InNumberSequence Then
                Exit For
            End If
        End If
    Next i
    
    If LastValidPos > 0 Then
        result = Trim(Left(result, LastValidPos))
    End If
    
    ExtractCorrespondentName = result
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

    ReportPath = SaveFolder & "\Report_MissingAddresses_" & Replace(CurrentDate, ".", "") & ".txt"
    
    FileNum = FreeFile
    Open ReportPath For Output As FileNum
    
    Print #FileNum, "MISSING CORRESPONDENT ADDRESSES REPORT"
    Print #FileNum, "Creation Date: " & Format(Now, "dd.mm.yyyy HH:mm:ss")
    Print #FileNum, "Author: Evgeniy Kerzhaev, FKU ""95 FES"" MO RF"
    Print #FileNum, ""
    Print #FileNum, "The following correspondents were not found in the addresses file:"
    Print #FileNum, "======================================================"
    Print #FileNum, ""
    Print #FileNum, NotFoundList
    Print #FileNum, ""
    Print #FileNum, "======================================================"
    Print #FileNum, "Please add these addresses to the address file"
    Print #FileNum, "to enable letter generation."
    
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
    
    Message = "=== LETTER GENERATION RESULTS ===" & vbCrLf & vbCrLf & _
              "Total correspondent groups: " & GroupCount & vbCrLf & _
              "Word documents created: " & FoundCount & vbCrLf & _
              "Addresses not found: " & NotFoundCount & vbCrLf & vbCrLf & _
              "Files saved in: " & SaveFolder & vbCrLf & vbCrLf
    
    If NotFoundCount > 0 Then
        Message = Message & "A missing addresses report has been created." & vbCrLf & _
                            "Add them to the address file and repeat the procedure." & vbCrLf & vbCrLf
    End If
    
    MsgBox Message, vbInformation, "Generation Completed"
End Sub

Private Sub DiagnoseBookmarkIssues(WordDoc As Object, BookmarkName As String)
    Debug.Print "  [Warn] Bookmark not found: " & BookmarkName
End Sub

Private Function InitializeLetterNumbering() As Boolean
    Dim StartNumber As String
    Dim LetterDateInput As String
    Dim InputLetterDate As Date
    
    On Error GoTo InitError
    
    InitializeLetterNumbering = False
    
    StartNumber = InputBox("Enter starting letter number:", "Outgoing Letter Number", "1")
    If StartNumber = "" Then Exit Function
    
    LetterDateInput = InputBox("Enter letter date (dd.mm.yyyy):", "Outgoing Letter Date", Format(Date, "dd.mm.yyyy"))
    If LetterDateInput = "" Then Exit Function
    
    On Error Resume Next
    InputLetterDate = CDate(LetterDateInput)
    If Err.Number <> 0 Then
        MsgBox "Invalid date format! Using current date.", vbExclamation, "Date Error"
        InputLetterDate = Date
    End If
    On Error GoTo InitError
    
    CurrentLetterNumber = CLng(StartNumber)
    LetterPrefix = "7/"
    CurrentDate = InputLetterDate
    FormattedCurrentDate = FormatLetterDate(CurrentDate)
    
    InitializeLetterNumbering = True
    Exit Function
    
InitError:
    MsgBox "Letter numbering initialization error: " & Err.description, vbCritical
    InitializeLetterNumbering = False
End Function

Private Function FormatLetterDate(ByVal InputDate As Date) As String
    Dim MonthNames As Variant
    MonthNames = Array("", "January", "February", "March", "April", "May", "June", _
                       "July", "August", "September", "October", "November", "December")
    
    FormatLetterDate = Day(InputDate) & " " & MonthNames(Month(InputDate)) & " " & Year(InputDate)
End Function

Private Function FormatShortDate(ByVal InputDate As Date) As String
    On Error Resume Next
    FormatShortDate = Format(InputDate, "dd.mm.yy")
    If Err.Number <> 0 Then
        FormatShortDate = Format(Now, "dd.mm.yy")
    End If
    On Error GoTo 0
End Function

Private Function ShouldIncludeDoverennost(WsDover As Worksheet, RowNumber As Long) As Boolean
    Dim OperationDateCol As Long
    Dim OperationNumberCol As Long
    Dim ExistingDate As Date
    Dim ExistingNumber As String
    Dim MonthsDiff As Long
    
    ShouldIncludeDoverennost = True
    
    OperationDateCol = FindColumnByName(WsDover, "OPERATION DATE")
    OperationNumberCol = FindColumnByName(WsDover, "OPERATION NUMBER")
    
    If OperationDateCol = 0 Or OperationNumberCol = 0 Then Exit Function
    
    ExistingNumber = Trim(CStr(WsDover.Cells(RowNumber, OperationNumberCol).value))
    
    If Len(ExistingNumber) > 0 And InStr(ExistingNumber, "/") > 0 Then
        On Error Resume Next
        ExistingDate = CDate(WsDover.Cells(RowNumber, OperationDateCol).value)
        On Error GoTo 0
        
        If ExistingDate > DateSerial(1900, 1, 2) Then
            MonthsDiff = DateDiff("m", ExistingDate, CurrentDate)
            If MonthsDiff < 3 Then
                ShouldIncludeDoverennost = False
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

Private Sub UpdateDoverennostiFile(WbDover As Workbook, GroupData As CorrespondentGroup, LetterInfo As LetterInfo)
    Dim WsDover As Worksheet
    Dim OperationDateCol As Long
    Dim OperationNumberCol As Long
    Dim i As Long
    
    On Error GoTo UpdateError
    
    Set WsDover = WbDover.Worksheets(1)
    
    OperationDateCol = FindColumnByName(WsDover, "OPERATION DATE")
    OperationNumberCol = FindColumnByName(WsDover, "OPERATION NUMBER")
    
    If OperationDateCol = 0 Or OperationNumberCol = 0 Then
        MsgBox "ERROR: Operation columns not found in Doverennosti file!", vbCritical, "Update Error"
        Exit Sub
    End If
    
    For i = 0 To GroupData.Count - 1
        With GroupData.Doverennosti(i)
            WsDover.Cells(.RowNumber, OperationNumberCol).NumberFormat = "@"
            WsDover.Cells(.RowNumber, OperationNumberCol).value = LetterInfo.Number
            WsDover.Cells(.RowNumber, OperationDateCol).value = LetterInfo.Date
        End With
    Next i
    
    On Error Resume Next
    WbDover.Save
    If Err.Number <> 0 Then
        Dim BackupPath As String
        BackupPath = Replace(WbDover.FullName, ".xlsx", "_updated_" & Format(Now, "ddmmyyyy_hhmmss") & ".xlsx")
        WbDover.SaveAs BackupPath
        
        MsgBox "WARNING!" & vbCrLf & vbCrLf & _
               "Original file locked." & vbCrLf & _
               "Saved as copy:" & vbCrLf & vbCrLf & _
               BackupPath, vbExclamation, "Saved as copy"
    End If
    On Error GoTo 0
    Exit Sub
    
UpdateError:
    MsgBox "CRITICAL ERROR updating Doverennosti file!" & vbCrLf & vbCrLf & _
           "Error: " & Err.description, vbCritical, "Critical Error"
End Sub

