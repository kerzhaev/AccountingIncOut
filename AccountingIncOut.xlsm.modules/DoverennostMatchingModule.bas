Attribute VB_Name = "DoverennostMatchingModule"
'==============================================
' DoverennostMatchingModule - THREE-PASS SYSTEM
' Version: 1.2.2
' Date: 07.09.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' =============================================
' DATA STRUCTURES
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
' SMART SEARCH STRUCTURES
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

' Structure for number+date pair
Public Type NumberDatePair
    Number As String
    DateStr As String
    OtPosition As Long  ' Position of word "FROM" for debugging
End Type

' =============================================
' MAIN PROCEDURE
' =============================================

Public Sub ProcessDoverennostMatching()
    Dim FileDoverennosti As String
    Dim FileOperations As String
    Dim WbDoverennosti As Workbook
    Dim WbOperations As Workbook
    
    On Error GoTo ProcessError
    
    Debug.Print "=== START OF THREE-PASS SYSTEM WITH FIXED REGEX ==="
    
    FileDoverennosti = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        LocalizationManager.GetText("Select Doverennosti journal file"))
    If FileDoverennosti = "False" Then Exit Sub
    
    FileOperations = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        LocalizationManager.GetText("Select 1C operations journal file"))
    If FileOperations = "False" Then Exit Sub
    
    Set WbDoverennosti = Workbooks.Open(FileDoverennosti, ReadOnly:=False)
    Set WbOperations = Workbooks.Open(FileOperations, ReadOnly:=True)
    
    Call PerformThreePassMatching(WbDoverennosti, WbOperations)
    
    WbOperations.Close False
    WbDoverennosti.Save
    WbDoverennosti.Close
    
    Debug.Print "=== THREE-PASS SYSTEM WITH FIXED REGEX COMPLETED ==="
    
    Exit Sub
    
ProcessError:
    On Error Resume Next
    If Not WbOperations Is Nothing Then WbOperations.Close False
    If Not WbDoverennosti Is Nothing Then WbDoverennosti.Close False
    
    Debug.Print "Error: " & Err.description
    MsgBox LocalizationManager.GetText("Error: ") & Err.description, vbCritical
End Sub

' =============================================
' @author Evgeniy Kerzhaev, FKU "95 FES" MO RF
' @description Master scenario: user selects (1) master file, (2) period export, (3) 1C export. Macro merges without duplicates and matches with 1C. Cache is stored in master file.
' =============================================
Public Sub ProcessDoverennostMasterMatching()
    Dim FileMaster As String
    Dim FilePeriod As String
    Dim FileOperations As String
    
    Dim WbMaster As Workbook
    Dim WbPeriod As Workbook
    Dim WbOperations As Workbook
    
    Dim AddedCount As Long
    
    On Error GoTo ProcessError
    
    Debug.Print "=== MASTER: Doverennosti Matching (Merge + 1C) ==="
    
    FileMaster = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        LocalizationManager.GetText("Select MASTER Doverennosti file (cumulative)"))
    If FileMaster = "False" Then Exit Sub
    
    FilePeriod = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        LocalizationManager.GetText("Select period Doverennosti file (export)"))
    If FilePeriod = "False" Then Exit Sub
    
    FileOperations = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv", , _
        LocalizationManager.GetText("Select 1C operations journal file (export)"))
    If FileOperations = "False" Then Exit Sub
    
    Application.StatusBar = LocalizationManager.GetText("Opening files...")
    Set WbMaster = Workbooks.Open(FileMaster, ReadOnly:=False)
    Set WbPeriod = Workbooks.Open(FilePeriod, ReadOnly:=True)
    Set WbOperations = Workbooks.Open(FileOperations, ReadOnly:=True)
    
    Application.StatusBar = LocalizationManager.GetText("Merging Doverennosti into master...")
    Call MergePeriodDoverennostiIntoMaster(WbMaster.Worksheets(1), WbPeriod.Worksheets(1), AddedCount)
    
    Application.StatusBar = LocalizationManager.GetText("Matching master file with 1C operations...")
    Call PerformThreePassMatching(WbMaster, WbOperations)
    
    Application.StatusBar = LocalizationManager.GetText("Saving master file...")
    WbOperations.Close False
    WbPeriod.Close False
    WbMaster.Save
    WbMaster.Close
    
    Application.StatusBar = False
    
    MsgBox LocalizationManager.GetText("Done.") & vbCrLf & vbCrLf & _
           LocalizationManager.GetText("Added new records to master: ") & AddedCount, vbInformation, LocalizationManager.GetText("Doverennosti Matching")
    
    Exit Sub
    
ProcessError:
    On Error Resume Next
    If Not WbOperations Is Nothing Then WbOperations.Close False
    If Not WbPeriod Is Nothing Then WbPeriod.Close False
    If Not WbMaster Is Nothing Then WbMaster.Close False
    Application.StatusBar = False
    
    Debug.Print "Master matching error: " & Err.description
    MsgBox LocalizationManager.GetText("Error: ") & Err.description, vbCritical
End Sub

' =============================================
' THREE-PASS MATCHING FUNCTION
' =============================================

Private Sub PerformThreePassMatching(WbDover As Workbook, WbOper As Workbook)
    Dim WsDover As Worksheet, WsOper As Worksheet
    Dim LastRowDover As Long, LastRowOper As Long
    Dim i As Long
    Dim ProcessedCount As Long, SkippedEmpty As Long, SkippedFound As Long
    
    ' Pass statistics
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
    
    ' Filter operations
    Dim FilteredOperData As Variant
    Dim FilteredCount As Long
    FilteredOperData = FilterOperationsByCorrespondent(OperData, FilteredCount, SkippedEmpty)
    
    Call AddThreePassResultColumns(WsDover)
    
    ' Define result column for checking already found
    Dim ResultColumn As Long
    ResultColumn = FindResultColumn(WsDover)
    
    ' Define correspondent column in doverennosti
    Dim CorrespondentColumn As Long
    CorrespondentColumn = FindCorrespondentColumn(WsDover)
    
    Debug.Print "THREE-PASS SYSTEM: " & (LastRowDover - 1) & " records vs " & FilteredCount & " operations"
    Debug.Print "Skipped " & SkippedEmpty & " operations with empty correspondent"
    
    If CorrespondentColumn > 0 Then
        Debug.Print "Correspondent column found: " & CorrespondentColumn
    Else
        Debug.Print "WARNING: Correspondent column not found, correspondent check will be skipped"
    End If
    
    ' Arrays to store unfound records
    Dim Pass1NotFound() As Long, Pass1NotFoundData() As ParsedNaryad
    Dim Pass2NotFound() As Long, Pass2NotFoundData() As ParsedNaryad
    Dim Pass1Count As Long, Pass2Count As Long
    
    Pass1Count = 0
    Pass2Count = 0
    
    Dim DoverComment As String
    Dim ParsedDover As ParsedNaryad
    Dim MatchResult As MatchResult
    Dim CurrentRow As Long
    
    ' MAIN PROCESSING LOOP
    For i = 1 To UBound(DoverData, 1)
        CurrentRow = i + 1
        DoverComment = Trim(CStr(DoverData(i, 9)))
        
        If i Mod 25 = 0 Then
            Application.StatusBar = LocalizationManager.GetText("3-pass system running: ") & i & LocalizationManager.GetText(" of ") & UBound(DoverData, 1)
        End If
        
        ' Check if row was already processed and FOUND
        If IsAlreadyProcessed(WsDover, CurrentRow, ResultColumn) Then
            SkippedFound = SkippedFound + 1
            Debug.Print "Skipping row " & CurrentRow & " - already found"
            GoTo NextIteration
        End If
        
        If DoverComment <> "" Then
            ' Extract doverennost correspondent
            Dim DoverCorrespondent As String
            DoverCorrespondent = ""
            If CorrespondentColumn > 0 Then
                DoverCorrespondent = Trim(CStr(WsDover.Cells(CurrentRow, CorrespondentColumn).value))
            End If
            
            ' === 1ST PASS: EXACT ORDERS AND OTHER DOCUMENTS ===
            ParsedDover = ParseNaryadForSubstring(DoverComment)
            
            ' Check if parsed as valid document type for 1st pass
            If ParsedDover.IsValid Then
                Dim ValidForFirstPass As Boolean
                Dim DocTypeUpper As String
                DocTypeUpper = UCase(ParsedDover.DocumentType)
                ValidForFirstPass = (DocTypeUpper = "ORDER" Or DocTypeUpper = "CHT" Or _
                                     DocTypeUpper = "ALLOTMENT" Or DocTypeUpper = "ORDER-REQUEST" Or _
                                     DocTypeUpper = "REQUEST ORDER")
                
                If ValidForFirstPass Then
                    MatchResult = FindBySubstringEnhanced(ParsedDover, FilteredOperData, DoverCorrespondent)
                    MatchResult.PassNumber = 1
                    
                    If MatchResult.FoundMatch Then
                        Pass1Found = Pass1Found + 1
                        MatchResult.MatchDetails = LocalizationManager.GetText("1st pass: ") & MatchResult.MatchDetails
                    Else
                        ' Save for 2nd pass
                        Pass1Count = Pass1Count + 1
                        ReDim Preserve Pass1NotFound(Pass1Count - 1)
                        ReDim Preserve Pass1NotFoundData(Pass1Count - 1)
                        Pass1NotFound(Pass1Count - 1) = CurrentRow
                        Pass1NotFoundData(Pass1Count - 1) = ParsedDover
                        MatchResult.MatchDetails = LocalizationManager.GetText("Awaiting 2nd pass")
                    End If
                Else
                    ' Other document types - save for 3rd pass
                    Pass2Count = Pass2Count + 1
                    ReDim Preserve Pass2NotFound(Pass2Count - 1)
                    ReDim Preserve Pass2NotFoundData(Pass2Count - 1)
                    Pass2NotFound(Pass2Count - 1) = CurrentRow
                    Pass2NotFoundData(Pass2Count - 1) = ParsedDover
                    
                    MatchResult = CreateEmptyMatchResult(LocalizationManager.GetText("Awaiting 3rd pass"))
                    MatchResult.PassNumber = 0
                End If
            Else
                ' If not parsed - save for 3rd pass
                Pass2Count = Pass2Count + 1
                ReDim Preserve Pass2NotFound(Pass2Count - 1)
                ReDim Preserve Pass2NotFoundData(Pass2Count - 1)
                Pass2NotFound(Pass2Count - 1) = CurrentRow
                Pass2NotFoundData(Pass2Count - 1).OriginalText = DoverComment
                Pass2NotFoundData(Pass2Count - 1).IsValid = False
                
                MatchResult = CreateEmptyMatchResult(LocalizationManager.GetText("Awaiting 3rd pass"))
                MatchResult.PassNumber = 0
            End If
            
            ProcessedCount = ProcessedCount + 1
        Else
            MatchResult = CreateEmptyMatchResult(LocalizationManager.GetText("Empty comment"))
            MatchResult.PassNumber = 0
        End If
        
        Call WriteThreePassResult(WsDover, CurrentRow, MatchResult)
        
NextIteration:
    Next i

    ' === 2ND PASS: ADAPTIVE REGEX ===
    If Pass1Count > 0 Then
        Debug.Print "=== 2ND PASS: ADAPTIVE REGEX ==="
        Pass2Found = ProcessSecondPassWithSupplierService(WsDover, FilteredOperData, Pass1NotFound, Pass1NotFoundData, Pass1Count, Pass2NotFound, Pass2NotFoundData, Pass2Count, CorrespondentColumn)
    End If
    
    ' === 3RD PASS: UNIVERSAL DOCUMENTS ===
    If Pass2Count > 0 Then
        Debug.Print "=== 3RD PASS: UNIVERSAL DOCUMENTS ==="
        Pass3Found = ProcessThirdPass(WsDover, FilteredOperData, Pass2NotFound, Pass2NotFoundData, Pass2Count, Pass3ByType)
        Debug.Print "3rd pass found: " & Pass3Found
    End If
    
    ' Restore settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    Dim TotalTime As Double
    TotalTime = Timer - StartTime
    
    Debug.Print "=== RESULTS OF THREE-PASS SYSTEM ==="
    Debug.Print "Execution time: " & Format(TotalTime, "0.0") & " sec"
    Debug.Print "Processed: " & ProcessedCount
    Debug.Print "Skipped already found: " & SkippedFound
    Debug.Print "1st pass (exact orders): " & Pass1Found
    Debug.Print "2nd pass (regex orders): " & Pass2Found
    Debug.Print "3rd pass (other docs): " & Pass3Found
    Debug.Print "Total found: " & (Pass1Found + Pass2Found + Pass3Found)
    Debug.Print "Skipped operations (empty correspondent): " & SkippedEmpty
    
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
    
    Debug.Print "Three-pass system error: " & Err.description
    MsgBox LocalizationManager.GetText("Error: ") & Err.description, vbCritical
End Sub

' =============================================
' @author Evgeniy Kerzhaev, FKU "95 FES" MO RF
' @description Merge period export into master sheet without duplicates.
' =============================================
Private Sub MergePeriodDoverennostiIntoMaster(WsMaster As Worksheet, WsPeriod As Worksheet, ByRef AddedCount As Long)
    Dim LastRowMaster As Long, LastRowPeriod As Long
    Dim MasterData As Variant, PeriodData As Variant
    Dim Dict As Object
    Dim i As Long
    Dim key As String
    
    On Error GoTo ErrorHandler
    
    AddedCount = 0
    
    LastRowMaster = WsMaster.Cells(WsMaster.Rows.Count, 1).End(xlUp).Row
    LastRowPeriod = WsPeriod.Cells(WsPeriod.Rows.Count, 1).End(xlUp).Row
    
    If LastRowPeriod < 2 Then Exit Sub
    
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = 1 ' TextCompare
    
    If LastRowMaster >= 2 Then
        MasterData = WsMaster.Range("A2:I" & LastRowMaster).Value2
        For i = 1 To UBound(MasterData, 1)
            key = BuildMasterRowKey(MasterData, i)
            If key <> "" Then
                If Not Dict.Exists(key) Then Dict.Add key, True
            End If
        Next i
    End If
    
    PeriodData = WsPeriod.Range("A2:I" & LastRowPeriod).Value2
    
    For i = 1 To UBound(PeriodData, 1)
        key = BuildMasterRowKey(PeriodData, i)
        If key <> "" Then
            If Not Dict.Exists(key) Then
                LastRowMaster = LastRowMaster + 1
                WsMaster.Range("A" & LastRowMaster & ":I" & LastRowMaster).Value2 = GetRowAs2DArray(PeriodData, i, 9)
                Dict.Add key, True
                AddedCount = AddedCount + 1
            End If
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Master file merge error: " & Err.description
    Err.Raise Err.Number, , "Master file merge error: " & Err.description
End Sub

' =============================================
' @author Evgeniy Kerzhaev, FKU "95 FES" MO RF
' @description Generates normalized row key.
' =============================================
Private Function BuildMasterRowKey(DataArr As Variant, RowIndex As Long) As String
    Dim j As Long
    Dim key As String
    
    On Error GoTo ErrorHandler
    
    For j = 1 To 9
        If j = 1 Then
            key = NormalizeKeyPart(DataArr(RowIndex, j))
        Else
            key = key & "|" & NormalizeKeyPart(DataArr(RowIndex, j))
        End If
    Next j
    
    BuildMasterRowKey = key
    Exit Function
    
ErrorHandler:
    BuildMasterRowKey = ""
End Function

' =============================================
' @author Evgeniy Kerzhaev, FKU "95 FES" MO RF
' @description Normalizes key part element.
' =============================================
Private Function NormalizeKeyPart(ByVal v As Variant) As String
    Dim s As String
    
    On Error GoTo ErrorHandler
    
    If IsError(v) Then
        NormalizeKeyPart = ""
        Exit Function
    End If
    
    If IsEmpty(v) Or IsNull(v) Then
        NormalizeKeyPart = ""
        Exit Function
    End If
    
    If IsDate(v) Then
        NormalizeKeyPart = Format(CDate(v), "yyyy-mm-dd hh:nn:ss")
        Exit Function
    End If
    
    s = CStr(v)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim(s)
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeKeyPart = UCase(s)
    Exit Function
    
ErrorHandler:
    NormalizeKeyPart = ""
End Function

' =============================================
' @author Evgeniy Kerzhaev, FKU "95 FES" MO RF
' @description Returns row from 2D array as 1xN 2D array.
' =============================================
Private Function GetRowAs2DArray(DataArr As Variant, RowIndex As Long, ColCount As Long) As Variant
    Dim OutArr() As Variant
    Dim j As Long
    
    ReDim OutArr(1 To 1, 1 To ColCount)
    For j = 1 To ColCount
        OutArr(1, j) = DataArr(RowIndex, j)
    Next j
    
    GetRowAs2DArray = OutArr
End Function

' =============================================
' 2ND PASS: SMART COMPONENT SEARCH
' =============================================

Private Function ProcessSecondPassWithSmartComponent(WsDover As Worksheet, OperData As Variant, NotFoundRows() As Long, NotFoundData() As ParsedNaryad, NotFoundCount As Long, ByRef Pass2NotFound() As Long, ByRef Pass2NotFoundData() As ParsedNaryad, ByRef Pass2Count As Long) As Long
    Dim i As Long
    Dim SecondPassFound As Long
    Dim MatchResult As MatchResult
    
    SecondPassFound = 0
    
    Debug.Print "=== 2ND PASS: SMART COMPONENT SEARCH ==="
    
    For i = 0 To NotFoundCount - 1
        Debug.Print "2nd pass SMART for: " & NotFoundData(i).DocumentNumber & " from " & NotFoundData(i).DocumentDate
        
        MatchResult = FindBySmartComponents(NotFoundData(i), OperData)
        MatchResult.PassNumber = 2
        
        If MatchResult.FoundMatch Then
            SecondPassFound = SecondPassFound + 1
            MatchResult.MatchDetails = LocalizationManager.GetText("2nd pass smart (") & Format(MatchResult.MatchScore, "0") & "%): " & MatchResult.MatchDetails
            
            Call WriteThreePassResult(WsDover, NotFoundRows(i), MatchResult)
            
            Debug.Print "  [OK] Found by SMART 2nd pass! Confidence: " & MatchResult.MatchScore & "%"
        Else
            ' Save for 3rd pass
            Pass2Count = Pass2Count + 1
            ReDim Preserve Pass2NotFound(Pass2Count - 1)
            ReDim Preserve Pass2NotFoundData(Pass2Count - 1)
            Pass2NotFound(Pass2Count - 1) = NotFoundRows(i)
            Pass2NotFoundData(Pass2Count - 1) = NotFoundData(i)
            
            Debug.Print "  [X] Not found by smart search, moves to 3rd pass"
        End If
    Next i
    
    ProcessSecondPassWithSmartComponent = SecondPassFound
End Function

Private Function FindBySmartComponents(DoverData As ParsedNaryad, OperData As Variant) As MatchResult
    Dim result As MatchResult
    Dim i As Long
    Dim OperComment As String
    
    result.FoundMatch = False
    result.MatchScore = 0
    
    ' Check each operation
    For i = 1 To UBound(OperData, 1)
        OperComment = Trim(CStr(OperData(i, 9)))
        
        If OperComment <> "" Then
            Dim ComponentMatch As ComponentMatchResult
            ComponentMatch = AnalyzeComponents(OperComment, DoverData.DocumentNumber, DoverData.DocumentDate)
            
            If ComponentMatch.IsMatch And ComponentMatch.Confidence >= 25 Then
                Dim CorrespondentBonus As Double
                CorrespondentBonus = CheckCorrespondentMatch(OperData(i, 6), DoverData.OriginalText, i + 1)
                ComponentMatch.Confidence = ComponentMatch.Confidence + CorrespondentBonus
                
                If CorrespondentBonus > 0 Then
                    ComponentMatch.MatchDetails = ComponentMatch.MatchDetails & LocalizationManager.GetText(" corr:+") & Format(CorrespondentBonus, "0") & "%"
                End If
                
                If ComponentMatch.Confidence >= 40 Then
                    result.FoundMatch = True
                    result.MatchScore = ComponentMatch.Confidence
                    result.OperationRow = i + 1
                    result.OperationNumber = Trim(CStr(OperData(i, 3)))
                    
                    ComponentMatch.MatchDetails = ComponentMatch.MatchDetails & " | 1C: " & OperComment
                    result.MatchDetails = ComponentMatch.MatchDetails
                    
                    On Error Resume Next
                    result.OperationDate = CDate(OperData(i, 2))
                    On Error GoTo 0
                    
                    Exit For
                End If
            End If
        End If
    Next i
    
    If Not result.FoundMatch Then
        result.MatchDetails = LocalizationManager.GetText("smart search did not find suitable components")
    End If
    
    FindBySmartComponents = result
End Function

Private Sub CreateFlexibleTextPatterns(DocumentNumber As String, DocumentDate As String, ByRef Patterns() As String)
    Dim PatternCount As Long
    PatternCount = 0
    
    ReDim Patterns(15)
    
    Debug.Print "    Creating SIMPLIFIED patterns for: No '" & DocumentNumber & "' date '" & DocumentDate & "'"
    
    Dim SimpleNumber As String
    Dim SimpleDatePatterns() As String
    
    SimpleNumber = EscapeMinimal(DocumentNumber)
    Call CreateSimpleDatePatterns(DocumentDate, SimpleDatePatterns)
    
    Dim DateVar As Long
    For DateVar = 0 To UBound(SimpleDatePatterns)
        If PatternCount >= 15 Then Exit For
        Patterns(PatternCount) = "order\s*No.?\s*" & SimpleNumber & "\s*from\s*" & SimpleDatePatterns(DateVar)
        PatternCount = PatternCount + 1
    Next DateVar
    
    If InStr(DocumentNumber, "/") > 0 And PatternCount < 14 Then
        Patterns(PatternCount) = CreateSlashPattern(DocumentNumber, DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    If InStr(DocumentNumber, "-") > 0 And PatternCount < 14 Then
        Patterns(PatternCount) = CreateDashPattern(DocumentNumber, DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    If ContainsMilitaryCode(DocumentNumber) And PatternCount < 14 Then
        Patterns(PatternCount) = CreateMilitaryPattern(DocumentNumber, DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    If IsNumericOnly(DocumentNumber) And PatternCount < 14 Then
        Patterns(PatternCount) = CreateNumericPattern(DocumentNumber, DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    If PatternCount < 15 Then
        Patterns(PatternCount) = "(?i)order.*?" & EscapeMinimal(DocumentNumber) & ".*?from.*?" & EscapeMinimal(DocumentDate)
        PatternCount = PatternCount + 1
    End If
    
    ReDim Preserve Patterns(PatternCount - 1)
End Sub

Private Function FindPatternInText(TextToSearch As String, Patterns() As String, ByRef MatchedPattern As String, ByRef MatchedText As String) As Boolean
    Dim RegEx As Object
    Dim Matches As Object
    Dim Match As Object
    
    Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        .IgnoreCase = True
        .Global = True
        .Multiline = True
    End With
    
    Dim i As Long
    
    For i = 0 To UBound(Patterns)
        RegEx.Pattern = Patterns(i)
        
        On Error Resume Next
        Set Matches = RegEx.Execute(TextToSearch)
        
        If Matches.Count > 0 Then
            Set Match = Matches(0)
            MatchedPattern = Patterns(i)
            MatchedText = Match.value
            
            FindPatternInText = True
            Exit Function
        End If
        On Error GoTo 0
    Next i
    
    FindPatternInText = False
End Function

Private Sub CreateSimpleDatePatterns(OriginalDate As String, ByRef DatePatterns() As String)
    Dim PatternCount As Long
    PatternCount = 0
    
    ReDim DatePatterns(5)
    
    Dim DateParts() As String
    DateParts = Split(OriginalDate, ".")
    
    If UBound(DateParts) >= 2 Then
        Dim Day As String, Month As String, Year As String
        Day = DateParts(0)
        Month = DateParts(1)
        Year = DateParts(2)
        
        DatePatterns(PatternCount) = Day & "\." & Month & "\." & Year & "(?:\s*y\.?)?"
        PatternCount = PatternCount + 1
        
        DatePatterns(PatternCount) = CLng(Day) & "\." & CLng(Month) & "\." & Year & "(?:\s*y\.?)?"
        PatternCount = PatternCount + 1
        
        If Len(Year) = 4 Then
            DatePatterns(PatternCount) = CLng(Day) & "\." & CLng(Month) & "\." & Right(Year, 2) & "(?:\s*y\.?)?"
            PatternCount = PatternCount + 1
        End If
        
        DatePatterns(PatternCount) = "\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2})(?:\s*y\.?)?"
        PatternCount = PatternCount + 1
        
    Else
        DatePatterns(PatternCount) = EscapeMinimal(OriginalDate)
        PatternCount = PatternCount + 1
    End If
    
    ReDim Preserve DatePatterns(PatternCount - 1)
End Sub

Private Function EscapeMinimal(Text As String) As String
    Dim result As String
    result = Text
    
    result = Replace(result, ".", "\.")
    result = Replace(result, "(", "\(")
    result = Replace(result, ")", "\)")
    result = Replace(result, "[", "\[")
    result = Replace(result, "]", "\]")
    result = Replace(result, "*", "\*")
    result = Replace(result, "+", "\+")
    result = Replace(result, "?", "\?")
    result = Replace(result, "^", "\^")
    result = Replace(result, "$", "\$")
    result = Replace(result, "|", "\|")
    
    EscapeMinimal = result
End Function

Public Sub AddCustomPatternForText(ByRef Patterns() As String, ByRef PatternCount As Long, DocumentNumber As String, DocumentDate As String)
    ' Custom pattern hook
End Sub

Private Function ExtractAllNumberDatePairs(Text As String) As NumberDatePair()
    Dim Pairs() As NumberDatePair
    Dim PairCount As Long
    Dim UpperText As String
    Dim SearchPos As Long
    Dim OtPos As Long
    
    UpperText = UCase(Text)
    PairCount = 0
    ReDim Pairs(0 To 9)
    
    SearchPos = 1
    
    Do
        OtPos = InStr(SearchPos, UpperText, " FROM ")
        If OtPos = 0 Then
            OtPos = InStr(SearchPos, UpperText, "FROM ")
            If OtPos > 0 And OtPos = SearchPos Then
                ' Valid
            ElseIf OtPos > 0 Then
                If OtPos > 1 And Mid(UpperText, OtPos - 1, 1) <> " " Then
                    OtPos = 0
                End If
            End If
        End If
        
        If OtPos = 0 Then Exit Do
        
        Dim NumberBefore As String
        NumberBefore = ExtractNumberBeforeOt(Text, OtPos)
        
        Dim DateAfter As String
        DateAfter = ExtractDateAfterOt(Text, OtPos)
        
        If NumberBefore <> "" And DateAfter <> "" Then
            Pairs(PairCount).Number = NumberBefore
            Pairs(PairCount).DateStr = DateAfter
            Pairs(PairCount).OtPosition = OtPos
            PairCount = PairCount + 1
            
            If PairCount >= 10 Then Exit Do
        End If
        
        SearchPos = OtPos + 4
        
    Loop
    
    ReDim Preserve Pairs(0 To IIf(PairCount > 0, PairCount - 1, 0))
    ExtractAllNumberDatePairs = Pairs
End Function

Private Function ExtractNumberBeforeOt(Text As String, OtPos As Long) As String
    Dim BeforeOt As String
    Dim CleanNumber As String
    Dim KnownPrefixes As Variant
    Dim j As Long
    Dim i As Long
    Dim Char As String
    Dim HasValidChar As Boolean
    Dim StartPos As Long
    Dim UpperText As String
    
    On Error GoTo ErrorHandler
    
    UpperText = UCase(Text)
    
    StartPos = OtPos - 1
    Do While StartPos > 0
        Char = Mid(UpperText, StartPos, 1)
        If Char = " " Or Char = "," Or Char = ";" Or Char = "." Then
            Exit Do
        End If
        StartPos = StartPos - 1
    Loop
    If StartPos > 0 Then StartPos = StartPos + 1
    
    BeforeOt = Trim(Mid(Text, StartPos, OtPos - StartPos))
    
    KnownPrefixes = Array("ORDER", "VS ORDER", "GSM REQUEST", "UAV TELEGRAM", "RAV ORDER", "BT ORDER", "PS ORDER", "SS ORDER", _
                          "VS", "GSM", "RAV", "BT", "PS", "SS", "UAV", "AS SEIZURE ACT", "MC EVALUATION ACT", "ORDER VMU YuVO", _
                          "ORDER EXTRACT", "DEFECT LIST", "AS", "SEIZURE ACT", "EVALUATION ACT", "VMU", "YuVO", "EXTRACT", "ORDER", "DEFECT", "LIST", "ACT")
    
    CleanNumber = UCase(BeforeOt)
    
    For j = 0 To UBound(KnownPrefixes)
        If Left(CleanNumber, Len(KnownPrefixes(j))) = KnownPrefixes(j) Then
            CleanNumber = Trim(Mid(CleanNumber, Len(KnownPrefixes(j)) + 1))
            Exit For
        End If
    Next j
    
    CleanNumber = Replace(CleanNumber, "No.", "")
    CleanNumber = Replace(CleanNumber, "#", "")
    CleanNumber = Trim(CleanNumber)
    
    Do While InStr(CleanNumber, "  ") > 0
        CleanNumber = Replace(CleanNumber, "  ", " ")
    Loop
    
    If CleanNumber <> "" Then
        HasValidChar = False
        For i = 1 To Len(CleanNumber)
            Char = Mid(CleanNumber, i, 1)
            If (Char >= "0" And Char <= "9") Or _
               (Char >= "A" And Char <= "Z") Or _
               Char = "/" Or Char = "-" Then
                HasValidChar = True
                Exit For
            End If
        Next i
        
        If HasValidChar Then
            ExtractNumberBeforeOt = CleanNumber
            Exit Function
        End If
    End If
    
    ExtractNumberBeforeOt = ""
    Exit Function
    
ErrorHandler:
    ExtractNumberBeforeOt = ""
End Function

Private Function ExtractDateAfterOt(Text As String, OtPos As Long) As String
    Dim DateStart As Long
    Dim DateStr As String
    Dim i As Long
    Dim Char As String
    Dim UpperText As String
    
    On Error GoTo ErrorHandler
    
    UpperText = UCase(Text)
    
    DateStart = OtPos + 4
    If Mid(UpperText, OtPos, 6) = " FROM " Then
        DateStart = OtPos + 6
    End If
    
    Do While DateStart <= Len(Text) And Mid(Text, DateStart, 1) = " "
        DateStart = DateStart + 1
    Loop
    
    For i = DateStart To Len(Text)
        Char = Mid(Text, i, 1)
        If (Char >= "0" And Char <= "9") Or Char = "." Then
            DateStr = DateStr & Char
        ElseIf DateStr <> "" And (Char = " " Or Char = "y" Or Char = "Y" Or Char = "," Or Char = ";" Or Char = ".") Then
            Exit For
        End If
    Next i
    
    If Len(DateStr) >= 8 And InStr(DateStr, ".") > 0 Then
        ExtractDateAfterOt = NormalizeDateForSubstring(DateStr)
    Else
        ExtractDateAfterOt = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractDateAfterOt = ""
End Function

Private Function ParseNaryadForSubstring(Text As String) As ParsedNaryad
    Dim result As ParsedNaryad
    Dim UpperText As String
    Dim DocumentTypes As Variant
    Dim i As Long
    Dim FoundType As String
    
    result.OriginalText = Text
    result.IsValid = False
    result.DocumentType = ""
    
    On Error GoTo SubstringParseError
    
    UpperText = UCase(Text)
    
    DocumentTypes = Array("ORDER", "CHT", "CERTIFICATE", "ALLOTMENT", "ORDER-REQUEST", "REQUEST ORDER", "TELEGRAM")
    
    For i = 0 To UBound(DocumentTypes)
        If InStr(1, UpperText, CStr(DocumentTypes(i))) > 0 Then
            FoundType = CStr(DocumentTypes(i))
            Exit For
        End If
    Next i
    
    If FoundType = "" Then
        result.MatchDetails = LocalizationManager.GetText("No keywords found")
        ParseNaryadForSubstring = result
        Exit Function
    End If
    
    result.DocumentType = FoundType
    result.DocumentNumber = ExtractSubstringNumber(Text, FoundType)
    result.DocumentDate = ExtractSubstringDate(Text)
    
    If result.DocumentNumber <> "" And result.DocumentDate <> "" Then
        result.IsValid = True
        result.MatchDetails = LocalizationManager.GetText("Substring parse: success")
    Else
        result.MatchDetails = LocalizationManager.GetText("Substring parse: no number or date")
    End If
    
    ParseNaryadForSubstring = result
    Exit Function
    
SubstringParseError:
    result.IsValid = False
    result.MatchDetails = LocalizationManager.GetText("Parse error: ") & Err.description
    ParseNaryadForSubstring = result
End Function

Private Function ExtractSubstringNumber(Text As String, DocumentType As String) As String
    Dim UpperText As String
    Dim TypePos As Long
    Dim FromPos As Long
    Dim BetweenText As String
    
    UpperText = UCase(Text)
    
    TypePos = InStr(UpperText, UCase(DocumentType))
    If TypePos = 0 Then
        ExtractSubstringNumber = ""
        Exit Function
    End If
    
    FromPos = InStr(TypePos, UpperText, " FROM ")
    If FromPos = 0 Then FromPos = InStr(TypePos, UpperText, "FROM ")
    
    If FromPos > 0 Then
        BetweenText = Trim(Mid(Text, TypePos + Len(DocumentType), FromPos - TypePos - Len(DocumentType)))
        
        BetweenText = Replace(BetweenText, "No.", "")
        BetweenText = Replace(BetweenText, "#", "")
        BetweenText = Trim(BetweenText)
        
        If BetweenText <> "" Then
            ExtractSubstringNumber = BetweenText
            Exit Function
        End If
    End If
    
    ExtractSubstringNumber = ""
End Function

Private Function ExtractSubstringDate(Text As String) As String
    Dim UpperText As String
    Dim FromPos As Long
    Dim DateStart As Long
    Dim DateStr As String
    Dim i As Long, Char As String
    
    UpperText = UCase(Text)
    
    FromPos = InStr(UpperText, " FROM ")
    If FromPos = 0 Then FromPos = InStr(UpperText, "FROM ")
    
    If FromPos > 0 Then
        DateStart = FromPos + 5
        If Mid(UpperText, FromPos, 6) = " FROM " Then DateStart = FromPos + 6
        
        Do While DateStart <= Len(Text) And Mid(Text, DateStart, 1) = " "
            DateStart = DateStart + 1
        Loop
        
        For i = DateStart To Len(Text)
            Char = Mid(Text, i, 1)
            If (Char >= "0" And Char <= "9") Or Char = "." Then
                DateStr = DateStr & Char
            ElseIf DateStr <> "" And (Char = " " Or Char = "y" Or Char = "Y") Then
                Exit For
            End If
        Next i
        
        If Len(DateStr) >= 8 And InStr(DateStr, ".") > 0 Then
            DateStr = NormalizeDateForSubstring(DateStr)
            ExtractSubstringDate = DateStr
            Exit Function
        End If
    End If
    
    ExtractSubstringDate = ""
End Function

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
    Else
        NormalizeDateForSubstring = DateStr
    End If
    
    Exit Function
    
NormalizeError:
    NormalizeDateForSubstring = DateStr
End Function

Private Function FindCorrespondentColumn(Ws As Worksheet) As Long
    Dim i As Long
    Dim HeaderNames As Variant
    Dim HeaderValue As String
    
    HeaderNames = Array("CORRESPONDENT", "RECEIVED FROM", "ADDRESSEE", "SUPPLIER")
    
    FindCorrespondentColumn = 0
    
    For i = 1 To 20
        HeaderValue = Trim(UCase(CStr(Ws.Cells(1, i).value)))
        Dim j As Long
        For j = 0 To UBound(HeaderNames)
            If HeaderValue = CStr(HeaderNames(j)) Then
                FindCorrespondentColumn = i
                Exit Function
            End If
        Next j
    Next i
End Function

Private Function CompareCorrespondents(DoverCorrespondent As String, OperCorrespondent As String) As Boolean
    Dim CleanDover As String, CleanOper As String
    
    CleanDover = CleanCorrespondentName(DoverCorrespondent)
    CleanOper = CleanCorrespondentName(OperCorrespondent)
    
    If UCase(Trim(CleanDover)) = UCase(Trim(CleanOper)) Then
        CompareCorrespondents = True
        Exit Function
    End If
    
    If Len(CleanDover) > 0 And Len(CleanOper) > 0 Then
        If InStr(1, UCase(CleanOper), UCase(CleanDover), vbTextCompare) > 0 Or _
           InStr(1, UCase(CleanDover), UCase(CleanOper), vbTextCompare) > 0 Then
            CompareCorrespondents = True
            Exit Function
        End If
    End If
    
    CompareCorrespondents = False
End Function

Private Function FindBySubstringEnhanced(DoverData As ParsedNaryad, OperArray As Variant, DoverCorrespondent As String) As MatchResult
    Dim result As MatchResult
    Dim i As Long
    Dim OperComment As String
    Dim OperCorrespondent As String
    Dim HasNumber As Boolean, HasDate As Boolean, HasCorrespondent As Boolean
    Dim BestPartialMatch As MatchResult
    Dim FoundPartialMatch As Boolean
    
    result.FoundMatch = False
    result.MatchScore = 0
    BestPartialMatch.FoundMatch = False
    BestPartialMatch.MatchScore = 0
    FoundPartialMatch = False
    
    For i = 1 To UBound(OperArray, 1)
        OperComment = Trim(CStr(OperArray(i, 9)))
        OperCorrespondent = Trim(CStr(OperArray(i, 6)))
        
        If OperComment <> "" Then
            HasNumber = (InStr(1, OperComment, DoverData.DocumentNumber, vbTextCompare) > 0)
            HasDate = (InStr(1, OperComment, DoverData.DocumentDate, vbTextCompare) > 0)
            
            If Not HasDate Then
                Dim ShortDate As String
                ShortDate = ConvertToShortDate(DoverData.DocumentDate)
                If ShortDate <> "" Then
                    HasDate = (InStr(1, OperComment, ShortDate, vbTextCompare) > 0)
                End If
            End If
            
            HasCorrespondent = False
            If DoverCorrespondent <> "" And OperCorrespondent <> "" Then
                HasCorrespondent = CompareCorrespondents(DoverCorrespondent, OperCorrespondent)
            End If
            
            If HasNumber And HasDate And HasCorrespondent Then
                result.FoundMatch = True
                result.MatchScore = 100
                result.OperationRow = i + 1
                result.OperationNumber = Trim(CStr(OperArray(i, 3)))
                result.MatchDetails = LocalizationManager.GetText("found number, date and correspondent | 1C: ") & OperComment
                
                On Error Resume Next
                result.OperationDate = CDate(OperArray(i, 2))
                On Error GoTo 0
                
                Exit For
            ElseIf HasNumber And HasDate Then
                If Not FoundPartialMatch Then
                    BestPartialMatch.FoundMatch = True
                    BestPartialMatch.MatchScore = 50
                    BestPartialMatch.OperationRow = i + 1
                    BestPartialMatch.OperationNumber = Trim(CStr(OperArray(i, 3)))
                    BestPartialMatch.MatchDetails = LocalizationManager.GetText("found number and date, but correspondent mismatch | 1C: ") & OperComment
                    
                    On Error Resume Next
                    BestPartialMatch.OperationDate = CDate(OperArray(i, 2))
                    On Error GoTo 0
                    
                    FoundPartialMatch = True
                End If
            End If
        End If
    Next i
    
    If Not result.FoundMatch And FoundPartialMatch Then
        result = BestPartialMatch
    End If
    
    If Not result.FoundMatch Then
        result.MatchDetails = LocalizationManager.GetText("number, date or correspondent not found")
    End If
    
    FindBySubstringEnhanced = result
End Function

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

Private Function NormalizeCommentForSearch(Comment As String) As String
    Dim result As String
    
    result = Comment
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    result = UCase(result)
    NormalizeCommentForSearch = result
End Function

Private Function CheckDateInComment(OperComment As String, TargetDate As String) As Boolean
    Dim HasDate As Boolean
    Dim DateVariants() As String
    Dim i As Long
    
    HasDate = False
    
    On Error GoTo ErrorHandler
    
    DateVariants = CreateDateVariants(TargetDate)
    
    If UBound(DateVariants) < 0 Then
        HasDate = (InStr(1, OperComment, TargetDate, vbTextCompare) > 0)
        If Not HasDate Then
            Dim ShortDate As String
            ShortDate = ConvertToShortDate(TargetDate)
            If ShortDate <> "" Then
                HasDate = (InStr(1, OperComment, ShortDate, vbTextCompare) > 0)
            End If
        End If
        CheckDateInComment = HasDate
        Exit Function
    End If
    
    For i = 0 To UBound(DateVariants)
        If DateVariants(i) <> "" Then
            If InStr(1, OperComment, DateVariants(i), vbTextCompare) > 0 Then
                HasDate = True
                Exit For
            End If
        End If
    Next i
    
    CheckDateInComment = HasDate
    Exit Function
    
ErrorHandler:
    HasDate = (InStr(1, OperComment, TargetDate, vbTextCompare) > 0)
    CheckDateInComment = HasDate
End Function

Private Function CreateDateVariants(FullDate As String) As String()
    Dim Variants() As String
    Dim DateParts() As String
    Dim Day As String, Month As String, Year As String
    Dim VariantCount As Long
    
    VariantCount = 0
    ReDim Variants(0 To 7)
    
    On Error GoTo ErrorHandler
    
    DateParts = Split(FullDate, ".")
    
    If UBound(DateParts) >= 2 Then
        Day = DateParts(0)
        Month = DateParts(1)
        Year = DateParts(2)
        
        Variants(VariantCount) = FullDate
        VariantCount = VariantCount + 1
        
        If Left(Day, 1) = "0" Then
            Variants(VariantCount) = CLng(Day) & "." & Month & "." & Year
            VariantCount = VariantCount + 1
        End If
        
        If Left(Month, 1) = "0" Then
            Variants(VariantCount) = Day & "." & CLng(Month) & "." & Year
            VariantCount = VariantCount + 1
        End If
        
        If Left(Day, 1) = "0" And Left(Month, 1) = "0" Then
            Variants(VariantCount) = CLng(Day) & "." & CLng(Month) & "." & Year
            VariantCount = VariantCount + 1
        End If
        
        If Len(Year) = 4 Then
            Variants(VariantCount) = Day & "." & Month & "." & Right(Year, 2)
            VariantCount = VariantCount + 1
        End If
        
        If Left(Day, 1) = "0" And Left(Month, 1) = "0" And Len(Year) = 4 Then
            Variants(VariantCount) = CLng(Day) & "." & CLng(Month) & "." & Right(Year, 2)
            VariantCount = VariantCount + 1
        End If
        
        Variants(VariantCount) = FullDate & " y."
        VariantCount = VariantCount + 1
        
        If Left(Day, 1) = "0" And Left(Month, 1) = "0" Then
            Variants(VariantCount) = CLng(Day) & "." & CLng(Month) & "." & Year & " y."
            VariantCount = VariantCount + 1
        End If
    End If
    
    ReDim Preserve Variants(0 To IIf(VariantCount > 0, VariantCount - 1, 0))
    CreateDateVariants = Variants
    Exit Function
    
ErrorHandler:
    ReDim Variants(0 To 0)
    Variants(0) = FullDate
    CreateDateVariants = Variants
End Function

Private Function ProcessThirdPass(WsDover As Worksheet, OperData As Variant, NotFoundRows() As Long, NotFoundData() As ParsedNaryad, NotFoundCount As Long, ByRef Pass3ByType() As Long) As Long
    Dim i As Long
    Dim ThirdPassFound As Long
    Dim MatchResult As MatchResult
    Dim UniversalDoc As ParsedNaryad
    
    ThirdPassFound = 0
    
    For i = 0 To NotFoundCount - 1
        UniversalDoc = ParseUniversalDocument(NotFoundData(i).OriginalText)
        
        If UniversalDoc.IsValid Then
            MatchResult = FindByUniversalDocument(UniversalDoc, OperData)
            MatchResult.PassNumber = 3
            
            If MatchResult.FoundMatch Then
                ThirdPassFound = ThirdPassFound + 1
                MatchResult.MatchDetails = LocalizationManager.GetText("3rd pass: ") & UniversalDoc.DocumentType & " No " & UniversalDoc.DocumentNumber
                
                Select Case UCase(UniversalDoc.DocumentType)
                    Case "ALLOTMENT": Pass3ByType(0) = Pass3ByType(0) + 1
                    Case "TELEGRAM": Pass3ByType(1) = Pass3ByType(1) + 1
                    Case "CERTIFICATE": Pass3ByType(2) = Pass3ByType(2) + 1
                    Case "REQUEST": Pass3ByType(3) = Pass3ByType(3) + 1
                End Select
                
                Call WriteThreePassResult(WsDover, NotFoundRows(i), MatchResult)
            End If
        End If
    Next i
    
    ProcessThirdPass = ThirdPassFound
End Function

Private Function ParseUniversalDocument(Text As String) As ParsedNaryad
    Dim result As ParsedNaryad
    Dim DocumentTypes(3) As String
    Dim i As Long
    Dim FoundType As String
    
    result.OriginalText = Text
    result.IsValid = False
    
    DocumentTypes(0) = "ALLOTMENT"
    DocumentTypes(1) = "TELEGRAM"
    DocumentTypes(2) = "CERTIFICATE"
    DocumentTypes(3) = "REQUEST"
    
    For i = 0 To UBound(DocumentTypes)
        If InStr(1, UCase(Text), DocumentTypes(i)) > 0 Then
            FoundType = DocumentTypes(i)
            Exit For
        End If
    Next i
    
    If FoundType = "" Then
        result.MatchDetails = LocalizationManager.GetText("Document type not found")
        ParseUniversalDocument = result
        Exit Function
    End If
    
    result.DocumentType = FoundType
    result.DocumentNumber = ExtractSubstringNumber(Text, FoundType)
    result.DocumentDate = ExtractSubstringDate(Text)
    
    If result.DocumentNumber <> "" And result.DocumentDate <> "" Then
        result.IsValid = True
        result.MatchDetails = LocalizationManager.GetText("3rd pass: success")
    Else
        result.MatchDetails = LocalizationManager.GetText("3rd pass: no number or date")
    End If
    
    ParseUniversalDocument = result
End Function

Private Function FindByUniversalDocument(UniversalDoc As ParsedNaryad, OperData As Variant) As MatchResult
    Dim result As MatchResult
    Dim i As Long
    Dim OperComment As String
    
    result.FoundMatch = False
    result.MatchScore = 0
    
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
                result.FoundMatch = True
                result.MatchScore = 100
                result.OperationRow = i + 1
                result.OperationNumber = Trim(CStr(OperData(i, 3)))
                result.MatchDetails = LocalizationManager.GetText("found ") & UniversalDoc.DocumentType
                
                On Error Resume Next
                result.OperationDate = CDate(OperData(i, 2))
                On Error GoTo 0
                Exit For
            End If
        End If
    Next i
    
    If Not result.FoundMatch Then
        result.MatchDetails = LocalizationManager.GetText("number or date not found")
    End If
    
    FindByUniversalDocument = result
End Function

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

Private Sub AddThreePassResultColumns(Ws As Worksheet)
    Dim LastCol As Long
    Dim i As Long
    Dim AlreadyExists As Boolean
    
    For i = 1 To 25
        If Trim(CStr(Ws.Cells(1, i).value)) = LocalizationManager.GetText("Three-pass search") Then
            AlreadyExists = True
            Exit For
        End If
    Next i
    
    If Not AlreadyExists Then
        LastCol = Ws.Cells(1, Ws.Columns.Count).End(xlToLeft).Column
        
        With Ws
            .Cells(1, LastCol + 1).value = LocalizationManager.GetText("Three-pass search")
            .Cells(1, LastCol + 2).value = LocalizationManager.GetText("Operation Number")
            .Cells(1, LastCol + 3).value = LocalizationManager.GetText("Operation Date")
            .Cells(1, LastCol + 4).value = LocalizationManager.GetText("Details")
            .Cells(1, LastCol + 5).value = LocalizationManager.GetText("1C Comment")
            
            .Range(.Cells(1, LastCol + 1), .Cells(1, LastCol + 5)).Font.Bold = True
            .Range(.Cells(1, LastCol + 1), .Cells(1, LastCol + 5)).Interior.Color = RGB(220, 220, 255)
            .Range(.Cells(1, LastCol + 1), .Cells(1, LastCol + 5)).Borders.LineStyle = xlContinuous
            .Range(.Cells(1, LastCol + 1), .Cells(1, LastCol + 5)).HorizontalAlignment = xlGeneral
            .Columns(LastCol + 1).AutoFit
            .Columns(LastCol + 2).AutoFit
            .Columns(LastCol + 3).AutoFit
            .Columns(LastCol + 4).AutoFit
            .Columns(LastCol + 5).AutoFit
        End With
    End If
End Sub

Private Sub WriteThreePassResult(Ws As Worksheet, Row As Long, MatchResult As MatchResult)
    Dim ResultCol As Long
    Static CachedCol As Long
    
    If CachedCol = 0 Then
        For ResultCol = 10 To 30
            If Trim(CStr(Ws.Cells(1, ResultCol).value)) = LocalizationManager.GetText("Three-pass search") Then
                CachedCol = ResultCol
                Exit For
            End If
        Next ResultCol
    End If
    
    If CachedCol > 0 Then
        If MatchResult.FoundMatch Then
            Dim ResultText As String
            If MatchResult.MatchScore = 100 Then
                ResultText = LocalizationManager.GetText("FOUND")
            Else
                ResultText = LocalizationManager.GetText("REQUIRES VERIFICATION")
            End If
            
            Ws.Cells(Row, CachedCol).value = ResultText
            Ws.Cells(Row, CachedCol + 1).value = MatchResult.OperationNumber
            Ws.Cells(Row, CachedCol + 2).value = MatchResult.OperationDate
            
            Dim Comment1C As String
            Dim DetailsWithoutComment As String
            Dim CommentStart As Long
            CommentStart = InStr(MatchResult.MatchDetails, "| 1C: ")

            If CommentStart > 0 Then
                Comment1C = Mid(MatchResult.MatchDetails, CommentStart + 6)
                Comment1C = Trim(Comment1C)
                DetailsWithoutComment = Trim(Left(MatchResult.MatchDetails, CommentStart - 1))
            Else
                DetailsWithoutComment = MatchResult.MatchDetails
            End If
            
            Ws.Cells(Row, CachedCol + 3).value = DetailsWithoutComment
            
            If Comment1C <> "" Then
                Call WriteFullComment(Ws, Row, CachedCol + 4, Comment1C)
            End If
            
            Select Case MatchResult.PassNumber
                Case 1:
                    If MatchResult.MatchScore = 100 Then
                        Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 4)).Interior.Color = RGB(0, 200, 0)
                    Else
                        Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 4)).Interior.Color = RGB(255, 255, 150)
                    End If
                Case 2:
                    If MatchResult.MatchScore = 100 Then
                        Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 4)).Interior.Color = RGB(150, 255, 150)
                    Else
                        Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 4)).Interior.Color = RGB(255, 255, 150)
                    End If
                Case 3:
                    Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 4)).Interior.Color = RGB(150, 200, 255)
            End Select
            
            Ws.Cells(Row, CachedCol).HorizontalAlignment = xlGeneral
            Ws.Cells(Row, CachedCol + 1).HorizontalAlignment = xlGeneral
            Ws.Cells(Row, CachedCol + 2).HorizontalAlignment = xlGeneral
            Ws.Cells(Row, CachedCol + 3).HorizontalAlignment = xlGeneral
        Else
            Dim DetailsWithoutCommentNotFound As String
            Dim CommentStartNotFound As Long
            CommentStartNotFound = InStr(MatchResult.MatchDetails, "| 1C: ")
            
            If CommentStartNotFound > 0 Then
                DetailsWithoutCommentNotFound = Trim(Left(MatchResult.MatchDetails, CommentStartNotFound - 1))
            Else
                DetailsWithoutCommentNotFound = MatchResult.MatchDetails
            End If
            
            Ws.Cells(Row, CachedCol).value = LocalizationManager.GetText("NOT FOUND")
            Ws.Cells(Row, CachedCol + 3).value = DetailsWithoutCommentNotFound
            Ws.Cells(Row, CachedCol).HorizontalAlignment = xlGeneral
            Ws.Cells(Row, CachedCol + 3).HorizontalAlignment = xlGeneral
            
            If InStr(MatchResult.MatchDetails, "pass") > 0 Then
                Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 4)).Interior.Color = RGB(255, 255, 150)
            Else
                Ws.Range(Ws.Cells(Row, CachedCol), Ws.Cells(Row, CachedCol + 4)).Interior.Color = RGB(255, 200, 200)
            End If
        End If
    End If
End Sub

Private Function CreateEmptyMatchResult(Message As String) As MatchResult
    Dim result As MatchResult
    result.FoundMatch = False
    result.MatchDetails = Message
    result.MatchScore = 0
    result.PassNumber = 0
    CreateEmptyMatchResult = result
End Function

Private Sub ShowThreePassStatistics(Processed As Long, Pass1 As Long, Pass2 As Long, Pass3 As Long, Pass3ByType() As Long, SkippedEmpty As Long, SkippedFound As Long)
    Dim TotalFound As Long
    Dim Message As String
    
    TotalFound = Pass1 + Pass2 + Pass3
    
    Message = LocalizationManager.GetText("=== THREE-PASS SYSTEM RESULTS ===") & vbCrLf & vbCrLf & _
              LocalizationManager.GetText("Total processed: ") & Processed & vbCrLf & vbCrLf & _
              LocalizationManager.GetText("1st pass (exact orders): ") & Pass1 & " (" & Format(Pass1 / Processed * 100, "0.0") & "%)" & vbCrLf & _
              LocalizationManager.GetText("2nd pass (regex orders): ") & Pass2 & " (" & Format(Pass2 / Processed * 100, "0.0") & "%)" & vbCrLf & _
              LocalizationManager.GetText("3rd pass (other documents): ") & Pass3 & " (" & Format(Pass3 / Processed * 100, "0.0") & "%)" & vbCrLf & _
              LocalizationManager.GetText("   - Allotments: ") & Pass3ByType(0) & vbCrLf & _
              LocalizationManager.GetText("   - Telegrams: ") & Pass3ByType(1) & vbCrLf & _
              LocalizationManager.GetText("   - Certificates: ") & Pass3ByType(2) & vbCrLf & _
              LocalizationManager.GetText("   - Requests: ") & Pass3ByType(3) & vbCrLf & vbCrLf & _
              LocalizationManager.GetText("Overall result: ") & TotalFound & " (" & Format(TotalFound / Processed * 100, "0.0") & "%)" & vbCrLf & _
              LocalizationManager.GetText("Not found: ") & (Processed - TotalFound) & vbCrLf & vbCrLf & _
              LocalizationManager.GetText("Skipped operations with empty correspondent: ") & SkippedEmpty & vbCrLf & vbCrLf & _
              LocalizationManager.GetText("COLOR CODING:") & vbCrLf & _
              LocalizationManager.GetText("Dark Green = 1st pass (exact orders)") & vbCrLf & _
              LocalizationManager.GetText("Light Green = 2nd pass (regex orders)") & vbCrLf & _
              LocalizationManager.GetText("Blue = 3rd pass (other documents)") & vbCrLf & vbCrLf & _
              LocalizationManager.GetText("Results in 'Three-pass search' columns")
    
    MsgBox Message, vbInformation, LocalizationManager.GetText("Three-pass system completed")
End Sub

Private Function ContainsMilitaryCode(Number As String) As Boolean
    Dim MilitaryCodes As Variant
    Dim i As Long
    Dim UpperNumber As String
    
    UpperNumber = UCase(Number)
    MilitaryCodes = Array("YuVO", "AT", "TsVO", "VVO", "SVO", "BT", "PVO", "VDV", "VMF", "ZCh", "I-", "A-", "B-", "V-", "S-")
    
    For i = 0 To UBound(MilitaryCodes)
        If InStr(UpperNumber, CStr(MilitaryCodes(i))) > 0 Then
            ContainsMilitaryCode = True
            Exit Function
        End If
    Next i
    
    ContainsMilitaryCode = False
End Function

Private Function ContainsLetters(Text As String) As Boolean
    Dim i As Long
    Dim Char As String
    
    For i = 1 To Len(Text)
        Char = Mid(Text, i, 1)
        If (Char >= "A" And Char <= "Z") Or (Char >= "a" And Char <= "z") Then
            ContainsLetters = True
            Exit Function
        End If
    Next i
    
    ContainsLetters = False
End Function

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

Private Function CreateSlashPattern(Number As String, DateStr As String) As String
    Dim FlexNumber As String
    FlexNumber = Replace(Number, "/", "\s*/\s*")
    CreateSlashPattern = "(?i)order\s*No.?\s*" & FlexNumber & "\s*from\s*" & EscapeMinimal(DateStr) & "(?:\s*y\.?)?"
End Function

Private Function CreateDashPattern(Number As String, DateStr As String) As String
    Dim FlexNumber As String
    FlexNumber = Replace(Number, "-", "\s*-\s*")
    CreateDashPattern = "(?i)order\s*No.?\s*" & FlexNumber & "\s*from\s*" & EscapeMinimal(DateStr) & "(?:\s*y\.?)?"
End Function

Private Function CreateMilitaryPattern(Number As String, DateStr As String) As String
    Dim FlexNumber As String
    FlexNumber = Number
    
    FlexNumber = Replace(FlexNumber, "YuVO", "(?:YuVO|yuvo)")
    FlexNumber = Replace(FlexNumber, "BT", "(?:BT|bt)")
    FlexNumber = Replace(FlexNumber, "ZCh", "(?:ZCh|zch)")
    
    FlexNumber = Replace(FlexNumber, "/", "\s*/\s*")
    FlexNumber = Replace(FlexNumber, "-", "\s*-\s*")
    
    CreateMilitaryPattern = "(?i)order\s*No.?\s*" & FlexNumber & "\s*from\s*" & EscapeMinimal(DateStr) & "(?:\s*y\.?)?"
End Function

Private Function CreateNumericPattern(Number As String, DateStr As String) As String
    CreateNumericPattern = "(?i)order\s*No.?\s*" & Number & "\s*from\s*" & EscapeMinimal(DateStr) & "(?:\s*y\.?)?"
End Function

Private Function AnalyzeComponents(CommentText As String, TargetNumber As String, TargetDate As String) As ComponentMatchResult
    Dim result As ComponentMatchResult
    
    result.IsMatch = False
    result.Confidence = 0
    
    Dim NumberCandidates() As NumberCandidate
    Dim NumberCount As Long
    Call ExtractNumberCandidates(CommentText, NumberCandidates, NumberCount)
    
    Dim DateCandidates() As DateCandidate
    Dim DateCount As Long
    Call ExtractDateCandidates(CommentText, DateCandidates, DateCount)
    
    If NumberCount > 0 And DateCount > 0 Then
        result = FindBestComponentMatch(CommentText, NumberCandidates, NumberCount, DateCandidates, DateCount, TargetNumber, TargetDate)
    End If
    
    AnalyzeComponents = result
End Function

Private Sub ExtractNumberCandidates(Text As String, ByRef NumberCandidates() As NumberCandidate, ByRef Count As Long)
    Dim i As Long, j As Long
    Dim Char As String
    Dim CurrentNumber As String
    Dim InNumber As Boolean
    Dim StartPos As Long
    
    Count = 0
    ReDim NumberCandidates(50)
    
    For i = 1 To Len(Text) - 1
        Char = Mid(Text, i, 1)
        
        If Not InNumber Then
            If IsNumberStart(Text, i) Then
                InNumber = True
                StartPos = i
                CurrentNumber = ""
            End If
        End If
        
        If InNumber Then
            If IsNumberChar(Char) Then
                CurrentNumber = CurrentNumber & Char
            ElseIf IsNumberSeparator(Char) Then
                CurrentNumber = CurrentNumber & Char
            Else
                If Len(CurrentNumber) > 0 Then
                    Call AddNumberCandidate(NumberCandidates, Count, CurrentNumber, StartPos, Len(CurrentNumber))
                End If
                InNumber = False
                CurrentNumber = ""
            End If
        End If
    Next i
    
    If InNumber And Len(CurrentNumber) > 0 Then
        Call AddNumberCandidate(NumberCandidates, Count, CurrentNumber, StartPos, Len(CurrentNumber))
    End If
End Sub

Private Sub ExtractDateCandidates(Text As String, ByRef DateCandidates() As DateCandidate, ByRef Count As Long)
    Dim RegEx As Object
    Dim Matches As Object
    Dim Match As Object
    Dim i As Long
    
    Set RegEx = CreateObject("VBScript.RegExp")
    
    Count = 0
    ReDim DateCandidates(20)
    
    With RegEx
        .IgnoreCase = True
        .Global = True
        .Pattern = "\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2})(?:\s*y\.?)?"
    End With
    
    On Error Resume Next
    Set Matches = RegEx.Execute(Text)
    
    If Matches.Count > 0 Then
        For i = 0 To Matches.Count - 1
            Set Match = Matches(i)
            Call AddDateCandidate(DateCandidates, Count, Match.value, Match.FirstIndex + 1, Match.Length)
        Next i
    End If
    On Error GoTo 0
End Sub

Private Function FindBestComponentMatch(Text As String, NumberCandidates() As NumberCandidate, NumberCount As Long, DateCandidates() As DateCandidate, DateCount As Long, TargetNumber As String, TargetDate As String) As ComponentMatchResult
    Dim result As ComponentMatchResult
    Dim BestScore As Double
    Dim i As Long, j As Long
    Dim TempResult As ComponentMatchResult
    
    result.IsMatch = False
    BestScore = 0
    
    For i = 0 To NumberCount - 1
        For j = 0 To DateCount - 1
            TempResult = EvaluateComponentPair(Text, NumberCandidates(i), DateCandidates(j), TargetNumber, TargetDate)
            
            If TempResult.Confidence > BestScore Then
                BestScore = TempResult.Confidence
                result = TempResult
            End If
        Next j
    Next i
    
    If BestScore >= 40 Then
        result.IsMatch = True
    End If
    
    FindBestComponentMatch = result
End Function

Private Function EvaluateComponentPair(Text As String, NumberCand As NumberCandidate, DateCand As DateCandidate, TargetNumber As String, TargetDate As String) As ComponentMatchResult
    Dim result As ComponentMatchResult
    Dim Score As Double
    
    Score = 0
    result.FoundNumber = NumberCand.Text
    result.FoundDate = DateCand.Text
    result.NumberPosition = NumberCand.Position
    result.DatePosition = DateCand.Position
    
    Dim NumberScore As Double
    NumberScore = CompareNumbers(TargetNumber, NumberCand.Text)
    Score = Score + NumberScore
    
    Dim DateScore As Double
    DateScore = CompareDates(TargetDate, DateCand.Text)
    Score = Score + DateScore
    
    Dim PositionScore As Double
    PositionScore = EvaluatePositions(Text, NumberCand.Position, DateCand.Position)
    Score = Score + PositionScore
    
    result.Confidence = Score
    result.MatchDetails = "num:" & NumberCand.Text & "(" & Format(NumberScore, "0") & "%) date:" & DateCand.Text & "(" & Format(DateScore, "0") & "%) pos:(" & Format(PositionScore, "0") & "%)"
    
    EvaluateComponentPair = result
End Function

Private Function CompareNumbers(Target As String, Found As String) As Double
    Dim Score As Double
    
    If UCase(Trim(Target)) = UCase(Trim(Found)) Then
        Score = 50
    ElseIf InStr(1, UCase(Found), UCase(Target), vbTextCompare) > 0 Then
        Score = 35
    ElseIf InStr(1, UCase(Target), UCase(Found), vbTextCompare) > 0 Then
        Score = 25
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
    
    If TargetDate = FoundDate And TargetDate > 0 And FoundDate > 0 Then
        Score = 40
    ElseIf Abs(TargetDate - FoundDate) = 1 And TargetDate > 0 And FoundDate > 0 Then
        Score = 25
    Else
        Score = 0
    End If
    
    CompareDates = Score
End Function

Private Function EvaluatePositions(Text As String, NumberPos As Long, DatePos As Long) As Double
    Dim Score As Double
    
    If NumberPos < DatePos Then
        Score = Score + 5
        
        Dim Distance As Long
        Distance = DatePos - NumberPos
        
        If Distance <= 20 Then
            Score = Score + 5
        ElseIf Distance <= 50 Then
            Score = Score + 3
        Else
            Score = Score + 1
        End If
    End If
    
    EvaluatePositions = Score
End Function

Private Function IsNumberStart(Text As String, Position As Long) As Boolean
    Dim Context As String
    Dim StartCheck As Long
    
    StartCheck = Position - 10
    If StartCheck < 1 Then StartCheck = 1
    
    Context = UCase(Mid(Text, StartCheck, 15))
    
    IsNumberStart = (InStr(Context, "ORDER") > 0) Or _
                    (InStr(Context, "NO.") > 0) Or _
                    (InStr(Context, "#") > 0)
End Function

Private Function IsNumberChar(Char As String) As Boolean
    IsNumberChar = (Char >= "0" And Char <= "9") Or _
                  (Char >= "A" And Char <= "Z") Or _
                  (Char >= "a" And Char <= "z")
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
    
    CleanDate = Replace(CleanDate, "y.", "")
    CleanDate = Replace(CleanDate, "y", "")
    CleanDate = Trim(CleanDate)
    
    NormalizeDateString = CleanDate
End Function

Private Function CalculatePartialNumberMatch(Target As String, Found As String) As Double
    Dim CommonChars As Long
    Dim i As Long, j As Long
    Dim MaxLen As Long
    
    MaxLen = IIf(Len(Target) > Len(Found), Len(Target), Len(Found))
    
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

Private Function CheckCorrespondentMatch(OperCorrespondent As Variant, DoverText As String, OperationRow As Long) As Double
    Dim OperCorr As String, DoverCorr As String
    Dim Score As Double
    
    Score = 0
    OperCorr = Trim(CStr(OperCorrespondent))
    
    If OperCorr = "" Then
        CheckCorrespondentMatch = 0
        Exit Function
    End If
    
    DoverCorr = ExtractCorrespondentFromComment(DoverText)
    
    If DoverCorr <> "" Then
        Dim CleanOperCorr As String, CleanDoverCorr As String
        CleanOperCorr = CleanCorrespondentName(OperCorr)
        CleanDoverCorr = CleanCorrespondentName(DoverCorr)
        
        If UCase(CleanOperCorr) = UCase(CleanDoverCorr) Then
            Score = 50
        ElseIf InStr(1, UCase(CleanOperCorr), UCase(CleanDoverCorr), vbTextCompare) > 0 Or _
               InStr(1, UCase(CleanDoverCorr), UCase(CleanOperCorr), vbTextCompare) > 0 Then
            Score = 25
        End If
    End If
    
    CheckCorrespondentMatch = Score
End Function

Private Function CleanCorrespondentName(FullName As String) As String
    Dim result As String
    Dim BracketPos As Long
    
    result = Trim(FullName)
    
    BracketPos = InStr(result, "(")
    If BracketPos > 0 Then
        result = Trim(Left(result, BracketPos - 1))
    End If
    
    CleanCorrespondentName = result
End Function

Private Function ExtractCorrespondentFromComment(Comment As String) As String
    ExtractCorrespondentFromComment = ""
End Function

Private Sub WriteFullComment(Ws As Worksheet, Row As Long, Col As Long, FullComment As String)
    With Ws.Cells(Row, Col)
        .value = FullComment
        .WrapText = True
    End With
    
    Ws.Rows(Row).AutoFit
    
    If Ws.Rows(Row).RowHeight > 100 Then
        Ws.Rows(Row).RowHeight = 100
    End If
End Sub

Private Function IsAlreadyProcessed(Ws As Worksheet, Row As Long, ResultColumn As Long) As Boolean
    Dim ResultValue As String
    
    IsAlreadyProcessed = False
    
    If ResultColumn > 0 Then
        ResultValue = Trim(UCase(CStr(Ws.Cells(Row, ResultColumn).value)))
        
        If ResultValue = LocalizationManager.GetText("FOUND") Then
            IsAlreadyProcessed = True
        End If
    End If
End Function

Private Function FindResultColumn(Ws As Worksheet) As Long
    Dim i As Long
    
    FindResultColumn = 0
    
    For i = 1 To 30
        If Trim(UCase(CStr(Ws.Cells(1, i).value))) = UCase(LocalizationManager.GetText("Three-pass search")) Then
            FindResultColumn = i
            Exit For
        End If
    Next i
End Function

Private Function ProcessSecondPassWithSupplierService(WsDover As Worksheet, OperData As Variant, NotFoundRows() As Long, NotFoundData() As ParsedNaryad, NotFoundCount As Long, ByRef Pass2NotFound() As Long, ByRef Pass2NotFoundData() As ParsedNaryad, ByRef Pass2Count As Long, CorrespondentColumn As Long) As Long
    Dim i As Long
    Dim SecondPassFound As Long
    Dim MatchResult As MatchResult
    Dim DoverSupplier As String
    Dim CurrentRow As Long
    
    SecondPassFound = 0
    
    For i = 0 To NotFoundCount - 1
        CurrentRow = NotFoundRows(i)
        
        DoverSupplier = ""
        If CorrespondentColumn > 0 Then
            DoverSupplier = Trim(CStr(WsDover.Cells(CurrentRow, CorrespondentColumn).value))
        End If
        
        MatchResult = FindBySupplierAndService(NotFoundData(i), OperData, DoverSupplier)
        MatchResult.PassNumber = 2
        
        If MatchResult.FoundMatch Then
            SecondPassFound = SecondPassFound + 1
            MatchResult.MatchDetails = LocalizationManager.GetText("2nd pass (") & Format(MatchResult.MatchScore, "0") & "%): " & MatchResult.MatchDetails
            
            Call WriteThreePassResult(WsDover, NotFoundRows(i), MatchResult)
        Else
            Pass2Count = Pass2Count + 1
            ReDim Preserve Pass2NotFound(Pass2Count - 1)
            ReDim Preserve Pass2NotFoundData(Pass2Count - 1)
            Pass2NotFound(Pass2Count - 1) = NotFoundRows(i)
            Pass2NotFoundData(Pass2Count - 1) = NotFoundData(i)
        End If
    Next i
    
    ProcessSecondPassWithSupplierService = SecondPassFound
End Function

Private Function FindBySupplierAndService(DoverData As ParsedNaryad, OperData As Variant, DoverSupplier As String) As MatchResult
    Dim result As MatchResult
    Dim i As Long
    Dim OperComment As String, OperSupplier As String
    Dim DoverService As String
    
    result.FoundMatch = False
    result.MatchScore = 0
    
    DoverService = ExtractServiceFromComment(DoverData.OriginalText)
    
    For i = 1 To UBound(OperData, 1)
        OperComment = Trim(CStr(OperData(i, 9)))
        OperSupplier = Trim(CStr(OperData(i, 6)))
        
        If OperComment <> "" And OperSupplier <> "" Then
            Dim MatchScore As Double
            Dim MatchDetails As String
            MatchScore = AnalyzeSupplierServiceMatch(OperComment, OperSupplier, DoverData.DocumentNumber, DoverSupplier, DoverService, MatchDetails)
            
            If MatchScore >= 60 Then
                result.FoundMatch = True
                result.MatchScore = MatchScore
                result.OperationRow = i + 1
                result.OperationNumber = Trim(CStr(OperData(i, 3)))
                result.MatchDetails = MatchDetails & " | 1C: " & OperComment
                
                On Error Resume Next
                result.OperationDate = CDate(OperData(i, 2))
                On Error GoTo 0
                Exit For
            End If
        End If
    Next i
    
    If Not result.FoundMatch Then
        result.MatchDetails = LocalizationManager.GetText("supplier+service mismatch")
    End If
    
    FindBySupplierAndService = result
End Function

Private Function AnalyzeSupplierServiceMatch(OperComment As String, OperSupplier As String, DoverNumber As String, DoverSupplier As String, DoverService As String, ByRef MatchDetails As String) As Double
    Dim TotalScore As Double
    Dim NumberScore As Double, SupplierScore As Double, ServiceScore As Double
    
    TotalScore = 0
    MatchDetails = ""
    
    NumberScore = CheckNumberInComment(OperComment, DoverNumber)
    TotalScore = TotalScore + NumberScore
    
    SupplierScore = CheckSupplierMatch(OperSupplier, DoverSupplier)
    TotalScore = TotalScore + SupplierScore
    
    If DoverService <> "" Then
        ServiceScore = CheckServiceInComment(OperComment, DoverService)
        TotalScore = TotalScore + ServiceScore
    End If
    
    MatchDetails = "num:" & Format(NumberScore, "0") & "% supplier:" & Format(SupplierScore, "0") & "%"
    If DoverService <> "" Then
        MatchDetails = MatchDetails & " service:" & Format(ServiceScore, "0") & "%"
    End If
    
    AnalyzeSupplierServiceMatch = TotalScore
End Function

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
        
        KnownServices = Array("MS", "VS", "PS", "TS", "IS", "SS", "AS", "GSM", "RTS", "VVS", "VMF", "RVSN")
        
        For i = 0 To UBound(KnownServices)
            If FirstWord = CStr(KnownServices(i)) Then
                ExtractServiceFromComment = FirstWord
                Exit Function
            End If
        Next i
    End If
End Function

Private Function CheckNumberInComment(OperComment As String, DoverNumber As String) As Double
    Dim Score As Double
    
    Score = 0
    
    If InStr(1, OperComment, DoverNumber, vbTextCompare) > 0 Then
        Score = 40
    ElseIf InStr(DoverNumber, "/") > 0 Or InStr(DoverNumber, "-") > 0 Then
        Score = CheckPartialNumberMatch(OperComment, DoverNumber)
    End If
    
    CheckNumberInComment = Score
End Function

Private Function CheckSupplierMatch(OperSupplier As String, DoverSupplier As String) As Double
    Dim Score As Double
    Score = 0
    
    If DoverSupplier = "" Or OperSupplier = "" Then
        CheckSupplierMatch = 0
        Exit Function
    End If
    
    If CompareCorrespondents(DoverSupplier, OperSupplier) Then
        Score = 40
    End If
    
    CheckSupplierMatch = Score
End Function

Private Function CheckServiceInComment(OperComment As String, DoverService As String) As Double
    Dim Score As Double
    Score = 0
    
    If DoverService <> "" Then
        If InStr(1, UCase(OperComment), DoverService, vbTextCompare) > 0 Then
            Score = 20
        End If
    End If
    
    CheckServiceInComment = Score
End Function

Private Function CleanSupplierName(SupplierName As String) As String
    Dim result As String
    
    result = Trim(SupplierName)
    result = Replace(result, """", "")
    result = Replace(result, "'", "")
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    CleanSupplierName = result
End Function

Private Function CheckPartialNumberMatch(OperComment As String, DoverNumber As String) As Double
    Dim Score As Double
    Dim NumberParts() As String
    Dim i As Long
    Dim FoundParts As Long
    
    Score = 0
    FoundParts = 0
    
    If InStr(DoverNumber, "/") > 0 Then
        NumberParts = Split(DoverNumber, "/")
    ElseIf InStr(DoverNumber, "-") > 0 Then
        NumberParts = Split(DoverNumber, "-")
    Else
        CheckPartialNumberMatch = 0
        Exit Function
    End If
    
    For i = 0 To UBound(NumberParts)
        If Len(Trim(NumberParts(i))) >= 2 Then
            If InStr(1, OperComment, Trim(NumberParts(i)), vbTextCompare) > 0 Then
                FoundParts = FoundParts + 1
            End If
        End If
    Next i
    
    If FoundParts > 0 Then
        Score = (FoundParts / (UBound(NumberParts) + 1)) * 25
    End If
    
    CheckPartialNumberMatch = Score
End Function

' =============================================
' TEST FUNCTIONS
' =============================================

Public Sub TestProblematicNumbers()
    Debug.Print "=== PROBLEMATIC NUMBERS TEST ==="
    Debug.Print ""
    
    Call QuickTest("901/243450", "25.10.2024", "GSM Order No. 901/243450 from 25.10.2024")
    Call QuickTest("561/22/176-a-25", "25.10.2024", "Order No. 561/22/176-a-25 from 25.10.2024")
    Call QuickTest("PS-4", "25.10.2024", "Order No. PS-4 from 25.10.2024")
    Call QuickTest("60/YuVO", "25.10.2024", "Order No. 60/YuVO from 25.10.2024")
    
    Debug.Print "=== END OF TEST ==="
End Sub

Public Sub QuickTest(TestNumber As String, TestDate As String, TestComment As String)
    Debug.Print "TEST: " & TestNumber
    
    Dim Patterns() As String
    Call CreateFlexibleTextPatterns(TestNumber, TestDate, Patterns)
    
    Dim MatchedPattern As String, MatchedText As String
    
    If FindPatternInText(TestComment, Patterns, MatchedPattern, MatchedText) Then
        Debug.Print "  [OK] FOUND!"
    Else
        Debug.Print "  [X] NOT FOUND"
    End If
End Sub

