Attribute VB_Name = "ProvodkaIntegrationModule"
'==============================================
' 1C INTEGRATION MODULE - ProvodkaIntegrationModule
' Purpose: Automatic matching of 1C postings with IncOut documents
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.0.1
' Date: 21.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' Structure for search result
Public Type MatchResult
    Found As Boolean
    ProvodkaNumber As String
    ProvodkaDate As Date
    MatchCount As Long
    StatusMessage As String
    candidatesList As String
End Type

' Mass processing of all IncOut records with 1C export file
Public Sub MassProcessWithFileSelection()
    Dim filePath As String
    Dim wb1C As Workbook
    Dim ws1C As Worksheet
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    Dim ProcessedCount As Long
    Dim FoundCount As Long
    Dim SkippedCount As Long
    Dim MultipleCount As Long
    Dim errorCount As Long
    
    On Error GoTo MassProcessError
    
    ' Select 1C export file
    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        , LocalizationManager.GetText("Select 1C export file"))
    
    If filePath = "False" Then Exit Sub
    
    ' Open 1C export file
    Application.StatusBar = LocalizationManager.GetText("Opening 1C export file...")
    Set wb1C = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws1C = wb1C.Worksheets(1) ' First sheet of the file
    
    ' Get IncOut table
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    Application.StatusBar = LocalizationManager.GetText("Starting mass processing...")
    Application.ScreenUpdating = False
    
    ' Process each IncOut record
    Dim i As Long
    Dim currentSuma As Double
    Dim currentCorrespondent As String
    Dim currentOtmetka As String
    Dim MatchResult As MatchResult
    Dim childProcessed As Long
    Dim childFound As Long
    Dim childCandidates As Long
    Dim childNotFound As Long
    Dim parentPrimaryStatus As String
    
    For i = 1 To tblData.ListRows.Count
        
        ' Get data of current IncOut record
        On Error Resume Next
        currentSuma = CDbl(tblData.DataBodyRange.Cells(i, 6).value)          ' Document amount
        currentCorrespondent = CStr(tblData.DataBodyRange.Cells(i, 9).value) ' Received from
        currentOtmetka = Trim(CStr(tblData.DataBodyRange.Cells(i, 18).value)) ' Execution mark
        On Error GoTo MassProcessError
        
        If PackageDocumentsManager.ShouldUseChildDocumentsForMatching(i) Then
            If PackageDocumentsManager.CountPendingPackageChildMatches(i) > 0 Then
                childProcessed = 0
                childFound = 0
                childCandidates = 0
                childNotFound = 0

                Call PackageDocumentsManager.ProcessPackageChildMatches(i, ws1C, childProcessed, childFound, childCandidates, childNotFound)

                If childProcessed > 0 Then
                    parentPrimaryStatus = PackageDocumentsManager.GetPackagePrimary1CStatus(i)
                    Select Case parentPrimaryStatus
                        Case "exact"
                            FoundCount = FoundCount + 1
                        Case "candidate", "manual"
                            MultipleCount = MultipleCount + 1
                    End Select
                    ProcessedCount = ProcessedCount + 1
                Else
                    SkippedCount = SkippedCount + 1
                End If
            Else
                SkippedCount = SkippedCount + 1
            End If
        ElseIf currentOtmetka = "" Then
            MatchResult = FindMatchInFile(currentSuma, currentCorrespondent, ws1C)

            If MatchResult.Found Then
                tblData.DataBodyRange.Cells(i, 18).value = MatchResult.ProvodkaNumber
                FoundCount = FoundCount + 1
            ElseIf MatchResult.MatchCount > 1 Then
                MultipleCount = MultipleCount + 1
            End If

            ProcessedCount = ProcessedCount + 1
        Else
            SkippedCount = SkippedCount + 1
        End If
        
        ' Update progress every 25 records
        If (ProcessedCount + SkippedCount) Mod 25 = 0 Then
            Application.StatusBar = LocalizationManager.GetText("Processed ") & (ProcessedCount + SkippedCount) & LocalizationManager.GetText(" of ") & tblData.ListRows.Count & LocalizationManager.GetText(" records")
        End If
        
    Next i
    
    ' Close 1C export file
    wb1C.Close False
    
    Application.ScreenUpdating = True
    
    ' Show processing results
    MsgBox LocalizationManager.GetText("MASS PROCESSING COMPLETED:") & vbCrLf & vbCrLf & _
           LocalizationManager.GetText("--- STATISTICS:") & vbCrLf & _
           LocalizationManager.GetText("Total records in table: ") & tblData.ListRows.Count & vbCrLf & _
           LocalizationManager.GetText("Processed (without mark): ") & ProcessedCount & vbCrLf & _
           LocalizationManager.GetText("Skipped (already filled): ") & SkippedCount & vbCrLf & vbCrLf & _
           LocalizationManager.GetText("--- RESULTS:") & vbCrLf & _
           LocalizationManager.GetText("Found automatically: ") & FoundCount & vbCrLf & _
           LocalizationManager.GetText("Multiple matches: ") & MultipleCount & vbCrLf & _
           LocalizationManager.GetText("Not found: ") & (ProcessedCount - FoundCount - MultipleCount) & vbCrLf & vbCrLf & _
           LocalizationManager.GetText("--- Success rate: ") & Format(IIf(ProcessedCount > 0, FoundCount / ProcessedCount, 0), "0.0%"), _
           vbInformation, LocalizationManager.GetText("1C Integration Results")
           
    Application.StatusBar = LocalizationManager.GetText("Integration completed. Found ") & FoundCount & LocalizationManager.GetText(" matches out of ") & ProcessedCount & LocalizationManager.GetText(" records.")
    
    Exit Sub
    
MassProcessError:
    Application.ScreenUpdating = True
    If Not wb1C Is Nothing Then wb1C.Close False
    
    MsgBox LocalizationManager.GetText("Mass processing error:") & vbCrLf & _
           LocalizationManager.GetText("Error: ") & Err.Number & " - " & Err.description & vbCrLf & _
           LocalizationManager.GetText("Processed records: ") & ProcessedCount, _
           vbCritical, LocalizationManager.GetText("Critical Error")
           
    Application.StatusBar = LocalizationManager.GetText("Mass processing error")
End Sub

' Search for matching posting in 1C export file
Private Function FindMatchInFile(suma As Double, Correspondent As String, ws1C As Worksheet) As MatchResult
    Dim result As MatchResult
    Dim LastRow As Long
    Dim i As Long
    
    ' Data from 1C export
    Dim currentStatus As String
    Dim currentSuma As Double
    Dim currentCorrespondent As String
    Dim CurrentNumber As String
    Dim CurrentDate As Date
    
    Dim CandidatesCount As Long
    Dim candidatesList As String
    Dim bestCandidate As String
    Dim bestCandidateDate As Date
    
    On Error GoTo FindError
    
    ' Initialize result
    result.Found = False
    result.MatchCount = 0
    result.StatusMessage = LocalizationManager.GetText("Not found")
    result.candidatesList = ""
    
    ' Determine last row with data
    LastRow = ws1C.Cells(ws1C.Rows.Count, 1).End(xlUp).Row
    
    ' Check data availability
    If LastRow < 2 Then
        result.StatusMessage = LocalizationManager.GetText("Export file is empty")
        FindMatchInFile = result
        Exit Function
    End If
    
    ' Search through all rows of 1C export (starting from 2nd row, 1st - headers)
    For i = 2 To LastRow
        
        On Error Resume Next
        ' Read data from 1C export
        currentStatus = CStr(ws1C.Cells(i, 1).value)        ' Column A - Status
        currentSuma = CDbl(ws1C.Cells(i, 5).value)          ' Column E - Amount (presumably)
        currentCorrespondent = CStr(ws1C.Cells(i, 6).value) ' Column F - Correspondent (presumably)
        CurrentNumber = CStr(ws1C.Cells(i, 3).value)        ' Column C - Number
        CurrentDate = CDate(ws1C.Cells(i, 2).value)         ' Column B - Date
        On Error GoTo FindError
        
        ' CHECK MATCH CRITERIA
        ' Exclude unposted documents (status = 1)
        ' Exact amount match (with 0.01 tolerance)
        ' Normalized correspondent match
        If (currentStatus <> "1") And _
           (Abs(currentSuma - suma) < 0.01) And _
           (CorrespondentsMatch(currentCorrespondent, Correspondent)) Then
            
            CandidatesCount = CandidatesCount + 1
            
            ' Save candidate info
            If CandidatesCount = 1 Then
                bestCandidate = CurrentNumber
                bestCandidateDate = CurrentDate
                candidatesList = CurrentNumber & " (" & Format(CurrentDate, "dd.mm.yyyy") & ")"
            Else
                candidatesList = candidatesList & "; " & CurrentNumber & " (" & Format(CurrentDate, "dd.mm.yyyy") & ")"
                
                ' Choose earlier date as best candidate
                If CurrentDate < bestCandidateDate Then
                    bestCandidate = CurrentNumber
                    bestCandidateDate = CurrentDate
                End If
            End If
            
        End If
        
    Next i
    
    ' Determine search result
    result.MatchCount = CandidatesCount
    result.candidatesList = candidatesList
    
    If CandidatesCount = 1 Then
        result.Found = True
        result.ProvodkaNumber = bestCandidate
        result.ProvodkaDate = bestCandidateDate
        result.StatusMessage = LocalizationManager.GetText("Single match found")
        
    ElseIf CandidatesCount > 1 Then
        result.Found = False ' Requires manual choice
        result.ProvodkaNumber = bestCandidate
        result.ProvodkaDate = bestCandidateDate
        result.StatusMessage = LocalizationManager.GetText("Found ") & CandidatesCount & LocalizationManager.GetText(" variants (selected by date)")
        
    Else
        result.Found = False
        result.StatusMessage = LocalizationManager.GetText("No match found")
    End If
    
    FindMatchInFile = result
    Exit Function
    
FindError:
    result.Found = False
    result.MatchCount = 0
    result.StatusMessage = LocalizationManager.GetText("Search error: ") & Err.description
    FindMatchInFile = result
End Function

Public Sub FindMatchDetailsInFile(ByVal suma As Double, ByVal Correspondent As String, ByVal ws1C As Worksheet, ByRef found As Boolean, ByRef provodkaNumber As String, ByRef provodkaDate As Variant, ByRef matchCount As Long, ByRef statusMessage As String, ByRef candidatesList As String)
    Dim result As MatchResult

    result = FindMatchInFile(suma, Correspondent, ws1C)
    found = result.Found
    provodkaNumber = result.ProvodkaNumber
    matchCount = result.MatchCount
    statusMessage = result.StatusMessage
    candidatesList = result.candidatesList

    If result.Found Or result.MatchCount > 1 Then
        provodkaDate = result.ProvodkaDate
    Else
        provodkaDate = vbNullString
    End If
End Sub

Private Function CorrespondentsMatch(ByVal leftText As String, ByVal rightText As String) As Boolean
    Dim leftMilitaryKey As String
    Dim rightMilitaryKey As String
    Dim leftNormalized As String
    Dim rightNormalized As String

    leftMilitaryKey = ExtractMilitaryUnitKey(leftText)
    rightMilitaryKey = ExtractMilitaryUnitKey(rightText)

    If Len(leftMilitaryKey) > 0 And Len(rightMilitaryKey) > 0 Then
        CorrespondentsMatch = (StrComp(leftMilitaryKey, rightMilitaryKey, vbTextCompare) = 0)
        Exit Function
    End If

    leftNormalized = NormalizeCorrespondentForMatch(leftText)
    rightNormalized = NormalizeCorrespondentForMatch(rightText)

    If Len(leftNormalized) = 0 Or Len(rightNormalized) = 0 Then Exit Function

    If StrComp(leftNormalized, rightNormalized, vbTextCompare) = 0 Then
        CorrespondentsMatch = True
    ElseIf InStr(1, leftNormalized, rightNormalized, vbTextCompare) > 0 Then
        CorrespondentsMatch = True
    ElseIf InStr(1, rightNormalized, leftNormalized, vbTextCompare) > 0 Then
        CorrespondentsMatch = True
    End If
End Function

Private Function NormalizeCorrespondentForMatch(ByVal sourceText As String) As String
    Dim normalized As String
    Dim i As Long
    Dim currentChar As String

    normalized = UCase$(Trim$(sourceText))
    normalized = Replace(normalized, "Ё", "Е")
    normalized = Replace(normalized, vbCr, " ")
    normalized = Replace(normalized, vbLf, " ")
    normalized = Replace(normalized, Chr$(34), " ")

    For i = 1 To Len(normalized)
        currentChar = Mid$(normalized, i, 1)
        If (currentChar >= "A" And currentChar <= "Z") Or _
           (currentChar >= "А" And currentChar <= "Я") Or _
           (currentChar >= "0" And currentChar <= "9") Then
            NormalizeCorrespondentForMatch = NormalizeCorrespondentForMatch & currentChar
        Else
            NormalizeCorrespondentForMatch = NormalizeCorrespondentForMatch & " "
        End If
    Next i

    Do While InStr(NormalizeCorrespondentForMatch, "  ") > 0
        NormalizeCorrespondentForMatch = Replace(NormalizeCorrespondentForMatch, "  ", " ")
    Loop

    NormalizeCorrespondentForMatch = Trim$(NormalizeCorrespondentForMatch)
End Function

Private Function ExtractMilitaryUnitKey(ByVal sourceText As String) As String
    Dim normalized As String
    Dim digitsBuffer As String
    Dim i As Long
    Dim currentChar As String
    Dim hasMilitaryMarker As Boolean

    normalized = UCase$(Trim$(sourceText))
    normalized = Replace(normalized, "Ё", "Е")

    If IsStandaloneMilitaryDigits(normalized) Then
        ExtractMilitaryUnitKey = normalized
        Exit Function
    End If

    hasMilitaryMarker = (InStr(normalized, "ВОЙСКОВ") > 0) Or _
                        (InStr(normalized, "ВОИНСК") > 0) Or _
                        (InStr(normalized, "В/Ч") > 0) Or _
                        (InStr(normalized, "ВЧ ") > 0) Or _
                        (Left$(normalized, 3) = "ВЧ ")

    If Not hasMilitaryMarker Then Exit Function

    For i = 1 To Len(normalized)
        currentChar = Mid$(normalized, i, 1)
        If currentChar >= "0" And currentChar <= "9" Then
            digitsBuffer = digitsBuffer & currentChar
        Else
            If Len(digitsBuffer) >= 4 And Len(digitsBuffer) <= 6 Then
                ExtractMilitaryUnitKey = digitsBuffer
                Exit Function
            End If
            digitsBuffer = vbNullString
        End If
    Next i

    If Len(digitsBuffer) >= 4 And Len(digitsBuffer) <= 6 Then
        ExtractMilitaryUnitKey = digitsBuffer
    End If
End Function

Private Function IsStandaloneMilitaryDigits(ByVal sourceText As String) As Boolean
    Dim i As Long
    Dim currentChar As String

    If Len(sourceText) < 4 Or Len(sourceText) > 6 Then Exit Function

    For i = 1 To Len(sourceText)
        currentChar = Mid$(sourceText, i, 1)
        If currentChar < "0" Or currentChar > "9" Then Exit Function
    Next i

    IsStandaloneMilitaryDigits = True
End Function

' Process single record (for manual search from form)
Public Sub ProcessSingleRecord(RowIndex As Long)
    Dim filePath As String
    Dim wb1C As Workbook
    Dim ws1C As Worksheet
    Dim wsData As Worksheet
    Dim tblData As ListObject
    
    Dim currentSuma As Double
    Dim currentCorrespondent As String
    Dim MatchResult As MatchResult
    Dim childProcessed As Long
    Dim childFound As Long
    Dim childCandidates As Long
    Dim childNotFound As Long
    Dim parentPrimaryStatus As String
    Dim parentPrimaryOperationNumber As String
    
    On Error GoTo SingleProcessError
    
    ' Get IncOut table
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    ' Check row number validity
    If RowIndex < 1 Or RowIndex > tblData.ListRows.Count Then
        MsgBox LocalizationManager.GetText("Invalid record number: ") & RowIndex, vbExclamation, LocalizationManager.GetText("Error")
        Exit Sub
    End If
    
    ' Get record data
    currentSuma = CDbl(tblData.DataBodyRange.Cells(RowIndex, 6).value)
    currentCorrespondent = CStr(tblData.DataBodyRange.Cells(RowIndex, 9).value)
    
    ' Select 1C export file
    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        , LocalizationManager.GetText("Select 1C export file for posting search"))
    
    If filePath = "False" Then Exit Sub
    
    ' Open 1C export file
    Application.StatusBar = LocalizationManager.GetText("Searching for posting in 1C file...")
    Set wb1C = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws1C = wb1C.Worksheets(1)
    
    If PackageDocumentsManager.ShouldUseChildDocumentsForMatching(RowIndex) Then
        If PackageDocumentsManager.CountPendingPackageChildMatches(RowIndex) = 0 Then
            wb1C.Close False
            MsgBox LocalizationManager.GetText("All child documents are already matched."), vbInformation, LocalizationManager.GetText("Package child matching completed")
            Application.StatusBar = LocalizationManager.GetText("Posting search completed")
            Exit Sub
        End If

        childProcessed = 0
        childFound = 0
        childCandidates = 0
        childNotFound = 0

        Call PackageDocumentsManager.ProcessPackageChildMatches(RowIndex, ws1C, childProcessed, childFound, childCandidates, childNotFound)
        wb1C.Close False

        parentPrimaryStatus = PackageDocumentsManager.GetPackagePrimary1CStatus(RowIndex)
        parentPrimaryOperationNumber = PackageDocumentsManager.GetPackagePrimary1COperationNumber(RowIndex)

        MsgBox LocalizationManager.GetText("Package child documents processed.") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Child documents processed: ") & childProcessed & vbCrLf & _
               LocalizationManager.GetText("Found automatically: ") & childFound & vbCrLf & _
               LocalizationManager.GetText("Multiple matches: ") & childCandidates & vbCrLf & _
               LocalizationManager.GetText("Not found: ") & childNotFound & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Package-level 1C status: ") & parentPrimaryStatus & IIf(Len(Trim$(parentPrimaryOperationNumber)) > 0, vbCrLf & LocalizationManager.GetText("Posting Number: ") & parentPrimaryOperationNumber, vbNullString), _
               vbInformation, LocalizationManager.GetText("Package child matching completed")

        If TableEventHandler.IsFormOpen("UserFormVhIsh") Then
            Call PackageDocumentsManager.RefreshPackageIndicatorsOnMainForm(UserFormVhIsh, RowIndex)
        End If
    Else
        MatchResult = FindMatchInFile(currentSuma, currentCorrespondent, ws1C)
        wb1C.Close False

        If MatchResult.Found Then
            tblData.DataBodyRange.Cells(RowIndex, 18).value = MatchResult.ProvodkaNumber

            MsgBox LocalizationManager.GetText("[OK] POSTING FOUND!") & vbCrLf & vbCrLf & _
                   LocalizationManager.GetText("Posting Number: ") & MatchResult.ProvodkaNumber & vbCrLf & _
                   LocalizationManager.GetText("Posting Date: ") & Format(MatchResult.ProvodkaDate, "dd.mm.yyyy") & vbCrLf & _
                   LocalizationManager.GetText("Amount: ") & currentSuma & vbCrLf & _
                   LocalizationManager.GetText("Correspondent: ") & currentCorrespondent & vbCrLf & vbCrLf & _
                   LocalizationManager.GetText("Posting number written to 'Execution Mark'"), _
                   vbInformation, LocalizationManager.GetText("Posting Search")

            If TableEventHandler.IsFormOpen("UserFormVhIsh") Then
                UserFormVhIsh.txtOtmetkaIspolnenie.Text = MatchResult.ProvodkaNumber
                UserFormVhIsh.txtOtmetkaIspolnenie.BackColor = RGB(200, 255, 200)
            End If
        ElseIf MatchResult.MatchCount > 1 Then
            Call ShowMultipleChoiceDialog(RowIndex, MatchResult, currentSuma, currentCorrespondent)
        Else
            MsgBox LocalizationManager.GetText("[WARN] POSTING NOT FOUND") & vbCrLf & vbCrLf & _
                   LocalizationManager.GetText("Search criteria:") & vbCrLf & _
                   LocalizationManager.GetText("Amount: ") & currentSuma & vbCrLf & _
                   LocalizationManager.GetText("Correspondent: ") & currentCorrespondent & vbCrLf & vbCrLf & _
                   LocalizationManager.GetText("Possible reasons:") & vbCrLf & _
                   LocalizationManager.GetText("- Document not yet posted in 1C") & vbCrLf & _
                   LocalizationManager.GetText("- Amount or correspondent name differs") & vbCrLf & _
                   LocalizationManager.GetText("- Document reversed in 1C"), _
                   vbExclamation, LocalizationManager.GetText("Posting Search")
        End If
    End If
    
    Application.StatusBar = LocalizationManager.GetText("Posting search completed")
    Exit Sub
    
SingleProcessError:
    If Not wb1C Is Nothing Then wb1C.Close False
    
    MsgBox LocalizationManager.GetText("Posting search error:") & vbCrLf & _
           LocalizationManager.GetText("Error: ") & Err.Number & " - " & Err.description, _
           vbCritical, LocalizationManager.GetText("Error")
           
    Application.StatusBar = LocalizationManager.GetText("Posting search error")
End Sub

' Dialog to choose from multiple variants
Private Sub ShowMultipleChoiceDialog(RowIndex As Long, MatchResult As MatchResult, suma As Double, Correspondent As String)
    Dim userChoice As String
    Dim selectedProvodka As String
    
    userChoice = InputBox( _
        LocalizationManager.GetText("[INFO] MULTIPLE POSTING VARIANTS FOUND:") & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Search criteria:") & vbCrLf & _
        LocalizationManager.GetText("Amount: ") & suma & vbCrLf & _
        LocalizationManager.GetText("Correspondent: ") & Correspondent & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Found variants:") & vbCrLf & _
        MatchResult.candidatesList & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Enter posting number to save") & vbCrLf & _
        LocalizationManager.GetText("(or leave empty to cancel):"), _
        LocalizationManager.GetText("Posting Selection"), _
        MatchResult.ProvodkaNumber)
    
    If Trim(userChoice) <> "" Then
        ' Write selected posting
        Dim wsData As Worksheet
        Dim tblData As ListObject
        
        Set wsData = ThisWorkbook.Worksheets("IncOut")
        Set tblData = wsData.ListObjects("TableIncOut")
        
        tblData.DataBodyRange.Cells(RowIndex, 18).value = Trim(userChoice)
        
        MsgBox LocalizationManager.GetText("[OK] Posting saved: ") & Trim(userChoice), vbInformation, LocalizationManager.GetText("Selection Saved")
        
        ' Update form if open
        If TableEventHandler.IsFormOpen("UserFormVhIsh") Then
            UserFormVhIsh.txtOtmetkaIspolnenie.Text = Trim(userChoice)
            UserFormVhIsh.txtOtmetkaIspolnenie.BackColor = RGB(255, 255, 200) ' Light yellow (manual input)
        End If
    End If
End Sub

' Search posting for current record from form
Public Sub FindProvodkaForCurrentRecord()
    If RecordOperations.CurrentRecordRow > 0 Then
        Call ProcessSingleRecord(RecordOperations.CurrentRecordRow)
    Else
        MsgBox LocalizationManager.GetText("Select a record in the table or open form with record"), vbExclamation, LocalizationManager.GetText("Posting Search")
    End If
End Sub

' Get matching statistics
Public Sub ShowMatchingStatistics()
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    
    Dim totalRecords As Long
    Dim filledRecords As Long
    Dim emptyRecords As Long
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    totalRecords = tblData.ListRows.Count
    
    For i = 1 To totalRecords
        If Trim(CStr(tblData.DataBodyRange.Cells(i, 18).value)) <> "" Then
            filledRecords = filledRecords + 1
        Else
            emptyRecords = emptyRecords + 1
        End If
    Next i
    
    MsgBox LocalizationManager.GetText("--- 1C MATCHING STATISTICS:") & vbCrLf & vbCrLf & _
           LocalizationManager.GetText("Total records in system: ") & totalRecords & vbCrLf & _
           LocalizationManager.GetText("With execution mark: ") & filledRecords & " (" & Format(filledRecords / totalRecords, "0.0%") & ")" & vbCrLf & _
           LocalizationManager.GetText("Without execution mark: ") & emptyRecords & " (" & Format(emptyRecords / totalRecords, "0.0%") & ")" & vbCrLf & vbCrLf & _
           LocalizationManager.GetText("Recommendation: Use 'Mass processing' to") & vbCrLf & _
           LocalizationManager.GetText("automatically fill empty marks."), _
           vbInformation, LocalizationManager.GetText("Integration Statistics")
End Sub

' Clear all execution marks (for reprocessing)
Public Sub ClearAllProvodkaMarks()
    Dim response As VbMsgBoxResult
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    Dim clearedCount As Long
    
    response = MsgBox(LocalizationManager.GetText("[WARN] ATTENTION!") & vbCrLf & vbCrLf & _
                      LocalizationManager.GetText("You are about to clear ALL execution marks") & vbCrLf & _
                      LocalizationManager.GetText("in the TableIncOut table.") & vbCrLf & vbCrLf & _
                      LocalizationManager.GetText("This action cannot be undone!") & vbCrLf & _
                      LocalizationManager.GetText("Continue?"), _
                      vbYesNo + vbExclamation + vbDefaultButton2, LocalizationManager.GetText("Clear Confirmation"))
    
    If response = vbNo Then Exit Sub
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    Application.StatusBar = LocalizationManager.GetText("Clearing execution marks...")
    
    For i = 1 To tblData.ListRows.Count
        If Trim(CStr(tblData.DataBodyRange.Cells(i, 18).value)) <> "" Then
            tblData.DataBodyRange.Cells(i, 18).value = ""
            clearedCount = clearedCount + 1
        End If
    Next i
    
    MsgBox LocalizationManager.GetText("[OK] Clearing completed!") & vbCrLf & _
           LocalizationManager.GetText("Records cleared: ") & clearedCount, _
           vbInformation, LocalizationManager.GetText("Clearing Executed")
           
    Application.StatusBar = LocalizationManager.GetText("Marks clearing completed. Cleared ") & clearedCount & LocalizationManager.GetText(" records.")
End Sub

