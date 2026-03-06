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
    
    For i = 1 To tblData.ListRows.Count
        
        ' Get data of current IncOut record
        On Error Resume Next
        currentSuma = CDbl(tblData.DataBodyRange.Cells(i, 6).value)          ' Document amount
        currentCorrespondent = CStr(tblData.DataBodyRange.Cells(i, 9).value) ' Received from
        currentOtmetka = Trim(CStr(tblData.DataBodyRange.Cells(i, 18).value)) ' Execution mark
        On Error GoTo MassProcessError
        
        ' Skip records with already filled execution mark
        If currentOtmetka = "" Then
            
            ' Search for match in 1C export
            MatchResult = FindMatchInFile(currentSuma, currentCorrespondent, ws1C)
            
            If MatchResult.Found Then
                ' Write posting number to execution mark
                tblData.DataBodyRange.Cells(i, 18).value = MatchResult.ProvodkaNumber
                FoundCount = FoundCount + 1
                
            ElseIf MatchResult.MatchCount > 1 Then
                ' Multiple matches - require manual processing
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
        ' Partial correspondent match
        If (currentStatus <> "1") And _
           (Abs(currentSuma - suma) < 0.01) And _
           (InStr(UCase(currentCorrespondent), UCase(Correspondent)) > 0) Then
            
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
    
    ' Search for match
    MatchResult = FindMatchInFile(currentSuma, currentCorrespondent, ws1C)
    
    ' Close file
    wb1C.Close False
    
    ' Process result
    If MatchResult.Found Then
        ' Automatically fill field
        tblData.DataBodyRange.Cells(RowIndex, 18).value = MatchResult.ProvodkaNumber
        
        MsgBox LocalizationManager.GetText("[OK] POSTING FOUND!") & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Posting Number: ") & MatchResult.ProvodkaNumber & vbCrLf & _
               LocalizationManager.GetText("Posting Date: ") & Format(MatchResult.ProvodkaDate, "dd.mm.yyyy") & vbCrLf & _
               LocalizationManager.GetText("Amount: ") & currentSuma & vbCrLf & _
               LocalizationManager.GetText("Correspondent: ") & currentCorrespondent & vbCrLf & vbCrLf & _
               LocalizationManager.GetText("Posting number written to 'Execution Mark'"), _
               vbInformation, LocalizationManager.GetText("Posting Search")
               
        ' If form is open, update display
        If TableEventHandler.IsFormOpen("UserFormVhIsh") Then
            UserFormVhIsh.txtOtmetkaIspolnenie.Text = MatchResult.ProvodkaNumber
            UserFormVhIsh.txtOtmetkaIspolnenie.BackColor = RGB(200, 255, 200) ' Light green
        End If
        
    ElseIf MatchResult.MatchCount > 1 Then
        ' Show dialog to choose from multiple variants
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

