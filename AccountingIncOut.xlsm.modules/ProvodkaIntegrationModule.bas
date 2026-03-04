Attribute VB_Name = "ProvodkaIntegrationModule"
'==============================================
' 1C INTEGRATION MODULE - ProvodkaIntegrationModule
' Purpose: Automatic matching of 1C postings with IncOut documents
' State: NEW MODULE FOR 1C EXPORT INTEGRATION
' Version: 1.0.0
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
        , "Select 1C export file")
    
    If filePath = "False" Then Exit Sub
    
    ' Open 1C export file
    Application.StatusBar = "Opening 1C export file..."
    Set wb1C = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws1C = wb1C.Worksheets(1) ' First sheet of the file
    
    ' Get IncOut table
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    Application.StatusBar = "Starting mass processing..."
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
            Application.StatusBar = "Processed " & (ProcessedCount + SkippedCount) & " of " & tblData.ListRows.Count & " records"
        End If
        
    Next i
    
    ' Close 1C export file
    wb1C.Close False
    
    Application.ScreenUpdating = True
    
    ' Show processing results
    MsgBox "MASS PROCESSING COMPLETED:" & vbCrLf & vbCrLf & _
           "--- STATISTICS:" & vbCrLf & _
           "Total records in table: " & tblData.ListRows.Count & vbCrLf & _
           "Processed (without mark): " & ProcessedCount & vbCrLf & _
           "Skipped (already filled): " & SkippedCount & vbCrLf & vbCrLf & _
           "--- RESULTS:" & vbCrLf & _
           "Found automatically: " & FoundCount & vbCrLf & _
           "Multiple matches: " & MultipleCount & vbCrLf & _
           "Not found: " & (ProcessedCount - FoundCount - MultipleCount) & vbCrLf & vbCrLf & _
           "--- Success rate: " & Format(IIf(ProcessedCount > 0, FoundCount / ProcessedCount, 0), "0.0%"), _
           vbInformation, "1C Integration Results"
           
    Application.StatusBar = "Integration completed. Found " & FoundCount & " matches out of " & ProcessedCount & " records."
    
    Exit Sub
    
MassProcessError:
    Application.ScreenUpdating = True
    If Not wb1C Is Nothing Then wb1C.Close False
    
    MsgBox "Mass processing error:" & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.description & vbCrLf & _
           "Processed records: " & ProcessedCount, _
           vbCritical, "Critical Error"
           
    Application.StatusBar = "Mass processing error"
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
    result.StatusMessage = "Not found"
    result.candidatesList = ""
    
    ' Determine last row with data
    LastRow = ws1C.Cells(ws1C.Rows.Count, 1).End(xlUp).Row
    
    ' Check data availability
    If LastRow < 2 Then
        result.StatusMessage = "Export file is empty"
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
        result.StatusMessage = "Single match found"
        
    ElseIf CandidatesCount > 1 Then
        result.Found = False ' Requires manual choice
        result.ProvodkaNumber = bestCandidate
        result.ProvodkaDate = bestCandidateDate
        result.StatusMessage = "Found " & CandidatesCount & " variants (selected by date)"
        
    Else
        result.Found = False
        result.StatusMessage = "No match found"
    End If
    
    FindMatchInFile = result
    Exit Function
    
FindError:
    result.Found = False
    result.MatchCount = 0
    result.StatusMessage = "Search error: " & Err.description
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
        MsgBox "Invalid record number: " & RowIndex, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Get record data
    currentSuma = CDbl(tblData.DataBodyRange.Cells(RowIndex, 6).value)
    currentCorrespondent = CStr(tblData.DataBodyRange.Cells(RowIndex, 9).value)
    
    ' Select 1C export file
    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        , "Select 1C export file for posting search")
    
    If filePath = "False" Then Exit Sub
    
    ' Open 1C export file
    Application.StatusBar = "Searching for posting in 1C file..."
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
        
        MsgBox "[OK] POSTING FOUND!" & vbCrLf & vbCrLf & _
               "Posting Number: " & MatchResult.ProvodkaNumber & vbCrLf & _
               "Posting Date: " & Format(MatchResult.ProvodkaDate, "dd.mm.yyyy") & vbCrLf & _
               "Amount: " & currentSuma & vbCrLf & _
               "Correspondent: " & currentCorrespondent & vbCrLf & vbCrLf & _
               "Posting number written to 'Execution Mark'", _
               vbInformation, "Posting Search"
               
        ' If form is open, update display
        If TableEventHandler.IsFormOpen("UserFormVhIsh") Then
            UserFormVhIsh.txtOtmetkaIspolnenie.Text = MatchResult.ProvodkaNumber
            UserFormVhIsh.txtOtmetkaIspolnenie.BackColor = RGB(200, 255, 200) ' Light green
        End If
        
    ElseIf MatchResult.MatchCount > 1 Then
        ' Show dialog to choose from multiple variants
        Call ShowMultipleChoiceDialog(RowIndex, MatchResult, currentSuma, currentCorrespondent)
        
    Else
        MsgBox "[WARN] POSTING NOT FOUND" & vbCrLf & vbCrLf & _
               "Search criteria:" & vbCrLf & _
               "Amount: " & currentSuma & vbCrLf & _
               "Correspondent: " & currentCorrespondent & vbCrLf & vbCrLf & _
               "Possible reasons:" & vbCrLf & _
               "- Document not yet posted in 1C" & vbCrLf & _
               "- Amount or correspondent name differs" & vbCrLf & _
               "- Document reversed in 1C", _
               vbExclamation, "Posting Search"
    End If
    
    Application.StatusBar = "Posting search completed"
    Exit Sub
    
SingleProcessError:
    If Not wb1C Is Nothing Then wb1C.Close False
    
    MsgBox "Posting search error:" & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.description, _
           vbCritical, "Error"
           
    Application.StatusBar = "Posting search error"
End Sub

' Dialog to choose from multiple variants
Private Sub ShowMultipleChoiceDialog(RowIndex As Long, MatchResult As MatchResult, suma As Double, Correspondent As String)
    Dim userChoice As String
    Dim selectedProvodka As String
    
    userChoice = InputBox( _
        "[INFO] MULTIPLE POSTING VARIANTS FOUND:" & vbCrLf & vbCrLf & _
        "Search criteria:" & vbCrLf & _
        "Amount: " & suma & vbCrLf & _
        "Correspondent: " & Correspondent & vbCrLf & vbCrLf & _
        "Found variants:" & vbCrLf & _
        MatchResult.candidatesList & vbCrLf & vbCrLf & _
        "Enter posting number to save" & vbCrLf & _
        "(or leave empty to cancel):", _
        "Posting Selection", _
        MatchResult.ProvodkaNumber)
    
    If Trim(userChoice) <> "" Then
        ' Write selected posting
        Dim wsData As Worksheet
        Dim tblData As ListObject
        
        Set wsData = ThisWorkbook.Worksheets("IncOut")
        Set tblData = wsData.ListObjects("TableIncOut")
        
        tblData.DataBodyRange.Cells(RowIndex, 18).value = Trim(userChoice)
        
        MsgBox "[OK] Posting saved: " & Trim(userChoice), vbInformation, "Selection Saved"
        
        ' Update form if open
        If TableEventHandler.IsFormOpen("UserFormVhIsh") Then
            UserFormVhIsh.txtOtmetkaIspolnenie.Text = Trim(userChoice)
            UserFormVhIsh.txtOtmetkaIspolnenie.BackColor = RGB(255, 255, 200) ' Light yellow (manual input)
        End If
    End If
End Sub

' Search posting for current record from form
Public Sub FindProvodkaForCurrentRecord()
    If DataManager.CurrentRecordRow > 0 Then
        Call ProcessSingleRecord(DataManager.CurrentRecordRow)
    Else
        MsgBox "Select a record in the table or open form with record", vbExclamation, "Posting Search"
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
    
    MsgBox "--- 1C MATCHING STATISTICS:" & vbCrLf & vbCrLf & _
           "Total records in system: " & totalRecords & vbCrLf & _
           "With execution mark: " & filledRecords & " (" & Format(filledRecords / totalRecords, "0.0%") & ")" & vbCrLf & _
           "Without execution mark: " & emptyRecords & " (" & Format(emptyRecords / totalRecords, "0.0%") & ")" & vbCrLf & vbCrLf & _
           "Recommendation: Use 'Mass processing' to" & vbCrLf & _
           "automatically fill empty marks.", _
           vbInformation, "Integration Statistics"
End Sub

' Clear all execution marks (for reprocessing)
Public Sub ClearAllProvodkaMarks()
    Dim response As VbMsgBoxResult
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long
    Dim clearedCount As Long
    
    response = MsgBox("[WARN] ATTENTION!" & vbCrLf & vbCrLf & _
                      "You are about to clear ALL execution marks" & vbCrLf & _
                      "in the TableIncOut table." & vbCrLf & vbCrLf & _
                      "This action cannot be undone!" & vbCrLf & _
                      "Continue?", _
                      vbYesNo + vbExclamation + vbDefaultButton2, "Clear Confirmation")
    
    If response = vbNo Then Exit Sub
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    Application.StatusBar = "Clearing execution marks..."
    
    For i = 1 To tblData.ListRows.Count
        If Trim(CStr(tblData.DataBodyRange.Cells(i, 18).value)) <> "" Then
            tblData.DataBodyRange.Cells(i, 18).value = ""
            clearedCount = clearedCount + 1
        End If
    Next i
    
    MsgBox "[OK] Clearing completed!" & vbCrLf & _
           "Records cleared: " & clearedCount, _
           vbInformation, "Clearing Executed"
           
    Application.StatusBar = "Marks clearing completed. Cleared " & clearedCount & " records."
End Sub

