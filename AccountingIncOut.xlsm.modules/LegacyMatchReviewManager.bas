Attribute VB_Name = "LegacyMatchReviewManager"
Option Explicit

Private Const LEGACY_REVIEW_SHEET_NAME As String = "LegacyMatchReview"
Private Const LEGACY_REVIEW_TABLE_NAME As String = "TableLegacyMatchReview"

Private Const LEGACY_COLUMN_REVIEW_ID As String = "ReviewId"
Private Const LEGACY_COLUMN_INCOUT_ROW As String = "IncOutRowIndex"
Private Const LEGACY_COLUMN_RECORD_NUMBER As String = "RecordNumber"
Private Const LEGACY_COLUMN_SERVICE As String = "Service"
Private Const LEGACY_COLUMN_DOCUMENT_TYPE As String = "DocumentType"
Private Const LEGACY_COLUMN_DOCUMENT_NUMBER As String = "DocumentNumber"
Private Const LEGACY_COLUMN_DOCUMENT_DATE As String = "DocumentDate"
Private Const LEGACY_COLUMN_AMOUNT As String = "Amount"
Private Const LEGACY_COLUMN_COUNTERPARTY As String = "Counterparty"
Private Const LEGACY_COLUMN_BEST_NUMBER As String = "BestCandidateNumber"
Private Const LEGACY_COLUMN_BEST_DATE As String = "BestCandidateDate"
Private Const LEGACY_COLUMN_BEST_COMMENT As String = "BestCandidateComment"
Private Const LEGACY_COLUMN_CANDIDATES_COUNT As String = "CandidatesCount"
Private Const LEGACY_COLUMN_CANDIDATES_LIST As String = "CandidatesList"
Private Const LEGACY_COLUMN_USE_BEST As String = "UseBestCandidate"
Private Const LEGACY_COLUMN_SELECTED_NUMBER As String = "SelectedOperationNumber"
Private Const LEGACY_COLUMN_SELECTED_DATE As String = "SelectedOperationDate"
Private Const LEGACY_COLUMN_REVIEW_STATUS As String = "ReviewStatus"
Private Const LEGACY_COLUMN_APPLIED_AT As String = "AppliedAt"
Private Const LEGACY_COLUMN_APPLY_ERROR As String = "ApplyError"

Private Type LegacyReviewMatchResult
    Found As Boolean
    BestNumber As String
    BestDate As Variant
    BestComment As String
    MatchCount As Long
    StatusMessage As String
    CandidatesList As String
End Type

Public Sub EnsureLegacyMatchReviewSchema()
    On Error GoTo SchemaError

    Dim ws As Worksheet
    Dim reviewTable As ListObject

    Set ws = CommonUtilities.GetWorksheetSafe(LEGACY_REVIEW_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LEGACY_REVIEW_SHEET_NAME
    End If

    Set reviewTable = CommonUtilities.GetListObjectSafe(ws, LEGACY_REVIEW_TABLE_NAME)
    If reviewTable Is Nothing Then
        Call CreateLegacyMatchReviewTable(ws)
        Set reviewTable = CommonUtilities.GetListObjectSafe(ws, LEGACY_REVIEW_TABLE_NAME)
    End If

    If reviewTable Is Nothing Then Exit Sub

    Call EnsureLegacyReviewColumns(reviewTable)
    Call FormatLegacyReviewSheet(ws, reviewTable)
    Exit Sub

SchemaError:
    Debug.Print "EnsureLegacyMatchReviewSchema error: " & Err.Description
End Sub

Public Sub BuildLegacyMatchReviewQueueWithFileSelection()
    Dim filePath As String
    Dim resultText As String

    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        , LocalizationManager.GetText("Select 1C export file for legacy review queue"))

    If filePath = "False" Then Exit Sub

    resultText = BuildLegacyMatchReviewQueueFromFile(CStr(filePath))
    MsgBox resultText, vbInformation, LocalizationManager.GetText("Legacy match review")
End Sub

Public Function BuildLegacyMatchReviewQueueFromFile(ByVal filePath As String) As String
    Dim wb1C As Workbook
    Dim ws1C As Worksheet

    On Error GoTo BuildError

    Set wb1C = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws1C = wb1C.Worksheets(1)
    BuildLegacyMatchReviewQueueFromFile = BuildLegacyMatchReviewQueueFromWorksheet(ws1C, False)

    Application.StatusBar = LocalizationManager.GetText("Legacy review queue completed.")
    Exit Function

BuildError:
    On Error Resume Next
    If Not wb1C Is Nothing Then wb1C.Close False
    Application.StatusBar = False
    BuildLegacyMatchReviewQueueFromFile = LocalizationManager.GetText("Legacy review build error: ") & Err.Description
End Function

Public Function BuildLegacyMatchReviewQueueFromWorksheet(ByVal ws1C As Worksheet, Optional ByVal skipBackfillCandidates As Boolean = False, Optional ByRef processedCount As Long = 0, Optional ByRef exactCount As Long = 0, Optional ByRef multipleCount As Long = 0, Optional ByRef notFoundCount As Long = 0) As String
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim reviewTable As ListObject
    Dim rowIndex As Long
    Dim currentAmount As Double
    Dim currentCorrespondent As String
    Dim currentExecutionMark As String
    Dim reviewResult As LegacyReviewMatchResult

    On Error GoTo BuildError

    Call EnsureLegacyMatchReviewSchema
    Set reviewTable = GetLegacyReviewTable()
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    If reviewTable Is Nothing Or tblData Is Nothing Or ws1C Is Nothing Then Exit Function

    Application.StatusBar = LocalizationManager.GetText("Building legacy review queue...")

    For rowIndex = 1 To tblData.ListRows.Count
        If PackageDocumentsManager.ShouldUseChildDocumentsForMatching(rowIndex) Then GoTo ContinueLoop
        If skipBackfillCandidates Then
            If LegacyPackageBackfillManager.IsLegacyPackageBackfillCandidateRowIndex(rowIndex) Then GoTo ContinueLoop
        End If

        currentExecutionMark = Trim$(CStr(tblData.DataBodyRange.Cells(rowIndex, 18).Value))
        If Len(currentExecutionMark) > 0 Then
            Call MarkLegacyReviewRowsResolved(reviewTable, rowIndex, "resolved_execution")
            GoTo ContinueLoop
        End If

        If Not IsNumeric(tblData.DataBodyRange.Cells(rowIndex, 6).Value) Then GoTo ContinueLoop
        currentAmount = CDbl(tblData.DataBodyRange.Cells(rowIndex, 6).Value)
        currentCorrespondent = Trim$(CStr(tblData.DataBodyRange.Cells(rowIndex, 9).Value))
        If Len(currentCorrespondent) = 0 Then GoTo ContinueLoop

        reviewResult = FindLegacyReviewMatchesInFile(currentAmount, currentCorrespondent, ws1C)
        processedCount = processedCount + 1

        If reviewResult.Found Then
            tblData.DataBodyRange.Cells(rowIndex, 18).Value = reviewResult.BestNumber
            Call MarkLegacyReviewRowsResolved(reviewTable, rowIndex, "resolved_exact")
            exactCount = exactCount + 1
        ElseIf reviewResult.MatchCount > 1 Then
            Call UpsertLegacyReviewRow(reviewTable, tblData, rowIndex, reviewResult)
            multipleCount = multipleCount + 1
        Else
            notFoundCount = notFoundCount + 1
        End If

ContinueLoop:
    Next rowIndex

    Call FormatLegacyReviewSheet(reviewTable.Parent, reviewTable)

    BuildLegacyMatchReviewQueueFromWorksheet = LocalizationManager.GetText("Legacy review queue completed.") & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Legacy records processed: ") & processedCount & vbCrLf & _
        LocalizationManager.GetText("Exact matches written: ") & exactCount & vbCrLf & _
        LocalizationManager.GetText("Multiple matches queued: ") & multipleCount & vbCrLf & _
        LocalizationManager.GetText("Not found: ") & notFoundCount & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Queue sheet opened for manual review.")
    Exit Function

BuildError:
    BuildLegacyMatchReviewQueueFromWorksheet = LocalizationManager.GetText("Legacy review build error: ") & Err.Description
End Function

Public Function ApplyLegacyMatchReviewSelections() As String
    Dim reviewTable As ListObject
    Dim dataTable As ListObject
    Dim rowIndex As Long
    Dim appliedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    Dim rowApplied As Boolean
    Dim rowSkipped As Boolean
    Dim applyErrorText As String

    On Error GoTo ApplyError

    Call EnsureLegacyMatchReviewSchema
    Set reviewTable = GetLegacyReviewTable()
    Set dataTable = ThisWorkbook.Worksheets("IncOut").ListObjects("TableIncOut")
    If reviewTable Is Nothing Or dataTable Is Nothing Then Exit Function

    If reviewTable.DataBodyRange Is Nothing Then
        ApplyLegacyMatchReviewSelections = LocalizationManager.GetText("Review queue is empty.")
        Exit Function
    End If

    For rowIndex = 1 To reviewTable.ListRows.Count
        If Not IsPendingLegacyReviewRow(reviewTable, rowIndex) Then GoTo ContinueLoop

        rowApplied = ApplyLegacyReviewRow(reviewTable, dataTable, rowIndex, rowSkipped, applyErrorText)
        If rowApplied Then
            appliedCount = appliedCount + 1
        ElseIf rowSkipped Then
            skippedCount = skippedCount + 1
        Else
            errorCount = errorCount + 1
        End If

ContinueLoop:
    Next rowIndex

    ApplyLegacyMatchReviewSelections = LocalizationManager.GetText("Legacy review selections applied.") & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Applied rows: ") & appliedCount & vbCrLf & _
        LocalizationManager.GetText("Skipped rows: ") & skippedCount & vbCrLf & _
        LocalizationManager.GetText("Apply errors: ") & errorCount
    Exit Function

ApplyError:
    ApplyLegacyMatchReviewSelections = LocalizationManager.GetText("Legacy review apply error: ") & Err.Description
End Function

Public Sub OpenLegacyMatchReviewForm()
    On Error GoTo OpenError

    Call EnsureLegacyMatchReviewSchema
    Load UserFormLegacyMatchReview
    UserFormLegacyMatchReview.InitializeForReview
    UserFormLegacyMatchReview.Show vbModeless
    Exit Sub

OpenError:
    MsgBox LocalizationManager.GetText("Legacy review apply error: ") & Err.Description, vbExclamation, LocalizationManager.GetText("Legacy match review")
End Sub

Public Function GetNextPendingLegacyReviewRow(Optional ByVal currentRowIndex As Long = 0) As Long
    Dim reviewTable As ListObject
    Dim rowIndex As Long

    Set reviewTable = GetLegacyReviewTable()
    If reviewTable Is Nothing Then Exit Function
    If reviewTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = currentRowIndex + 1 To reviewTable.ListRows.Count
        If IsPendingLegacyReviewRow(reviewTable, rowIndex) Then
            GetNextPendingLegacyReviewRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    For rowIndex = 1 To currentRowIndex
        If IsPendingLegacyReviewRow(reviewTable, rowIndex) Then
            GetNextPendingLegacyReviewRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Public Function GetPendingLegacyReviewCount() As Long
    Dim reviewTable As ListObject
    Dim rowIndex As Long

    Set reviewTable = GetLegacyReviewTable()
    If reviewTable Is Nothing Then Exit Function
    If reviewTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To reviewTable.ListRows.Count
        If IsPendingLegacyReviewRow(reviewTable, rowIndex) Then
            GetPendingLegacyReviewCount = GetPendingLegacyReviewCount + 1
        End If
    Next rowIndex
End Function

Public Sub BindLegacyReviewForm(ByVal frm As Object, Optional ByVal reviewRowIndex As Long = 0)
    Dim reviewTable As ListObject
    Dim rowToLoad As Long
    Dim summaryText As String

    On Error GoTo BindError

    If frm Is Nothing Then Exit Sub

    Set reviewTable = GetLegacyReviewTable()
    If reviewTable Is Nothing Then Exit Sub

    rowToLoad = reviewRowIndex
    If rowToLoad <= 0 Then rowToLoad = GetNextPendingLegacyReviewRow(0)

    summaryText = LocalizationManager.GetText("Pending Rows:") & " " & GetPendingLegacyReviewCount()
    If rowToLoad <= 0 Then
        frm.lblQueueSummary.Caption = summaryText & " | " & LocalizationManager.GetText("No pending legacy review rows.")
        frm.txtQueueSummary.Text = frm.lblQueueSummary.Caption
        Call ClearLegacyReviewForm(frm)
        Exit Sub
    End If

    frm.lblQueueSummary.Caption = summaryText & " | " & LocalizationManager.GetText("Current Review") & ": " & rowToLoad
    frm.txtQueueSummary.Text = frm.lblQueueSummary.Caption
    Call LoadLegacyReviewRowIntoForm(frm, reviewTable, rowToLoad)
    Exit Sub

BindError:
    Debug.Print "BindLegacyReviewForm error: " & Err.Description
End Sub

Public Sub ApplyLegacyReviewFromForm(ByVal frm As Object, Optional ByVal moveNext As Boolean = False)
    Dim reviewTable As ListObject
    Dim dataTable As ListObject
    Dim reviewRowIndex As Long
    Dim nextRowIndex As Long
    Dim rowApplied As Boolean
    Dim rowSkipped As Boolean
    Dim applyErrorText As String

    On Error GoTo ApplyFormError

    If frm Is Nothing Then Exit Sub

    Set reviewTable = GetLegacyReviewTable()
    Set dataTable = ThisWorkbook.Worksheets("IncOut").ListObjects("TableIncOut")
    If reviewTable Is Nothing Or dataTable Is Nothing Then Exit Sub

    reviewRowIndex = CLng(Val(CStr(frm.txtReviewRowIndex.Text)))
    If reviewRowIndex < 1 Then Exit Sub

    Call SaveLegacyReviewDecisionFromForm(frm, reviewTable, reviewRowIndex)

    rowApplied = ApplyLegacyReviewRow(reviewTable, dataTable, reviewRowIndex, rowSkipped, applyErrorText)
    If Not rowApplied And Not rowSkipped Then
        MsgBox LocalizationManager.GetText("Legacy review apply error: ") & applyErrorText, vbExclamation, LocalizationManager.GetText("Legacy match review")
    End If

    If moveNext Then
        nextRowIndex = GetNextPendingLegacyReviewRow(reviewRowIndex)
    Else
        nextRowIndex = reviewRowIndex
    End If

    Call BindLegacyReviewForm(frm, nextRowIndex)
    Exit Sub

ApplyFormError:
    MsgBox LocalizationManager.GetText("Legacy review apply error: ") & Err.Description, vbExclamation, LocalizationManager.GetText("Legacy match review")
End Sub

Public Sub MoveLegacyReviewFormNext(ByVal frm As Object)
    Dim currentRowIndex As Long
    Dim nextRowIndex As Long

    If frm Is Nothing Then Exit Sub

    currentRowIndex = CLng(Val(CStr(frm.txtReviewRowIndex.Text)))
    nextRowIndex = GetNextPendingLegacyReviewRow(currentRowIndex)
    Call BindLegacyReviewForm(frm, nextRowIndex)
End Sub

Public Sub SelectLegacyCandidateFromForm(ByVal frm As Object)
    Dim selectedLine As String
    Dim delimiterPos As Long
    Dim remainderText As String
    Dim secondDelimiterPos As Long

    If frm Is Nothing Then Exit Sub
    If frm.lstCandidates.ListIndex < 0 Then Exit Sub

    selectedLine = CStr(frm.lstCandidates.List(frm.lstCandidates.ListIndex))
    delimiterPos = InStr(1, selectedLine, " | ", vbTextCompare)
    If delimiterPos <= 0 Then Exit Sub

    frm.txtSelectedOperationNumber.Text = Left$(selectedLine, delimiterPos - 1)
    remainderText = Mid$(selectedLine, delimiterPos + 3)
    secondDelimiterPos = InStr(1, remainderText, " | ", vbTextCompare)

    If secondDelimiterPos > 0 Then
        frm.txtSelectedOperationDate.Text = Left$(remainderText, secondDelimiterPos - 1)
        frm.txtCandidateComment.Text = Mid$(remainderText, secondDelimiterPos + 3)
    Else
        frm.txtSelectedOperationDate.Text = remainderText
        frm.txtCandidateComment.Text = vbNullString
    End If

    frm.chkUseBestCandidate.Value = False
End Sub

Private Function FindLegacyReviewMatchesInFile(ByVal amountValue As Double, ByVal correspondent As String, ByVal ws1C As Worksheet) As LegacyReviewMatchResult
    Dim result As LegacyReviewMatchResult
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim currentStatus As String
    Dim currentAmount As Double
    Dim currentCorrespondent As String
    Dim currentNumber As String
    Dim currentDate As Variant
    Dim currentComment As String

    On Error GoTo FindError

    lastRow = ws1C.Cells(ws1C.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        result.StatusMessage = LocalizationManager.GetText("Export file is empty")
        FindLegacyReviewMatchesInFile = result
        Exit Function
    End If

    For rowIndex = 2 To lastRow
        On Error Resume Next
        currentStatus = CStr(ws1C.Cells(rowIndex, 1).Value)
        currentAmount = CDbl(ws1C.Cells(rowIndex, 5).Value)
        currentCorrespondent = CStr(ws1C.Cells(rowIndex, 6).Value)
        currentNumber = CStr(ws1C.Cells(rowIndex, 3).Value)
        currentDate = ws1C.Cells(rowIndex, 2).Value
        currentComment = CStr(ws1C.Cells(rowIndex, 9).Value)
        On Error GoTo FindError

        If currentStatus <> "1" Then
            If Abs(currentAmount - amountValue) < 0.01 Then
                If CommonUtilities.CorrespondentTextsMatch(currentCorrespondent, correspondent) Then
                    Call AddLegacyCandidate(result, currentNumber, currentDate, currentComment)
                End If
            End If
        End If
    Next rowIndex

    If result.MatchCount = 1 Then
        result.Found = True
        result.StatusMessage = LocalizationManager.GetText("Single match found")
    ElseIf result.MatchCount > 1 Then
        result.StatusMessage = LocalizationManager.GetText("Found ") & result.MatchCount & LocalizationManager.GetText(" variants (selected by date)")
    Else
        result.StatusMessage = LocalizationManager.GetText("No match found")
    End If

    FindLegacyReviewMatchesInFile = result
    Exit Function

FindError:
    result.Found = False
    result.MatchCount = 0
    result.StatusMessage = LocalizationManager.GetText("Search error: ") & Err.Description
    FindLegacyReviewMatchesInFile = result
End Function

Private Sub AddLegacyCandidate(ByRef result As LegacyReviewMatchResult, ByVal operationNumber As String, ByVal operationDate As Variant, ByVal operationComment As String)
    Dim candidateLine As String
    Dim bestDateValue As Date
    Dim currentDateValue As Date

    result.MatchCount = result.MatchCount + 1
    candidateLine = operationNumber
    If IsDate(operationDate) Then candidateLine = candidateLine & " | " & Format$(CDate(operationDate), "dd.mm.yyyy")
    If Len(Trim$(operationComment)) > 0 Then candidateLine = candidateLine & " | " & Trim$(operationComment)

    If result.MatchCount = 1 Then
        result.BestNumber = operationNumber
        result.BestDate = operationDate
        result.BestComment = operationComment
        result.CandidatesList = candidateLine
        Exit Sub
    End If

    result.CandidatesList = result.CandidatesList & vbLf & candidateLine

    If IsDate(result.BestDate) And IsDate(operationDate) Then
        bestDateValue = CDate(result.BestDate)
        currentDateValue = CDate(operationDate)
        If currentDateValue < bestDateValue Then
            result.BestNumber = operationNumber
            result.BestDate = operationDate
            result.BestComment = operationComment
        End If
    End If
End Sub

Private Sub AddLegacyReviewRow(ByVal reviewTable As ListObject, ByVal dataTable As ListObject, ByVal rowIndex As Long, ByRef reviewResult As LegacyReviewMatchResult)
    Dim newRow As ListRow
    Dim reviewRowIndex As Long

    Set newRow = reviewTable.ListRows.Add
    reviewRowIndex = newRow.Index

    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_REVIEW_ID, CreateLegacyReviewId())
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_INCOUT_ROW, rowIndex)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_RECORD_NUMBER, dataTable.DataBodyRange.Cells(rowIndex, 1).Value)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_SERVICE, dataTable.DataBodyRange.Cells(rowIndex, 2).Value)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_DOCUMENT_TYPE, dataTable.DataBodyRange.Cells(rowIndex, 4).Value)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_DOCUMENT_NUMBER, dataTable.DataBodyRange.Cells(rowIndex, 5).Value)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_DOCUMENT_DATE, dataTable.DataBodyRange.Cells(rowIndex, 8).Value)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_AMOUNT, dataTable.DataBodyRange.Cells(rowIndex, 6).Value)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_COUNTERPARTY, dataTable.DataBodyRange.Cells(rowIndex, 9).Value)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_BEST_NUMBER, reviewResult.BestNumber)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_BEST_DATE, reviewResult.BestDate)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_BEST_COMMENT, reviewResult.BestComment)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_CANDIDATES_COUNT, reviewResult.MatchCount)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_CANDIDATES_LIST, reviewResult.CandidatesList)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_USE_BEST, False)
    Call SetLegacyReviewValue(reviewTable, reviewRowIndex, LEGACY_COLUMN_REVIEW_STATUS, "pending")
End Sub

Private Sub UpsertLegacyReviewRow(ByVal reviewTable As ListObject, ByVal dataTable As ListObject, ByVal rowIndex As Long, ByRef reviewResult As LegacyReviewMatchResult)
    Dim existingRowIndex As Long

    existingRowIndex = FindExistingLegacyReviewRow(reviewTable, rowIndex)
    If existingRowIndex > 0 Then
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_RECORD_NUMBER, dataTable.DataBodyRange.Cells(rowIndex, 1).Value)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_SERVICE, dataTable.DataBodyRange.Cells(rowIndex, 2).Value)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_DOCUMENT_TYPE, dataTable.DataBodyRange.Cells(rowIndex, 4).Value)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_DOCUMENT_NUMBER, dataTable.DataBodyRange.Cells(rowIndex, 5).Value)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_DOCUMENT_DATE, dataTable.DataBodyRange.Cells(rowIndex, 8).Value)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_AMOUNT, dataTable.DataBodyRange.Cells(rowIndex, 6).Value)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_COUNTERPARTY, dataTable.DataBodyRange.Cells(rowIndex, 9).Value)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_BEST_NUMBER, reviewResult.BestNumber)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_BEST_DATE, reviewResult.BestDate)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_BEST_COMMENT, reviewResult.BestComment)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_CANDIDATES_COUNT, reviewResult.MatchCount)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_CANDIDATES_LIST, reviewResult.CandidatesList)
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_REVIEW_STATUS, "pending")
        Call SetLegacyReviewValue(reviewTable, existingRowIndex, LEGACY_COLUMN_APPLY_ERROR, vbNullString)
    Else
        Call AddLegacyReviewRow(reviewTable, dataTable, rowIndex, reviewResult)
    End If
End Sub

Private Function FindExistingLegacyReviewRow(ByVal reviewTable As ListObject, ByVal incOutRowIndex As Long) As Long
    Dim rowIndex As Long

    If reviewTable Is Nothing Then Exit Function
    If reviewTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To reviewTable.ListRows.Count
        If CLng(Val(CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_INCOUT_ROW)))) = incOutRowIndex Then
            FindExistingLegacyReviewRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub MarkLegacyReviewRowsResolved(ByVal reviewTable As ListObject, ByVal incOutRowIndex As Long, ByVal targetStatus As String)
    Dim rowIndex As Long
    Dim currentStatus As String

    If reviewTable Is Nothing Then Exit Sub
    If reviewTable.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To reviewTable.ListRows.Count
        If CLng(Val(CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_INCOUT_ROW)))) = incOutRowIndex Then
            currentStatus = LCase$(Trim$(CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_REVIEW_STATUS))))
            If currentStatus <> "applied" Then
                Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_REVIEW_STATUS, targetStatus)
                Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_APPLY_ERROR, vbNullString)
            End If
        End If
    Next rowIndex
End Sub

Private Sub ClearLegacyReviewQueue(ByVal reviewTable As ListObject)
    On Error Resume Next
    If Not reviewTable Is Nothing Then
        If Not reviewTable.DataBodyRange Is Nothing Then reviewTable.DataBodyRange.Rows.Delete
    End If
    On Error GoTo 0
End Sub

Private Function GetLegacyReviewTable() As ListObject
    Dim ws As Worksheet

    Set ws = CommonUtilities.GetWorksheetSafe(LEGACY_REVIEW_SHEET_NAME)
    If ws Is Nothing Then Exit Function
    Set GetLegacyReviewTable = CommonUtilities.GetListObjectSafe(ws, LEGACY_REVIEW_TABLE_NAME)
End Function

Private Sub CreateLegacyMatchReviewTable(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim columnCount As Long
    Dim index As Long
    Dim tableRange As Range
    Dim reviewTable As ListObject

    headers = GetLegacyReviewHeaders()
    columnCount = UBound(headers) - LBound(headers) + 1

    For index = LBound(headers) To UBound(headers)
        ws.Cells(1, index - LBound(headers) + 1).Value = CStr(headers(index))
    Next index

    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(2, columnCount))
    Set reviewTable = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    reviewTable.Name = LEGACY_REVIEW_TABLE_NAME

    On Error Resume Next
    If Not reviewTable.DataBodyRange Is Nothing Then reviewTable.DataBodyRange.Rows.Delete
    On Error GoTo 0
End Sub

Private Sub EnsureLegacyReviewColumns(ByVal reviewTable As ListObject)
    Dim headers As Variant
    Dim index As Long

    headers = GetLegacyReviewHeaders()
    For index = LBound(headers) To UBound(headers)
        Call EnsureLegacyReviewColumn(reviewTable, CStr(headers(index)))
    Next index
End Sub

Private Function GetLegacyReviewHeaders() As Variant
    Dim headers(0 To 19) As String

    headers(0) = LEGACY_COLUMN_REVIEW_ID
    headers(1) = LEGACY_COLUMN_INCOUT_ROW
    headers(2) = LEGACY_COLUMN_RECORD_NUMBER
    headers(3) = LEGACY_COLUMN_SERVICE
    headers(4) = LEGACY_COLUMN_DOCUMENT_TYPE
    headers(5) = LEGACY_COLUMN_DOCUMENT_NUMBER
    headers(6) = LEGACY_COLUMN_DOCUMENT_DATE
    headers(7) = LEGACY_COLUMN_AMOUNT
    headers(8) = LEGACY_COLUMN_COUNTERPARTY
    headers(9) = LEGACY_COLUMN_BEST_NUMBER
    headers(10) = LEGACY_COLUMN_BEST_DATE
    headers(11) = LEGACY_COLUMN_BEST_COMMENT
    headers(12) = LEGACY_COLUMN_CANDIDATES_COUNT
    headers(13) = LEGACY_COLUMN_CANDIDATES_LIST
    headers(14) = LEGACY_COLUMN_USE_BEST
    headers(15) = LEGACY_COLUMN_SELECTED_NUMBER
    headers(16) = LEGACY_COLUMN_SELECTED_DATE
    headers(17) = LEGACY_COLUMN_REVIEW_STATUS
    headers(18) = LEGACY_COLUMN_APPLIED_AT
    headers(19) = LEGACY_COLUMN_APPLY_ERROR

    GetLegacyReviewHeaders = headers
End Function

Private Sub EnsureLegacyReviewColumn(ByVal reviewTable As ListObject, ByVal columnName As String)
    If reviewTable Is Nothing Then Exit Sub
    If CommonUtilities.GetListColumnSafe(reviewTable, columnName) Is Nothing Then
        reviewTable.ListColumns.Add.Name = columnName
    End If
End Sub

Private Sub FormatLegacyReviewSheet(ByVal ws As Worksheet, ByVal reviewTable As ListObject)
    Dim candidatesColumn As ListColumn
    Dim commentColumn As ListColumn

    If ws Is Nothing Or reviewTable Is Nothing Then Exit Sub

    ws.Cells.WrapText = False
    ws.Columns("A:T").EntireColumn.AutoFit
    ws.Columns(CommonUtilities.GetListColumnSafe(reviewTable, LEGACY_COLUMN_COUNTERPARTY).Range.Column).ColumnWidth = 34
    ws.Columns(CommonUtilities.GetListColumnSafe(reviewTable, LEGACY_COLUMN_BEST_COMMENT).Range.Column).ColumnWidth = 42
    ws.Columns(CommonUtilities.GetListColumnSafe(reviewTable, LEGACY_COLUMN_CANDIDATES_LIST).Range.Column).ColumnWidth = 70

    Set candidatesColumn = CommonUtilities.GetListColumnSafe(reviewTable, LEGACY_COLUMN_CANDIDATES_LIST)
    Set commentColumn = CommonUtilities.GetListColumnSafe(reviewTable, LEGACY_COLUMN_BEST_COMMENT)
    If Not candidatesColumn Is Nothing Then
        If Not candidatesColumn.DataBodyRange Is Nothing Then candidatesColumn.DataBodyRange.WrapText = True
    End If
    If Not commentColumn Is Nothing Then
        If Not commentColumn.DataBodyRange Is Nothing Then commentColumn.DataBodyRange.WrapText = True
    End If
    ws.Activate
End Sub

Private Function CreateLegacyReviewId() As String
    Randomize
    CreateLegacyReviewId = "LEG-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int((Rnd() * 9000) + 1000), "0000")
End Function

Private Function GetLegacyReviewValue(ByVal reviewTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim listColumn As ListColumn

    Set listColumn = CommonUtilities.GetListColumnSafe(reviewTable, columnName)
    If listColumn Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > reviewTable.ListRows.Count Then Exit Function

    GetLegacyReviewValue = reviewTable.DataBodyRange.Cells(rowIndex, listColumn.Index).Value
End Function

Private Sub SetLegacyReviewValue(ByVal reviewTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueToWrite As Variant)
    Dim listColumn As ListColumn

    Set listColumn = CommonUtilities.GetListColumnSafe(reviewTable, columnName)
    If listColumn Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > reviewTable.ListRows.Count Then Exit Sub

    reviewTable.DataBodyRange.Cells(rowIndex, listColumn.Index).Value = valueToWrite
End Sub

Private Function ApplyLegacyReviewRow(ByVal reviewTable As ListObject, ByVal dataTable As ListObject, ByVal rowIndex As Long, ByRef rowSkipped As Boolean, ByRef applyErrorText As String) As Boolean
    Dim incOutRowIndex As Long
    Dim selectedNumber As String
    Dim selectedDate As Variant
    Dim useBestCandidate As Boolean
    Dim bestNumber As String
    Dim bestDate As Variant

    incOutRowIndex = CLng(Val(CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_INCOUT_ROW))))
    useBestCandidate = IsTrueMarker(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_USE_BEST))
    bestNumber = Trim$(CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_BEST_NUMBER)))
    bestDate = GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_BEST_DATE)
    selectedNumber = Trim$(CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SELECTED_NUMBER)))
    selectedDate = GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SELECTED_DATE)

    If useBestCandidate And Len(selectedNumber) = 0 Then
        selectedNumber = bestNumber
        selectedDate = bestDate
    End If

    If Len(selectedNumber) = 0 Then
        rowSkipped = True
        Exit Function
    End If

    If incOutRowIndex < 1 Or incOutRowIndex > dataTable.ListRows.Count Then
        applyErrorText = LocalizationManager.GetText("IncOut row index is no longer valid.")
        Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_REVIEW_STATUS, "failed")
        Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_APPLY_ERROR, applyErrorText)
        Exit Function
    End If

    dataTable.DataBodyRange.Cells(incOutRowIndex, 18).Value = selectedNumber
    Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SELECTED_NUMBER, selectedNumber)
    Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SELECTED_DATE, selectedDate)
    Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_REVIEW_STATUS, "applied")
    Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_APPLIED_AT, Now)
    Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_APPLY_ERROR, vbNullString)
    ApplyLegacyReviewRow = True
End Function

Private Sub LoadLegacyReviewRowIntoForm(ByVal frm As Object, ByVal reviewTable As ListObject, ByVal rowIndex As Long)
    Dim candidatesText As String
    Dim candidateLines As Variant
    Dim i As Long

    Call ClearLegacyReviewForm(frm)

    frm.txtReviewRowIndex.Text = CStr(rowIndex)
    frm.txtReviewId.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_REVIEW_ID))
    frm.txtIncOutRow.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_INCOUT_ROW))
    frm.txtRecordNumber.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_RECORD_NUMBER))
    frm.txtService.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SERVICE))
    frm.txtDocumentType.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_DOCUMENT_TYPE))
    frm.txtDocumentNumber.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_DOCUMENT_NUMBER))
    frm.txtDocumentDate.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_DOCUMENT_DATE))
    frm.txtAmount.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_AMOUNT))
    frm.txtCounterparty.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_COUNTERPARTY))
    frm.txtBestCandidateNumber.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_BEST_NUMBER))
    frm.txtBestCandidateDate.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_BEST_DATE))
    frm.txtBestCandidateComment.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_BEST_COMMENT))
    frm.chkUseBestCandidate.Value = IsTrueMarker(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_USE_BEST))
    frm.txtSelectedOperationNumber.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SELECTED_NUMBER))
    frm.txtSelectedOperationDate.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SELECTED_DATE))
    frm.txtCurrentStatus.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_REVIEW_STATUS))

    candidatesText = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_CANDIDATES_LIST))
    If Len(candidatesText) > 0 Then
        candidateLines = Split(candidatesText, vbLf)
        For i = LBound(candidateLines) To UBound(candidateLines)
            If Len(Trim$(CStr(candidateLines(i)))) > 0 Then
                frm.lstCandidates.AddItem CStr(candidateLines(i))
            End If
        Next i
    End If

    frm.txtCandidateComment.Text = CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_BEST_COMMENT))
End Sub

Private Sub SaveLegacyReviewDecisionFromForm(ByVal frm As Object, ByVal reviewTable As ListObject, ByVal rowIndex As Long)
    Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_USE_BEST, frm.chkUseBestCandidate.Value)
    Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SELECTED_NUMBER, Trim$(CStr(frm.txtSelectedOperationNumber.Text)))
    Call SetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_SELECTED_DATE, Trim$(CStr(frm.txtSelectedOperationDate.Text)))
End Sub

Private Sub ClearLegacyReviewForm(ByVal frm As Object)
    If frm Is Nothing Then Exit Sub

    frm.txtReviewRowIndex.Text = vbNullString
    frm.txtReviewId.Text = vbNullString
    frm.txtIncOutRow.Text = vbNullString
    frm.txtRecordNumber.Text = vbNullString
    frm.txtService.Text = vbNullString
    frm.txtDocumentType.Text = vbNullString
    frm.txtDocumentNumber.Text = vbNullString
    frm.txtDocumentDate.Text = vbNullString
    frm.txtAmount.Text = vbNullString
    frm.txtCounterparty.Text = vbNullString
    frm.txtBestCandidateNumber.Text = vbNullString
    frm.txtBestCandidateDate.Text = vbNullString
    frm.txtBestCandidateComment.Text = vbNullString
    frm.txtSelectedOperationNumber.Text = vbNullString
    frm.txtSelectedOperationDate.Text = vbNullString
    frm.txtCandidateComment.Text = vbNullString
    frm.txtCurrentStatus.Text = vbNullString
    frm.chkUseBestCandidate.Value = False
    frm.lstCandidates.Clear
End Sub

Private Function IsPendingLegacyReviewRow(ByVal reviewTable As ListObject, ByVal rowIndex As Long) As Boolean
    Dim statusValue As String

    statusValue = LCase$(Trim$(CStr(GetLegacyReviewValue(reviewTable, rowIndex, LEGACY_COLUMN_REVIEW_STATUS))))
    IsPendingLegacyReviewRow = (statusValue = "pending" Or statusValue = "failed")
End Function

Private Function IsTrueMarker(ByVal sourceValue As Variant) As Boolean
    Dim marker As String

    marker = LCase$(Trim$(CStr(sourceValue)))
    Select Case marker
        Case "true", "1", "yes", "y", "x", "+"
            IsTrueMarker = True
    End Select
End Function
