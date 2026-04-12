Attribute VB_Name = "LegacyPackageBackfillManager"
Option Explicit

Private Const BACKFILL_SHEET_NAME As String = "LegacyPackageBackfill"
Private Const BACKFILL_TABLE_NAME As String = "TableLegacyPackageBackfill"

Private Const BACKFILL_COLUMN_QUEUE_ID As String = "BackfillId"
Private Const BACKFILL_COLUMN_INCOUT_ROW As String = "IncOutRowIndex"
Private Const BACKFILL_COLUMN_PACKAGE_ID As String = "PackageId"
Private Const BACKFILL_COLUMN_RECORD_NUMBER As String = "RecordNumber"
Private Const BACKFILL_COLUMN_PARENT_TYPE As String = "ParentDocumentType"
Private Const BACKFILL_COLUMN_PARENT_NUMBER As String = "ParentDocumentNumber"
Private Const BACKFILL_COLUMN_PARENT_DATE As String = "ParentDocumentDate"
Private Const BACKFILL_COLUMN_PARENT_AMOUNT As String = "ParentAmount"
Private Const BACKFILL_COLUMN_COUNTERPARTY As String = "Counterparty"
Private Const BACKFILL_COLUMN_OPERATION_TYPE As String = "OperationType"
Private Const BACKFILL_COLUMN_OPERATION_NUMBER As String = "Matched1COperationNumber"
Private Const BACKFILL_COLUMN_OPERATION_DATE As String = "Matched1COperationDate"
Private Const BACKFILL_COLUMN_OPERATION_AMOUNT As String = "Matched1COperationAmount"
Private Const BACKFILL_COLUMN_OPERATION_COMMENT As String = "Matched1CComment"
Private Const BACKFILL_COLUMN_CHILD_TYPE As String = "ProposedChildType"
Private Const BACKFILL_COLUMN_ASSET_CATEGORY As String = "DerivedAssetCategory"
Private Const BACKFILL_COLUMN_CHILD_NUMBER As String = "ProposedChildNumber"
Private Const BACKFILL_COLUMN_CHILD_DATE As String = "ProposedChildDate"
Private Const BACKFILL_COLUMN_CHILD_AMOUNT As String = "ProposedChildAmount"
Private Const BACKFILL_COLUMN_CHILD_DESCRIPTION As String = "ProposedDescription"
Private Const BACKFILL_COLUMN_ORDER_HINT As String = "OrderHint"
Private Const BACKFILL_COLUMN_GROUP_TOTAL As String = "GroupProposedTotal"
Private Const BACKFILL_COLUMN_GROUP_COUNT As String = "GroupItemCount"
Private Const BACKFILL_COLUMN_GROUP_AMOUNT_STATUS As String = "GroupAmountCheckStatus"
Private Const BACKFILL_COLUMN_CONFIDENCE As String = "Confidence"
Private Const BACKFILL_COLUMN_USE_PROPOSAL As String = "UseProposal"
Private Const BACKFILL_COLUMN_REVIEW_STATUS As String = "ReviewStatus"
Private Const BACKFILL_COLUMN_APPLIED_ITEM_ID As String = "AppliedItemId"
Private Const BACKFILL_COLUMN_APPLIED_AT As String = "AppliedAt"
Private Const BACKFILL_COLUMN_APPLY_ERROR As String = "ApplyError"

Private Const SOURCE_COLUMN_DATE As Long = 2
Private Const SOURCE_COLUMN_NUMBER As Long = 3
Private Const SOURCE_COLUMN_TYPE As Long = 4
Private Const SOURCE_COLUMN_AMOUNT As Long = 5
Private Const SOURCE_COLUMN_COUNTERPARTY As Long = 6
Private Const SOURCE_COLUMN_COMMENT As Long = 9

Private Type BackfillProposal
    OperationType As String
    OperationNumber As String
    OperationDate As Variant
    OperationAmount As Double
    OperationComment As String
    ProposedChildType As String
    DerivedAssetCategory As String
    ProposedChildNumber As String
    ProposedChildDate As Variant
    ProposedChildAmount As Double
    ProposedDescription As String
    OrderHint As String
    Confidence As Long
End Type

Public Sub EnsureLegacyPackageBackfillSchema()
    On Error GoTo SchemaError

    Dim ws As Worksheet
    Dim backfillTable As ListObject

    Set ws = CommonUtilities.GetWorksheetSafe(BACKFILL_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = BACKFILL_SHEET_NAME
    End If

    Set backfillTable = CommonUtilities.GetListObjectSafe(ws, BACKFILL_TABLE_NAME)
    If backfillTable Is Nothing Then
        Call CreateLegacyPackageBackfillTable(ws)
        Set backfillTable = CommonUtilities.GetListObjectSafe(ws, BACKFILL_TABLE_NAME)
    End If

    If backfillTable Is Nothing Then Exit Sub

    Call EnsureLegacyPackageBackfillColumns(backfillTable)
    Call FormatLegacyPackageBackfillSheet(ws, backfillTable)
    Exit Sub

SchemaError:
    Debug.Print "EnsureLegacyPackageBackfillSchema error: " & Err.Description
End Sub

Public Sub BuildLegacyPackageBackfillQueueWithFileSelection()
    Dim filePath As String
    Dim resultText As String

    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        , LocalizationManager.GetText("Select 1C export file for legacy package backfill"))

    If filePath = "False" Then Exit Sub

    resultText = BuildLegacyPackageBackfillQueueFromFile(CStr(filePath))
    MsgBox resultText, vbInformation, LocalizationManager.GetText("Legacy package backfill")
End Sub

Public Function BuildLegacyPackageBackfillQueueFromFile(ByVal filePath As String) As String
    Dim wb1C As Workbook
    Dim ws1C As Worksheet
    Dim wsData As Worksheet
    Dim parentTable As ListObject
    Dim backfillTable As ListObject
    Dim rowIndex As Long
    Dim parentCount As Long
    Dim proposalCount As Long
    Dim skippedCount As Long
    Dim packageId As String
    Dim parentDocumentNumber As String
    Dim parentCorrespondent As String
    Dim parentAmount As Double
    Dim parentDateValue As Variant
    Dim reusedCount As Long

    On Error GoTo BuildError

    Call EnsurePackageDocumentsSchema
    Call EnsureLegacyPackageBackfillSchema

    Set backfillTable = GetLegacyPackageBackfillTable()
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set parentTable = wsData.ListObjects(PACKAGE_PARENT_TABLE_NAME)
    If backfillTable Is Nothing Or parentTable Is Nothing Then Exit Function

    Application.StatusBar = LocalizationManager.GetText("Building legacy package backfill queue...")
    Set wb1C = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws1C = wb1C.Worksheets(1)

    For rowIndex = 1 To parentTable.ListRows.Count
        packageId = EnsurePackageIdForParentRow(parentTable, rowIndex)
        If Len(packageId) = 0 Then
            skippedCount = skippedCount + 1
            GoTo ContinueLoop
        End If

        If ShouldSkipParentForBackfill(parentTable, rowIndex, packageId) Then
            Call MarkBackfillRowsClosedForPackage(backfillTable, packageId)
            skippedCount = skippedCount + 1
            GoTo ContinueLoop
        End If

        parentDocumentNumber = NormalizeLegacyCellText(parentTable.DataBodyRange.Cells(rowIndex, 5).Value)
        parentCorrespondent = NormalizeLegacyCellText(parentTable.DataBodyRange.Cells(rowIndex, 9).Value)
        parentDateValue = parentTable.DataBodyRange.Cells(rowIndex, 8).Value
        If Len(parentDocumentNumber) = 0 Or Len(parentCorrespondent) = 0 Then
            skippedCount = skippedCount + 1
            GoTo ContinueLoop
        End If
        If Not IsNumeric(parentTable.DataBodyRange.Cells(rowIndex, 6).Value) Then
            skippedCount = skippedCount + 1
            GoTo ContinueLoop
        End If

        parentAmount = CDbl(parentTable.DataBodyRange.Cells(rowIndex, 6).Value)
        proposalCount = proposalCount + AddBackfillRowsForParent(backfillTable, parentTable, rowIndex, packageId, parentDocumentNumber, parentDateValue, parentAmount, parentCorrespondent, ws1C, reusedCount)
        parentCount = parentCount + 1

ContinueLoop:
    Next rowIndex

    wb1C.Close False
    Call UpdateBackfillGroupMetrics(backfillTable)
    Call FormatLegacyPackageBackfillSheet(backfillTable.Parent, backfillTable)
    backfillTable.Parent.Activate

    BuildLegacyPackageBackfillQueueFromFile = LocalizationManager.GetText("Legacy package backfill completed.") & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Legacy parent rows processed: ") & parentCount & vbCrLf & _
        LocalizationManager.GetText("Backfill proposals created: ") & proposalCount & vbCrLf & _
        LocalizationManager.GetText("Backfill proposals reused: ") & reusedCount & vbCrLf & _
        LocalizationManager.GetText("Skipped rows: ") & skippedCount & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Queue sheet opened for review.")

    Application.StatusBar = LocalizationManager.GetText("Legacy package backfill completed.")
    Exit Function

BuildError:
    On Error Resume Next
    If Not wb1C Is Nothing Then wb1C.Close False
    Application.StatusBar = False
    BuildLegacyPackageBackfillQueueFromFile = LocalizationManager.GetText("Legacy package backfill error: ") & Err.Description
End Function

Public Function ApplyLegacyPackageBackfillSelections() As String
    Dim backfillTable As ListObject
    Dim appliedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    Dim rowIndex As Long
    Dim rowApplied As Boolean
    Dim rowSkipped As Boolean

    On Error GoTo ApplyError

    Call EnsurePackageDocumentsSchema
    Call EnsureLegacyPackageBackfillSchema

    Set backfillTable = GetLegacyPackageBackfillTable()
    If backfillTable Is Nothing Then Exit Function

    If backfillTable.DataBodyRange Is Nothing Then
        ApplyLegacyPackageBackfillSelections = LocalizationManager.GetText("Backfill queue is empty.")
        Exit Function
    End If

    For rowIndex = 1 To backfillTable.ListRows.Count
        If Not IsPendingBackfillRow(backfillTable, rowIndex) Then GoTo ContinueLoop

        rowApplied = ApplyLegacyPackageBackfillRow(backfillTable, rowIndex, rowSkipped)
        If rowApplied Then
            appliedCount = appliedCount + 1
        ElseIf rowSkipped Then
            skippedCount = skippedCount + 1
        Else
            errorCount = errorCount + 1
        End If

ContinueLoop:
    Next rowIndex

    ApplyLegacyPackageBackfillSelections = LocalizationManager.GetText("Legacy package backfill applied.") & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Applied rows: ") & appliedCount & vbCrLf & _
        LocalizationManager.GetText("Skipped rows: ") & skippedCount & vbCrLf & _
        LocalizationManager.GetText("Apply errors: ") & errorCount
    Exit Function

ApplyError:
    ApplyLegacyPackageBackfillSelections = LocalizationManager.GetText("Legacy package backfill error: ") & Err.Description
End Function

Public Sub OpenLegacyPackageBackfillReviewForm()
    On Error GoTo OpenError

    Call EnsureLegacyPackageBackfillSchema
    Load UserFormLegacyMatchReview
    UserFormLegacyMatchReview.InitializeForBackfillReview
    UserFormLegacyMatchReview.Show vbModeless
    Exit Sub

OpenError:
    MsgBox LocalizationManager.GetText("Legacy package backfill error: ") & Err.Description, vbExclamation, LocalizationManager.GetText("Legacy package backfill")
End Sub

Public Function GetNextPendingLegacyPackageBackfillRow(Optional ByVal currentRowIndex As Long = 0) As Long
    Dim backfillTable As ListObject
    Dim rowIndex As Long

    Set backfillTable = GetLegacyPackageBackfillTable()
    If backfillTable Is Nothing Then Exit Function
    If backfillTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = currentRowIndex + 1 To backfillTable.ListRows.Count
        If IsPendingBackfillRow(backfillTable, rowIndex) Then
            GetNextPendingLegacyPackageBackfillRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    For rowIndex = 1 To currentRowIndex
        If IsPendingBackfillRow(backfillTable, rowIndex) Then
            GetNextPendingLegacyPackageBackfillRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Public Function GetPendingLegacyPackageBackfillCount() As Long
    Dim backfillTable As ListObject
    Dim rowIndex As Long

    Set backfillTable = GetLegacyPackageBackfillTable()
    If backfillTable Is Nothing Then Exit Function
    If backfillTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To backfillTable.ListRows.Count
        If IsPendingBackfillRow(backfillTable, rowIndex) Then
            GetPendingLegacyPackageBackfillCount = GetPendingLegacyPackageBackfillCount + 1
        End If
    Next rowIndex
End Function

Public Sub BindLegacyPackageBackfillForm(ByVal frm As Object, Optional ByVal reviewRowIndex As Long = 0)
    Dim backfillTable As ListObject
    Dim rowToLoad As Long
    Dim summaryText As String

    On Error GoTo BindError

    If frm Is Nothing Then Exit Sub

    Set backfillTable = GetLegacyPackageBackfillTable()
    If backfillTable Is Nothing Then Exit Sub

    rowToLoad = reviewRowIndex
    If rowToLoad <= 0 Then rowToLoad = GetNextPendingLegacyPackageBackfillRow(0)

    summaryText = LocalizationManager.GetText("Pending Rows:") & " " & GetPendingLegacyPackageBackfillCount()
    If rowToLoad <= 0 Then
        frm.lblQueueSummary.Caption = summaryText
        frm.txtQueueSummary.Text = summaryText & vbCrLf & LocalizationManager.GetText("No pending legacy package backfill rows.")
        Call ClearLegacyPackageBackfillForm(frm)
        Exit Sub
    End If

    frm.lblQueueSummary.Caption = summaryText
    frm.txtQueueSummary.Text = BuildBackfillQueueSummaryText(backfillTable, rowToLoad)
    Call LoadLegacyPackageBackfillRowIntoForm(frm, backfillTable, rowToLoad)
    Exit Sub

BindError:
    Debug.Print "BindLegacyPackageBackfillForm error: " & Err.Description
End Sub

Public Sub OpenCurrentBackfillIncOutRow(ByVal frm As Object)
    Dim rowIndex As Long
    Dim ws As Worksheet
    Dim parentTable As ListObject

    On Error GoTo OpenRowError

    If frm Is Nothing Then Exit Sub

    rowIndex = CLng(Val(CStr(frm.txtIncOutRow.Text)))
    If rowIndex < 1 Then Exit Sub

    Set ws = CommonUtilities.GetWorksheetSafe("IncOut")
    If ws Is Nothing Then Exit Sub

    Set parentTable = CommonUtilities.GetListObjectSafe(ws, PACKAGE_PARENT_TABLE_NAME)
    If parentTable Is Nothing Then Exit Sub
    If parentTable.DataBodyRange Is Nothing Then Exit Sub
    If rowIndex > parentTable.ListRows.Count Then Exit Sub

    ws.Activate
    parentTable.DataBodyRange.Cells(rowIndex, 1).Select
    Exit Sub

OpenRowError:
    MsgBox LocalizationManager.GetText("Legacy package backfill error: ") & Err.Description, vbExclamation, LocalizationManager.GetText("Legacy package backfill")
End Sub

Public Sub ApplyLegacyPackageBackfillFromForm(ByVal frm As Object, Optional ByVal moveNext As Boolean = False)
    Dim backfillTable As ListObject
    Dim reviewRowIndex As Long
    Dim nextRowIndex As Long
    Dim rowApplied As Boolean
    Dim rowSkipped As Boolean
    Dim currentPackageId As String

    On Error GoTo ApplyFormError

    If frm Is Nothing Then Exit Sub

    Set backfillTable = GetLegacyPackageBackfillTable()
    If backfillTable Is Nothing Then Exit Sub

    reviewRowIndex = CLng(Val(CStr(frm.txtReviewRowIndex.Text)))
    If reviewRowIndex < 1 Then Exit Sub
    currentPackageId = CStr(GetBackfillValue(backfillTable, reviewRowIndex, BACKFILL_COLUMN_PACKAGE_ID))

    Call SaveLegacyPackageBackfillDecisionFromForm(frm, backfillTable, reviewRowIndex)

    rowApplied = ApplyLegacyPackageBackfillRow(backfillTable, reviewRowIndex, rowSkipped)
    If Not rowApplied And Not rowSkipped Then
        MsgBox LocalizationManager.GetText("Legacy package backfill error: ") & CStr(GetBackfillValue(backfillTable, reviewRowIndex, BACKFILL_COLUMN_APPLY_ERROR)), vbExclamation, LocalizationManager.GetText("Legacy package backfill")
    End If

    If rowApplied And Not moveNext Then
        nextRowIndex = GetNextPendingLegacyPackageBackfillRowInPackage(currentPackageId, reviewRowIndex)
        If nextRowIndex <= 0 Then nextRowIndex = GetNextPendingLegacyPackageBackfillRow(reviewRowIndex)
    ElseIf moveNext Then
        nextRowIndex = GetNextPendingLegacyPackageBackfillRow(reviewRowIndex)
    Else
        nextRowIndex = reviewRowIndex
    End If

    Call BindLegacyPackageBackfillForm(frm, nextRowIndex)
    Exit Sub

ApplyFormError:
    MsgBox LocalizationManager.GetText("Legacy package backfill error: ") & Err.Description, vbExclamation, LocalizationManager.GetText("Legacy package backfill")
End Sub

Private Function GetNextPendingLegacyPackageBackfillRowInPackage(ByVal packageId As String, Optional ByVal currentRowIndex As Long = 0) As Long
    Dim backfillTable As ListObject
    Dim rowIndex As Long

    If Len(Trim$(packageId)) = 0 Then Exit Function

    Set backfillTable = GetLegacyPackageBackfillTable()
    If backfillTable Is Nothing Then Exit Function
    If backfillTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = currentRowIndex + 1 To backfillTable.ListRows.Count
        If IsPendingBackfillRow(backfillTable, rowIndex) Then
            If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
                GetNextPendingLegacyPackageBackfillRowInPackage = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex

    For rowIndex = 1 To currentRowIndex
        If IsPendingBackfillRow(backfillTable, rowIndex) Then
            If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
                GetNextPendingLegacyPackageBackfillRowInPackage = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Public Sub MoveLegacyPackageBackfillFormNext(ByVal frm As Object)
    Dim currentRowIndex As Long
    Dim nextRowIndex As Long

    If frm Is Nothing Then Exit Sub

    currentRowIndex = CLng(Val(CStr(frm.txtReviewRowIndex.Text)))
    nextRowIndex = GetNextPendingLegacyPackageBackfillRow(currentRowIndex)
    Call BindLegacyPackageBackfillForm(frm, nextRowIndex)
End Sub

Public Sub SelectLegacyPackageProposalFromForm(ByVal frm As Object)
    Dim selectedLine As String
    Dim delimiterPos As Long
    Dim rowIndexText As String
    Dim rowIndex As Long

    If frm Is Nothing Then Exit Sub
    If frm.lstCandidates.ListIndex < 0 Then Exit Sub

    selectedLine = CStr(frm.lstCandidates.List(frm.lstCandidates.ListIndex))
    delimiterPos = InStr(1, selectedLine, " | ", vbTextCompare)
    If delimiterPos <= 0 Then Exit Sub

    rowIndexText = Trim$(Left$(selectedLine, delimiterPos - 1))
    rowIndex = CLng(Val(rowIndexText))
    If rowIndex > 0 Then Call BindLegacyPackageBackfillForm(frm, rowIndex)
End Sub

Private Function AddBackfillRowsForParent(ByVal backfillTable As ListObject, ByVal parentTable As ListObject, ByVal parentRowIndex As Long, ByVal packageId As String, ByVal parentDocumentNumber As String, ByVal parentDateValue As Variant, ByVal parentAmount As Double, ByVal parentCorrespondent As String, ByVal ws1C As Worksheet, ByRef reusedCount As Long) As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim currentStatus As String
    Dim currentType As String
    Dim currentNumber As String
    Dim currentDate As Variant
    Dim currentAmount As Double
    Dim currentCorrespondent As String
    Dim currentComment As String
    Dim proposal As BackfillProposal
    Dim existingQueueRow As Long

    If backfillTable Is Nothing Or parentTable Is Nothing Or ws1C Is Nothing Then Exit Function

    lastRow = ws1C.Cells(ws1C.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    For rowIndex = 2 To lastRow
        On Error Resume Next
        currentStatus = CStr(ws1C.Cells(rowIndex, 1).Value)
        currentDate = ws1C.Cells(rowIndex, SOURCE_COLUMN_DATE).Value
        currentNumber = Trim$(CStr(ws1C.Cells(rowIndex, SOURCE_COLUMN_NUMBER).Value))
        currentType = Trim$(CStr(ws1C.Cells(rowIndex, SOURCE_COLUMN_TYPE).Value))
        currentAmount = CDbl(ws1C.Cells(rowIndex, SOURCE_COLUMN_AMOUNT).Value)
        currentCorrespondent = Trim$(CStr(ws1C.Cells(rowIndex, SOURCE_COLUMN_COUNTERPARTY).Value))
        currentComment = Trim$(CStr(ws1C.Cells(rowIndex, SOURCE_COLUMN_COMMENT).Value))
        On Error GoTo 0

        If currentStatus = "1" Then GoTo ContinueLoop
        If Len(currentComment) = 0 Then GoTo ContinueLoop
        If Len(currentCorrespondent) = 0 Then GoTo ContinueLoop
        If Not CommonUtilities.CorrespondentTextsMatch(currentCorrespondent, parentCorrespondent) Then GoTo ContinueLoop
        If Not CommentContainsDocumentNumber(currentComment, parentDocumentNumber) Then GoTo ContinueLoop

        Call BuildBackfillProposal(currentType, currentNumber, currentDate, currentAmount, currentComment, parentDocumentNumber, parentDateValue, parentAmount, proposal)
        existingQueueRow = FindExistingBackfillQueueRow(backfillTable, packageId, proposal.OperationNumber)
        If existingQueueRow > 0 Then
            reusedCount = reusedCount + 1
        Else
            Call AddLegacyPackageBackfillRow(backfillTable, parentTable, parentRowIndex, packageId, proposal)
            AddBackfillRowsForParent = AddBackfillRowsForParent + 1
        End If

ContinueLoop:
    Next rowIndex
End Function

Private Sub BuildBackfillProposal(ByVal operationType As String, ByVal operationNumber As String, ByVal operationDate As Variant, ByVal operationAmount As Double, ByVal operationComment As String, ByVal parentDocumentNumber As String, ByVal parentDateValue As Variant, ByVal parentAmount As Double, ByRef proposal As BackfillProposal)
    Dim childTypeText As String
    Dim childNumberText As String
    Dim childDateValue As Variant
    Dim confidenceValue As Long
    Dim parentDateMatched As Boolean

    proposal.OperationType = operationType
    proposal.OperationNumber = operationNumber
    proposal.OperationDate = operationDate
    proposal.OperationAmount = operationAmount
    proposal.OperationComment = operationComment

    Call ExtractChildReferenceFromComment(operationComment, parentDocumentNumber, operationType, childTypeText, childNumberText, childDateValue)

    proposal.ProposedChildType = GetBackfillChildTypeDisplay(childTypeText, operationType)
    proposal.DerivedAssetCategory = GetBackfillAssetCategoryValue(proposal.ProposedChildType, operationType)
    proposal.ProposedChildNumber = childNumberText
    proposal.ProposedChildDate = childDateValue
    proposal.ProposedChildAmount = operationAmount
    proposal.ProposedDescription = operationComment
    proposal.OrderHint = ExtractOrderHintFromComment(operationComment)

    parentDateMatched = CommentContainsDateVariant(operationComment, parentDateValue)
    confidenceValue = 40
    If parentDateMatched Then confidenceValue = confidenceValue + 20
    If Len(childNumberText) > 0 Then confidenceValue = confidenceValue + 20
    If IsDate(childDateValue) Then confidenceValue = confidenceValue + 10
    If Abs(operationAmount - parentAmount) < 0.01 Then confidenceValue = confidenceValue + 10
    If Len(proposal.DerivedAssetCategory) > 0 Then confidenceValue = confidenceValue + 10
    If confidenceValue > 100 Then confidenceValue = 100
    proposal.Confidence = confidenceValue
End Sub

Private Function GetBackfillChildTypeDisplay(ByVal extractedTypeText As String, ByVal operationType As String) As String
    Dim effectiveTypeText As String
    Dim normalizedOperationText As String

    effectiveTypeText = Trim$(extractedTypeText)
    normalizedOperationText = BuildBackfillTypeKey(operationType)

    If Len(effectiveTypeText) > 0 Then
        If Not IsGenericBackfillTypeLabel(effectiveTypeText) Then
            GetBackfillChildTypeDisplay = effectiveTypeText
            Exit Function
        End If
    End If

    If IsBackfillMaterialType(normalizedOperationText) Then
        GetBackfillChildTypeDisplay = BuildMaterialBackfillTypeText()
    ElseIf IsBackfillFixedAssetType(normalizedOperationText) Then
        GetBackfillChildTypeDisplay = Trim$(operationType)
    ElseIf IsBackfillAmbiguousAccountingType(normalizedOperationText) Then
        GetBackfillChildTypeDisplay = Trim$(operationType)
    ElseIf Len(effectiveTypeText) > 0 Then
        GetBackfillChildTypeDisplay = effectiveTypeText
    Else
        GetBackfillChildTypeDisplay = Trim$(operationType)
    End If
End Function

Private Function GetBackfillAssetCategoryValue(ByVal childTypeDisplay As String, ByVal operationType As String) As String
    Dim normalizedChildType As String
    Dim normalizedOperationType As String

    normalizedChildType = BuildBackfillTypeKey(childTypeDisplay)
    normalizedOperationType = BuildBackfillTypeKey(operationType)

    If IsBackfillMaterialType(normalizedChildType) Or IsBackfillMaterialType(normalizedOperationType) Then
        GetBackfillAssetCategoryValue = "inventory"
    ElseIf IsBackfillFixedAssetType(normalizedChildType) Or IsBackfillFixedAssetType(normalizedOperationType) Then
        GetBackfillAssetCategoryValue = "fixed_assets"
    ElseIf IsBackfillAmbiguousAccountingType(normalizedChildType) Or IsBackfillAmbiguousAccountingType(normalizedOperationType) Then
        GetBackfillAssetCategoryValue = vbNullString
    End If
End Function

Private Function BuildBackfillTypeKey(ByVal sourceText As String) As String
    BuildBackfillTypeKey = UCase$(Trim$(Replace(Replace(Replace(CStr(sourceText), vbCrLf, " "), vbTab, " "), " ", "_")))
End Function

Private Function IsGenericBackfillTypeLabel(ByVal sourceText As String) As Boolean
    Dim normalizedText As String

    normalizedText = BuildBackfillTypeKey(sourceText)
    If Len(normalizedText) = 0 Then Exit Function

    Select Case normalizedText
        Case "АКТ", "ДОКУМЕНТ", "ДОК", "НАКЛ", "НАКЛАДНАЯ"
            IsGenericBackfillTypeLabel = True
    End Select
End Function

Private Function IsBackfillMaterialType(ByVal typeKey As String) As Boolean
    IsBackfillMaterialType = (InStr(typeKey, ChrW$(1053) & ChrW$(1040) & ChrW$(1050) & ChrW$(1051) & ChrW$(1040) & ChrW$(1044) & ChrW$(1053)) > 0) _
        Or (InStr(typeKey, ChrW$(1052) & ChrW$(1040) & ChrW$(1058) & ChrW$(1045) & ChrW$(1056) & ChrW$(1048) & ChrW$(1040) & ChrW$(1051)) > 0) _
        Or (InStr(typeKey, ChrW$(1055) & ChrW$(1054) & ChrW$(1057) & ChrW$(1058) & ChrW$(1059) & ChrW$(1055) & ChrW$(1051) & ChrW$(1045) & ChrW$(1053) & ChrW$(1048) & ChrW$(1045) & "_" & ChrW$(1052) & ChrW$(1047)) > 0) _
        Or (InStr(typeKey, ChrW$(1055) & ChrW$(1056) & ChrW$(1048) & ChrW$(1045) & ChrW$(1052) & ChrW$(1050) & ChrW$(1048) & "_" & ChrW$(1052) & ChrW$(1040) & ChrW$(1058) & ChrW$(1045) & ChrW$(1056) & ChrW$(1048) & ChrW$(1040) & ChrW$(1051)) > 0)
End Function

Private Function IsBackfillFixedAssetType(ByVal typeKey As String) As Boolean
    IsBackfillFixedAssetType = (InStr(typeKey, ChrW$(1055) & ChrW$(1045) & ChrW$(1056) & ChrW$(1045) & ChrW$(1044) & ChrW$(1040) & ChrW$(1063) & ChrW$(1040) & "_" & ChrW$(1054) & ChrW$(1041) & ChrW$(1066) & ChrW$(1045) & ChrW$(1050) & ChrW$(1058) & ChrW$(1054) & ChrW$(1042)) > 0) _
        Or (InStr(typeKey, ChrW$(1055) & ChrW$(1056) & ChrW$(1048) & ChrW$(1053) & ChrW$(1071) & ChrW$(1058) & ChrW$(1048) & ChrW$(1045) & "_" & ChrW$(1050) & "_" & ChrW$(1059) & ChrW$(1063) & ChrW$(1045) & ChrW$(1058) & ChrW$(1059) & "_" & ChrW$(1054) & ChrW$(1057)) > 0) _
        Or (InStr(typeKey, ChrW$(1055) & ChrW$(1054) & ChrW$(1057) & ChrW$(1058) & ChrW$(1059) & ChrW$(1055) & ChrW$(1051) & ChrW$(1045) & ChrW$(1053) & ChrW$(1048) & ChrW$(1045) & "_" & ChrW$(1054) & ChrW$(1057)) > 0)
End Function

Private Function IsBackfillAmbiguousAccountingType(ByVal typeKey As String) As Boolean
    IsBackfillAmbiguousAccountingType = (InStr(typeKey, ChrW$(1054) & ChrW$(1055) & ChrW$(1045) & ChrW$(1056) & ChrW$(1040) & ChrW$(1062) & ChrW$(1048) & ChrW$(1071) & "_" & ChrW$(1041) & ChrW$(1059) & ChrW$(1061) & ChrW$(1043) & ChrW$(1040) & ChrW$(1051) & ChrW$(1058) & ChrW$(1045) & ChrW$(1056) & ChrW$(1057) & ChrW$(1050) & ChrW$(1040) & ChrW$(1071)) > 0)
End Function

Private Function BuildMaterialBackfillTypeText() As String
    BuildMaterialBackfillTypeText = _
        ChrW$(1053) & ChrW$(1072) & ChrW$(1082) & ChrW$(1083) & ChrW$(1072) & ChrW$(1076) & ChrW$(1085) & ChrW$(1072) & ChrW$(1103) & " " & _
        ChrW$(1085) & ChrW$(1072) & " " & ChrW$(1086) & ChrW$(1090) & ChrW$(1087) & ChrW$(1091) & ChrW$(1089) & ChrW$(1082) & " " & _
        ChrW$(1084) & ChrW$(1072) & ChrW$(1090) & ChrW$(1077) & ChrW$(1088) & ChrW$(1080) & ChrW$(1072) & ChrW$(1083) & ChrW$(1086) & ChrW$(1074) & " " & _
        ChrW$(1085) & ChrW$(1072) & " " & ChrW$(1089) & ChrW$(1090) & ChrW$(1086) & ChrW$(1088) & ChrW$(1086) & ChrW$(1085) & ChrW$(1091)
End Function

Private Sub ExtractChildReferenceFromComment(ByVal operationComment As String, ByVal parentDocumentNumber As String, ByVal fallbackType As String, ByRef childTypeText As String, ByRef childNumberText As String, ByRef childDateValue As Variant)
    Dim regex As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim parentSeen As Boolean
    Dim currentNumber As String
    Dim currentTypeText As String
    Dim currentDateText As String

    childTypeText = Trim$(fallbackType)
    childNumberText = vbNullString
    childDateValue = vbNullString

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "([^" & ChrW$(8470) & vbCrLf & ",;]{0,40})" & ChrW$(8470) & "\s*([^\s,;|]+)(?:[^0-9]{0,10}(\d{2}\.\d{2}\.\d{2,4}))?"

    If Not regex.Test(operationComment) Then Exit Sub

    Set matches = regex.Execute(operationComment)
    For Each matchItem In matches
        currentTypeText = CleanReferenceLabel(CStr(matchItem.SubMatches(0)))
        currentNumber = Trim$(CStr(matchItem.SubMatches(1)))
        currentDateText = Trim$(CStr(matchItem.SubMatches(2)))

        If NumbersMatch(currentNumber, parentDocumentNumber) Then
            parentSeen = True
        ElseIf parentSeen Then
            childTypeText = ChooseBackfillTypeText(currentTypeText, fallbackType)
            childNumberText = currentNumber
            If Len(currentDateText) > 0 Then childDateValue = NormalizeDateTextValue(currentDateText)
            Exit Sub
        ElseIf Len(childNumberText) = 0 Then
            childTypeText = ChooseBackfillTypeText(currentTypeText, fallbackType)
            childNumberText = currentNumber
            If Len(currentDateText) > 0 Then childDateValue = NormalizeDateTextValue(currentDateText)
        End If
    Next matchItem

    If NumbersMatch(childNumberText, parentDocumentNumber) Then
        childNumberText = vbNullString
        childDateValue = vbNullString
        childTypeText = Trim$(fallbackType)
    End If
End Sub

Private Function ChooseBackfillTypeText(ByVal extractedText As String, ByVal fallbackText As String) As String
    If Len(Trim$(extractedText)) > 0 Then
        ChooseBackfillTypeText = extractedText
    Else
        ChooseBackfillTypeText = Trim$(fallbackText)
    End If
End Function

Private Function CleanReferenceLabel(ByVal sourceText As String) As String
    Dim cleanedText As String

    cleanedText = Replace(sourceText, vbTab, " ")
    cleanedText = Replace(cleanedText, ".", " ")
    cleanedText = Replace(cleanedText, ":", " ")
    cleanedText = Replace(cleanedText, "(", " ")
    cleanedText = Replace(cleanedText, ")", " ")
    cleanedText = Replace(cleanedText, "-", " ")
    cleanedText = Trim$(cleanedText)

    Do While InStr(cleanedText, "  ") > 0
        cleanedText = Replace(cleanedText, "  ", " ")
    Loop

    CleanReferenceLabel = cleanedText
End Function

Private Function CommentContainsDocumentNumber(ByVal operationComment As String, ByVal documentNumber As String) As Boolean
    Dim normalizedComment As String
    Dim normalizedDocumentNumber As String

    normalizedComment = NormalizeBackfillNumberToken(operationComment)
    normalizedDocumentNumber = NormalizeBackfillNumberToken(documentNumber)
    If Len(normalizedDocumentNumber) = 0 Then Exit Function

    CommentContainsDocumentNumber = (InStr(1, normalizedComment, normalizedDocumentNumber, vbTextCompare) > 0)
End Function

Private Function CommentContainsDateVariant(ByVal operationComment As String, ByVal documentDateValue As Variant) As Boolean
    Dim longDateText As String
    Dim shortDateText As String

    If Not IsDate(documentDateValue) Then Exit Function

    longDateText = Format$(CDate(documentDateValue), "dd.mm.yyyy")
    shortDateText = Format$(CDate(documentDateValue), "dd.mm.yy")

    If InStr(1, operationComment, longDateText, vbTextCompare) > 0 Then
        CommentContainsDateVariant = True
        Exit Function
    End If

    CommentContainsDateVariant = (InStr(1, operationComment, shortDateText, vbTextCompare) > 0)
End Function

Private Function NormalizeBackfillNumberToken(ByVal sourceText As String) As String
    Dim normalizedText As String
    Dim index As Long
    Dim currentChar As String

    sourceText = TrimNumericDecimalSuffix(Trim$(sourceText))

    For index = 1 To Len(sourceText)
        currentChar = Mid$(sourceText, index, 1)
        If IsBackfillTokenCharacter(currentChar) Then normalizedText = normalizedText & UCase$(currentChar)
    Next index

    NormalizeBackfillNumberToken = normalizedText
End Function

Private Function NormalizeLegacyCellText(ByVal sourceValue As Variant) As String
    NormalizeLegacyCellText = TrimNumericDecimalSuffix(Trim$(CStr(sourceValue)))
End Function

Private Function TrimNumericDecimalSuffix(ByVal sourceText As String) As String
    Dim regex As Object
    Dim matches As Object

    If Len(sourceText) = 0 Then Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "^\s*(\d+)([.,]0+)\s*$"

    If regex.Test(sourceText) Then
        Set matches = regex.Execute(sourceText)
        TrimNumericDecimalSuffix = CStr(matches(0).SubMatches(0))
    Else
        TrimNumericDecimalSuffix = sourceText
    End If
End Function

Private Function IsBackfillTokenCharacter(ByVal sourceChar As String) As Boolean
    Dim charCode As Long

    If Len(sourceChar) = 0 Then Exit Function
    If sourceChar = "/" Or sourceChar = "-" Then
        IsBackfillTokenCharacter = True
        Exit Function
    End If
    If sourceChar Like "[0-9A-Za-z]" Then
        IsBackfillTokenCharacter = True
        Exit Function
    End If

    charCode = AscW(sourceChar)
    If (charCode >= 1040 And charCode <= 1103) Or charCode = 1025 Or charCode = 1105 Then
        IsBackfillTokenCharacter = True
    End If
End Function

Private Function NumbersMatch(ByVal leftNumber As String, ByVal rightNumber As String) As Boolean
    Dim normalizedLeft As String
    Dim normalizedRight As String

    normalizedLeft = NormalizeBackfillNumberToken(leftNumber)
    normalizedRight = NormalizeBackfillNumberToken(rightNumber)

    If Len(normalizedLeft) = 0 Or Len(normalizedRight) = 0 Then Exit Function
    NumbersMatch = (StrComp(normalizedLeft, normalizedRight, vbTextCompare) = 0)
End Function

Private Function NormalizeDateTextValue(ByVal sourceDateText As String) As Variant
    Dim normalizedText As String

    normalizedText = Trim$(sourceDateText)
    If Len(normalizedText) = 8 Then normalizedText = Left$(normalizedText, 6) & "20" & Right$(normalizedText, 2)
    If IsDate(normalizedText) Then
        NormalizeDateTextValue = CDate(normalizedText)
    Else
        NormalizeDateTextValue = sourceDateText
    End If
End Function

Private Function ExtractOrderHintFromComment(ByVal operationComment As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim hintText As String

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "(" & GetOrderKeywordPattern() & "\s*" & ChrW$(8470) & "\s*[^,;]+(?:\s+" & GetOrderDateKeywordPattern() & "\s+\d{2}\.\d{2}\.\d{2,4})?)"

    If regex.Test(operationComment) Then
        Set matches = regex.Execute(operationComment)
        hintText = Trim$(CStr(matches(0).SubMatches(0)))
        ExtractOrderHintFromComment = hintText
    End If
End Function

Private Function GetOrderKeywordPattern() As String
    GetOrderKeywordPattern = ChrW$(1053) & ChrW$(1072) & ChrW$(1088) & ChrW$(1103) & ChrW$(1076)
End Function

Private Function GetOrderDateKeywordPattern() As String
    GetOrderDateKeywordPattern = ChrW$(1086) & ChrW$(1090)
End Function

Private Sub AddLegacyPackageBackfillRow(ByVal backfillTable As ListObject, ByVal parentTable As ListObject, ByVal parentRowIndex As Long, ByVal packageId As String, ByRef proposal As BackfillProposal)
    Dim newRow As ListRow
    Dim backfillRowIndex As Long

    Set newRow = backfillTable.ListRows.Add
    backfillRowIndex = newRow.Index

    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_QUEUE_ID, CreateLegacyPackageBackfillId())
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_INCOUT_ROW, parentRowIndex)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_PACKAGE_ID, packageId)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_RECORD_NUMBER, parentTable.DataBodyRange.Cells(parentRowIndex, 1).Value)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_PARENT_TYPE, parentTable.DataBodyRange.Cells(parentRowIndex, 4).Value)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_PARENT_NUMBER, parentTable.DataBodyRange.Cells(parentRowIndex, 5).Value)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_PARENT_DATE, parentTable.DataBodyRange.Cells(parentRowIndex, 8).Value)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_PARENT_AMOUNT, parentTable.DataBodyRange.Cells(parentRowIndex, 6).Value)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_COUNTERPARTY, parentTable.DataBodyRange.Cells(parentRowIndex, 9).Value)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_OPERATION_TYPE, proposal.OperationType)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_OPERATION_NUMBER, proposal.OperationNumber)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_OPERATION_DATE, proposal.OperationDate)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_OPERATION_AMOUNT, proposal.OperationAmount)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_OPERATION_COMMENT, proposal.OperationComment)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_CHILD_TYPE, proposal.ProposedChildType)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_ASSET_CATEGORY, proposal.DerivedAssetCategory)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_CHILD_NUMBER, proposal.ProposedChildNumber)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_CHILD_DATE, proposal.ProposedChildDate)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_CHILD_AMOUNT, proposal.ProposedChildAmount)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_CHILD_DESCRIPTION, proposal.ProposedDescription)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_ORDER_HINT, proposal.OrderHint)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_CONFIDENCE, proposal.Confidence)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_USE_PROPOSAL, False)
    Call SetBackfillValue(backfillTable, backfillRowIndex, BACKFILL_COLUMN_REVIEW_STATUS, "pending")
End Sub

Private Function ApplyLegacyPackageBackfillRow(ByVal backfillTable As ListObject, ByVal rowIndex As Long, ByRef rowSkipped As Boolean) As Boolean
    Dim parentTable As ListObject
    Dim itemsTable As ListObject
    Dim parentRowIndex As Long
    Dim packageId As String
    Dim itemId As String
    Dim existingItemId As String
    Dim selectedFlag As Boolean
    Dim childAmount As Double

    On Error GoTo ApplyRowError

    selectedFlag = IsTrueMarker(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_USE_PROPOSAL))
    If Not selectedFlag Then
        rowSkipped = True
        Exit Function
    End If

    parentRowIndex = CLng(Val(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_INCOUT_ROW))))
    If parentRowIndex < 1 Then
        Call SetBackfillFailure(backfillTable, rowIndex, LocalizationManager.GetText("IncOut row index is no longer valid."))
        Exit Function
    End If

    Set parentTable = ThisWorkbook.Worksheets("IncOut").ListObjects(PACKAGE_PARENT_TABLE_NAME)
    Set itemsTable = CommonUtilities.GetListObjectSafe(CommonUtilities.GetWorksheetSafe(PACKAGE_ITEMS_SHEET_NAME), PACKAGE_ITEMS_TABLE_NAME)
    If parentTable Is Nothing Or itemsTable Is Nothing Then
        Call SetBackfillFailure(backfillTable, rowIndex, LocalizationManager.GetText("Package tables are not available."))
        Exit Function
    End If
    If parentRowIndex > parentTable.ListRows.Count Then
        Call SetBackfillFailure(backfillTable, rowIndex, LocalizationManager.GetText("IncOut row index is no longer valid."))
        Exit Function
    End If

    packageId = EnsurePackageIdForParentRow(parentTable, parentRowIndex)
    If Len(packageId) = 0 Then
        Call SetBackfillFailure(backfillTable, rowIndex, LocalizationManager.GetText("Unable to determine PackageId for the current record."))
        Exit Function
    End If

    existingItemId = FindBackfillExistingItemId(itemsTable, packageId, CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_NUMBER)))
    If Len(existingItemId) > 0 Then
        Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_APPLIED_ITEM_ID, existingItemId)
        Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS, "duplicate")
        Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_APPLIED_AT, Now)
        Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_APPLY_ERROR, vbNullString)
        ApplyLegacyPackageBackfillRow = True
        Exit Function
    End If

    If Not IsNumeric(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_AMOUNT)) Then
        childAmount = 0
    Else
        childAmount = CDbl(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_AMOUNT))
    End If

    itemId = PackageDocumentsManager.SavePackageItemRecord( _
        parentRowIndex, _
        packageId, _
        vbNullString, _
        CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_TYPE)), _
        CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_NUMBER)), _
        GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_DATE), _
        childAmount, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_DESCRIPTION)), _
        vbNullString, _
        CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_NUMBER)), _
        GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_DATE), _
        "manual", _
        CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_COMMENT)), _
        False)

    If Len(itemId) = 0 Then
        Call SetBackfillFailure(backfillTable, rowIndex, LocalizationManager.GetText("Unable to create package child item."))
        Exit Function
    End If

    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID, packageId)
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_APPLIED_ITEM_ID, itemId)
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS, "applied")
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_APPLIED_AT, Now)
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_APPLY_ERROR, vbNullString)
    ApplyLegacyPackageBackfillRow = True
    Exit Function

ApplyRowError:
    Call SetBackfillFailure(backfillTable, rowIndex, Err.Description)
End Function

Private Function FindBackfillExistingItemId(ByVal itemsTable As ListObject, ByVal packageId As String, ByVal operationNumber As String) As String
    Dim packageColumn As ListColumn
    Dim operationColumn As ListColumn
    Dim itemIdColumn As ListColumn
    Dim rowIndex As Long

    If itemsTable Is Nothing Then Exit Function
    If itemsTable.DataBodyRange Is Nothing Then Exit Function
    If Len(Trim$(operationNumber)) = 0 Then Exit Function

    Set packageColumn = CommonUtilities.GetListColumnSafe(itemsTable, "PackageId")
    Set operationColumn = CommonUtilities.GetListColumnSafe(itemsTable, "Matched1COperationNumber")
    Set itemIdColumn = CommonUtilities.GetListColumnSafe(itemsTable, "ItemId")
    If packageColumn Is Nothing Or operationColumn Is Nothing Or itemIdColumn Is Nothing Then Exit Function

    For rowIndex = 1 To itemsTable.ListRows.Count
        If StrComp(Trim$(CStr(itemsTable.DataBodyRange.Cells(rowIndex, packageColumn.Index).Value)), Trim$(packageId), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(itemsTable.DataBodyRange.Cells(rowIndex, operationColumn.Index).Value)), Trim$(operationNumber), vbTextCompare) = 0 Then
                FindBackfillExistingItemId = Trim$(CStr(itemsTable.DataBodyRange.Cells(rowIndex, itemIdColumn.Index).Value))
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Sub SetBackfillFailure(ByVal backfillTable As ListObject, ByVal rowIndex As Long, ByVal errorText As String)
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS, "failed")
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_APPLY_ERROR, errorText)
End Sub

Private Function BuildBackfillGroupSummary(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As String
    Dim itemCount As String
    Dim totalText As String
    Dim statusText As String

    itemCount = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_COUNT))
    totalText = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_TOTAL))
    statusText = GetBackfillAmountCheckDisplayText(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_AMOUNT_STATUS)))

    If Len(itemCount) = 0 And Len(totalText) = 0 And Len(statusText) = 0 Then Exit Function

    BuildBackfillGroupSummary = " | " & LocalizationManager.GetText("Items") & ": " & itemCount & _
        " | " & LocalizationManager.GetText("Children Total") & ": " & totalText & _
        " | " & LocalizationManager.GetText("Amount Check") & ": " & statusText
End Function

Private Function BuildBackfillQueueSummaryText(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As String
    Dim packageId As String
    Dim groupCount As String
    Dim groupTotal As String
    Dim amountStatus As String
    Dim orderHint As String
    Dim reviewStatus As String
    Dim assetMix As String
    Dim targetStageText As String

    packageId = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))
    groupCount = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_COUNT))
    groupTotal = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_TOTAL))
    amountStatus = GetBackfillAmountCheckDisplayText(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_AMOUNT_STATUS)))
    orderHint = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ORDER_HINT))
    reviewStatus = GetBackfillReviewStatusDisplayText(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS)))
    assetMix = BuildBackfillAssetMixSummary(backfillTable, packageId)
    targetStageText = GetBackfillTargetStageDisplayText(backfillTable, rowIndex)

    BuildBackfillQueueSummaryText = LocalizationManager.GetText("Pending Rows:") & " " & GetPendingLegacyPackageBackfillCount() & _
        vbCrLf & LocalizationManager.GetText("Current Review") & ": " & rowIndex & _
        vbCrLf & LocalizationManager.GetText("Package ID") & ": " & packageId & _
        vbCrLf & LocalizationManager.GetText("Review status") & ": " & reviewStatus & _
        vbCrLf & LocalizationManager.GetText("Target stage") & ": " & targetStageText & _
        vbCrLf & LocalizationManager.GetText("Items") & ": " & groupCount & " | " & _
        LocalizationManager.GetText("Children Total") & ": " & groupTotal & " | " & _
        LocalizationManager.GetText("Amount Check") & ": " & amountStatus

    If Len(assetMix) > 0 Then
        BuildBackfillQueueSummaryText = BuildBackfillQueueSummaryText & vbCrLf & LocalizationManager.GetText("Asset mix") & ": " & assetMix
    End If
    If Len(Trim$(orderHint)) > 0 Then
        BuildBackfillQueueSummaryText = BuildBackfillQueueSummaryText & vbCrLf & LocalizationManager.GetText("Order hint") & ": " & orderHint
    End If
End Function

Private Sub LoadLegacyPackageBackfillRowIntoForm(ByVal frm As Object, ByVal backfillTable As ListObject, ByVal rowIndex As Long)
    On Error Resume Next
    frm.SetCandidateClickSuspended True
    On Error GoTo 0

    Call ClearLegacyPackageBackfillForm(frm)

    frm.txtReviewRowIndex.Text = CStr(rowIndex)
    frm.txtReviewId.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_QUEUE_ID))
    frm.txtIncOutRow.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_INCOUT_ROW))
    frm.txtRecordNumber.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_RECORD_NUMBER))
    frm.txtService.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))
    frm.txtDocumentType.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PARENT_TYPE))
    frm.txtDocumentNumber.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PARENT_NUMBER))
    frm.txtDocumentDate.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PARENT_DATE))
    frm.txtAmount.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PARENT_AMOUNT))
    frm.txtCounterparty.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_COUNTERPARTY))
    frm.txtBestCandidateNumber.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_NUMBER))
    frm.txtBestCandidateDate.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_DATE))
    frm.txtBestCandidateComment.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_COMMENT))
    frm.chkUseBestCandidate.Value = IsTrueMarker(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_USE_PROPOSAL))
    frm.txtSelectedOperationNumber.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_NUMBER))
    frm.txtSelectedOperationDate.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_DATE))
    frm.txtCurrentStatus.Text = GetBackfillReviewStatusDisplayText(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS)))
    frm.txtCandidateComment.Text = CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_DESCRIPTION))
    frm.lblCandidateComment.Caption = BuildBackfillSelectedProposalCaption(backfillTable, rowIndex)

    Call PopulateBackfillProposalListForPackage(frm, backfillTable, rowIndex)

    On Error Resume Next
    frm.SetCandidateClickSuspended False
    On Error GoTo 0
End Sub

Private Sub PopulateBackfillProposalListForPackage(ByVal frm As Object, ByVal backfillTable As ListObject, ByVal currentRowIndex As Long)
    Dim packageId As String
    Dim rowIndex As Long
    Dim lineText As String
    Dim selectedListIndex As Long

    packageId = Trim$(CStr(GetBackfillValue(backfillTable, currentRowIndex, BACKFILL_COLUMN_PACKAGE_ID)))
    If Len(packageId) = 0 Then Exit Sub

    selectedListIndex = -1
    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            If IsPendingBackfillRow(backfillTable, rowIndex) Then
                lineText = BuildBackfillProposalListLine(backfillTable, rowIndex)
                frm.lstCandidates.AddItem lineText
                If rowIndex = currentRowIndex Then selectedListIndex = frm.lstCandidates.ListCount - 1
            End If
        End If
    Next rowIndex

    If selectedListIndex >= 0 Then
        frm.lstCandidates.ListIndex = selectedListIndex
    ElseIf frm.lstCandidates.ListCount > 0 Then
        frm.lstCandidates.ListIndex = 0
    End If

    frm.lblCandidates.Caption = LocalizationManager.GetText("Package proposals") & " (" & frm.lstCandidates.ListCount & ")"
End Sub

Private Function BuildBackfillSelectedProposalCaption(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As String
    Dim childTypeText As String
    Dim childAmountText As String
    Dim assetCategoryText As String

    childTypeText = Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_TYPE)))
    childAmountText = Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_AMOUNT)))
    assetCategoryText = GetBackfillAssetCategoryDisplayText(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ASSET_CATEGORY))))

    BuildBackfillSelectedProposalCaption = LocalizationManager.GetText("Proposed description")
    If Len(childTypeText) > 0 Then
        BuildBackfillSelectedProposalCaption = BuildBackfillSelectedProposalCaption & " - " & childTypeText
    End If
    If Len(childAmountText) > 0 Then
        BuildBackfillSelectedProposalCaption = BuildBackfillSelectedProposalCaption & " | " & LocalizationManager.GetText("Amount") & ": " & childAmountText
    End If
    If Len(assetCategoryText) > 0 Then
        BuildBackfillSelectedProposalCaption = BuildBackfillSelectedProposalCaption & " | " & LocalizationManager.GetText("Asset category") & ": " & assetCategoryText
    End If
End Function

Private Function BuildBackfillProposalListLine(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As String
    Dim lineText As String

    lineText = CStr(rowIndex) & " | " & CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_NUMBER))
    If Len(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_DATE)))) > 0 Then
        lineText = lineText & " | " & CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_DATE))
    End If
    lineText = lineText & " | " & CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_AMOUNT))
    lineText = lineText & " | " & CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_TYPE))
    If Len(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ASSET_CATEGORY)))) > 0 Then
        lineText = lineText & " | " & GetBackfillAssetCategoryDisplayText(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ASSET_CATEGORY)))
    End If
    BuildBackfillProposalListLine = lineText
End Function

Private Sub SaveLegacyPackageBackfillDecisionFromForm(ByVal frm As Object, ByVal backfillTable As ListObject, ByVal rowIndex As Long)
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_USE_PROPOSAL, True)
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_NUMBER, Trim$(CStr(frm.txtSelectedOperationNumber.Text)))
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_DATE, Trim$(CStr(frm.txtSelectedOperationDate.Text)))
    Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_DESCRIPTION, Trim$(CStr(frm.txtCandidateComment.Text)))
End Sub

Private Sub ClearLegacyPackageBackfillForm(ByVal frm As Object)
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
    frm.lblCandidateComment.Caption = LocalizationManager.GetText("Proposed description")
    frm.lblCandidates.Caption = LocalizationManager.GetText("Package proposals")
    frm.chkUseBestCandidate.Value = False
    frm.lstCandidates.Clear
End Sub

Private Sub UpdateBackfillGroupMetrics(ByVal backfillTable As ListObject)
    Dim rowIndex As Long
    Dim packageId As String
    Dim totalAmount As Double
    Dim itemCount As Long
    Dim parentAmount As Double
    Dim statusValue As String
    Dim orderHint As String
    Dim reviewStatus As String

    If backfillTable Is Nothing Then Exit Sub
    If backfillTable.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To backfillTable.ListRows.Count
        packageId = Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID)))
        If Len(packageId) = 0 Then GoTo ContinueLoop

        totalAmount = GetBackfillGroupTotal(backfillTable, packageId)
        itemCount = GetBackfillGroupCount(backfillTable, packageId)
        parentAmount = 0
        If IsNumeric(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PARENT_AMOUNT)) Then
            parentAmount = CDbl(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PARENT_AMOUNT))
        End If
        statusValue = GetBackfillAmountCheckStatus(parentAmount, totalAmount)
        orderHint = GetBackfillGroupOrderHint(backfillTable, packageId)
        reviewStatus = GetBackfillGroupReviewStatus(backfillTable, rowIndex, packageId, statusValue)

        Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_TOTAL, totalAmount)
        Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_COUNT, itemCount)
        Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_GROUP_AMOUNT_STATUS, statusValue)
        Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ORDER_HINT, orderHint)
        If Not IsTerminalBackfillStatus(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS))) Then
            Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS, reviewStatus)
        End If

ContinueLoop:
    Next rowIndex
End Sub

Private Function GetBackfillGroupReviewStatus(ByVal backfillTable As ListObject, ByVal rowIndex As Long, ByVal packageId As String, ByVal amountStatus As String) As String
    Dim ambiguousCount As Long
    Dim confidentCount As Long
    Dim itemCount As Long
    Dim distinctOrderHintCount As Long

    itemCount = GetBackfillGroupCount(backfillTable, packageId)
    ambiguousCount = GetBackfillGroupAmbiguousCount(backfillTable, packageId)
    confidentCount = GetBackfillGroupConfidentCount(backfillTable, packageId, 70)
    distinctOrderHintCount = GetBackfillDistinctOrderHintCount(backfillTable, packageId)

    Select Case LCase$(Trim$(amountStatus))
        Case "mismatch"
            GetBackfillGroupReviewStatus = "amount_mismatch"
        Case "partial"
            GetBackfillGroupReviewStatus = "partial_group"
        Case Else
            If distinctOrderHintCount > 1 Or ambiguousCount > 0 Then
                GetBackfillGroupReviewStatus = "needs_review"
            ElseIf itemCount > 0 And confidentCount = itemCount Then
                GetBackfillGroupReviewStatus = GetBackfillReadyStatus(backfillTable, rowIndex)
            Else
                GetBackfillGroupReviewStatus = "needs_review"
            End If
    End Select
End Function

Private Function GetBackfillGroupTotal(ByVal backfillTable As ListObject, ByVal packageId As String) As Double
    Dim rowIndex As Long

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            If IsNumeric(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_AMOUNT)) Then
                GetBackfillGroupTotal = GetBackfillGroupTotal + CDbl(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CHILD_AMOUNT))
            End If
        End If
    Next rowIndex
End Function

Private Function GetBackfillGroupCount(ByVal backfillTable As ListObject, ByVal packageId As String) As Long
    Dim rowIndex As Long

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            GetBackfillGroupCount = GetBackfillGroupCount + 1
        End If
    Next rowIndex
End Function

Private Function GetBackfillAmountCheckStatus(ByVal parentAmount As Double, ByVal groupTotal As Double) As String
    If Abs(parentAmount - groupTotal) < 0.01 Then
        GetBackfillAmountCheckStatus = "match"
    ElseIf groupTotal < parentAmount Then
        GetBackfillAmountCheckStatus = "partial"
    Else
        GetBackfillAmountCheckStatus = "mismatch"
    End If
End Function

Private Function GetBackfillGroupOrderHint(ByVal backfillTable As ListObject, ByVal packageId As String) As String
    Dim rowIndex As Long
    Dim currentHint As String

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            currentHint = Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ORDER_HINT)))
            If Len(currentHint) > 0 Then
                GetBackfillGroupOrderHint = currentHint
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Function GetBackfillDistinctOrderHintCount(ByVal backfillTable As ListObject, ByVal packageId As String) As Long
    Dim rowIndex As Long
    Dim currentHint As String
    Dim seenHints As String

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            currentHint = Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ORDER_HINT)))
            If Len(currentHint) > 0 Then
                If InStr(1, "|" & seenHints & "|", "|" & currentHint & "|", vbTextCompare) = 0 Then
                    seenHints = seenHints & "|" & currentHint
                    GetBackfillDistinctOrderHintCount = GetBackfillDistinctOrderHintCount + 1
                End If
            End If
        End If
    Next rowIndex
End Function

Private Function GetBackfillGroupAmbiguousCount(ByVal backfillTable As ListObject, ByVal packageId As String) As Long
    Dim rowIndex As Long
    Dim assetCategoryValue As String

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            assetCategoryValue = Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ASSET_CATEGORY)))
            If Len(assetCategoryValue) = 0 Then GetBackfillGroupAmbiguousCount = GetBackfillGroupAmbiguousCount + 1
        End If
    Next rowIndex
End Function

Private Function GetBackfillGroupConfidentCount(ByVal backfillTable As ListObject, ByVal packageId As String, ByVal minConfidence As Long) As Long
    Dim rowIndex As Long
    Dim currentConfidence As Long

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            If IsNumeric(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CONFIDENCE)) Then
                currentConfidence = CLng(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_CONFIDENCE))
                If currentConfidence >= minConfidence Then GetBackfillGroupConfidentCount = GetBackfillGroupConfidentCount + 1
            End If
        End If
    Next rowIndex
End Function

Private Function BuildBackfillAssetMixSummary(ByVal backfillTable As ListObject, ByVal packageId As String) As String
    Dim inventoryCount As Long
    Dim fixedAssetCount As Long
    Dim ambiguousCount As Long

    inventoryCount = GetBackfillGroupCategoryCount(backfillTable, packageId, "inventory")
    fixedAssetCount = GetBackfillGroupCategoryCount(backfillTable, packageId, "fixed_assets")
    ambiguousCount = GetBackfillGroupCategoryCount(backfillTable, packageId, vbNullString)

    BuildBackfillAssetMixSummary = LocalizationManager.GetText("Inventory") & "=" & inventoryCount & " | " & _
        LocalizationManager.GetText("Fixed assets") & "=" & fixedAssetCount
    If ambiguousCount > 0 Then
        BuildBackfillAssetMixSummary = BuildBackfillAssetMixSummary & " | " & _
            LocalizationManager.GetText("Ambiguous") & "=" & ambiguousCount
    End If
End Function

Private Function GetBackfillReadyStatus(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As String
    If IsBackfillOutgoingRow(backfillTable, rowIndex) Then
        If HasConfirmedBackfillStatus(backfillTable, rowIndex) Then
            GetBackfillReadyStatus = "ready_confirmed"
        Else
            GetBackfillReadyStatus = "ready_outgoing"
        End If
    Else
        GetBackfillReadyStatus = "ready_incoming"
    End If
End Function

Private Function GetBackfillTargetStageDisplayText(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As String
    Dim statusValue As String

    statusValue = LCase$(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS))))
    Select Case statusValue
        Case "ready_confirmed"
            GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Confirmed by counterparty")
        Case "ready_outgoing"
            GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Awaiting confirmation")
        Case "ready_incoming"
            GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Matched in 1C")
        Case "partial_group"
            GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Partially matched")
        Case "amount_mismatch", "needs_review", "failed"
            GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Needs review")
        Case "applied"
            GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Applied")
        Case "duplicate"
            GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Duplicate")
        Case "closed_package"
            GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Closed package")
        Case Else
            If IsBackfillOutgoingRow(backfillTable, rowIndex) Then
                GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Awaiting confirmation")
            Else
                GetBackfillTargetStageDisplayText = LocalizationManager.GetText("Matched in 1C")
            End If
    End Select
End Function

Private Function IsBackfillOutgoingRow(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As Boolean
    Dim directionKey As String

    directionKey = Replace$(LCase$(Trim$(GetBackfillParentDirectionKey(backfillTable, rowIndex))), ".", vbNullString)
    IsBackfillOutgoingRow = (InStr(directionKey, "out") > 0) _
        Or (InStr(directionKey, "исх") > 0) _
        Or (directionKey = "èñõ")
End Function

Private Function HasConfirmedBackfillStatus(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As Boolean
    Dim confirmationStatus As String
    Dim normalizedStatus As String

    confirmationStatus = GetBackfillParentConfirmationStatus(backfillTable, rowIndex)
    normalizedStatus = UCase$(Trim$(confirmationStatus))
    If Len(normalizedStatus) = 0 Then Exit Function

    If InStr(normalizedStatus, "CONFIRM") > 0 Then
        HasConfirmedBackfillStatus = True
    ElseIf InStr(normalizedStatus, ChrW$(1055) & ChrW$(1054) & ChrW$(1044) & ChrW$(1058) & ChrW$(1042) & ChrW$(1045) & ChrW$(1056)) > 0 Then
        HasConfirmedBackfillStatus = True
    End If
End Function

Private Function GetBackfillParentDirectionKey(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As String
    Dim parentTable As ListObject
    Dim parentRowIndex As Long

    Set parentTable = CommonUtilities.GetListObjectSafe(CommonUtilities.GetWorksheetSafe("IncOut"), PACKAGE_PARENT_TABLE_NAME)
    If parentTable Is Nothing Then Exit Function

    parentRowIndex = CLng(Val(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_INCOUT_ROW))))
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function

    GetBackfillParentDirectionKey = CStr(parentTable.DataBodyRange.Cells(parentRowIndex, 3).Value)
End Function

Private Function GetBackfillParentConfirmationStatus(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As String
    Dim parentTable As ListObject
    Dim parentRowIndex As Long

    Set parentTable = CommonUtilities.GetListObjectSafe(CommonUtilities.GetWorksheetSafe("IncOut"), PACKAGE_PARENT_TABLE_NAME)
    If parentTable Is Nothing Then Exit Function

    parentRowIndex = CLng(Val(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_INCOUT_ROW))))
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function

    GetBackfillParentConfirmationStatus = CStr(parentTable.DataBodyRange.Cells(parentRowIndex, 19).Value)
End Function

Private Function GetBackfillAssetCategoryDisplayText(ByVal assetCategoryValue As String) As String
    Select Case LCase$(Trim$(assetCategoryValue))
        Case "inventory"
            GetBackfillAssetCategoryDisplayText = LocalizationManager.GetText("Inventory")
        Case "fixed_assets"
            GetBackfillAssetCategoryDisplayText = LocalizationManager.GetText("Fixed assets")
        Case Else
            GetBackfillAssetCategoryDisplayText = Trim$(assetCategoryValue)
    End Select
End Function

Private Function GetBackfillReviewStatusDisplayText(ByVal statusValue As String) As String
    Select Case LCase$(Trim$(statusValue))
        Case "pending"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Pending")
        Case "ready_incoming"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Ready for 1C closure")
        Case "ready_outgoing"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Ready for confirmation")
        Case "ready_confirmed"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Ready and confirmed")
        Case "needs_review"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Needs review")
        Case "partial_group"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Partial group")
        Case "amount_mismatch"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Amount mismatch")
        Case "applied"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Applied")
        Case "duplicate"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Duplicate")
        Case "closed_package"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Closed package")
        Case "failed"
            GetBackfillReviewStatusDisplayText = LocalizationManager.GetText("Failed")
        Case Else
            GetBackfillReviewStatusDisplayText = Trim$(statusValue)
    End Select
End Function

Private Function GetBackfillAmountCheckDisplayText(ByVal amountStatus As String) As String
    Select Case LCase$(Trim$(amountStatus))
        Case "match"
            GetBackfillAmountCheckDisplayText = LocalizationManager.GetText("Amount matches")
        Case "partial"
            GetBackfillAmountCheckDisplayText = LocalizationManager.GetText("Partial group")
        Case "mismatch"
            GetBackfillAmountCheckDisplayText = LocalizationManager.GetText("Amount mismatch")
        Case Else
            GetBackfillAmountCheckDisplayText = Trim$(amountStatus)
    End Select
End Function

Private Function GetBackfillGroupCategoryCount(ByVal backfillTable As ListObject, ByVal packageId As String, ByVal assetCategoryValue As String) As Long
    Dim rowIndex As Long
    Dim currentCategory As String

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            currentCategory = Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_ASSET_CATEGORY)))
            If Len(assetCategoryValue) = 0 Then
                If Len(currentCategory) = 0 Then GetBackfillGroupCategoryCount = GetBackfillGroupCategoryCount + 1
            ElseIf StrComp(currentCategory, assetCategoryValue, vbTextCompare) = 0 Then
                GetBackfillGroupCategoryCount = GetBackfillGroupCategoryCount + 1
            End If
        End If
    Next rowIndex
End Function

Private Function IsTerminalBackfillStatus(ByVal statusValue As String) As Boolean
    Select Case LCase$(Trim$(statusValue))
        Case "applied", "duplicate", "closed_package"
            IsTerminalBackfillStatus = True
    End Select
End Function

Private Function IsPendingBackfillRow(ByVal backfillTable As ListObject, ByVal rowIndex As Long) As Boolean
    Dim statusValue As String

    statusValue = LCase$(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS))))
    Select Case statusValue
        Case "pending", "failed", "ready_incoming", "ready_outgoing", "ready_confirmed", "needs_review", "partial_group", "amount_mismatch"
            IsPendingBackfillRow = True
    End Select
End Function

Private Function FindExistingBackfillQueueRow(ByVal backfillTable As ListObject, ByVal packageId As String, ByVal operationNumber As String) As Long
    Dim rowIndex As Long

    If backfillTable Is Nothing Then Exit Function
    If backfillTable.DataBodyRange Is Nothing Then Exit Function
    If Len(Trim$(operationNumber)) = 0 Then Exit Function

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_OPERATION_NUMBER))), Trim$(operationNumber), vbTextCompare) = 0 Then
                FindExistingBackfillQueueRow = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Function ShouldSkipParentForBackfill(ByVal parentTable As ListObject, ByVal parentRowIndex As Long, ByVal packageId As String) As Boolean
    Dim childCount As Long
    Dim pendingCount As Long
    Dim amountStatus As String
    Dim stageValue As String

    If parentTable Is Nothing Then Exit Function
    If Len(Trim$(packageId)) = 0 Then Exit Function

    childCount = PackageDocumentsManager.GetPackageChildDocumentCount(parentRowIndex)
    If childCount <= 0 Then Exit Function

    pendingCount = PackageDocumentsManager.CountPendingPackageChildMatches(parentRowIndex)
    amountStatus = LCase$(Trim$(CStr(GetParentPackageValue(parentTable, parentRowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS))))
    stageValue = LCase$(Trim$(CStr(GetParentPackageValue(parentTable, parentRowIndex, PACKAGE_COLUMN_DOCUMENT_STAGE))))

    If amountStatus <> "match" Then Exit Function
    If pendingCount > 0 Then Exit Function

    Select Case stageValue
        Case "matched_in_1c", "awaiting_confirmation", "confirmed_by_counterparty"
            ShouldSkipParentForBackfill = True
    End Select
End Function

Private Function GetParentPackageValue(ByVal parentTable As ListObject, ByVal parentRowIndex As Long, ByVal columnName As String) As Variant
    Dim targetColumn As ListColumn

    If parentTable Is Nothing Then Exit Function
    If parentTable.DataBodyRange Is Nothing Then Exit Function
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function

    Set targetColumn = CommonUtilities.GetListColumnSafe(parentTable, columnName)
    If targetColumn Is Nothing Then Exit Function

    GetParentPackageValue = parentTable.DataBodyRange.Cells(parentRowIndex, targetColumn.Index).Value
End Function

Private Sub MarkBackfillRowsClosedForPackage(ByVal backfillTable As ListObject, ByVal packageId As String)
    Dim rowIndex As Long

    If backfillTable Is Nothing Then Exit Sub
    If backfillTable.DataBodyRange Is Nothing Then Exit Sub
    If Len(Trim$(packageId)) = 0 Then Exit Sub

    For rowIndex = 1 To backfillTable.ListRows.Count
        If StrComp(Trim$(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
            If Not IsTerminalBackfillStatus(CStr(GetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS))) Then
                Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_REVIEW_STATUS, "closed_package")
                Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_USE_PROPOSAL, False)
                Call SetBackfillValue(backfillTable, rowIndex, BACKFILL_COLUMN_APPLY_ERROR, vbNullString)
            End If
        End If
    Next rowIndex
End Sub

Private Function GetLegacyPackageBackfillTable() As ListObject
    Dim ws As Worksheet

    Set ws = CommonUtilities.GetWorksheetSafe(BACKFILL_SHEET_NAME)
    If ws Is Nothing Then Exit Function

    Set GetLegacyPackageBackfillTable = CommonUtilities.GetListObjectSafe(ws, BACKFILL_TABLE_NAME)
End Function

Private Sub ClearLegacyPackageBackfillQueue(ByVal backfillTable As ListObject)
    On Error Resume Next
    If Not backfillTable Is Nothing Then
        If Not backfillTable.DataBodyRange Is Nothing Then backfillTable.DataBodyRange.Rows.Delete
    End If
    On Error GoTo 0
End Sub

Private Sub CreateLegacyPackageBackfillTable(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim columnCount As Long
    Dim index As Long
    Dim tableRange As Range
    Dim backfillTable As ListObject

    headers = GetLegacyPackageBackfillHeaders()
    columnCount = UBound(headers) - LBound(headers) + 1

    For index = LBound(headers) To UBound(headers)
        ws.Cells(1, index - LBound(headers) + 1).Value = CStr(headers(index))
    Next index

    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(2, columnCount))
    Set backfillTable = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    backfillTable.Name = BACKFILL_TABLE_NAME

    On Error Resume Next
    If Not backfillTable.DataBodyRange Is Nothing Then backfillTable.DataBodyRange.Rows.Delete
    On Error GoTo 0
End Sub

Private Sub EnsureLegacyPackageBackfillColumns(ByVal backfillTable As ListObject)
    Dim headers As Variant
    Dim index As Long

    headers = GetLegacyPackageBackfillHeaders()
    For index = LBound(headers) To UBound(headers)
        Call EnsureLegacyPackageBackfillColumn(backfillTable, CStr(headers(index)))
    Next index
End Sub

Private Function GetLegacyPackageBackfillHeaders() As Variant
    Dim headers(0 To 29) As String

    headers(0) = BACKFILL_COLUMN_QUEUE_ID
    headers(1) = BACKFILL_COLUMN_INCOUT_ROW
    headers(2) = BACKFILL_COLUMN_PACKAGE_ID
    headers(3) = BACKFILL_COLUMN_RECORD_NUMBER
    headers(4) = BACKFILL_COLUMN_PARENT_TYPE
    headers(5) = BACKFILL_COLUMN_PARENT_NUMBER
    headers(6) = BACKFILL_COLUMN_PARENT_DATE
    headers(7) = BACKFILL_COLUMN_PARENT_AMOUNT
    headers(8) = BACKFILL_COLUMN_COUNTERPARTY
    headers(9) = BACKFILL_COLUMN_OPERATION_TYPE
    headers(10) = BACKFILL_COLUMN_OPERATION_NUMBER
    headers(11) = BACKFILL_COLUMN_OPERATION_DATE
    headers(12) = BACKFILL_COLUMN_OPERATION_AMOUNT
    headers(13) = BACKFILL_COLUMN_OPERATION_COMMENT
    headers(14) = BACKFILL_COLUMN_CHILD_TYPE
    headers(15) = BACKFILL_COLUMN_ASSET_CATEGORY
    headers(16) = BACKFILL_COLUMN_CHILD_NUMBER
    headers(17) = BACKFILL_COLUMN_CHILD_DATE
    headers(18) = BACKFILL_COLUMN_CHILD_AMOUNT
    headers(19) = BACKFILL_COLUMN_CHILD_DESCRIPTION
    headers(20) = BACKFILL_COLUMN_ORDER_HINT
    headers(21) = BACKFILL_COLUMN_GROUP_TOTAL
    headers(22) = BACKFILL_COLUMN_GROUP_COUNT
    headers(23) = BACKFILL_COLUMN_GROUP_AMOUNT_STATUS
    headers(24) = BACKFILL_COLUMN_CONFIDENCE
    headers(25) = BACKFILL_COLUMN_USE_PROPOSAL
    headers(26) = BACKFILL_COLUMN_REVIEW_STATUS
    headers(27) = BACKFILL_COLUMN_APPLIED_ITEM_ID
    headers(28) = BACKFILL_COLUMN_APPLIED_AT
    headers(29) = BACKFILL_COLUMN_APPLY_ERROR

    GetLegacyPackageBackfillHeaders = headers
End Function

Private Sub EnsureLegacyPackageBackfillColumn(ByVal backfillTable As ListObject, ByVal columnName As String)
    If backfillTable Is Nothing Then Exit Sub
    If Len(Trim$(columnName)) = 0 Then Exit Sub

    If CommonUtilities.GetListColumnSafe(backfillTable, columnName) Is Nothing Then
        backfillTable.ListColumns.Add.Name = columnName
    End If
End Sub

Private Sub FormatLegacyPackageBackfillSheet(ByVal ws As Worksheet, ByVal backfillTable As ListObject)
    Dim counterpartyColumn As ListColumn
    Dim operationCommentColumn As ListColumn
    Dim descriptionColumn As ListColumn

    If ws Is Nothing Or backfillTable Is Nothing Then Exit Sub

    Set counterpartyColumn = CommonUtilities.GetListColumnSafe(backfillTable, BACKFILL_COLUMN_COUNTERPARTY)
    Set operationCommentColumn = CommonUtilities.GetListColumnSafe(backfillTable, BACKFILL_COLUMN_OPERATION_COMMENT)
    Set descriptionColumn = CommonUtilities.GetListColumnSafe(backfillTable, BACKFILL_COLUMN_CHILD_DESCRIPTION)

    ws.Cells.WrapText = False
    ws.Cells.WrapText = False
    ws.Columns("A:AD").EntireColumn.AutoFit

    If Not counterpartyColumn Is Nothing Then ws.Columns(counterpartyColumn.Range.Column).ColumnWidth = 28
    If Not operationCommentColumn Is Nothing Then ws.Columns(operationCommentColumn.Range.Column).ColumnWidth = 58
    If Not descriptionColumn Is Nothing Then ws.Columns(descriptionColumn.Range.Column).ColumnWidth = 58
    If Not CommonUtilities.GetListColumnSafe(backfillTable, BACKFILL_COLUMN_ASSET_CATEGORY) Is Nothing Then
        ws.Columns(CommonUtilities.GetListColumnSafe(backfillTable, BACKFILL_COLUMN_ASSET_CATEGORY).Range.Column).ColumnWidth = 16
    End If
    If Not CommonUtilities.GetListColumnSafe(backfillTable, BACKFILL_COLUMN_ORDER_HINT) Is Nothing Then
        ws.Columns(CommonUtilities.GetListColumnSafe(backfillTable, BACKFILL_COLUMN_ORDER_HINT).Range.Column).ColumnWidth = 28
    End If

    If Not operationCommentColumn Is Nothing Then
        If Not operationCommentColumn.DataBodyRange Is Nothing Then operationCommentColumn.DataBodyRange.WrapText = True
    End If
    If Not descriptionColumn Is Nothing Then
        If Not descriptionColumn.DataBodyRange Is Nothing Then descriptionColumn.DataBodyRange.WrapText = True
    End If
    ws.Activate
End Sub

Private Function GetBackfillValue(ByVal backfillTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim listColumn As ListColumn

    Set listColumn = CommonUtilities.GetListColumnSafe(backfillTable, columnName)
    If listColumn Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > backfillTable.ListRows.Count Then Exit Function

    GetBackfillValue = backfillTable.DataBodyRange.Cells(rowIndex, listColumn.Index).Value
End Function

Private Sub SetBackfillValue(ByVal backfillTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueToWrite As Variant)
    Dim listColumn As ListColumn

    Set listColumn = CommonUtilities.GetListColumnSafe(backfillTable, columnName)
    If listColumn Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > backfillTable.ListRows.Count Then Exit Sub

    backfillTable.DataBodyRange.Cells(rowIndex, listColumn.Index).Value = valueToWrite
End Sub

Private Function CreateLegacyPackageBackfillId() As String
    Randomize
    CreateLegacyPackageBackfillId = "LPB-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int((Rnd() * 9000) + 1000), "0000")
End Function

Private Function IsTrueMarker(ByVal sourceValue As Variant) As Boolean
    Dim normalizedText As String

    normalizedText = LCase$(Trim$(CStr(sourceValue)))
    IsTrueMarker = (normalizedText = "true" Or normalizedText = "1" Or normalizedText = "yes")
End Function
