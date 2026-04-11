Attribute VB_Name = "PackageDocumentsManager"
Option Explicit

Private Const PARENT_SOURCE_SERVICE_COLUMN As Long = 2
Private Const PARENT_SOURCE_DIRECTION_COLUMN As Long = 3
Private Const PARENT_SOURCE_DOCUMENT_TYPE_COLUMN As Long = 4
Private Const PARENT_SOURCE_DOCUMENT_NUMBER_COLUMN As Long = 5
Private Const PARENT_SOURCE_AMOUNT_COLUMN As Long = 6
Private Const PARENT_SOURCE_FRP_NUMBER_COLUMN As Long = 7
Private Const PARENT_SOURCE_FRP_DATE_COLUMN As Long = 8
Private Const PARENT_SOURCE_COUNTERPARTY_COLUMN As Long = 9
Private Const PARENT_SOURCE_EXECUTOR_COLUMN As Long = 11
Private Const PARENT_SOURCE_ORDER_INFO_COLUMN As Long = 20

Private Const ITEM_COLUMN_ITEM_ID As String = "ItemId"
Private Const ITEM_COLUMN_PACKAGE_ID As String = "PackageId"
Private Const ITEM_COLUMN_ITEM_ORDER As String = "ItemOrder"
Private Const ITEM_COLUMN_DOCUMENT_TYPE As String = "ItemDocumentType"
Private Const ITEM_COLUMN_DOCUMENT_TYPE_DISPLAY As String = "ItemDocumentTypeDisplay"
Private Const ITEM_COLUMN_DOCUMENT_NUMBER As String = "ItemDocumentNumber"
Private Const ITEM_COLUMN_DOCUMENT_DATE As String = "ItemDocumentDate"
Private Const ITEM_COLUMN_AMOUNT As String = "ItemAmount"
Private Const ITEM_COLUMN_COUNTERPARTY_NAME As String = "CounterpartyName"
Private Const ITEM_COLUMN_COUNTERPARTY_NORMALIZED As String = "CounterpartyNormalized"
Private Const ITEM_COLUMN_DIRECTION As String = "Direction"
Private Const ITEM_COLUMN_SERVICE As String = "Service"
Private Const ITEM_COLUMN_EXECUTOR As String = "Executor"
Private Const ITEM_COLUMN_ORDER_INFO As String = "OrderInfo"
Private Const ITEM_COLUMN_FRP_NUMBER As String = "FRPNumber"
Private Const ITEM_COLUMN_FRP_DATE As String = "FRPDate"
Private Const ITEM_COLUMN_ASSET_CATEGORY As String = "ItemAssetCategory"
Private Const ITEM_COLUMN_QUANTITY As String = "ItemQuantity"
Private Const ITEM_COLUMN_UNIT As String = "ItemUnit"
Private Const ITEM_COLUMN_DESCRIPTION As String = "ItemDescription"
Private Const ITEM_COLUMN_BASE_DOCUMENT_TYPE As String = "BaseDocumentType"
Private Const ITEM_COLUMN_BASE_DOCUMENT_NUMBER As String = "BaseDocumentNumber"
Private Const ITEM_COLUMN_BASE_DOCUMENT_DATE As String = "BaseDocumentDate"
Private Const ITEM_COLUMN_MATCHED_OPERATION_NUMBER As String = "Matched1COperationNumber"
Private Const ITEM_COLUMN_MATCHED_OPERATION_DATE As String = "Matched1COperationDate"
Private Const ITEM_COLUMN_MATCHED_STATUS As String = "Matched1CMatchStatus"
Private Const ITEM_COLUMN_MATCHED_CONFIDENCE As String = "Matched1CConfidence"
Private Const ITEM_COLUMN_MATCHED_COMMENT As String = "Matched1CComment"
Private Const ITEM_COLUMN_MATCHED_MODE As String = "Matched1CMode"
Private Const ITEM_COLUMN_IS_POSTED_SEPARATELY As String = "IsPostedSeparately"
Private Const ITEM_COLUMN_CREATED_AT As String = "CreatedAt"
Private Const ITEM_COLUMN_UPDATED_AT As String = "UpdatedAt"
Private Const ITEM_COLUMN_NOTES As String = "Notes"

Public Sub OpenPackageDocumentsForParentRow(ByVal parentRowIndex As Long)
    On Error GoTo OpenError

    Dim parentTable As ListObject
    Dim packageId As String

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Sub
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then
        MsgBox LocalizationManager.GetText("Unable to determine the current record row."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    packageId = EnsurePackageIdForParentRow(parentTable, parentRowIndex)
    If Len(packageId) = 0 Then
        MsgBox LocalizationManager.GetText("Unable to determine PackageId for the current record."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    Load UserFormPackageDocuments
    UserFormPackageDocuments.OpenForParentRow parentRowIndex, packageId
    Exit Sub

OpenError:
    MsgBox LocalizationManager.GetText("Error opening package documents: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Public Function PackageHasChildDocuments(ByVal parentRowIndex As Long) As Boolean
    Dim parentTable As ListObject

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Function
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function

    PackageHasChildDocuments = (LCase$(Trim$(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_HAS_CHILD_DOCUMENTS))) = "true")
End Function

Public Function GetPackageChildDocumentCount(ByVal parentRowIndex As Long) As Long
    Dim parentTable As ListObject
    Dim childCountText As String

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Function
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function

    childCountText = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT)
    If IsNumeric(childCountText) Then
        GetPackageChildDocumentCount = CLng(childCountText)
    End If
End Function

Public Function ShouldUseChildDocumentsForMatching(ByVal parentRowIndex As Long) As Boolean
    ShouldUseChildDocumentsForMatching = (GetPackageChildDocumentCount(parentRowIndex) > 0)
End Function

Public Function CountPendingPackageChildMatches(ByVal parentRowIndex As Long) As Long
    Dim parentTable As ListObject
    Dim itemsTable As ListObject
    Dim packageId As String
    Dim rowIndex As Long

    Set parentTable = GetParentTable()
    Set itemsTable = GetItemsTable()
    If parentTable Is Nothing Or itemsTable Is Nothing Then Exit Function
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function
    If itemsTable.DataBodyRange Is Nothing Then Exit Function

    packageId = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PACKAGE_ID)
    If Len(packageId) = 0 Then Exit Function

    For rowIndex = 1 To itemsTable.ListRows.Count
        If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            If IsChildMatchPending(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS))) Then
                CountPendingPackageChildMatches = CountPendingPackageChildMatches + 1
            End If
        End If
    Next rowIndex
End Function

Public Function GetPackagePrimary1CStatus(ByVal parentRowIndex As Long) As String
    Dim parentTable As ListObject

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Function
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function

    GetPackagePrimary1CStatus = LCase$(Trim$(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_STATUS)))
End Function

Public Function GetPackagePrimary1COperationNumber(ByVal parentRowIndex As Long) As String
    Dim parentTable As ListObject

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Function
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function

    GetPackagePrimary1COperationNumber = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_NUMBER)
End Function

Public Sub ProcessPackageChildMatches(ByVal parentRowIndex As Long, ByVal ws1C As Worksheet, ByRef processedCount As Long, ByRef foundCount As Long, ByRef candidateCount As Long, ByRef notFoundCount As Long)
    Dim parentTable As ListObject
    Dim itemsTable As ListObject
    Dim packageId As String
    Dim rowIndex As Long
    Dim childAmount As Double
    Dim childCorrespondent As String
    Dim childFound As Boolean
    Dim childProvodkaNumber As String
    Dim childProvodkaDate As Variant
    Dim childMatchCount As Long
    Dim childStatusMessage As String
    Dim childCandidatesList As String

    Set parentTable = GetParentTable()
    Set itemsTable = GetItemsTable()
    If parentTable Is Nothing Or itemsTable Is Nothing Then Exit Sub
    If ws1C Is Nothing Then Exit Sub
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Sub
    If itemsTable.DataBodyRange Is Nothing Then Exit Sub

    packageId = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PACKAGE_ID)
    If Len(packageId) = 0 Then Exit Sub

    For rowIndex = 1 To itemsTable.ListRows.Count
        If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
            If IsChildMatchPending(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS))) Then
                childAmount = 0
                If IsNumeric(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_AMOUNT)) Then
                    childAmount = CDbl(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_AMOUNT))
                End If

                childCorrespondent = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_COUNTERPARTY_NAME))
                If Len(Trim$(childCorrespondent)) = 0 Then
                    childCorrespondent = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_COUNTERPARTY_NORMALIZED))
                End If

                Call ProvodkaIntegrationModule.FindMatchDetailsInFile(childAmount, childCorrespondent, ws1C, childFound, childProvodkaNumber, childProvodkaDate, childMatchCount, childStatusMessage, childCandidatesList)
                processedCount = processedCount + 1
                Call ApplyChildMatchResult(itemsTable, rowIndex, childFound, childProvodkaNumber, childProvodkaDate, childMatchCount, childStatusMessage, childCandidatesList)

                If childFound Then
                    foundCount = foundCount + 1
                ElseIf childMatchCount > 1 Then
                    candidateCount = candidateCount + 1
                Else
                    notFoundCount = notFoundCount + 1
                End If
            End If
        End If
    Next rowIndex

    Call RefreshParentPackageSummary(parentRowIndex)
End Sub

Public Sub RefreshPackageIndicatorsOnMainForm(ByVal frm As Object, ByVal parentRowIndex As Long)
    Dim parentTable As ListObject
    Dim childCount As String
    Dim childrenTotal As String
    Dim statusValue As String
    Dim statusText As String
    Dim primaryStatusValue As String
    Dim primaryStatusText As String
    Dim primaryOperationNumber As String
    Dim reviewSummaryText As String

    On Error GoTo IndicatorError

    If frm Is Nothing Then Exit Sub
    If parentRowIndex < 1 Then
        Call ClearPackageIndicatorsOnMainForm(frm)
        Exit Sub
    End If

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then
        Call ClearPackageIndicatorsOnMainForm(frm)
        Exit Sub
    End If
    If parentRowIndex > parentTable.ListRows.Count Then
        Call ClearPackageIndicatorsOnMainForm(frm)
        Exit Sub
    End If

    childCount = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT)
    If Len(childCount) = 0 Then childCount = "0"

    childrenTotal = FormatItemAmountValue(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT))
    statusValue = LCase$(Trim$(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS)))
    statusText = TranslateAmountCheckStatus(statusValue)
    primaryStatusValue = LCase$(Trim$(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_STATUS)))
    primaryStatusText = TranslatePrimaryMatchStatus(primaryStatusValue)
    primaryOperationNumber = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_NUMBER)
    reviewSummaryText = BuildPackageReviewSummary(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PACKAGE_ID))

    frm.lblPackageIndicators.Caption = LocalizationManager.GetText("Items:") & " " & childCount & " | " & _
        LocalizationManager.GetText("Children Total:") & " " & childrenTotal & vbCrLf & _
        LocalizationManager.GetText("Amount Check:") & " " & statusText & " | " & _
        LocalizationManager.GetText("1C Status") & ": " & primaryStatusText & IIf(Len(Trim$(primaryOperationNumber)) > 0, " | " & LocalizationManager.GetText("1C Operation No.") & ": " & primaryOperationNumber, vbNullString) & vbCrLf & _
        reviewSummaryText
        frm.lblPackageIndicators.Height = 54

    Select Case statusValue
        Case "match"
            frm.lblPackageIndicators.ForeColor = RGB(0, 102, 51)
        Case "mismatch"
            frm.lblPackageIndicators.ForeColor = RGB(156, 0, 6)
        Case Else
            frm.lblPackageIndicators.ForeColor = RGB(96, 96, 96)
    End Select
    Exit Sub

IndicatorError:
    Call ClearPackageIndicatorsOnMainForm(frm)
End Sub

Public Sub ClearPackageIndicatorsOnMainForm(ByVal frm As Object)
    On Error Resume Next
    If frm Is Nothing Then Exit Sub
    frm.lblPackageIndicators.Caption = ""
    frm.lblPackageIndicators.Height = 54
    frm.lblPackageIndicators.ForeColor = RGB(96, 96, 96)
End Sub

Public Function IsParentPackageAmountMismatch(ByVal parentRowIndex As Long) As Boolean
    Dim parentTable As ListObject
    Dim statusValue As String

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Function
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Function

    statusValue = LCase$(Trim$(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS)))
    IsParentPackageAmountMismatch = (statusValue = "mismatch")
End Function

Public Sub BindPackageDocumentsForm(ByVal frm As Object, ByVal parentRowIndex As Long, ByVal packageId As String)
    Call LoadParentSummaryToForm(frm, parentRowIndex, packageId)
    Call LoadPackageItemsToList(frm, packageId, GetActiveReviewFilter(frm))
    Call ClearPackageItemEditor(frm)
End Sub

Public Sub LoadPackageItemsToList(ByVal frm As Object, ByVal packageId As String, Optional ByVal reviewFilter As String = "")
    Dim itemsTable As ListObject
    Dim dataRow As Range
    Dim listIndex As Long
    Dim statusValue As String

    Set itemsTable = GetItemsTable()
    frm.lstPackageItems.Clear

    If itemsTable Is Nothing Then Exit Sub
    If itemsTable.DataBodyRange Is Nothing Then Exit Sub

    For Each dataRow In itemsTable.DataBodyRange.rows
        If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
            statusValue = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_MATCHED_STATUS))
            If Not IsItemVisibleForReviewFilter(statusValue, reviewFilter) Then GoTo ContinueLoop
            frm.lstPackageItems.AddItem CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_ITEM_ORDER))
            listIndex = frm.lstPackageItems.listCount - 1
            frm.lstPackageItems.List(listIndex, 1) = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_DOCUMENT_TYPE_DISPLAY))
            frm.lstPackageItems.List(listIndex, 2) = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_DOCUMENT_NUMBER))
            frm.lstPackageItems.List(listIndex, 3) = FormatItemDateValue(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_DOCUMENT_DATE))
            frm.lstPackageItems.List(listIndex, 4) = FormatItemAmountValue(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_AMOUNT))
            frm.lstPackageItems.List(listIndex, 5) = statusValue
            frm.lstPackageItems.List(listIndex, 6) = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_MATCHED_OPERATION_NUMBER))
            frm.lstPackageItems.List(listIndex, 7) = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_ITEM_ID))
        End If
ContinueLoop:
    Next dataRow
End Sub

Public Sub ApplyPackageReviewFilterFromForm(ByVal frm As Object, ByVal packageId As String)
    Call LoadPackageItemsToList(frm, packageId, GetActiveReviewFilter(frm))
    Call ClearPackageItemEditor(frm)
End Sub

Public Sub SetPackageReviewFilterFromForm(ByVal frm As Object, ByVal packageId As String, ByVal filterKey As String)
    If frm Is Nothing Then Exit Sub

    frm.cmbReviewFilter.value = LocalizationManager.GetText(filterKey)
    Call ApplyPackageReviewFilterFromForm(frm, packageId)
End Sub

Public Sub SelectNextReviewItemFromForm(ByVal frm As Object)
    Dim listIndex As Long
    Dim startIndex As Long

    If frm Is Nothing Then Exit Sub
    If frm.lstPackageItems.listCount = 0 Then Exit Sub

    startIndex = frm.lstPackageItems.listIndex + 1
    If startIndex < 0 Then startIndex = 0

    For listIndex = startIndex To frm.lstPackageItems.listCount - 1
        If IsReviewTargetStatus(CStr(frm.lstPackageItems.List(listIndex, 5))) Then
            frm.lstPackageItems.listIndex = listIndex
            Call LoadSelectedPackageItemIntoForm(frm, vbNullString)
            Exit Sub
        End If
    Next listIndex

    For listIndex = 0 To startIndex - 1
        If IsReviewTargetStatus(CStr(frm.lstPackageItems.List(listIndex, 5))) Then
            frm.lstPackageItems.listIndex = listIndex
            Call LoadSelectedPackageItemIntoForm(frm, vbNullString)
            Exit Sub
        End If
    Next listIndex

    MsgBox LocalizationManager.GetText("No more review items in the current list."), vbInformation, LocalizationManager.GetText("Package Documents")
End Sub

Public Sub MarkSelectedPackageItemManualFromForm(ByVal frm As Object, ByVal parentRowIndex As Long, ByVal packageId As String)
    Dim itemId As String

    itemId = GetSelectedPackageItemIdFromForm(frm)
    If Len(itemId) = 0 Then
        MsgBox LocalizationManager.GetText("Select a package document item first."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    Call UpdatePackageItemMatchState(itemId, "manual", True)
    Call RefreshParentPackageSummary(parentRowIndex)
    Call BindPackageDocumentsForm(frm, parentRowIndex, packageId)
    Call SelectPackageItemInList(frm, itemId)
End Sub

Public Sub ResetSelectedPackageItemMatchFromForm(ByVal frm As Object, ByVal parentRowIndex As Long, ByVal packageId As String)
    Dim itemId As String

    itemId = GetSelectedPackageItemIdFromForm(frm)
    If Len(itemId) = 0 Then
        MsgBox LocalizationManager.GetText("Select a package document item first."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    Call UpdatePackageItemMatchState(itemId, "not_checked", False)
    Call RefreshParentPackageSummary(parentRowIndex)
    Call BindPackageDocumentsForm(frm, parentRowIndex, packageId)
    Call SelectPackageItemInList(frm, itemId)
End Sub

Public Sub LoadSelectedPackageItemIntoForm(ByVal frm As Object, ByVal packageId As String)
    On Error GoTo LoadError

    Dim selectedItemId As String
    Dim itemsTable As ListObject
    Dim rowIndex As Long

    If frm.lstPackageItems.listIndex < 0 Then Exit Sub
    selectedItemId = CStr(frm.lstPackageItems.List(frm.lstPackageItems.listIndex, 7))
    If Len(selectedItemId) = 0 Then Exit Sub

    Set itemsTable = GetItemsTable()
    rowIndex = FindItemRowIndexById(itemsTable, selectedItemId)
    If rowIndex = 0 Then Exit Sub

    frm.txtItemId.Text = selectedItemId
    frm.cmbItemDocumentTypeDisplay.value = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_TYPE_DISPLAY))
    frm.txtItemDocumentNumber.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_NUMBER))
    frm.txtItemDocumentDate.Text = FormatItemDateValue(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_DATE))
    frm.txtItemAmount.Text = FormatEditorAmountValue(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_AMOUNT))
    frm.txtItemDescription.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DESCRIPTION))
    frm.txtMatched1COperationNumber.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER))
    frm.txtMatched1COperationDate.Text = FormatItemDateValue(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE))
    frm.cmbMatched1CStatus.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS))
    frm.txtMatched1CComment.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_COMMENT))
    Call UpdatePackageItemReviewHint(frm)
    Exit Sub

LoadError:
    MsgBox LocalizationManager.GetText("Error loading package document item: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Public Sub SavePackageItemFromForm(ByVal frm As Object, ByVal parentRowIndex As Long, ByVal packageId As String, ByVal updateExisting As Boolean)
    On Error GoTo SaveError

    Dim amountValue As Double
    Dim itemDateValue As Variant
    Dim itemId As String
    Dim matchedOperationDateValue As Variant

    If Not ValidatePackageItemForm(frm, amountValue, itemDateValue) Then Exit Sub
    If Not ValidateMatchedFields(frm, matchedOperationDateValue) Then Exit Sub

    itemId = SavePackageItemRecord( _
        parentRowIndex, _
        packageId, _
        Trim$(frm.txtItemId.Text), _
        Trim$(frm.cmbItemDocumentTypeDisplay.value), _
        Trim$(frm.txtItemDocumentNumber.Text), _
        itemDateValue, _
        amountValue, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        Trim$(frm.txtItemDescription.Text), _
        vbNullString, _
        Trim$(frm.txtMatched1COperationNumber.Text), _
        matchedOperationDateValue, _
        Trim$(frm.cmbMatched1CStatus.Text), _
        Trim$(frm.txtMatched1CComment.Text), _
        updateExisting)

    If Len(itemId) = 0 Then Exit Sub

    Call BindPackageDocumentsForm(frm, parentRowIndex, packageId)

    If updateExisting Then
        MsgBox LocalizationManager.GetText("Package document item updated."), vbInformation, LocalizationManager.GetText("Package Documents")
    Else
        MsgBox LocalizationManager.GetText("Package document item added."), vbInformation, LocalizationManager.GetText("Package Documents")
    End If
    Exit Sub

SaveError:
    MsgBox LocalizationManager.GetText("Error saving package document item: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Public Sub DeleteSelectedPackageItemFromForm(ByVal frm As Object, ByVal parentRowIndex As Long, ByVal packageId As String)
    On Error GoTo DeleteError

    Dim itemId As String

    itemId = Trim$(frm.txtItemId.Text)
    If Len(itemId) = 0 Then
        MsgBox LocalizationManager.GetText("Select a package document item first."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    If MsgBox(LocalizationManager.GetText("Delete the selected package document item?"), vbYesNo + vbQuestion, LocalizationManager.GetText("Package Documents")) <> vbYes Then Exit Sub

    If Not DeletePackageItemRecord(packageId, itemId) Then Exit Sub

    Call RefreshParentPackageSummary(parentRowIndex)
    Call BindPackageDocumentsForm(frm, parentRowIndex, packageId)

    MsgBox LocalizationManager.GetText("Package document item deleted."), vbInformation, LocalizationManager.GetText("Package Documents")
    Exit Sub

DeleteError:
    MsgBox LocalizationManager.GetText("Error deleting package document item: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Public Sub DuplicateSelectedPackageItemFromForm(ByVal frm As Object, ByVal parentRowIndex As Long, ByVal packageId As String)
    On Error GoTo DuplicateError

    Dim itemId As String
    Dim newItemId As String

    itemId = Trim$(frm.txtItemId.Text)
    If Len(itemId) = 0 Then
        MsgBox LocalizationManager.GetText("Select a package document item first."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    newItemId = DuplicatePackageItemRecord(parentRowIndex, packageId, itemId)
    If Len(newItemId) = 0 Then
        MsgBox LocalizationManager.GetText("The selected package document item no longer exists."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    Call BindPackageDocumentsForm(frm, parentRowIndex, packageId)
    Call SelectPackageItemInList(frm, newItemId)
    MsgBox LocalizationManager.GetText("Package document item duplicated."), vbInformation, LocalizationManager.GetText("Package Documents")
    Exit Sub

DuplicateError:
    MsgBox LocalizationManager.GetText("Error duplicating package document item: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Public Sub FillPackageItemEditorFromParent(ByVal frm As Object, ByVal parentRowIndex As Long, Optional ByVal showMessages As Boolean = True)
    On Error GoTo ApplyError

    Dim parentTable As ListObject
    Dim itemsTable As ListObject
    Dim packageId As String
    Dim childCount As Long
    Dim selectedItemId As String
    Dim itemRowIndex As Long
    Dim parentAmount As Double
    Dim childAmount As Double
    Dim primaryStatus As String
    Dim primaryOperationNumber As String
    Dim primaryOperationDate As Variant

    Set parentTable = GetParentTable()
    Set itemsTable = GetItemsTable()
    If frm Is Nothing Or parentTable Is Nothing Or itemsTable Is Nothing Then Exit Sub

    packageId = EnsurePackageIdForParentRow(parentTable, parentRowIndex)
    childCount = GetPackageChildDocumentCount(parentRowIndex)
    If childCount <> 1 Then
        If showMessages Then MsgBox LocalizationManager.GetText("Package-level match can only be applied when the package has exactly one child document."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    selectedItemId = GetSelectedPackageItemIdFromForm(frm)
    If Len(selectedItemId) = 0 Then
        If frm.lstPackageItems.listCount = 1 Then
            frm.lstPackageItems.listIndex = 0
            selectedItemId = Trim$(CStr(frm.lstPackageItems.List(0, 7)))
        End If
    End If

    If Len(selectedItemId) = 0 Then
        If showMessages Then MsgBox LocalizationManager.GetText("Select a package document item first."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    itemRowIndex = FindItemRowIndexById(itemsTable, selectedItemId)
    If itemRowIndex = 0 Then Exit Sub

    If IsNumeric(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_AMOUNT_COLUMN)) Then
        parentAmount = CDbl(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_AMOUNT_COLUMN))
    End If
    If IsNumeric(GetItemCellValue(itemsTable, itemRowIndex, ITEM_COLUMN_AMOUNT)) Then
        childAmount = CDbl(GetItemCellValue(itemsTable, itemRowIndex, ITEM_COLUMN_AMOUNT))
    End If

    If Abs(parentAmount - childAmount) >= 0.01 Then
        If showMessages Then MsgBox LocalizationManager.GetText("Package-level match can only be applied when the single child amount matches the package amount."), vbExclamation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    primaryStatus = GetPackagePrimary1CStatus(parentRowIndex)
    primaryOperationNumber = GetPackagePrimary1COperationNumber(parentRowIndex)
    primaryOperationDate = GetPackagePrimary1COperationDate(parentRowIndex)

    If Len(Trim$(primaryOperationNumber)) = 0 Or Not IsApplicablePackageMatchStatus(primaryStatus) Then
        If showMessages Then MsgBox LocalizationManager.GetText("No package-level 1C match is available to apply."), vbInformation, LocalizationManager.GetText("Package Documents")
        Exit Sub
    End If

    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER, primaryOperationNumber
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE, primaryOperationDate
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_STATUS, primaryStatus
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_MODE, "package"
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_CONFIDENCE, GetPackageMatchConfidence(primaryStatus)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_COMMENT, LocalizationManager.GetText("Applied from package-level 1C match.")
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_UPDATED_AT, Now

    Call RefreshParentPackageSummary(parentRowIndex)
    Call BindPackageDocumentsForm(frm, parentRowIndex, packageId)
    Call SelectPackageItemInList(frm, selectedItemId)

    If showMessages Then MsgBox LocalizationManager.GetText("Package-level 1C match applied to the child document."), vbInformation, LocalizationManager.GetText("Package Documents")
    Exit Sub

ApplyError:
    If showMessages Then MsgBox LocalizationManager.GetText("Error applying package-level 1C match: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Public Function DuplicatePackageItemRecord(ByVal parentRowIndex As Long, ByVal packageId As String, ByVal itemId As String) As String
    Dim itemsTable As ListObject
    Dim rowIndex As Long

    Set itemsTable = GetItemsTable()
    rowIndex = FindItemRowIndexById(itemsTable, itemId)
    If rowIndex = 0 Then Exit Function

    DuplicatePackageItemRecord = SavePackageItemRecord( _
        parentRowIndex, _
        packageId, _
        vbNullString, _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_TYPE_DISPLAY)), _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_NUMBER)), _
        GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_DATE), _
        CDbl(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_AMOUNT)), _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_ASSET_CATEGORY)), _
        GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_QUANTITY), _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_UNIT)), _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_ORDER_INFO)), _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_FRP_NUMBER)), _
        GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_FRP_DATE), _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DESCRIPTION)), _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_NOTES)), _
        vbNullString, _
        vbNullString, _
        "not_checked", _
        vbNullString, _
        False)
End Function

Public Function SavePackageItemRecord(ByVal parentRowIndex As Long, ByVal packageId As String, ByVal itemId As String, ByVal documentTypeDisplay As String, ByVal documentNumber As String, ByVal itemDateValue As Variant, ByVal amountValue As Double, ByVal assetCategoryValue As String, ByVal quantityValue As Variant, ByVal unitValue As String, ByVal orderInfoValue As String, ByVal frpNumberValue As String, ByVal itemFrpDateValue As Variant, ByVal itemDescription As String, ByVal itemNotes As String, ByVal matchedOperationNumber As String, ByVal matchedOperationDateValue As Variant, ByVal matchedStatus As String, ByVal matchedComment As String, ByVal updateExisting As Boolean) As String
    Dim parentTable As ListObject
    Dim itemsTable As ListObject
    Dim itemRowIndex As Long
    Dim derivedAssetCategory As String
    Dim retainedQuantity As Variant
    Dim retainedUnit As Variant
    Dim retainedNotes As Variant
    Dim parentOrderInfo As Variant
    Dim parentFrpNumber As Variant
    Dim parentFrpDate As Variant

    Set parentTable = GetParentTable()
    Set itemsTable = GetItemsTable()
    If parentTable Is Nothing Or itemsTable Is Nothing Then Exit Function

    If updateExisting Then
        If Len(itemId) = 0 Then Exit Function
        itemRowIndex = FindItemRowIndexById(itemsTable, itemId)
        If itemRowIndex = 0 Then Exit Function
    Else
        itemId = CreatePackageItemId()
        itemsTable.ListRows.Add
        itemRowIndex = itemsTable.ListRows.Count
        SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_ITEM_ID, itemId
        SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_ITEM_ORDER, GetNextItemOrder(itemsTable, packageId)
        SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_CREATED_AT, Now
    End If

    derivedAssetCategory = GetDerivedAssetCategoryValue(documentTypeDisplay, GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_ASSET_CATEGORY))
    retainedQuantity = GetRetainedItemValue(itemsTable, itemRowIndex, ITEM_COLUMN_QUANTITY, updateExisting, vbNullString)
    retainedUnit = GetRetainedItemValue(itemsTable, itemRowIndex, ITEM_COLUMN_UNIT, updateExisting, vbNullString)
    retainedNotes = GetRetainedItemValue(itemsTable, itemRowIndex, ITEM_COLUMN_NOTES, updateExisting, itemNotes)
    parentOrderInfo = GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_ORDER_INFO_COLUMN)
    parentFrpNumber = GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_FRP_NUMBER_COLUMN)
    parentFrpDate = GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_FRP_DATE_COLUMN)

    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_PACKAGE_ID, packageId
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DOCUMENT_TYPE_DISPLAY, documentTypeDisplay
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DOCUMENT_TYPE, BuildItemTypeKey(documentTypeDisplay)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DOCUMENT_NUMBER, documentNumber
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DOCUMENT_DATE, itemDateValue
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_AMOUNT, amountValue
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_ASSET_CATEGORY, derivedAssetCategory
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_QUANTITY, retainedQuantity
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_UNIT, retainedUnit
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DESCRIPTION, itemDescription
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_NOTES, retainedNotes
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER, matchedOperationNumber
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE, matchedOperationDateValue
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_COMMENT, matchedComment

    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_COUNTERPARTY_NAME, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_COUNTERPARTY_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_COUNTERPARTY_NORMALIZED, GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_COUNTERPARTY_NORMALIZED)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DIRECTION, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_DIRECTION_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_SERVICE, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_SERVICE_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_EXECUTOR, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_EXECUTOR_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_ORDER_INFO, parentOrderInfo
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_FRP_NUMBER, parentFrpNumber
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_FRP_DATE, parentFrpDate
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_BASE_DOCUMENT_TYPE, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_DOCUMENT_TYPE_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_BASE_DOCUMENT_NUMBER, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_DOCUMENT_NUMBER_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_BASE_DOCUMENT_DATE, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_FRP_DATE_COLUMN)

    If Len(Trim$(matchedStatus)) > 0 Then
        SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_STATUS, matchedStatus
    ElseIf Len(Trim$(CStr(GetItemCellValue(itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_STATUS)))) = 0 Then
        SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_STATUS, "not_checked"
    End If
    If Len(Trim$(CStr(GetItemCellValue(itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_MODE)))) = 0 Then
        SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_MODE, "manual"
    End If
    If Len(Trim$(CStr(GetItemCellValue(itemsTable, itemRowIndex, ITEM_COLUMN_IS_POSTED_SEPARATELY)))) = 0 Then
        SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_IS_POSTED_SEPARATELY, "False"
    End If
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_UPDATED_AT, Now

    Call RefreshParentPackageSummary(parentRowIndex)
    SavePackageItemRecord = itemId
End Function

Public Function DeletePackageItemRecord(ByVal packageId As String, ByVal itemId As String) As Boolean
    Dim itemsTable As ListObject
    Dim itemRowIndex As Long

    Set itemsTable = GetItemsTable()
    itemRowIndex = FindItemRowIndexById(itemsTable, itemId)
    If itemRowIndex = 0 Then Exit Function

    itemsTable.ListRows(itemRowIndex).Delete
    Call ReindexPackageItemOrders(itemsTable, packageId)
    DeletePackageItemRecord = True
End Function

Public Sub ClearPackageItemEditor(ByVal frm As Object)
    frm.txtItemId.Text = ""
    frm.cmbItemDocumentTypeDisplay.value = ""
    frm.txtItemDocumentNumber.Text = ""
    frm.txtItemDocumentDate.Text = ""
    frm.txtItemAmount.Text = ""
    frm.txtItemDescription.Text = ""
    frm.txtMatched1COperationNumber.Text = ""
    frm.txtMatched1COperationDate.Text = ""
    frm.cmbMatched1CStatus.Text = "not_checked"
    frm.txtMatched1CComment.Text = ""
    If frm.lstPackageItems.listCount > 0 Then frm.lstPackageItems.listIndex = -1
    Call UpdatePackageItemReviewHint(frm)
End Sub

Public Sub UpdatePackageItemReviewHint(ByVal frm As Object)
    Dim statusValue As String
    Dim hintText As String
    Dim accentColor As Long
    Dim warningColor As Long
    Dim dangerColor As Long
    Dim neutralColor As Long

    If frm Is Nothing Then Exit Sub

    statusValue = LCase$(Trim$(CStr(frm.cmbMatched1CStatus.Text)))
    warningColor = RGB(255, 242, 204)
    dangerColor = RGB(255, 199, 206)
    accentColor = RGB(226, 239, 218)
    neutralColor = RGB(255, 255, 255)

    Select Case statusValue
        Case "candidate"
            hintText = LocalizationManager.GetText("Child match requires review. Confirm it as manual or reset it.")
            Call SetReviewFieldsColor(frm, warningColor)
            frm.lblReviewHint.ForeColor = RGB(156, 101, 0)
        Case "not_found"
            hintText = LocalizationManager.GetText("No 1C match found for this child document.")
            Call SetReviewFieldsColor(frm, dangerColor)
            frm.lblReviewHint.ForeColor = RGB(156, 0, 6)
        Case "manual"
            hintText = LocalizationManager.GetText("Child document was manually confirmed.")
            Call SetReviewFieldsColor(frm, accentColor)
            frm.lblReviewHint.ForeColor = RGB(0, 97, 0)
        Case "exact"
            hintText = LocalizationManager.GetText("Child document has an exact 1C match.")
            Call SetReviewFieldsColor(frm, accentColor)
            frm.lblReviewHint.ForeColor = RGB(0, 97, 0)
        Case Else
            hintText = LocalizationManager.GetText("Child document is pending 1C matching.")
            Call SetReviewFieldsColor(frm, neutralColor)
            frm.lblReviewHint.ForeColor = RGB(96, 96, 96)
    End Select

    frm.lblReviewHint.Caption = hintText
End Sub

Private Sub SelectPackageItemInList(ByVal frm As Object, ByVal itemId As String)
    Dim listIndex As Long

    If frm Is Nothing Then Exit Sub
    If Len(Trim$(itemId)) = 0 Then Exit Sub

    For listIndex = 0 To frm.lstPackageItems.listCount - 1
        If StrComp(CStr(frm.lstPackageItems.List(listIndex, 7)), itemId, vbTextCompare) = 0 Then
            frm.lstPackageItems.listIndex = listIndex
            Exit For
        End If
    Next listIndex
End Sub

Public Sub LoadParentSummaryToForm(ByVal frm As Object, ByVal parentRowIndex As Long, ByVal packageId As String)
    Dim parentTable As ListObject
    Dim counterparty As String
    Dim parentAmount As String
    Dim childrenTotal As String
    Dim amountCheckText As String
    Dim childCount As String
    Dim primaryStatusText As String
    Dim primaryOperationNumber As String
    Dim reviewSummaryText As String

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Sub

    counterparty = CStr(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_COUNTERPARTY_COLUMN))
    parentAmount = FormatItemAmountValue(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_AMOUNT_COLUMN))
    childrenTotal = FormatItemAmountValue(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT))
    childCount = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT)
    amountCheckText = TranslateAmountCheckStatus(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS))
    primaryStatusText = TranslatePrimaryMatchStatus(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_STATUS))
    primaryOperationNumber = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_NUMBER)
    reviewSummaryText = BuildPackageReviewSummary(packageId)

    frm.lblPackageSummary.Caption = LocalizationManager.GetText("Package:") & " " & packageId & vbCrLf & _
        LocalizationManager.GetText("Counterparty:") & " " & counterparty & vbCrLf & _
        LocalizationManager.GetText("Parent Amount:") & " " & parentAmount & "    " & _
        LocalizationManager.GetText("Children Total:") & " " & childrenTotal & "    " & _
        LocalizationManager.GetText("Items:") & " " & childCount & "    " & _
        LocalizationManager.GetText("Amount Check:") & " " & amountCheckText & vbCrLf & _
        LocalizationManager.GetText("1C Status") & ": " & primaryStatusText & IIf(Len(Trim$(primaryOperationNumber)) > 0, "    " & LocalizationManager.GetText("1C Operation No.") & ": " & primaryOperationNumber, vbNullString) & vbCrLf & _
        reviewSummaryText
End Sub

Public Sub RefreshParentPackageSummary(ByVal parentRowIndex As Long)
    On Error GoTo RefreshError

    Dim parentTable As ListObject
    Dim itemsTable As ListObject
    Dim packageId As String
    Dim parentAmount As Double
    Dim childrenTotal As Double
    Dim childCount As Long
    Dim rowIndex As Long
    Dim statusValue As String
    Dim exactCount As Long
    Dim candidateCount As Long
    Dim manualCount As Long
    Dim notFoundCount As Long
    Dim notCheckedCount As Long
    Dim childStatusValue As String
    Dim childOperationNumber As String
    Dim childOperationDate As Variant
    Dim primaryStatusValue As String
    Dim matchedOperationCount As Long
    Dim matchedOperationNumber As String
    Dim matchedOperationDate As Variant
    Dim existingPrimaryStatus As String
    Dim existingPrimaryNumber As String
    Dim existingPrimaryDate As Variant

    Set parentTable = GetParentTable()
    Set itemsTable = GetItemsTable()
    If parentTable Is Nothing Or itemsTable Is Nothing Then Exit Sub

    packageId = EnsurePackageIdForParentRow(parentTable, parentRowIndex)
    If Len(packageId) = 0 Then Exit Sub

    existingPrimaryStatus = LCase$(Trim$(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_STATUS)))
    existingPrimaryNumber = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_NUMBER)
    existingPrimaryDate = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_DATE)

    If IsNumeric(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_AMOUNT_COLUMN)) Then
        parentAmount = CDbl(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_AMOUNT_COLUMN))
    End If

    childrenTotal = 0
    childCount = 0

    If Not itemsTable.DataBodyRange Is Nothing Then
        For rowIndex = 1 To itemsTable.ListRows.Count
            If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_PACKAGE_ID))), packageId, vbTextCompare) = 0 Then
                childCount = childCount + 1
                If IsNumeric(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_AMOUNT)) Then
                    childrenTotal = childrenTotal + CDbl(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_AMOUNT))
                End If

                childStatusValue = LCase$(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS))))
                Select Case childStatusValue
                    Case "exact"
                        exactCount = exactCount + 1
                    Case "candidate"
                        candidateCount = candidateCount + 1
                    Case "manual"
                        manualCount = manualCount + 1
                    Case "not_found"
                        notFoundCount = notFoundCount + 1
                    Case Else
                        notCheckedCount = notCheckedCount + 1
                End Select

                childOperationNumber = Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER)))
                childOperationDate = GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE)
                If Len(childOperationNumber) > 0 And childStatusValue <> "not_checked" And childStatusValue <> "not_found" Then
                    matchedOperationCount = matchedOperationCount + 1
                    If matchedOperationCount = 1 Then
                        matchedOperationNumber = childOperationNumber
                        matchedOperationDate = childOperationDate
                    Else
                        matchedOperationNumber = ""
                        matchedOperationDate = ""
                    End If
                End If
            End If
        Next rowIndex
    End If

    If childCount = 0 Then
        statusValue = "not_checked"
    ElseIf Abs(parentAmount - childrenTotal) < 0.01 Then
        statusValue = "match"
    Else
        statusValue = "mismatch"
    End If

    If childCount = 0 Then
        primaryStatusValue = "not_checked"
    ElseIf exactCount = childCount Then
        primaryStatusValue = "exact"
    ElseIf manualCount = childCount Then
        primaryStatusValue = "manual"
    ElseIf notFoundCount = childCount Then
        primaryStatusValue = "not_found"
    ElseIf (exactCount + candidateCount + manualCount) > 0 Then
        primaryStatusValue = "candidate"
    ElseIf childCount > 0 And notCheckedCount = childCount And IsApplicablePackageMatchStatus(existingPrimaryStatus) And Len(Trim$(existingPrimaryNumber)) > 0 Then
        primaryStatusValue = existingPrimaryStatus
        matchedOperationNumber = existingPrimaryNumber
        matchedOperationDate = existingPrimaryDate
    Else
        primaryStatusValue = "not_checked"
    End If

    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_HAS_CHILD_DOCUMENTS, IIf(childCount > 0, "True", "False")
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT, childCount
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT, childrenTotal
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS, statusValue
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_STATUS, primaryStatusValue
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_NUMBER, matchedOperationNumber
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_DATE, matchedOperationDate
    Exit Sub

RefreshError:
    MsgBox LocalizationManager.GetText("Error refreshing package summary: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Private Function ValidatePackageItemForm(ByVal frm As Object, ByRef amountValue As Double, ByRef itemDateValue As Variant) As Boolean
    ValidatePackageItemForm = False

    If Len(Trim$(frm.cmbItemDocumentTypeDisplay.value)) = 0 Then
        MsgBox LocalizationManager.GetText("Document Type is required."), vbExclamation, LocalizationManager.GetText("Package Documents")
        frm.cmbItemDocumentTypeDisplay.SetFocus
        Exit Function
    End If

    If Len(Trim$(frm.txtItemDocumentNumber.Text)) = 0 Then
        MsgBox LocalizationManager.GetText("Document Number is required."), vbExclamation, LocalizationManager.GetText("Package Documents")
        frm.txtItemDocumentNumber.SetFocus
        Exit Function
    End If

    If Len(Trim$(frm.txtItemAmount.Text)) = 0 Or Not IsNumeric(frm.txtItemAmount.Text) Then
        MsgBox LocalizationManager.GetText("Amount must be numeric."), vbExclamation, LocalizationManager.GetText("Package Documents")
        frm.txtItemAmount.SetFocus
        Exit Function
    End If
    amountValue = CDbl(frm.txtItemAmount.Text)

    itemDateValue = ""
    If Len(Trim$(frm.txtItemDocumentDate.Text)) > 0 Then
        If Not CommonUtilities.IsValidDateFormat(frm.txtItemDocumentDate.Text) Then
            MsgBox LocalizationManager.GetText("Enter date in DD.MM.YY format or leave it blank."), vbExclamation, LocalizationManager.GetText("Package Documents")
            frm.txtItemDocumentDate.SetFocus
            Exit Function
        End If
        itemDateValue = ParseShortDateText(frm.txtItemDocumentDate.Text)
    End If

    ValidatePackageItemForm = True
End Function

Private Function ValidateMatchedFields(ByVal frm As Object, ByRef matchedOperationDateValue As Variant) As Boolean
    ValidateMatchedFields = False

    matchedOperationDateValue = ""
    If Len(Trim$(frm.txtMatched1COperationDate.Text)) > 0 Then
        If Not CommonUtilities.IsValidDateFormat(frm.txtMatched1COperationDate.Text) Then
            MsgBox LocalizationManager.GetText("Enter date in DD.MM.YY format or leave it blank."), vbExclamation, LocalizationManager.GetText("Package Documents")
            frm.txtMatched1COperationDate.SetFocus
            Exit Function
        End If
        matchedOperationDateValue = ParseShortDateText(frm.txtMatched1COperationDate.Text)
    End If

    If Len(Trim$(frm.cmbMatched1CStatus.Text)) = 0 Then
        frm.cmbMatched1CStatus.Text = "not_checked"
    End If

    ValidateMatchedFields = True
End Function

Private Function ParseShortDateText(ByVal shortDateText As String) As Date
    ParseShortDateText = CDate(Left$(shortDateText, 6) & "20" & Right$(shortDateText, 2))
End Function

Private Function GetDerivedAssetCategoryValue(ByVal documentTypeDisplay As String, ByVal parentAssetCategory As String) As String
    Dim documentTypeKey As String

    documentTypeKey = BuildItemTypeKey(documentTypeDisplay)

    If IsMaterialDocumentType(documentTypeKey) Then
        GetDerivedAssetCategoryValue = "inventory"
    ElseIf IsFixedAssetDocumentType(documentTypeKey) Then
        GetDerivedAssetCategoryValue = "fixed_assets"
    ElseIf IsAmbiguousAccountingOperationType(documentTypeKey) Then
        GetDerivedAssetCategoryValue = vbNullString
    Else
        GetDerivedAssetCategoryValue = Trim$(parentAssetCategory)
    End If
End Function

Private Function IsMaterialDocumentType(ByVal documentTypeKey As String) As Boolean
    IsMaterialDocumentType = (InStr(documentTypeKey, ChrW$(1053) & ChrW$(1040) & ChrW$(1050) & ChrW$(1051) & ChrW$(1040) & ChrW$(1044) & ChrW$(1053)) > 0) _
        Or (InStr(documentTypeKey, ChrW$(1052) & ChrW$(1040) & ChrW$(1058) & ChrW$(1045) & ChrW$(1056) & ChrW$(1048) & ChrW$(1040) & ChrW$(1051)) > 0) _
        Or (InStr(documentTypeKey, ChrW$(1055) & ChrW$(1054) & ChrW$(1057) & ChrW$(1058) & ChrW$(1059) & ChrW$(1055) & ChrW$(1051) & ChrW$(1045) & ChrW$(1053) & ChrW$(1048) & ChrW$(1045) & "_" & ChrW$(1052) & ChrW$(1047)) > 0) _
        Or (InStr(documentTypeKey, ChrW$(1055) & ChrW$(1056) & ChrW$(1048) & ChrW$(1045) & ChrW$(1052) & ChrW$(1050) & ChrW$(1048) & "_" & ChrW$(1052) & ChrW$(1040) & ChrW$(1058) & ChrW$(1045) & ChrW$(1056) & ChrW$(1048) & ChrW$(1040) & ChrW$(1051)) > 0)
End Function

Private Function IsFixedAssetDocumentType(ByVal documentTypeKey As String) As Boolean
    IsFixedAssetDocumentType = (InStr(documentTypeKey, ChrW$(1055) & ChrW$(1045) & ChrW$(1056) & ChrW$(1045) & ChrW$(1044) & ChrW$(1040) & ChrW$(1063) & ChrW$(1040) & "_" & ChrW$(1054) & ChrW$(1041) & ChrW$(1066) & ChrW$(1045) & ChrW$(1050) & ChrW$(1058) & ChrW$(1054) & ChrW$(1042)) > 0) _
        Or (InStr(documentTypeKey, ChrW$(1055) & ChrW$(1056) & ChrW$(1048) & ChrW$(1053) & ChrW$(1071) & ChrW$(1058) & ChrW$(1048) & ChrW$(1045) & "_" & ChrW$(1050) & "_" & ChrW$(1059) & ChrW$(1063) & ChrW$(1045) & ChrW$(1058) & ChrW$(1059) & "_" & ChrW$(1054) & ChrW$(1057)) > 0) _
        Or (InStr(documentTypeKey, ChrW$(1055) & ChrW$(1054) & ChrW$(1057) & ChrW$(1058) & ChrW$(1059) & ChrW$(1055) & ChrW$(1051) & ChrW$(1045) & ChrW$(1053) & ChrW$(1048) & ChrW$(1045) & "_" & ChrW$(1054) & ChrW$(1057)) > 0)
End Function

Private Function IsAmbiguousAccountingOperationType(ByVal documentTypeKey As String) As Boolean
    IsAmbiguousAccountingOperationType = (InStr(documentTypeKey, ChrW$(1054) & ChrW$(1055) & ChrW$(1045) & ChrW$(1056) & ChrW$(1040) & ChrW$(1062) & ChrW$(1048) & ChrW$(1071) & "_" & ChrW$(1041) & ChrW$(1059) & ChrW$(1061) & ChrW$(1043) & ChrW$(1040) & ChrW$(1051) & ChrW$(1058) & ChrW$(1045) & ChrW$(1056) & ChrW$(1057) & ChrW$(1050) & ChrW$(1040) & ChrW$(1071)) > 0)
End Function

Private Function GetEffectiveTextValue(ByVal currentValue As String, ByVal fallbackValue As Variant) As String
    If Len(Trim$(currentValue)) > 0 Then
        GetEffectiveTextValue = Trim$(currentValue)
    Else
        GetEffectiveTextValue = Trim$(CStr(fallbackValue))
    End If
End Function

Private Function GetRetainedItemValue(ByVal itemsTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal updateExisting As Boolean, ByVal fallbackValue As Variant) As Variant
    If updateExisting Then
        GetRetainedItemValue = GetItemCellValue(itemsTable, rowIndex, columnName)
    Else
        GetRetainedItemValue = fallbackValue
    End If
End Function

Private Function GetEffectiveDateValue(ByVal currentValue As Variant, ByVal fallbackValue As Variant) As Variant
    If Len(Trim$(CStr(currentValue))) > 0 Then
        GetEffectiveDateValue = currentValue
    Else
        GetEffectiveDateValue = fallbackValue
    End If
End Function

Private Function GetOptionalControl(ByVal frm As Object, ByVal controlName As String) As Object
    On Error Resume Next
    Set GetOptionalControl = frm.Controls(controlName)
    On Error GoTo 0
End Function

Private Function GetOptionalFormText(ByVal frm As Object, ByVal controlName As String) As String
    Dim ctrl As Object
    Set ctrl = GetOptionalControl(frm, controlName)
    If ctrl Is Nothing Then Exit Function
    On Error Resume Next
    GetOptionalFormText = Trim$(CStr(ctrl.Text))
    If Err.Number <> 0 Then
        Err.Clear
        GetOptionalFormText = Trim$(CStr(ctrl.value))
    End If
    On Error GoTo 0
End Function

Private Function GetOptionalFormValue(ByVal frm As Object, ByVal controlName As String) As String
    Dim ctrl As Object
    Set ctrl = GetOptionalControl(frm, controlName)
    If ctrl Is Nothing Then Exit Function
    On Error Resume Next
    GetOptionalFormValue = Trim$(CStr(ctrl.value))
    On Error GoTo 0
End Function

Private Sub SetOptionalFormText(ByVal frm As Object, ByVal controlName As String, ByVal textValue As String)
    Dim ctrl As Object
    Set ctrl = GetOptionalControl(frm, controlName)
    If ctrl Is Nothing Then Exit Sub
    On Error Resume Next
    ctrl.Text = textValue
    If Err.Number <> 0 Then
        Err.Clear
        ctrl.value = textValue
    End If
    On Error GoTo 0
End Sub

Private Sub SetOptionalFormValue(ByVal frm As Object, ByVal controlName As String, ByVal valueText As String)
    Dim ctrl As Object
    Set ctrl = GetOptionalControl(frm, controlName)
    If ctrl Is Nothing Then Exit Sub
    On Error Resume Next
    ctrl.value = valueText
    On Error GoTo 0
End Sub

Private Sub SetFocusIfExists(ByVal frm As Object, ByVal controlName As String)
    Dim ctrl As Object
    Set ctrl = GetOptionalControl(frm, controlName)
    If ctrl Is Nothing Then Exit Sub
    On Error Resume Next
    ctrl.SetFocus
    On Error GoTo 0
End Sub

Private Function FormatEditorNumberValue(ByVal rawValue As Variant) As String
    Dim numericValue As Double

    If Len(Trim$(CStr(rawValue))) = 0 Then Exit Function
    If IsNumeric(rawValue) Then
        numericValue = CDbl(rawValue)
        If numericValue = CLng(numericValue) Then
            FormatEditorNumberValue = CStr(CLng(numericValue))
        Else
            FormatEditorNumberValue = Replace(Trim$(CStr(numericValue)), ",", ".")
        End If
    Else
        FormatEditorNumberValue = Trim$(CStr(rawValue))
    End If
End Function

Private Function GetParentTable() As ListObject
    Dim ws As Worksheet
    Set ws = CommonUtilities.GetWorksheetSafe("IncOut")
    Set GetParentTable = CommonUtilities.GetListObjectSafe(ws, PACKAGE_PARENT_TABLE_NAME)
End Function

Private Function GetItemsTable() As ListObject
    Dim ws As Worksheet
    Set ws = CommonUtilities.GetWorksheetSafe(PACKAGE_ITEMS_SHEET_NAME)
    Set GetItemsTable = CommonUtilities.GetListObjectSafe(ws, PACKAGE_ITEMS_TABLE_NAME)
End Function

Private Function IsChildMatchPending(ByVal statusValue As String) As Boolean
    Select Case LCase$(Trim$(statusValue))
        Case "exact", "manual"
            IsChildMatchPending = False
        Case Else
            IsChildMatchPending = True
    End Select
End Function

Private Function IsReviewTargetStatus(ByVal statusValue As String) As Boolean
    Select Case LCase$(Trim$(statusValue))
        Case "candidate", "not_found", "not_checked", vbNullString
            IsReviewTargetStatus = True
    End Select
End Function

Private Function IsItemVisibleForReviewFilter(ByVal statusValue As String, ByVal reviewFilter As String) As Boolean
    Dim filterValue As String
    Dim normalizedStatus As String

    filterValue = LCase$(Trim$(reviewFilter))
    normalizedStatus = LCase$(Trim$(statusValue))

    If Len(filterValue) = 0 Or filterValue = LCase$(LocalizationManager.GetText("All")) Then
        IsItemVisibleForReviewFilter = True
        Exit Function
    End If

    Select Case filterValue
        Case LCase$(LocalizationManager.GetText("Needs review"))
            IsItemVisibleForReviewFilter = IsReviewTargetStatus(normalizedStatus)
        Case LCase$(LocalizationManager.GetText("Pending"))
            IsItemVisibleForReviewFilter = (normalizedStatus = "not_checked" Or Len(normalizedStatus) = 0)
        Case LCase$(LocalizationManager.GetText("Candidate"))
            IsItemVisibleForReviewFilter = (normalizedStatus = "candidate")
        Case LCase$(LocalizationManager.GetText("Not found"))
            IsItemVisibleForReviewFilter = (normalizedStatus = "not_found")
        Case Else
            IsItemVisibleForReviewFilter = True
    End Select
End Function

Private Function GetActiveReviewFilter(ByVal frm As Object) As String
    On Error Resume Next
    GetActiveReviewFilter = Trim$(CStr(frm.cmbReviewFilter.value))
    On Error GoTo 0
End Function

Private Sub SetReviewFieldsColor(ByVal frm As Object, ByVal backColorValue As Long)
    On Error Resume Next
    frm.txtMatched1COperationNumber.BackColor = backColorValue
    frm.txtMatched1COperationDate.BackColor = backColorValue
    frm.cmbMatched1CStatus.BackColor = backColorValue
    frm.txtMatched1CComment.BackColor = backColorValue
    On Error GoTo 0
End Sub

Private Function BuildPackageReviewSummary(ByVal packageId As String) As String
    Dim itemsTable As ListObject
    Dim rowIndex As Long
    Dim candidateCount As Long
    Dim notFoundCount As Long
    Dim pendingCount As Long
    Dim manualCount As Long
    Dim statusValue As String

    Set itemsTable = GetItemsTable()
    If itemsTable Is Nothing Then
        BuildPackageReviewSummary = LocalizationManager.GetText("Review:") & " " & _
            LocalizationManager.GetText("Candidate") & ": 0 | " & _
            LocalizationManager.GetText("Not found") & ": 0 | " & _
            LocalizationManager.GetText("Pending") & ": 0 | " & _
            LocalizationManager.GetText("Manual") & ": 0"
        Exit Function
    End If

    If Not itemsTable.DataBodyRange Is Nothing Then
        For rowIndex = 1 To itemsTable.ListRows.Count
            If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
                statusValue = LCase$(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS))))
                Select Case statusValue
                    Case "candidate"
                        candidateCount = candidateCount + 1
                    Case "not_found"
                        notFoundCount = notFoundCount + 1
                    Case "manual"
                        manualCount = manualCount + 1
                    Case Else
                        If statusValue <> "exact" Then pendingCount = pendingCount + 1
                End Select
            End If
        Next rowIndex
    End If

    BuildPackageReviewSummary = LocalizationManager.GetText("Review:") & " " & _
        LocalizationManager.GetText("Candidate") & ": " & candidateCount & " | " & _
        LocalizationManager.GetText("Not found") & ": " & notFoundCount & " | " & _
        LocalizationManager.GetText("Pending") & ": " & pendingCount & " | " & _
        LocalizationManager.GetText("Manual") & ": " & manualCount
End Function

Private Function GetSelectedPackageItemIdFromForm(ByVal frm As Object) As String
    If frm Is Nothing Then Exit Function

    GetSelectedPackageItemIdFromForm = Trim$(CStr(frm.txtItemId.Text))
    If Len(GetSelectedPackageItemIdFromForm) > 0 Then Exit Function

    If frm.lstPackageItems.listIndex >= 0 Then
        GetSelectedPackageItemIdFromForm = Trim$(CStr(frm.lstPackageItems.List(frm.lstPackageItems.listIndex, 7)))
    End If
End Function

Private Sub UpdatePackageItemMatchState(ByVal itemId As String, ByVal statusValue As String, ByVal keepExistingOperation As Boolean)
    Dim itemsTable As ListObject
    Dim rowIndex As Long

    Set itemsTable = GetItemsTable()
    rowIndex = FindItemRowIndexById(itemsTable, itemId)
    If rowIndex = 0 Then Exit Sub

    Select Case LCase$(Trim$(statusValue))
        Case "manual"
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS, "manual"
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_MODE, "manual"
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_CONFIDENCE, 100
        Case Else
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER, vbNullString
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE, vbNullString
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_COMMENT, vbNullString
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS, "not_checked"
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_MODE, vbNullString
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_CONFIDENCE, 0
    End Select

    If Not keepExistingOperation Then
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER, vbNullString
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE, vbNullString
    End If

    SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_UPDATED_AT, Now
End Sub

Private Sub ApplyChildMatchResult(ByVal itemsTable As ListObject, ByVal rowIndex As Long, ByVal childFound As Boolean, ByVal childProvodkaNumber As String, ByVal childProvodkaDate As Variant, ByVal childMatchCount As Long, ByVal childStatusMessage As String, ByVal childCandidatesList As String)
    Dim commentText As String

    If itemsTable Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > itemsTable.ListRows.Count Then Exit Sub

    commentText = BuildChildMatchComment(childStatusMessage, childMatchCount, childCandidatesList)

    If childFound Then
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER, childProvodkaNumber
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE, childProvodkaDate
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS, "exact"
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_MODE, "auto"
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_CONFIDENCE, 100
    ElseIf childMatchCount > 1 Then
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER, childProvodkaNumber
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE, childProvodkaDate
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS, "candidate"
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_MODE, "auto"
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_CONFIDENCE, 50
    Else
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_NUMBER, vbNullString
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_OPERATION_DATE, vbNullString
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_STATUS, "not_found"
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_MODE, "auto"
        SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_CONFIDENCE, 0
    End If

    SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_MATCHED_COMMENT, commentText
    SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_UPDATED_AT, Now
End Sub

Private Function BuildChildMatchComment(ByVal childStatusMessage As String, ByVal childMatchCount As Long, ByVal childCandidatesList As String) As String
    BuildChildMatchComment = Trim$(childStatusMessage)

    If childMatchCount > 1 Then
        If Len(Trim$(childCandidatesList)) > 0 Then
            BuildChildMatchComment = BuildChildMatchComment & " | " & childCandidatesList
        End If
    End If
End Function

Private Function GetPackagePrimary1COperationDate(ByVal parentRowIndex As Long) As Variant
    Dim parentTable As ListObject

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then
        GetPackagePrimary1COperationDate = vbNullString
        Exit Function
    End If

    GetPackagePrimary1COperationDate = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_PRIMARY_1C_DATE)
End Function

Private Function IsApplicablePackageMatchStatus(ByVal statusValue As String) As Boolean
    Select Case LCase$(Trim$(statusValue))
        Case "exact", "candidate", "manual"
            IsApplicablePackageMatchStatus = True
    End Select
End Function

Private Function GetPackageMatchConfidence(ByVal statusValue As String) As Long
    Select Case LCase$(Trim$(statusValue))
        Case "exact", "manual"
            GetPackageMatchConfidence = 100
        Case "candidate"
            GetPackageMatchConfidence = 50
        Case Else
            GetPackageMatchConfidence = 0
    End Select
End Function

Private Function GetItemCellValue(ByVal itemsTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim listColumn As listColumn
    Set listColumn = CommonUtilities.GetListColumnSafe(itemsTable, columnName)
    If listColumn Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > itemsTable.ListRows.Count Then Exit Function
    GetItemCellValue = itemsTable.DataBodyRange.Cells(rowIndex, listColumn.Index).value
End Function

Private Sub SetItemCellValue(ByVal itemsTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueToWrite As Variant)
    Dim listColumn As listColumn
    Set listColumn = CommonUtilities.GetListColumnSafe(itemsTable, columnName)
    If listColumn Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > itemsTable.ListRows.Count Then Exit Sub
    itemsTable.DataBodyRange.Cells(rowIndex, listColumn.Index).value = valueToWrite
End Sub

Private Function GetParentSourceValue(ByVal parentTable As ListObject, ByVal rowIndex As Long, ByVal columnIndex As Long) As Variant
    If parentTable Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > parentTable.ListRows.Count Then Exit Function
    If columnIndex < 1 Or columnIndex > parentTable.ListColumns.Count Then Exit Function
    GetParentSourceValue = parentTable.DataBodyRange.Cells(rowIndex, columnIndex).value
End Function

Private Function GetParentPackageText(ByVal parentTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim listColumn As listColumn
    Set listColumn = CommonUtilities.GetListColumnSafe(parentTable, columnName)
    If listColumn Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > parentTable.ListRows.Count Then Exit Function
    GetParentPackageText = Trim$(CStr(parentTable.DataBodyRange.Cells(rowIndex, listColumn.Index).value))
End Function

Private Sub SetParentPackageValue(ByVal parentTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueToWrite As Variant)
    Dim listColumn As listColumn
    Set listColumn = CommonUtilities.GetListColumnSafe(parentTable, columnName)
    If listColumn Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > parentTable.ListRows.Count Then Exit Sub
    parentTable.DataBodyRange.Cells(rowIndex, listColumn.Index).value = valueToWrite
End Sub

Private Function FindItemRowIndexById(ByVal itemsTable As ListObject, ByVal itemId As String) As Long
    Dim rowIndex As Long
    If itemsTable Is Nothing Then Exit Function
    If itemsTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To itemsTable.ListRows.Count
        If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_ITEM_ID))), Trim$(itemId), vbTextCompare) = 0 Then
            FindItemRowIndexById = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function GetNextItemOrder(ByVal itemsTable As ListObject, ByVal packageId As String) As Long
    Dim rowIndex As Long
    Dim currentOrder As Long

    GetNextItemOrder = 1
    If itemsTable Is Nothing Then Exit Function
    If itemsTable.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To itemsTable.ListRows.Count
        If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
            If IsNumeric(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_ITEM_ORDER)) Then
                currentOrder = CLng(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_ITEM_ORDER))
                If currentOrder >= GetNextItemOrder Then GetNextItemOrder = currentOrder + 1
            End If
        End If
    Next rowIndex
End Function

Private Sub ReindexPackageItemOrders(ByVal itemsTable As ListObject, ByVal packageId As String)
    Dim rowIndex As Long
    Dim nextOrder As Long

    nextOrder = 1
    If itemsTable Is Nothing Then Exit Sub
    If itemsTable.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To itemsTable.ListRows.Count
        If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
            SetItemCellValue itemsTable, rowIndex, ITEM_COLUMN_ITEM_ORDER, nextOrder
            nextOrder = nextOrder + 1
        End If
    Next rowIndex
End Sub

Private Function CreatePackageItemId() As String
    Randomize
    CreatePackageItemId = "ITM-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int((Rnd() * 9000) + 1000), "0000")
End Function

Private Function BuildItemTypeKey(ByVal displayText As String) As String
    Dim keyText As String

    keyText = Trim$(displayText)
    keyText = Replace(keyText, vbCr, " ")
    keyText = Replace(keyText, vbLf, " ")
    keyText = Replace(keyText, "/", "_")
    keyText = Replace(keyText, "-", "_")
    keyText = Replace(keyText, ".", "_")
    keyText = Replace(keyText, ",", "_")
    Do While InStr(keyText, "  ") > 0
        keyText = Replace(keyText, "  ", " ")
    Loop
    keyText = Replace(keyText, " ", "_")
    Do While InStr(keyText, "__") > 0
        keyText = Replace(keyText, "__", "_")
    Loop

    If Len(keyText) = 0 Then keyText = "UNSPECIFIED_DOCUMENT"
    BuildItemTypeKey = UCase$(keyText)
End Function

Private Function FormatItemDateValue(ByVal valueToFormat As Variant) As String
    If IsDate(valueToFormat) Then
        FormatItemDateValue = Format$(CDate(valueToFormat), "dd.mm.yy")
    Else
        FormatItemDateValue = Trim$(CStr(valueToFormat))
    End If
End Function

Private Function FormatItemAmountValue(ByVal valueToFormat As Variant) As String
    If IsNumeric(valueToFormat) Then
        FormatItemAmountValue = Replace(Format$(CDbl(valueToFormat), "0.00"), ".00", "")
    Else
        FormatItemAmountValue = Trim$(CStr(valueToFormat))
    End If
End Function

Private Function FormatEditorAmountValue(ByVal valueToFormat As Variant) As String
    If IsNumeric(valueToFormat) Then
        FormatEditorAmountValue = Replace(Format$(CDbl(valueToFormat), "0.00"), ",", Application.DecimalSeparator)
    Else
        FormatEditorAmountValue = Trim$(CStr(valueToFormat))
    End If
End Function

Private Function TranslateAmountCheckStatus(ByVal statusValue As String) As String
    Select Case LCase$(Trim$(statusValue))
        Case "match"
            TranslateAmountCheckStatus = LocalizationManager.GetText("Amount matches")
        Case "mismatch"
            TranslateAmountCheckStatus = LocalizationManager.GetText("Amount mismatch")
        Case Else
            TranslateAmountCheckStatus = LocalizationManager.GetText("Not checked")
    End Select
End Function

Private Function TranslatePrimaryMatchStatus(ByVal statusValue As String) As String
    Select Case LCase$(Trim$(statusValue))
        Case "exact"
            TranslatePrimaryMatchStatus = LocalizationManager.GetText("Exact")
        Case "candidate"
            TranslatePrimaryMatchStatus = LocalizationManager.GetText("Candidate")
        Case "manual"
            TranslatePrimaryMatchStatus = LocalizationManager.GetText("Manual")
        Case "not_found"
            TranslatePrimaryMatchStatus = LocalizationManager.GetText("Not found")
        Case Else
            TranslatePrimaryMatchStatus = LocalizationManager.GetText("Not checked")
    End Select
End Function


