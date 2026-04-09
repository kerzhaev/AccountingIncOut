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
Private Const ITEM_COLUMN_DESCRIPTION As String = "ItemDescription"
Private Const ITEM_COLUMN_BASE_DOCUMENT_TYPE As String = "BaseDocumentType"
Private Const ITEM_COLUMN_BASE_DOCUMENT_NUMBER As String = "BaseDocumentNumber"
Private Const ITEM_COLUMN_BASE_DOCUMENT_DATE As String = "BaseDocumentDate"
Private Const ITEM_COLUMN_MATCHED_STATUS As String = "Matched1CMatchStatus"
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

Public Sub RefreshPackageIndicatorsOnMainForm(ByVal frm As Object, ByVal parentRowIndex As Long)
    Dim parentTable As ListObject
    Dim childCount As String
    Dim childrenTotal As String
    Dim statusValue As String
    Dim statusText As String

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

    frm.lblPackageIndicators.Caption = LocalizationManager.GetText("Items:") & " " & childCount & " | " & _
        LocalizationManager.GetText("Children Total:") & " " & childrenTotal & vbCrLf & _
        LocalizationManager.GetText("Amount Check:") & " " & statusText

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
    Call LoadPackageItemsToList(frm, packageId)
    Call ClearPackageItemEditor(frm)
End Sub

Public Sub LoadPackageItemsToList(ByVal frm As Object, ByVal packageId As String)
    Dim itemsTable As ListObject
    Dim dataRow As Range
    Dim listIndex As Long

    Set itemsTable = GetItemsTable()
    frm.lstPackageItems.Clear

    If itemsTable Is Nothing Then Exit Sub
    If itemsTable.DataBodyRange Is Nothing Then Exit Sub

    For Each dataRow In itemsTable.DataBodyRange.rows
        If StrComp(Trim$(CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_PACKAGE_ID))), Trim$(packageId), vbTextCompare) = 0 Then
            frm.lstPackageItems.AddItem CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_ITEM_ORDER))
            listIndex = frm.lstPackageItems.listCount - 1
            frm.lstPackageItems.List(listIndex, 1) = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_DOCUMENT_TYPE_DISPLAY))
            frm.lstPackageItems.List(listIndex, 2) = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_DOCUMENT_NUMBER))
            frm.lstPackageItems.List(listIndex, 3) = FormatItemDateValue(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_DOCUMENT_DATE))
            frm.lstPackageItems.List(listIndex, 4) = FormatItemAmountValue(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_AMOUNT))
            frm.lstPackageItems.List(listIndex, 5) = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_MATCHED_STATUS))
            frm.lstPackageItems.List(listIndex, 6) = CStr(GetItemCellValue(itemsTable, dataRow.Row - itemsTable.DataBodyRange.Row + 1, ITEM_COLUMN_ITEM_ID))
        End If
    Next dataRow
End Sub

Public Sub LoadSelectedPackageItemIntoForm(ByVal frm As Object, ByVal packageId As String)
    On Error GoTo LoadError

    Dim selectedItemId As String
    Dim itemsTable As ListObject
    Dim rowIndex As Long

    If frm.lstPackageItems.listIndex < 0 Then Exit Sub
    selectedItemId = CStr(frm.lstPackageItems.List(frm.lstPackageItems.listIndex, 6))
    If Len(selectedItemId) = 0 Then Exit Sub

    Set itemsTable = GetItemsTable()
    rowIndex = FindItemRowIndexById(itemsTable, selectedItemId)
    If rowIndex = 0 Then Exit Sub

    frm.txtItemId.Text = selectedItemId
    frm.txtItemDocumentTypeDisplay.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_TYPE_DISPLAY))
    frm.txtItemDocumentNumber.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_NUMBER))
    frm.txtItemDocumentDate.Text = FormatItemDateValue(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DOCUMENT_DATE))
    frm.txtItemAmount.Text = FormatEditorAmountValue(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_AMOUNT))
    frm.txtItemDescription.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DESCRIPTION))
    frm.txtItemNotes.Text = CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_NOTES))
    Exit Sub

LoadError:
    MsgBox LocalizationManager.GetText("Error loading package document item: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Public Sub SavePackageItemFromForm(ByVal frm As Object, ByVal parentRowIndex As Long, ByVal packageId As String, ByVal updateExisting As Boolean)
    On Error GoTo SaveError

    Dim amountValue As Double
    Dim itemDateValue As Variant
    Dim itemId As String

    If Not ValidatePackageItemForm(frm, amountValue, itemDateValue) Then Exit Sub

    itemId = SavePackageItemRecord( _
        parentRowIndex, _
        packageId, _
        Trim$(frm.txtItemId.Text), _
        Trim$(frm.txtItemDocumentTypeDisplay.Text), _
        Trim$(frm.txtItemDocumentNumber.Text), _
        itemDateValue, _
        amountValue, _
        Trim$(frm.txtItemDescription.Text), _
        Trim$(frm.txtItemNotes.Text), _
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

Public Sub FillPackageItemEditorFromParent(ByVal frm As Object, ByVal parentRowIndex As Long)
    Dim parentTable As ListObject
    Dim parentAmount As Variant

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Sub
    If parentRowIndex < 1 Or parentRowIndex > parentTable.ListRows.Count Then Exit Sub

    Call ClearPackageItemEditor(frm)

    frm.txtItemDocumentTypeDisplay.Text = CStr(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_DOCUMENT_TYPE_COLUMN))
    frm.txtItemDocumentNumber.Text = CStr(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_DOCUMENT_NUMBER_COLUMN))
    frm.txtItemDocumentDate.Text = FormatItemDateValue(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_FRP_DATE_COLUMN))

    parentAmount = GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_AMOUNT_COLUMN)
    If IsNumeric(parentAmount) Then
        frm.txtItemAmount.Text = FormatEditorAmountValue(parentAmount)
    Else
        frm.txtItemAmount.Text = ""
    End If

    frm.txtItemDescription.Text = LocalizationManager.GetText("Copied from package")
    frm.txtItemNotes.Text = ""
    frm.txtItemDocumentTypeDisplay.SetFocus
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
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_DESCRIPTION)), _
        CStr(GetItemCellValue(itemsTable, rowIndex, ITEM_COLUMN_NOTES)), _
        False)
End Function

Public Function SavePackageItemRecord(ByVal parentRowIndex As Long, ByVal packageId As String, ByVal itemId As String, ByVal documentTypeDisplay As String, ByVal documentNumber As String, ByVal itemDateValue As Variant, ByVal amountValue As Double, ByVal itemDescription As String, ByVal itemNotes As String, ByVal updateExisting As Boolean) As String
    Dim parentTable As ListObject
    Dim itemsTable As ListObject
    Dim itemRowIndex As Long

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

    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_PACKAGE_ID, packageId
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DOCUMENT_TYPE_DISPLAY, documentTypeDisplay
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DOCUMENT_TYPE, BuildItemTypeKey(documentTypeDisplay)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DOCUMENT_NUMBER, documentNumber
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DOCUMENT_DATE, itemDateValue
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_AMOUNT, amountValue
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DESCRIPTION, itemDescription
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_NOTES, itemNotes

    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_COUNTERPARTY_NAME, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_COUNTERPARTY_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_COUNTERPARTY_NORMALIZED, GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_COUNTERPARTY_NORMALIZED)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_DIRECTION, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_DIRECTION_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_SERVICE, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_SERVICE_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_EXECUTOR, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_EXECUTOR_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_ORDER_INFO, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_ORDER_INFO_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_FRP_NUMBER, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_FRP_NUMBER_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_FRP_DATE, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_FRP_DATE_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_BASE_DOCUMENT_TYPE, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_DOCUMENT_TYPE_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_BASE_DOCUMENT_NUMBER, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_DOCUMENT_NUMBER_COLUMN)
    SetItemCellValue itemsTable, itemRowIndex, ITEM_COLUMN_BASE_DOCUMENT_DATE, GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_FRP_DATE_COLUMN)

    If Len(Trim$(CStr(GetItemCellValue(itemsTable, itemRowIndex, ITEM_COLUMN_MATCHED_STATUS)))) = 0 Then
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
    frm.txtItemDocumentTypeDisplay.Text = ""
    frm.txtItemDocumentNumber.Text = ""
    frm.txtItemDocumentDate.Text = ""
    frm.txtItemAmount.Text = ""
    frm.txtItemDescription.Text = ""
    frm.txtItemNotes.Text = ""
    If frm.lstPackageItems.listCount > 0 Then frm.lstPackageItems.listIndex = -1
End Sub

Private Sub SelectPackageItemInList(ByVal frm As Object, ByVal itemId As String)
    Dim listIndex As Long

    If frm Is Nothing Then Exit Sub
    If Len(Trim$(itemId)) = 0 Then Exit Sub

    For listIndex = 0 To frm.lstPackageItems.listCount - 1
        If StrComp(CStr(frm.lstPackageItems.List(listIndex, 6)), itemId, vbTextCompare) = 0 Then
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

    Set parentTable = GetParentTable()
    If parentTable Is Nothing Then Exit Sub

    counterparty = CStr(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_COUNTERPARTY_COLUMN))
    parentAmount = FormatItemAmountValue(GetParentSourceValue(parentTable, parentRowIndex, PARENT_SOURCE_AMOUNT_COLUMN))
    childrenTotal = FormatItemAmountValue(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT))
    childCount = GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT)
    amountCheckText = TranslateAmountCheckStatus(GetParentPackageText(parentTable, parentRowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS))

    frm.lblPackageSummary.Caption = LocalizationManager.GetText("Package:") & " " & packageId & vbCrLf & _
        LocalizationManager.GetText("Counterparty:") & " " & counterparty & vbCrLf & _
        LocalizationManager.GetText("Parent Amount:") & " " & parentAmount & "    " & _
        LocalizationManager.GetText("Children Total:") & " " & childrenTotal & "    " & _
        LocalizationManager.GetText("Items:") & " " & childCount & "    " & _
        LocalizationManager.GetText("Amount Check:") & " " & amountCheckText
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

    Set parentTable = GetParentTable()
    Set itemsTable = GetItemsTable()
    If parentTable Is Nothing Or itemsTable Is Nothing Then Exit Sub

    packageId = EnsurePackageIdForParentRow(parentTable, parentRowIndex)
    If Len(packageId) = 0 Then Exit Sub

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

    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_HAS_CHILD_DOCUMENTS, IIf(childCount > 0, "True", "False")
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT, childCount
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT, childrenTotal
    SetParentPackageValue parentTable, parentRowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS, statusValue
    Exit Sub

RefreshError:
    MsgBox LocalizationManager.GetText("Error refreshing package summary: ") & Err.description, vbCritical, LocalizationManager.GetText("Package Documents")
End Sub

Private Function ValidatePackageItemForm(ByVal frm As Object, ByRef amountValue As Double, ByRef itemDateValue As Variant) As Boolean
    ValidatePackageItemForm = False

    If Len(Trim$(frm.txtItemDocumentTypeDisplay.Text)) = 0 Then
        MsgBox LocalizationManager.GetText("Document Type is required."), vbExclamation, LocalizationManager.GetText("Package Documents")
        frm.txtItemDocumentTypeDisplay.SetFocus
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

Private Function ParseShortDateText(ByVal shortDateText As String) As Date
    ParseShortDateText = CDate(Left$(shortDateText, 6) & "20" & Right$(shortDateText, 2))
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









