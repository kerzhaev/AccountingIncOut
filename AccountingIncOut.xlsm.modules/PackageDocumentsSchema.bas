Attribute VB_Name = "PackageDocumentsSchema"
Option Explicit

Public Const PACKAGE_PARENT_TABLE_NAME As String = "TableIncOut"
Public Const PACKAGE_ITEMS_SHEET_NAME As String = "IncOutItems"
Public Const PACKAGE_ITEMS_TABLE_NAME As String = "TableIncOutItems"

Public Const PACKAGE_COLUMN_PACKAGE_ID As String = "PackageId"
Public Const PACKAGE_COLUMN_PACKAGE_TYPE As String = "PackageType"
Public Const PACKAGE_COLUMN_ASSET_CATEGORY As String = "AssetCategory"
Public Const PACKAGE_COLUMN_DOCUMENT_STAGE As String = "DocumentStage"
Public Const PACKAGE_COLUMN_COUNTERPARTY_NORMALIZED As String = "CounterpartyNormalized"
Public Const PACKAGE_COLUMN_HAS_CHILD_DOCUMENTS As String = "HasChildDocuments"
Public Const PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT As String = "ChildDocumentsCount"
Public Const PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT As String = "ChildrenTotalAmount"
Public Const PACKAGE_COLUMN_AMOUNT_CHECK_STATUS As String = "PackageAmountCheckStatus"
Public Const PACKAGE_COLUMN_PRIMARY_1C_NUMBER As String = "Primary1COperationNumber"
Public Const PACKAGE_COLUMN_PRIMARY_1C_DATE As String = "Primary1COperationDate"
Public Const PACKAGE_COLUMN_PRIMARY_1C_STATUS As String = "Primary1CMatchStatus"

Public Sub EnsurePackageDocumentsSchema()
    On Error GoTo SchemaError

    Dim wsParent As Worksheet
    Dim parentTable As ListObject

    Set wsParent = CommonUtilities.GetWorksheetSafe("IncOut")
    If wsParent Is Nothing Then Exit Sub

    Set parentTable = CommonUtilities.GetListObjectSafe(wsParent, PACKAGE_PARENT_TABLE_NAME)
    If parentTable Is Nothing Then Exit Sub

    Call EnsureParentPackageColumns(parentTable)
    Call EnsurePackageItemsStorage
    Exit Sub

SchemaError:
    Debug.Print "EnsurePackageDocumentsSchema error: " & Err.Description
End Sub

Public Sub ApplyPackageDefaultsToParentRow(ByVal parentTable As ListObject, ByVal rowIndex As Long)
    On Error GoTo ApplyError

    Dim packageId As String
    Dim counterpartyText As String

    If parentTable Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub
    If rowIndex > parentTable.ListRows.Count Then Exit Sub

    packageId = Trim$(GetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_PACKAGE_ID))
    If Len(packageId) = 0 Then
        packageId = CreatePackageId()
        Call SetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_PACKAGE_ID, packageId)
    End If

    If Len(Trim$(GetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_HAS_CHILD_DOCUMENTS))) = 0 Then
        Call SetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_HAS_CHILD_DOCUMENTS, "False")
    End If

    If Len(Trim$(GetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT))) = 0 Then
        Call SetParentNumericValue(parentTable, rowIndex, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT, 0)
    End If

    If Len(Trim$(GetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT))) = 0 Then
        Call SetParentNumericValue(parentTable, rowIndex, PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT, 0)
    End If

    If Len(Trim$(GetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS))) = 0 Then
        Call SetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS, "not_checked")
    End If

    If Len(Trim$(GetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_PRIMARY_1C_STATUS))) = 0 Then
        Call SetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_PRIMARY_1C_STATUS, "not_checked")
    End If

    counterpartyText = GetSourceCounterpartyValue(parentTable, rowIndex)
    If Len(counterpartyText) > 0 Then
        Call SetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_COUNTERPARTY_NORMALIZED, NormalizeCounterparty(counterpartyText))
    End If

    Exit Sub

ApplyError:
    Debug.Print "ApplyPackageDefaultsToParentRow error: " & Err.Description
End Sub

Public Function EnsurePackageIdForParentRow(ByVal parentTable As ListObject, ByVal rowIndex As Long) As String
    On Error GoTo EnsureError

    Call ApplyPackageDefaultsToParentRow(parentTable, rowIndex)
    EnsurePackageIdForParentRow = GetParentTextValue(parentTable, rowIndex, PACKAGE_COLUMN_PACKAGE_ID)
    Exit Function

EnsureError:
    EnsurePackageIdForParentRow = ""
End Function

Private Sub EnsureParentPackageColumns(ByVal parentTable As ListObject)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_PACKAGE_ID)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_PACKAGE_TYPE)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_ASSET_CATEGORY)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_DOCUMENT_STAGE)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_COUNTERPARTY_NORMALIZED)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_HAS_CHILD_DOCUMENTS)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_CHILD_DOCUMENTS_COUNT)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_CHILDREN_TOTAL_AMOUNT)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_AMOUNT_CHECK_STATUS)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_PRIMARY_1C_NUMBER)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_PRIMARY_1C_DATE)
    Call EnsureTableColumn(parentTable, PACKAGE_COLUMN_PRIMARY_1C_STATUS)
End Sub

Private Sub EnsurePackageItemsStorage()
    Dim ws As Worksheet
    Dim itemsTable As ListObject

    Set ws = CommonUtilities.GetWorksheetSafe(PACKAGE_ITEMS_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PACKAGE_ITEMS_SHEET_NAME
    End If

    Set itemsTable = CommonUtilities.GetListObjectSafe(ws, PACKAGE_ITEMS_TABLE_NAME)
    If itemsTable Is Nothing Then
        Call CreatePackageItemsTable(ws)
        Set itemsTable = CommonUtilities.GetListObjectSafe(ws, PACKAGE_ITEMS_TABLE_NAME)
    End If

    If Not itemsTable Is Nothing Then
        Call EnsurePackageItemsColumns(itemsTable)
    End If

    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub CreatePackageItemsTable(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim tableRange As Range
    Dim itemsTable As ListObject
    Dim lastHeaderColumn As Long
    Dim i As Long

    headers = GetPackageItemsHeaders()
    lastHeaderColumn = UBound(headers) + 1

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = CStr(headers(i))
    Next i

    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(2, lastHeaderColumn))

    Set itemsTable = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    itemsTable.Name = PACKAGE_ITEMS_TABLE_NAME

    On Error Resume Next
    If Not itemsTable.DataBodyRange Is Nothing Then itemsTable.DataBodyRange.Rows.Delete
    On Error GoTo 0
End Sub

Private Sub EnsurePackageItemsColumns(ByVal itemsTable As ListObject)
    Dim headers As Variant
    Dim i As Long

    headers = GetPackageItemsHeaders()
    For i = LBound(headers) To UBound(headers)
        Call EnsureTableColumn(itemsTable, CStr(headers(i)))
    Next i
End Sub

Private Function GetPackageItemsHeaders() As Variant
    Dim headers(0 To 32) As String

    headers(0) = "ItemId"
    headers(1) = "PackageId"
    headers(2) = "ItemOrder"
    headers(3) = "ItemDocumentType"
    headers(4) = "ItemDocumentTypeDisplay"
    headers(5) = "ItemDocumentNumber"
    headers(6) = "ItemDocumentDate"
    headers(7) = "ItemAmount"
    headers(8) = "CounterpartyName"
    headers(9) = "CounterpartyNormalized"
    headers(10) = "Direction"
    headers(11) = "Service"
    headers(12) = "Executor"
    headers(13) = "OrderInfo"
    headers(14) = "FRPNumber"
    headers(15) = "FRPDate"
    headers(16) = "ItemAssetCategory"
    headers(17) = "ItemDescription"
    headers(18) = "ItemQuantity"
    headers(19) = "ItemUnit"
    headers(20) = "BaseDocumentType"
    headers(21) = "BaseDocumentNumber"
    headers(22) = "BaseDocumentDate"
    headers(23) = "Matched1COperationNumber"
    headers(24) = "Matched1COperationDate"
    headers(25) = "Matched1CMatchStatus"
    headers(26) = "Matched1CConfidence"
    headers(27) = "Matched1CComment"
    headers(28) = "Matched1CMode"
    headers(29) = "IsPostedSeparately"
    headers(30) = "Notes"
    headers(31) = "CreatedAt"
    headers(32) = "UpdatedAt"

    GetPackageItemsHeaders = headers
End Function

Private Sub EnsureTableColumn(ByVal targetTable As ListObject, ByVal columnName As String)
    If targetTable Is Nothing Then Exit Sub
    If Len(Trim$(columnName)) = 0 Then Exit Sub

    If FindListColumn(targetTable, columnName) Is Nothing Then
        targetTable.ListColumns.Add.Name = columnName
    End If
End Sub

Private Function FindListColumn(ByVal targetTable As ListObject, ByVal columnName As String) As ListColumn
    Dim listColumn As ListColumn

    If targetTable Is Nothing Then Exit Function

    For Each listColumn In targetTable.ListColumns
        If StrComp(Trim$(listColumn.Name), Trim$(columnName), vbTextCompare) = 0 Then
            Set FindListColumn = listColumn
            Exit Function
        End If
    Next listColumn
End Function

Private Function GetParentTextValue(ByVal parentTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim listColumn As ListColumn

    Set listColumn = FindListColumn(parentTable, columnName)
    If listColumn Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > parentTable.ListRows.Count Then Exit Function

    GetParentTextValue = Trim$(CStr(parentTable.DataBodyRange.Cells(rowIndex, listColumn.Index).Value))
End Function

Private Sub SetParentTextValue(ByVal parentTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal textValue As String)
    Dim listColumn As ListColumn

    Set listColumn = FindListColumn(parentTable, columnName)
    If listColumn Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > parentTable.ListRows.Count Then Exit Sub

    parentTable.DataBodyRange.Cells(rowIndex, listColumn.Index).Value = textValue
End Sub

Private Sub SetParentNumericValue(ByVal parentTable As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal numericValue As Double)
    Dim listColumn As ListColumn

    Set listColumn = FindListColumn(parentTable, columnName)
    If listColumn Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > parentTable.ListRows.Count Then Exit Sub

    parentTable.DataBodyRange.Cells(rowIndex, listColumn.Index).Value = numericValue
End Sub

Private Function CreatePackageId() As String
    Randomize
    CreatePackageId = "PKG-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int((Rnd() * 9000) + 1000), "0000")
End Function

Private Function GetSourceCounterpartyValue(ByVal parentTable As ListObject, ByVal rowIndex As Long) As String
    If rowIndex < 1 Or rowIndex > parentTable.ListRows.Count Then Exit Function
    If parentTable.ListColumns.Count < 9 Then Exit Function

    GetSourceCounterpartyValue = Trim$(CStr(parentTable.DataBodyRange.Cells(rowIndex, 9).Value))
End Function

Private Function NormalizeCounterparty(ByVal sourceText As String) As String
    Dim normalized As String

    normalized = UCase$(Trim$(sourceText))
    Do While InStr(normalized, "  ") > 0
        normalized = Replace(normalized, "  ", " ")
    Loop

    NormalizeCounterparty = normalized
End Function


