VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPackageDocuments 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserFormPackageDocuments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPackageDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const DESIGN_FORM_WIDTH As Single = 1020
Private Const DESIGN_FORM_HEIGHT As Single = 560
Private Const SCREEN_POINTS_PER_PIXEL As Double = 72 / 96
Private Const VIEWPORT_WIDTH_RATIO As Double = 0.92
Private Const VIEWPORT_HEIGHT_RATIO As Double = 0.88

Private mParentRowIndex As Long
Private mPackageId As String
Private mDocumentTypeItems As Variant
Private mIsReviewFilterInitializing As Boolean

Public Sub OpenForParentRow(ByVal parentRowIndex As Long, ByVal packageId As String)
    mParentRowIndex = parentRowIndex
    mPackageId = packageId
    Call BindPackageDocumentsForm(Me, mParentRowIndex, mPackageId)
    Me.Show vbModeless
    DoEvents
    Call ApplyPackageItemsColumnLayout
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Package Documents"
    Me.Width = DESIGN_FORM_WIDTH
    Me.Height = DESIGN_FORM_HEIGHT
    Call SetupPackageItemsList
    Call SetupMatchedStatusCombo
    Call SetupReviewFilterCombo
    Call LoadDocumentTypeComboData
    Call LocalizationManager.TranslateForm(Me)
    Call ApplyLocalizedCaptions
    Call ConfigureEditorFields
    Call ApplyFormLayout
    Call HideOptionalEditorControls
    Call ApplyEntryTabOrder
    Call ResizeAndCenterForm
End Sub

Private Sub UserForm_Activate()
    Me.Width = DESIGN_FORM_WIDTH
    Me.Height = DESIGN_FORM_HEIGHT
    Call ApplyLocalizedCaptions
    Call ConfigureEditorFields
    Call ApplyFormLayout
    Call HideOptionalEditorControls
    Call ApplyEntryTabOrder
    Call ResizeAndCenterForm
End Sub

Private Sub lstPackageItems_Click()
    Call LoadSelectedPackageItemIntoForm(Me, mPackageId)
End Sub

Private Sub cmbReviewFilter_Change()
    If mIsReviewFilterInitializing Then Exit Sub
    Call ApplyPackageReviewFilterFromForm(Me, mPackageId)
End Sub

Private Sub cmbMatched1CStatus_Change()
    Call UpdatePackageItemReviewHint(Me)
End Sub

Private Sub cmbItemDocumentTypeDisplay_DropButtonClick()
End Sub

Private Sub btnAddItem_Click()
    Call SavePackageItemFromForm(Me, mParentRowIndex, mPackageId, False)
End Sub

Private Sub btnUpdateItem_Click()
    Call SavePackageItemFromForm(Me, mParentRowIndex, mPackageId, True)
End Sub

Private Sub btnDeleteItem_Click()
    Call DeleteSelectedPackageItemFromForm(Me, mParentRowIndex, mPackageId)
End Sub

Private Sub btnDuplicateItem_Click()
    Call DuplicateSelectedPackageItemFromForm(Me, mParentRowIndex, mPackageId)
End Sub

Private Sub btnFillFromPackage_Click()
    Call FillPackageItemEditorFromParent(Me, mParentRowIndex)
End Sub

Private Sub btnMatchIn1C_Click()
    Call ProvodkaIntegrationModule.ProcessSingleRecord(mParentRowIndex)
    Call BindPackageDocumentsForm(Me, mParentRowIndex, mPackageId)
End Sub

Private Sub btnNextReview_Click()
    Call SelectNextReviewItemFromForm(Me)
End Sub

Private Sub btnFilterAll_Click()
    Call SetPackageReviewFilterFromForm(Me, mPackageId, "All")
End Sub

Private Sub btnFilterPending_Click()
    Call SetPackageReviewFilterFromForm(Me, mPackageId, "Pending")
End Sub

Private Sub btnFilterCandidate_Click()
    Call SetPackageReviewFilterFromForm(Me, mPackageId, "Candidate")
End Sub

Private Sub btnFilterNotFound_Click()
    Call SetPackageReviewFilterFromForm(Me, mPackageId, "Not found")
End Sub

Private Sub btnMarkManual_Click()
    Call MarkSelectedPackageItemManualFromForm(Me, mParentRowIndex, mPackageId)
End Sub

Private Sub btnResetMatch_Click()
    Call ResetSelectedPackageItemMatchFromForm(Me, mParentRowIndex, mPackageId)
End Sub

Private Sub btnClearItem_Click()
    Call ClearPackageItemEditor(Me)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub SetupPackageItemsList()
    With Me.lstPackageItems
        .ColumnCount = 8
        .MultiSelect = fmMultiSelectSingle
    End With
    Call ApplyPackageItemsColumnLayout
End Sub

Private Sub SetupMatchedStatusCombo()
    Call PackageDocumentsManager.PopulateMatchedStatusCombo(Me.cmbMatched1CStatus, "not_checked")
End Sub

Private Sub SetupReviewFilterCombo()
    mIsReviewFilterInitializing = True
    With Me.cmbReviewFilter
        .Clear
        .Style = fmStyleDropDownList
        .AddItem LocalizationManager.GetText("All")
        .AddItem LocalizationManager.GetText("Needs review")
        .AddItem LocalizationManager.GetText("Pending")
        .AddItem LocalizationManager.GetText("Candidate")
        .AddItem LocalizationManager.GetText("Not found")
        .listIndex = 0
    End With
    mIsReviewFilterInitializing = False
End Sub

Private Sub LoadDocumentTypeComboData()
    Dim wsSettings As Worksheet

    On Error GoTo LoadError
    Set wsSettings = ThisWorkbook.Worksheets("Dictionaries")
    On Error GoTo 0

    Call LoadComboData(Me.cmbItemDocumentTypeDisplay, wsSettings, "C:C", "Document Types")
    mDocumentTypeItems = GetComboItems(Me.cmbItemDocumentTypeDisplay)
    Exit Sub

LoadError:
    Me.cmbItemDocumentTypeDisplay.Clear
    mDocumentTypeItems = Empty
End Sub

Private Sub ApplyLocalizedCaptions()
    Me.Caption = LocalizationManager.GetText("Package Documents")
    Me.lblPackageItemsTitle.Caption = LocalizationManager.GetText("Package Items")
    Me.lblReviewFilter.Caption = LocalizationManager.GetText("Review filter")
    Me.btnFilterAll.Caption = LocalizationManager.GetText("All")
    Me.btnFilterCandidate.Caption = LocalizationManager.GetText("Candidate")
    Me.btnFilterPending.Caption = LocalizationManager.GetText("Pending")
    Me.btnFilterNotFound.Caption = LocalizationManager.GetText("Not found")
    Me.btnAddItem.Caption = LocalizationManager.GetText("Add Item")
    Me.btnUpdateItem.Caption = LocalizationManager.GetText("Update Item")
    Me.btnDeleteItem.Caption = LocalizationManager.GetText("Delete Item")
    Me.btnDuplicateItem.Caption = LocalizationManager.GetText("Duplicate")
    Me.btnFillFromPackage.Caption = LocalizationManager.GetText("Apply Package Match")
    Me.btnMatchIn1C.Caption = LocalizationManager.GetText("Match in 1C")
    Me.btnNextReview.Caption = LocalizationManager.GetText("Next review")
    Me.btnMarkManual.Caption = LocalizationManager.GetText("Mark as manual")
    Me.btnResetMatch.Caption = LocalizationManager.GetText("Reset match")
    Me.btnClearItem.Caption = LocalizationManager.GetText("Clear")
    Me.btnClose.Caption = LocalizationManager.GetText("Close")
    Me.lblMatched1COperationNumber.Caption = LocalizationManager.GetText("1C Operation No.")
    Me.lblMatched1COperationDate.Caption = LocalizationManager.GetText("1C Operation Date")
    Me.lblMatched1CStatus.Caption = LocalizationManager.GetText("1C Status")
End Sub

Private Sub ConfigureEditorFields()
    On Error Resume Next
    Me.txtItemDescription.MultiLine = True
    Me.txtItemDescription.WordWrap = True
    Me.txtMatched1CComment.MultiLine = True
    Me.txtMatched1CComment.WordWrap = True
    Me.txtMatched1CComment.EnterKeyBehavior = True
    Me.txtMatched1CComment.ScrollBars = fmScrollBarsVertical
    On Error GoTo 0
End Sub

Private Function GetControlByName(ByVal controlName As String) As Object
    On Error Resume Next
    Set GetControlByName = Me.Controls(controlName)
    On Error GoTo 0
End Function

Private Sub HideOptionalEditorControls()
    Call SetControlVisible("lblItemNotes", False)
    Call SetControlVisible("txtItemNotes", False)
    Call SetControlVisible("lblItemAssetCategory", False)
    Call SetControlVisible("cmbItemAssetCategory", False)
    Call SetControlVisible("lblItemQuantity", False)
    Call SetControlVisible("txtItemQuantity", False)
    Call SetControlVisible("lblItemUnit", False)
    Call SetControlVisible("txtItemUnit", False)
    Call SetControlVisible("lblItemOrderInfo", False)
    Call SetControlVisible("txtItemOrderInfo", False)
    Call SetControlVisible("lblItemFrpNumber", False)
    Call SetControlVisible("txtItemFrpNumber", False)
    Call SetControlVisible("lblItemFrpDate", False)
    Call SetControlVisible("txtItemFrpDate", False)
    Call SetTabStopIfExists("lstPackageItems", False)
    Call SetTabStopIfExists("cmbReviewFilter", False)
    Call SetTabStopIfExists("btnFilterAll", False)
    Call SetTabStopIfExists("btnFilterCandidate", False)
    Call SetTabStopIfExists("btnFilterPending", False)
    Call SetTabStopIfExists("btnFilterNotFound", False)
    Call SetTabStopIfExists("btnAddItem", False)
    Call SetTabStopIfExists("btnUpdateItem", False)
    Call SetTabStopIfExists("btnDeleteItem", False)
    Call SetTabStopIfExists("btnDuplicateItem", False)
    Call SetTabStopIfExists("btnFillFromPackage", False)
    Call SetTabStopIfExists("btnMatchIn1C", False)
    Call SetTabStopIfExists("btnNextReview", False)
    Call SetTabStopIfExists("btnMarkManual", False)
    Call SetTabStopIfExists("btnResetMatch", False)
    Call SetTabStopIfExists("btnClearItem", False)
    Call SetTabStopIfExists("btnClose", False)
    Call SetTabStopIfExists("txtItemId", False)
End Sub

Private Sub SetControlVisible(ByVal controlName As String, ByVal isVisible As Boolean)
    Dim ctrl As Object
    Set ctrl = GetControlByName(controlName)
    If ctrl Is Nothing Then Exit Sub
    ctrl.Visible = isVisible
    On Error Resume Next
    ctrl.TabStop = isVisible
    On Error GoTo 0
End Sub

Private Sub ApplyEntryTabOrder()
    Call SetTabIndexIfExists("cmbItemDocumentTypeDisplay", 0)
    Call SetTabIndexIfExists("txtItemDocumentNumber", 1)
    Call SetTabIndexIfExists("txtItemDocumentDate", 2)
    Call SetTabIndexIfExists("txtItemAmount", 3)
    Call SetTabIndexIfExists("txtItemDescription", 4)
    Call SetTabIndexIfExists("txtMatched1COperationNumber", 5)
    Call SetTabIndexIfExists("txtMatched1COperationDate", 6)
    Call SetTabIndexIfExists("cmbMatched1CStatus", 7)
    Call SetTabIndexIfExists("txtMatched1CComment", 8)
End Sub

Private Sub SetTabIndexIfExists(ByVal controlName As String, ByVal tabIndexValue As Integer)
    Dim ctrl As Object
    Set ctrl = GetControlByName(controlName)
    If ctrl Is Nothing Then Exit Sub
    On Error Resume Next
    ctrl.TabIndex = tabIndexValue
    ctrl.TabStop = True
    On Error GoTo 0
End Sub

Private Sub SetTabStopIfExists(ByVal controlName As String, ByVal tabStopValue As Boolean)
    Dim ctrl As Object
    Set ctrl = GetControlByName(controlName)
    If ctrl Is Nothing Then Exit Sub
    On Error Resume Next
    ctrl.TabStop = tabStopValue
    On Error GoTo 0
End Sub

Private Function GetComboItems(TargetCombo As MSForms.ComboBox) As Variant
    Dim result() As String
    Dim i As Long

    If TargetCombo.listCount = 0 Then
        GetComboItems = Empty
        Exit Function
    End If

    ReDim result(0 To TargetCombo.listCount - 1)
    For i = 0 To TargetCombo.listCount - 1
        result(i) = CStr(TargetCombo.List(i))
    Next i

    GetComboItems = result
End Function

Private Sub ResetDocumentTypeCombo()
    Dim i As Long

    If IsEmpty(mDocumentTypeItems) Then Exit Sub

    Me.cmbItemDocumentTypeDisplay.Clear
    For i = LBound(mDocumentTypeItems) To UBound(mDocumentTypeItems)
        Me.cmbItemDocumentTypeDisplay.AddItem mDocumentTypeItems(i)
    Next i
End Sub

Private Sub LoadComboData(TargetCombo As MSForms.ComboBox, SourceSheet As Worksheet, SourceColumn As String, PrimaryHeader As String)
    Dim searchRange As Range
    Dim headerCell As Range
    Dim currentCell As Range
    Dim startLoad As Boolean

    TargetCombo.Clear
    TargetCombo.Style = fmStyleDropDownCombo
    TargetCombo.MatchRequired = False
    TargetCombo.MatchEntry = fmMatchEntryComplete

    Set searchRange = SourceSheet.Range(SourceColumn)
    Set headerCell = searchRange.Find(What:=PrimaryHeader, LookIn:=xlValues, LookAt:=xlWhole)
    If headerCell Is Nothing Then Exit Sub

    startLoad = False
    For Each currentCell In searchRange.Cells
        If currentCell.Row <= headerCell.Row Then GoTo ContinueLoop

        If Trim$(CStr(currentCell.value)) <> "" Then
            TargetCombo.AddItem CStr(currentCell.value)
            startLoad = True
        ElseIf startLoad Then
            Exit For
        End If
ContinueLoop:
    Next currentCell
End Sub

Private Sub ApplyFormLayout()
    Dim marginX As Single
    Dim currentTop As Single
    Dim contentWidth As Single
    Dim innerWidth As Single
    Dim buttonTop As Single
    Dim reviewButtonTop As Single
    Dim buttonWidth As Single
    Dim buttonGap As Single
    Dim firstRowTop As Single
    Dim secondRowTop As Single
    Dim thirdRowTop As Single
    Dim fourthRowTop As Single
    Dim labelOffset As Single
    Dim filterTop As Single
    Dim comboWidth As Single
    Dim filterLabelWidth As Single
    Dim editorRightEdge As Single
    Dim statusColumnWidth As Single
    Dim operationDateWidth As Single
    Dim operationNumberWidth As Single
    Dim amountWidth As Single
    Dim dateWidth As Single
    Dim numberWidth As Single
    Dim typeWidth As Single
    Dim descriptionGap As Single
    Dim minFormWidth As Single
    Dim rightColumnEdge As Single
    Dim descriptionWidth As Single

    marginX = 12
    labelOffset = 14
    minFormWidth = DESIGN_FORM_WIDTH
    If Me.Width < minFormWidth Then Me.Width = minFormWidth

    innerWidth = DESIGN_FORM_WIDTH - 18
    contentWidth = innerWidth - (marginX * 2)
    editorRightEdge = marginX + contentWidth
    rightColumnEdge = editorRightEdge - 26

    Me.lblPackageSummary.Left = marginX
    Me.lblPackageSummary.Top = 12
    Me.lblPackageSummary.Width = contentWidth
    Me.lblPackageSummary.Height = 68

    currentTop = Me.lblPackageSummary.Top + Me.lblPackageSummary.Height + 6
    filterTop = currentTop
    Me.lblPackageItemsTitle.Left = marginX
    Me.lblPackageItemsTitle.Top = filterTop + 4
    Me.lblPackageItemsTitle.Width = 92

    Me.btnFilterAll.Left = marginX + Me.lblPackageItemsTitle.Width + 10
    Me.btnFilterAll.Top = filterTop
    Me.btnFilterAll.Width = 50
    Me.btnFilterCandidate.Left = Me.btnFilterAll.Left + Me.btnFilterAll.Width + 6
    Me.btnFilterCandidate.Top = filterTop
    Me.btnFilterCandidate.Width = 74
    Me.btnFilterPending.Left = Me.btnFilterCandidate.Left + Me.btnFilterCandidate.Width + 6
    Me.btnFilterPending.Top = filterTop
    Me.btnFilterPending.Width = 74
    Me.btnFilterNotFound.Left = Me.btnFilterPending.Left + Me.btnFilterPending.Width + 6
    Me.btnFilterNotFound.Top = filterTop
    Me.btnFilterNotFound.Width = 86

    comboWidth = 102
    filterLabelWidth = 86
    Me.cmbReviewFilter.Left = rightColumnEdge - comboWidth
    Me.lblReviewFilter.Left = Me.cmbReviewFilter.Left - filterLabelWidth - 6
    Me.lblReviewFilter.Top = filterTop + 4
    Me.lblReviewFilter.Width = filterLabelWidth
    Me.cmbReviewFilter.Top = filterTop
    Me.cmbReviewFilter.Width = comboWidth

    currentTop = filterTop + 30
    Me.lstPackageItems.Left = marginX
    Me.lstPackageItems.Top = currentTop
    Me.lstPackageItems.Width = contentWidth
    Me.lstPackageItems.Height = 112
    Call ApplyPackageItemsColumnLayout

    buttonTop = Me.lstPackageItems.Top + Me.lstPackageItems.Height + 8
    buttonWidth = 76
    buttonGap = 6
    Me.btnAddItem.Left = marginX
    Me.btnAddItem.Top = buttonTop
    Me.btnAddItem.Width = buttonWidth
    Me.btnUpdateItem.Left = Me.btnAddItem.Left + buttonWidth + buttonGap
    Me.btnUpdateItem.Top = buttonTop
    Me.btnUpdateItem.Width = buttonWidth
    Me.btnDeleteItem.Left = Me.btnUpdateItem.Left + buttonWidth + buttonGap
    Me.btnDeleteItem.Top = buttonTop
    Me.btnDeleteItem.Width = buttonWidth
    Me.btnDuplicateItem.Left = Me.btnDeleteItem.Left + buttonWidth + buttonGap
    Me.btnDuplicateItem.Top = buttonTop
    Me.btnDuplicateItem.Width = buttonWidth
    Me.btnFillFromPackage.Left = Me.btnDuplicateItem.Left + buttonWidth + buttonGap
    Me.btnFillFromPackage.Top = buttonTop
    Me.btnFillFromPackage.Width = 108
    Me.btnMatchIn1C.Left = Me.btnFillFromPackage.Left + Me.btnFillFromPackage.Width + buttonGap
    Me.btnMatchIn1C.Top = buttonTop
    Me.btnMatchIn1C.Width = 90

    reviewButtonTop = buttonTop + 30
    Me.btnNextReview.Left = marginX
    Me.btnNextReview.Top = reviewButtonTop
    Me.btnNextReview.Width = 124
    Me.btnMarkManual.Left = Me.btnNextReview.Left + Me.btnNextReview.Width + buttonGap
    Me.btnMarkManual.Top = reviewButtonTop
    Me.btnMarkManual.Width = 132
    Me.btnResetMatch.Left = Me.btnMarkManual.Left + Me.btnMarkManual.Width + buttonGap
    Me.btnResetMatch.Top = reviewButtonTop
    Me.btnResetMatch.Width = 132
    Me.btnClose.Width = 82
    Me.btnClose.Left = rightColumnEdge - Me.btnClose.Width
    Me.btnClose.Top = reviewButtonTop
    Me.btnClearItem.Width = 82
    Me.btnClearItem.Left = Me.btnClose.Left - buttonGap - Me.btnClearItem.Width
    Me.btnClearItem.Top = reviewButtonTop

    firstRowTop = reviewButtonTop + 28
    statusColumnWidth = 118
    operationDateWidth = 124
    operationNumberWidth = 102
    amountWidth = 92
    dateWidth = 82
    numberWidth = 92
    typeWidth = (rightColumnEdge - operationDateWidth - operationNumberWidth - amountWidth - dateWidth - numberWidth - (buttonGap * 5)) - marginX
    If typeWidth < 156 Then typeWidth = 156
    descriptionGap = 8

    Me.lblItemDocumentType.Left = marginX
    Me.lblItemDocumentType.Top = firstRowTop
    Me.cmbItemDocumentTypeDisplay.Left = marginX
    Me.cmbItemDocumentTypeDisplay.Top = firstRowTop + labelOffset
    Me.cmbItemDocumentTypeDisplay.Width = typeWidth

    Me.lblItemDocumentNumber.Left = Me.cmbItemDocumentTypeDisplay.Left + Me.cmbItemDocumentTypeDisplay.Width + buttonGap
    Me.lblItemDocumentNumber.Top = firstRowTop
    Me.txtItemDocumentNumber.Left = Me.lblItemDocumentNumber.Left
    Me.txtItemDocumentNumber.Top = firstRowTop + labelOffset
    Me.txtItemDocumentNumber.Width = numberWidth

    Me.lblItemDocumentDate.Left = Me.txtItemDocumentNumber.Left + Me.txtItemDocumentNumber.Width + buttonGap
    Me.lblItemDocumentDate.Top = firstRowTop
    Me.txtItemDocumentDate.Left = Me.lblItemDocumentDate.Left
    Me.txtItemDocumentDate.Top = firstRowTop + labelOffset
    Me.txtItemDocumentDate.Width = dateWidth

    Me.lblItemAmount.Left = Me.txtItemDocumentDate.Left + Me.txtItemDocumentDate.Width + buttonGap
    Me.lblItemAmount.Top = firstRowTop
    Me.txtItemAmount.Left = Me.lblItemAmount.Left
    Me.txtItemAmount.Top = firstRowTop + labelOffset
    Me.txtItemAmount.Width = amountWidth

    Me.lblMatched1COperationNumber.Left = Me.txtItemAmount.Left + Me.txtItemAmount.Width + buttonGap
    Me.lblMatched1COperationNumber.Top = firstRowTop
    Me.txtMatched1COperationNumber.Left = Me.lblMatched1COperationNumber.Left
    Me.txtMatched1COperationNumber.Top = firstRowTop + labelOffset
    Me.txtMatched1COperationNumber.Width = operationNumberWidth

    Me.txtMatched1COperationDate.Left = rightColumnEdge - operationDateWidth
    Me.lblMatched1COperationDate.Left = Me.txtMatched1COperationDate.Left
    Me.lblMatched1COperationDate.Top = firstRowTop
    Me.txtMatched1COperationDate.Top = firstRowTop + labelOffset
    Me.txtMatched1COperationDate.Width = operationDateWidth

    Me.txtMatched1COperationNumber.Left = Me.txtMatched1COperationDate.Left - buttonGap - operationNumberWidth
    Me.lblMatched1COperationNumber.Left = Me.txtMatched1COperationNumber.Left

    Me.txtItemAmount.Left = Me.txtMatched1COperationNumber.Left - buttonGap - amountWidth
    Me.lblItemAmount.Left = Me.txtItemAmount.Left

    Me.txtItemDocumentDate.Left = Me.txtItemAmount.Left - buttonGap - dateWidth
    Me.lblItemDocumentDate.Left = Me.txtItemDocumentDate.Left

    Me.txtItemDocumentNumber.Left = Me.txtItemDocumentDate.Left - buttonGap - numberWidth
    Me.lblItemDocumentNumber.Left = Me.txtItemDocumentNumber.Left

    secondRowTop = firstRowTop + 48
    descriptionWidth = Me.cmbMatched1CStatus.Left - marginX - descriptionGap
    Me.lblItemDescription.Left = marginX
    Me.lblItemDescription.Top = secondRowTop
    Me.txtItemDescription.Left = marginX
    Me.txtItemDescription.Top = secondRowTop + labelOffset
    Me.txtItemDescription.Width = descriptionWidth
    Me.txtItemDescription.Height = 34

    Me.cmbMatched1CStatus.Left = rightColumnEdge - statusColumnWidth
    Me.lblMatched1CStatus.Left = Me.cmbMatched1CStatus.Left
    Me.lblMatched1CStatus.Top = secondRowTop
    Me.cmbMatched1CStatus.Top = secondRowTop + labelOffset
    Me.cmbMatched1CStatus.Width = statusColumnWidth

    thirdRowTop = secondRowTop + 50
    Me.lblMatched1CComment.Left = marginX
    Me.lblMatched1CComment.Top = thirdRowTop
    Me.txtMatched1CComment.Left = marginX
    Me.txtMatched1CComment.Top = thirdRowTop + labelOffset
    Me.txtMatched1CComment.Width = descriptionWidth
    Me.txtMatched1CComment.Height = 60

    fourthRowTop = thirdRowTop + 78
    Me.lblReviewHint.Left = marginX
    Me.lblReviewHint.Top = fourthRowTop
    Me.lblReviewHint.Width = contentWidth
    Me.lblReviewHint.Height = 22
    Me.lblReviewHint.WordWrap = True

    Me.txtItemId.Left = contentWidth - 90
    Me.txtItemId.Top = Me.lblReviewHint.Top + Me.lblReviewHint.Height + 6
    Me.txtItemId.Width = 72
    Me.txtItemId.Visible = False
End Sub

Private Sub ResizeAndCenterForm()
    Dim screenWidthPoints As Double
    Dim screenHeightPoints As Double
    Dim viewportWidth As Single
    Dim viewportHeight As Single
    Dim contentWidth As Single
    Dim contentHeight As Single
    Dim usableWidth As Single
    Dim usableHeight As Single
    Dim scrollMode As fmScrollBars
    Dim needsHorizontalScroll As Boolean
    Dim needsVerticalScroll As Boolean

    screenWidthPoints = GetSystemMetrics(SM_CXSCREEN) * SCREEN_POINTS_PER_PIXEL
    screenHeightPoints = GetSystemMetrics(SM_CYSCREEN) * SCREEN_POINTS_PER_PIXEL
    usableWidth = CSng(screenWidthPoints * VIEWPORT_WIDTH_RATIO)
    usableHeight = CSng(screenHeightPoints * VIEWPORT_HEIGHT_RATIO)

    Me.ScrollBars = fmScrollBarsNone
    Me.KeepScrollBarsVisible = fmScrollBarsNone
    Me.ScrollLeft = 0
    Me.ScrollTop = 0
    Me.Width = DESIGN_FORM_WIDTH
    Me.Height = DESIGN_FORM_HEIGHT
    Call ApplyFormLayout
    Call HideOptionalEditorControls

    contentWidth = GetFormContentWidth
    contentHeight = GetFormContentHeight

    viewportWidth = contentWidth
    viewportHeight = contentHeight
    If viewportWidth > usableWidth Then viewportWidth = usableWidth
    If viewportHeight > usableHeight Then viewportHeight = usableHeight

    needsHorizontalScroll = (contentWidth > viewportWidth)
    needsVerticalScroll = (contentHeight > viewportHeight)

    scrollMode = fmScrollBarsNone
    If needsHorizontalScroll And needsVerticalScroll Then
        scrollMode = fmScrollBarsBoth
    ElseIf needsHorizontalScroll Then
        scrollMode = fmScrollBarsHorizontal
    ElseIf needsVerticalScroll Then
        scrollMode = fmScrollBarsVertical
    End If

    Me.Width = viewportWidth
    Me.Height = viewportHeight
    Me.ScrollBars = scrollMode
    Me.KeepScrollBarsVisible = scrollMode
    Me.ScrollWidth = contentWidth
    Me.ScrollHeight = contentHeight
    Me.ScrollLeft = 0
    Me.ScrollTop = 0

    Me.StartUpPosition = 0
    Me.Left = (screenWidthPoints - viewportWidth) / 2
    Me.Top = (screenHeightPoints - viewportHeight) / 2
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
    Call ApplyPackageItemsColumnLayout
End Sub

Private Sub ApplyPackageItemsColumnLayout()
    Dim availableWidth As Single
    Dim orderWidth As Single
    Dim numberWidth As Single
    Dim dateWidth As Single
    Dim amountWidth As Single
    Dim statusWidth As Single
    Dim documentTypeWidth As Single

    availableWidth = Me.lstPackageItems.Width - 18
    If availableWidth <= 0 Then Exit Sub

    orderWidth = 24
    numberWidth = 90
    dateWidth = 76
    amountWidth = 84
    statusWidth = 150
    documentTypeWidth = availableWidth - (orderWidth + numberWidth + dateWidth + amountWidth + statusWidth)
    If documentTypeWidth < 220 Then documentTypeWidth = 220

    Me.lstPackageItems.ColumnWidths = _
        orderWidth & " pt;" & _
        documentTypeWidth & " pt;" & _
        numberWidth & " pt;" & _
        dateWidth & " pt;" & _
        amountWidth & " pt;" & _
        statusWidth & " pt;0 pt;0 pt"
End Sub

Private Function GetFormContentWidth() As Single
    Dim ctrl As MSForms.Control
    Dim rightEdge As Single

    rightEdge = 0
    For Each ctrl In Me.Controls
        If ctrl.Visible Then
            If ctrl.Left + ctrl.Width > rightEdge Then rightEdge = ctrl.Left + ctrl.Width
        End If
    Next ctrl

    If rightEdge <= 0 Then rightEdge = DESIGN_FORM_WIDTH
    GetFormContentWidth = rightEdge + 18
End Function

Private Function GetFormContentHeight() As Single
    Dim ctrl As MSForms.Control
    Dim bottomEdge As Single

    bottomEdge = 0
    For Each ctrl In Me.Controls
        If ctrl.Visible Then
            If ctrl.Top + ctrl.Height > bottomEdge Then bottomEdge = ctrl.Top + ctrl.Height
        End If
    Next ctrl

    If bottomEdge <= 0 Then bottomEdge = DESIGN_FORM_HEIGHT
    GetFormContentHeight = bottomEdge + 36
End Function
