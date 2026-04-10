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

Private mParentRowIndex As Long
Private mPackageId As String
Private mDocumentTypeItems As Variant
Private mIsReviewFilterInitializing As Boolean

Public Sub OpenForParentRow(ByVal parentRowIndex As Long, ByVal packageId As String)
    mParentRowIndex = parentRowIndex
    mPackageId = packageId
    Call BindPackageDocumentsForm(Me, mParentRowIndex, mPackageId)
    Me.Show
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Package Documents"
    Me.Width = 900
    Me.Height = 700
    Call SetupPackageItemsList
    Call SetupMatchedStatusCombo
    Call SetupReviewFilterCombo
    Call LoadDocumentTypeComboData
    Call EnsureExtendedEditorControls
    Call SetupAssetCategoryCombo
    Call LocalizationManager.TranslateForm(Me)
    Call ApplyFormLayout
    Call ResizeAndCenterForm
End Sub

Private Sub UserForm_Activate()
    Call ApplyFormLayout
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
        .ColumnWidths = "28 pt;128 pt;96 pt;62 pt;74 pt;84 pt;116 pt;0 pt"
        .MultiSelect = fmMultiSelectSingle
    End With
End Sub

Private Sub SetupMatchedStatusCombo()
    With Me.cmbMatched1CStatus
        .Clear
        .AddItem "not_checked"
        .AddItem "exact"
        .AddItem "candidate"
        .AddItem "manual"
        .AddItem "not_found"
        .Text = "not_checked"
    End With
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

Private Sub EnsureExtendedEditorControls()
    Call EnsureLabelControl("lblItemAssetCategory", "Asset Category")
    Call EnsureComboBoxControl("cmbItemAssetCategory")
    Call EnsureLabelControl("lblItemQuantity", "Quantity")
    Call EnsureTextBoxControl("txtItemQuantity")
    Call EnsureLabelControl("lblItemUnit", "Unit")
    Call EnsureTextBoxControl("txtItemUnit")
    Call EnsureLabelControl("lblItemOrderInfo", "Order Info")
    Call EnsureTextBoxControl("txtItemOrderInfo")
    Call EnsureLabelControl("lblItemFrpNumber", "FRP Number")
    Call EnsureTextBoxControl("txtItemFrpNumber")
    Call EnsureLabelControl("lblItemFrpDate", "FRP Date")
    Call EnsureTextBoxControl("txtItemFrpDate")
End Sub

Private Sub SetupAssetCategoryCombo()
    Dim comboControl As Object

    Set comboControl = GetControlByName("cmbItemAssetCategory")
    If comboControl Is Nothing Then Exit Sub

    With comboControl
        .Clear
        .Style = fmStyleDropDownList
        .AddItem vbNullString
        .AddItem LocalizationManager.GetText("Inventory")
        .AddItem LocalizationManager.GetText("Fixed assets")
        .listIndex = 0
    End With
End Sub

Private Sub EnsureLabelControl(ByVal controlName As String, ByVal captionText As String)
    Dim ctrl As Object

    Set ctrl = GetControlByName(controlName)
    If ctrl Is Nothing Then
        Set ctrl = Me.Controls.Add("Forms.Label.1", controlName, True)
    End If

    ctrl.Caption = captionText
    ctrl.BackStyle = fmBackStyleTransparent
End Sub

Private Sub EnsureTextBoxControl(ByVal controlName As String)
    Dim ctrl As Object

    Set ctrl = GetControlByName(controlName)
    If ctrl Is Nothing Then
        Set ctrl = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    End If

    ctrl.MultiLine = False
    ctrl.EnterKeyBehavior = False
End Sub

Private Sub EnsureComboBoxControl(ByVal controlName As String)
    Dim ctrl As Object

    Set ctrl = GetControlByName(controlName)
    If ctrl Is Nothing Then
        Set ctrl = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    End If
End Sub

Private Function GetControlByName(ByVal controlName As String) As Object
    On Error Resume Next
    Set GetControlByName = Me.Controls(controlName)
    On Error GoTo 0
End Function

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
    Dim buttonTop As Single
    Dim reviewButtonTop As Single
    Dim buttonWidth As Single
    Dim buttonGap As Single
    Dim firstRowTop As Single
    Dim secondRowTop As Single
    Dim thirdRowTop As Single
    Dim fourthRowTop As Single
    Dim notesTop As Single
    Dim labelOffset As Single
    Dim filterTop As Single
    Dim comboWidth As Single
    Dim filterLabelWidth As Single

    marginX = 12
    labelOffset = 14
    contentWidth = Me.Width - (marginX * 2) - 18

    Me.lblPackageSummary.Left = marginX
    Me.lblPackageSummary.Top = 12
    Me.lblPackageSummary.Width = contentWidth
    Me.lblPackageSummary.Height = 74

    currentTop = Me.lblPackageSummary.Top + Me.lblPackageSummary.Height + 8
    filterTop = currentTop
    Me.lblPackageItemsTitle.Left = marginX
    Me.lblPackageItemsTitle.Top = filterTop + 4
    Me.lblPackageItemsTitle.Width = 92

    Me.btnFilterAll.Left = 108
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

    comboWidth = 112
    filterLabelWidth = 86
    Me.cmbReviewFilter.Left = marginX + contentWidth - comboWidth
    Me.cmbReviewFilter.Top = filterTop
    Me.cmbReviewFilter.Width = comboWidth
    Me.lblReviewFilter.Left = Me.cmbReviewFilter.Left - filterLabelWidth - 6
    Me.lblReviewFilter.Top = filterTop + 4
    Me.lblReviewFilter.Width = filterLabelWidth

    currentTop = filterTop + 30
    Me.lstPackageItems.Left = marginX
    Me.lstPackageItems.Top = currentTop
    Me.lstPackageItems.Width = contentWidth
    Me.lstPackageItems.Height = 126

    buttonTop = Me.lstPackageItems.Top + Me.lstPackageItems.Height + 8
    buttonWidth = 82
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
    Me.btnFillFromPackage.Width = 112
    Me.btnMatchIn1C.Left = Me.btnFillFromPackage.Left + Me.btnFillFromPackage.Width + buttonGap
    Me.btnMatchIn1C.Top = buttonTop
    Me.btnMatchIn1C.Width = 94

    reviewButtonTop = buttonTop + 34
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
    Me.btnClose.Left = marginX + contentWidth - Me.btnClose.Width
    Me.btnClose.Top = reviewButtonTop
    Me.btnClearItem.Width = 82
    Me.btnClearItem.Left = Me.btnClose.Left - buttonGap - Me.btnClearItem.Width
    Me.btnClearItem.Top = reviewButtonTop

    firstRowTop = reviewButtonTop + 40
    Me.lblItemDocumentType.Left = marginX
    Me.lblItemDocumentType.Top = firstRowTop
    Me.cmbItemDocumentTypeDisplay.Left = marginX
    Me.cmbItemDocumentTypeDisplay.Top = firstRowTop + labelOffset
    Me.cmbItemDocumentTypeDisplay.Width = 154

    Me.lblItemDocumentNumber.Left = 178
    Me.lblItemDocumentNumber.Top = firstRowTop
    Me.txtItemDocumentNumber.Left = 178
    Me.txtItemDocumentNumber.Top = firstRowTop + labelOffset
    Me.txtItemDocumentNumber.Width = 118

    Me.lblItemDocumentDate.Left = 310
    Me.lblItemDocumentDate.Top = firstRowTop
    Me.txtItemDocumentDate.Left = 310
    Me.txtItemDocumentDate.Top = firstRowTop + labelOffset
    Me.txtItemDocumentDate.Width = 82

    Me.lblItemAmount.Left = 406
    Me.lblItemAmount.Top = firstRowTop
    Me.txtItemAmount.Left = 406
    Me.txtItemAmount.Top = firstRowTop + labelOffset
    Me.txtItemAmount.Width = 74

    GetControlByName("lblItemAssetCategory").Left = 494
    GetControlByName("lblItemAssetCategory").Top = firstRowTop
    GetControlByName("cmbItemAssetCategory").Left = 494
    GetControlByName("cmbItemAssetCategory").Top = firstRowTop + labelOffset
    GetControlByName("cmbItemAssetCategory").Width = 118

    Me.lblMatched1COperationNumber.Left = 624
    Me.lblMatched1COperationNumber.Top = firstRowTop
    Me.txtMatched1COperationNumber.Left = 624
    Me.txtMatched1COperationNumber.Top = firstRowTop + labelOffset
    Me.txtMatched1COperationNumber.Width = 112

    Me.lblMatched1COperationDate.Left = 750
    Me.lblMatched1COperationDate.Top = firstRowTop
    Me.txtMatched1COperationDate.Left = 750
    Me.txtMatched1COperationDate.Top = firstRowTop + labelOffset
    Me.txtMatched1COperationDate.Width = 92

    secondRowTop = firstRowTop + 56
    GetControlByName("lblItemQuantity").Left = marginX
    GetControlByName("lblItemQuantity").Top = secondRowTop
    GetControlByName("txtItemQuantity").Left = marginX
    GetControlByName("txtItemQuantity").Top = secondRowTop + labelOffset
    GetControlByName("txtItemQuantity").Width = 88

    GetControlByName("lblItemUnit").Left = 114
    GetControlByName("lblItemUnit").Top = secondRowTop
    GetControlByName("txtItemUnit").Left = 114
    GetControlByName("txtItemUnit").Top = secondRowTop + labelOffset
    GetControlByName("txtItemUnit").Width = 88

    GetControlByName("lblItemOrderInfo").Left = 228
    GetControlByName("lblItemOrderInfo").Top = secondRowTop
    GetControlByName("txtItemOrderInfo").Left = 228
    GetControlByName("txtItemOrderInfo").Top = secondRowTop + labelOffset
    GetControlByName("txtItemOrderInfo").Width = 246

    GetControlByName("lblItemFrpNumber").Left = 500
    GetControlByName("lblItemFrpNumber").Top = secondRowTop
    GetControlByName("txtItemFrpNumber").Left = 500
    GetControlByName("txtItemFrpNumber").Top = secondRowTop + labelOffset
    GetControlByName("txtItemFrpNumber").Width = 156

    GetControlByName("lblItemFrpDate").Left = 682
    GetControlByName("lblItemFrpDate").Top = secondRowTop
    GetControlByName("txtItemFrpDate").Left = 682
    GetControlByName("txtItemFrpDate").Top = secondRowTop + labelOffset
    GetControlByName("txtItemFrpDate").Width = 92

    thirdRowTop = secondRowTop + 56
    Me.lblItemDescription.Left = marginX
    Me.lblItemDescription.Top = thirdRowTop
    Me.txtItemDescription.Left = marginX
    Me.txtItemDescription.Top = thirdRowTop + labelOffset
    Me.txtItemDescription.Width = 406
    Me.txtItemDescription.Height = 42

    Me.lblMatched1CStatus.Left = 440
    Me.lblMatched1CStatus.Top = thirdRowTop
    Me.cmbMatched1CStatus.Left = 440
    Me.cmbMatched1CStatus.Top = thirdRowTop + labelOffset
    Me.cmbMatched1CStatus.Width = 120

    Me.lblMatched1CComment.Left = 582
    Me.lblMatched1CComment.Top = thirdRowTop
    Me.txtMatched1CComment.Left = 582
    Me.txtMatched1CComment.Top = thirdRowTop + labelOffset
    Me.txtMatched1CComment.Width = 260
    Me.txtMatched1CComment.Height = 42

    fourthRowTop = thirdRowTop + 64
    Me.lblReviewHint.Left = marginX
    Me.lblReviewHint.Top = fourthRowTop
    Me.lblReviewHint.Width = contentWidth
    Me.lblReviewHint.Height = 22
    Me.lblReviewHint.WordWrap = True

    notesTop = fourthRowTop + 32
    Me.lblItemNotes.Left = marginX
    Me.lblItemNotes.Top = notesTop
    Me.txtItemNotes.Left = marginX
    Me.txtItemNotes.Top = notesTop + labelOffset
    Me.txtItemNotes.Width = contentWidth
    Me.txtItemNotes.Height = 66

    Me.txtItemId.Left = contentWidth - 90
    Me.txtItemId.Top = Me.txtItemNotes.Top + Me.txtItemNotes.Height + 6
    Me.txtItemId.Width = 72
    Me.txtItemId.Visible = False
End Sub

Private Sub ResizeAndCenterForm()
    Dim screenWidthPoints As Double
    Dim screenHeightPoints As Double
    Dim maxWidth As Double
    Dim maxHeight As Double

    screenWidthPoints = GetSystemMetrics(SM_CXSCREEN) * (72 / 96)
    screenHeightPoints = GetSystemMetrics(SM_CYSCREEN) * (72 / 96)
    maxWidth = screenWidthPoints * 0.9
    maxHeight = screenHeightPoints * 0.85

    If Me.Width > maxWidth Then Me.Width = maxWidth
    If Me.Height > maxHeight Then Me.Height = maxHeight

    Me.StartUpPosition = 0
    Me.Left = (screenWidthPoints - Me.Width) / 2
    Me.Top = (screenHeightPoints - Me.Height) / 2
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
End Sub

