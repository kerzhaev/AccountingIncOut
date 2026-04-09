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

Public Sub OpenForParentRow(ByVal parentRowIndex As Long, ByVal packageId As String)
    mParentRowIndex = parentRowIndex
    mPackageId = packageId
    Call BindPackageDocumentsForm(Me, mParentRowIndex, mPackageId)
    Me.Show
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Package Documents"
    Me.Width = 760
    Me.Height = 520
    Call SetupPackageItemsList
    Call SetupMatchedStatusCombo
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

Private Sub ApplyFormLayout()
    Dim marginX As Single
    Dim currentTop As Single
    Dim contentWidth As Single
    Dim buttonTop As Single
    Dim buttonWidth As Single
    Dim buttonGap As Single
    Dim firstRowTop As Single
    Dim secondRowTop As Single
    Dim thirdRowTop As Single
    Dim labelOffset As Single

    marginX = 12
    labelOffset = 14
    contentWidth = Me.Width - (marginX * 2) - 12

    Me.lblPackageSummary.Left = marginX
    Me.lblPackageSummary.Top = 12
    Me.lblPackageSummary.Width = contentWidth
    Me.lblPackageSummary.Height = 54

    currentTop = Me.lblPackageSummary.Top + Me.lblPackageSummary.Height + 8
    Me.lblPackageItemsTitle.Left = marginX
    Me.lblPackageItemsTitle.Top = currentTop
    Me.lblPackageItemsTitle.Width = 180

    currentTop = Me.lblPackageItemsTitle.Top + 16
    Me.lstPackageItems.Left = marginX
    Me.lstPackageItems.Top = currentTop
    Me.lstPackageItems.Width = contentWidth
    Me.lstPackageItems.Height = 120

    buttonTop = Me.lstPackageItems.Top + Me.lstPackageItems.Height + 8
    buttonWidth = 86
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
    Me.btnFillFromPackage.Width = 110
    Me.btnClearItem.Left = Me.btnFillFromPackage.Left + Me.btnFillFromPackage.Width + buttonGap
    Me.btnClearItem.Top = buttonTop
    Me.btnClearItem.Width = 86
    Me.btnClose.Left = Me.btnClearItem.Left + Me.btnClearItem.Width + buttonGap
    Me.btnClose.Top = buttonTop
    Me.btnClose.Width = 86

    firstRowTop = buttonTop + 44
    Me.lblItemDocumentType.Left = marginX
    Me.lblItemDocumentType.Top = firstRowTop
    Me.txtItemDocumentTypeDisplay.Left = marginX
    Me.txtItemDocumentTypeDisplay.Top = firstRowTop + labelOffset
    Me.txtItemDocumentTypeDisplay.Width = 170

    Me.lblItemDocumentNumber.Left = 194
    Me.lblItemDocumentNumber.Top = firstRowTop
    Me.txtItemDocumentNumber.Left = 194
    Me.txtItemDocumentNumber.Top = firstRowTop + labelOffset
    Me.txtItemDocumentNumber.Width = 120

    Me.lblItemDocumentDate.Left = 330
    Me.lblItemDocumentDate.Top = firstRowTop
    Me.txtItemDocumentDate.Left = 330
    Me.txtItemDocumentDate.Top = firstRowTop + labelOffset
    Me.txtItemDocumentDate.Width = 82

    Me.lblItemAmount.Left = 426
    Me.lblItemAmount.Top = firstRowTop
    Me.txtItemAmount.Left = 426
    Me.txtItemAmount.Top = firstRowTop + labelOffset
    Me.txtItemAmount.Width = 78

    Me.lblMatched1COperationNumber.Left = 522
    Me.lblMatched1COperationNumber.Top = firstRowTop
    Me.txtMatched1COperationNumber.Left = 522
    Me.txtMatched1COperationNumber.Top = firstRowTop + labelOffset
    Me.txtMatched1COperationNumber.Width = 102

    Me.lblMatched1COperationDate.Left = 636
    Me.lblMatched1COperationDate.Top = firstRowTop
    Me.txtMatched1COperationDate.Left = 636
    Me.txtMatched1COperationDate.Top = firstRowTop + labelOffset
    Me.txtMatched1COperationDate.Width = 84

    secondRowTop = firstRowTop + 50
    Me.lblItemDescription.Left = marginX
    Me.lblItemDescription.Top = secondRowTop
    Me.txtItemDescription.Left = marginX
    Me.txtItemDescription.Top = secondRowTop + labelOffset
    Me.txtItemDescription.Width = 360
    Me.txtItemDescription.Height = 42

    Me.lblMatched1CStatus.Left = 390
    Me.lblMatched1CStatus.Top = secondRowTop
    Me.cmbMatched1CStatus.Left = 390
    Me.cmbMatched1CStatus.Top = secondRowTop + labelOffset
    Me.cmbMatched1CStatus.Width = 110

    Me.lblMatched1CComment.Left = 516
    Me.lblMatched1CComment.Top = secondRowTop
    Me.txtMatched1CComment.Left = 516
    Me.txtMatched1CComment.Top = secondRowTop + labelOffset
    Me.txtMatched1CComment.Width = 204
    Me.txtMatched1CComment.Height = 42

    thirdRowTop = secondRowTop + 64
    Me.lblItemNotes.Left = marginX
    Me.lblItemNotes.Top = thirdRowTop
    Me.txtItemNotes.Left = marginX
    Me.txtItemNotes.Top = thirdRowTop + labelOffset
    Me.txtItemNotes.Width = contentWidth
    Me.txtItemNotes.Height = 52

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

