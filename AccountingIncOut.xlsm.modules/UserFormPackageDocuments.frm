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
    Me.Width = 615
    Me.Height = 385
    Call SetupPackageItemsList
    Call LocalizationManager.TranslateForm(Me)
    Call ResizeAndCenterForm
End Sub

Private Sub UserForm_Activate()
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
        .ColumnCount = 7
        .ColumnWidths = "24 pt;110 pt;90 pt;60 pt;70 pt;80 pt;0 pt"
        .MultiSelect = fmMultiSelectSingle
    End With
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

