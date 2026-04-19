VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormLegacyMatchReview 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserFormLegacyMatchReview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormLegacyMatchReview"
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
Private Const DESIGN_FORM_WIDTH As Single = 860
Private Const DESIGN_FORM_HEIGHT As Single = 560
Private Const SCREEN_POINTS_PER_PIXEL As Double = 72 / 96
Private Const VIEWPORT_WIDTH_RATIO As Double = 0.92
Private Const VIEWPORT_HEIGHT_RATIO As Double = 0.88
Private Const REVIEW_MODE_MATCH As String = "legacy_match"
Private Const REVIEW_MODE_BACKFILL As String = "legacy_backfill"

Private mReviewMode As String
Private mSuspendCandidateClick As Boolean
Private WithEvents mOpenSourceButton As MSForms.CommandButton

Public Sub InitializeForReview()
    mReviewMode = REVIEW_MODE_MATCH
    Call LocalizationManager.TranslateForm(Me)
    Call ConfigureEditorFields
    Call ApplyFormLayout
    Call ApplyEntryTabOrder
    Call LegacyMatchReviewManager.BindLegacyReviewForm(Me, 0)
    Call ResizeAndCenterForm
End Sub

Public Sub InitializeForBackfillReview()
    mReviewMode = REVIEW_MODE_BACKFILL
    Call LocalizationManager.TranslateForm(Me)
    Call ApplyBackfillCaptions
    Call ConfigureEditorFields
    Call ApplyFormLayout
    Call ApplyEntryTabOrder
    Call LegacyPackageBackfillManager.BindLegacyPackageBackfillForm(Me, 0)
    Call ResizeAndCenterForm
End Sub

Private Sub UserForm_Initialize()
    If Len(mReviewMode) = 0 Then mReviewMode = REVIEW_MODE_MATCH
    Me.Caption = "Legacy Match Review"
    Me.Width = DESIGN_FORM_WIDTH
    Me.Height = DESIGN_FORM_HEIGHT
End Sub

Private Sub UserForm_Activate()
    Me.Width = DESIGN_FORM_WIDTH
    Me.Height = DESIGN_FORM_HEIGHT
    Call LocalizationManager.TranslateForm(Me)
    If mReviewMode = REVIEW_MODE_BACKFILL Then Call ApplyBackfillCaptions
    Call ConfigureEditorFields
    Call ApplyFormLayout
    Call ApplyEntryTabOrder
    Call ResizeAndCenterForm
End Sub

Private Sub UserForm_Terminate()
    Set mOpenSourceButton = Nothing
End Sub

Private Sub lstCandidates_Click()
    If mSuspendCandidateClick Then Exit Sub

    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Call LegacyPackageBackfillManager.SelectLegacyPackageProposalFromForm(Me)
    Else
        Call LegacyMatchReviewManager.SelectLegacyCandidateFromForm(Me)
    End If
End Sub

Public Sub SetCandidateClickSuspended(ByVal isSuspended As Boolean)
    mSuspendCandidateClick = isSuspended
End Sub

Private Sub chkUseBestCandidate_Click()
    If Me.chkUseBestCandidate.value Then
        If mReviewMode = REVIEW_MODE_MATCH Then
            Me.txtSelectedOperationNumber.Text = vbNullString
            Me.txtSelectedOperationDate.Text = vbNullString
        End If
        Me.txtCandidateComment.Text = Me.txtBestCandidateComment.Text
    End If
End Sub

Private Sub btnApply_Click()
    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Call LegacyPackageBackfillManager.ApplyLegacyPackageBackfillFromForm(Me, False)
    Else
        Call LegacyMatchReviewManager.ApplyLegacyReviewFromForm(Me, False)
    End If
End Sub

Private Sub btnApplyNext_Click()
    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Call LegacyPackageBackfillManager.ApplyLegacyPackageBackfillFromForm(Me, True)
    Else
        Call LegacyMatchReviewManager.ApplyLegacyReviewFromForm(Me, True)
    End If
End Sub

Private Sub btnNextPending_Click()
    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Call LegacyPackageBackfillManager.MoveLegacyPackageBackfillFormNext(Me)
    Else
        Call LegacyMatchReviewManager.MoveLegacyReviewFormNext(Me)
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub mOpenSourceButton_Click()
    If mReviewMode <> REVIEW_MODE_BACKFILL Then Exit Sub
    Call LegacyPackageBackfillManager.OpenCurrentBackfillIncOutRow(Me)
End Sub

Private Sub ConfigureEditorFields()
    On Error Resume Next
    Me.lstCandidates.RowSource = vbNullString
    Me.lstCandidates.ColumnCount = 1
    Me.lstCandidates.Font.Name = "Segoe UI"
    Me.lstCandidates.Font.Size = 10
    Me.txtQueueSummary.Locked = True
    Me.txtQueueSummary.MultiLine = True
    Me.txtQueueSummary.WordWrap = True
    Me.txtCounterparty.MultiLine = True
    Me.txtCounterparty.WordWrap = True
    Me.txtBestCandidateComment.MultiLine = True
    Me.txtBestCandidateComment.WordWrap = True
    Me.txtBestCandidateComment.EnterKeyBehavior = True
    Me.txtBestCandidateComment.ScrollBars = fmScrollBarsVertical
    Me.txtCandidateComment.MultiLine = True
    Me.txtCandidateComment.WordWrap = True
    Me.txtCandidateComment.EnterKeyBehavior = True
    Me.txtCandidateComment.ScrollBars = fmScrollBarsVertical
    Me.txtReviewRowIndex.Visible = False
    MakeReadOnly Me.txtQueueSummary
    MakeReadOnly Me.txtReviewId
    MakeReadOnly Me.txtIncOutRow
    MakeReadOnly Me.txtRecordNumber
    MakeReadOnly Me.txtService
    MakeReadOnly Me.txtDocumentType
    MakeReadOnly Me.txtDocumentNumber
    MakeReadOnly Me.txtDocumentDate
    MakeReadOnly Me.txtAmount
    MakeReadOnly Me.txtCounterparty
    MakeReadOnly Me.txtBestCandidateNumber
    MakeReadOnly Me.txtBestCandidateDate
    MakeReadOnly Me.txtCurrentStatus
    MakeReadOnly Me.txtBestCandidateComment
    MakeReadOnly Me.txtCandidateComment
    On Error GoTo 0

    Call EnsureBackfillOpenRowButton

    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Me.chkUseBestCandidate.Visible = False
        Call ApplyBackfillFieldColors
    Else
        Me.chkUseBestCandidate.Visible = True
        Call ResetMatchFieldColors
    End If
End Sub

Private Sub ApplyBackfillFieldColors()
    Call SetControlBackColorIfExists("txtQueueSummary", RGB(245, 245, 245))
    Call SetControlBackColorIfExists("txtCurrentStatus", RGB(245, 245, 245))
    Call SetControlBackColorIfExists("txtBestCandidateNumber", RGB(245, 245, 245))
    Call SetControlBackColorIfExists("txtBestCandidateDate", RGB(245, 245, 245))
    Call SetControlBackColorIfExists("txtBestCandidateComment", RGB(245, 245, 245))
    Call SetControlBackColorIfExists("txtSelectedOperationNumber", RGB(255, 250, 205))
    Call SetControlBackColorIfExists("txtSelectedOperationDate", RGB(255, 250, 205))
    Call SetControlBackColorIfExists("txtCandidateComment", RGB(255, 250, 205))
End Sub

Private Sub ResetMatchFieldColors()
    Call SetControlBackColorIfExists("txtQueueSummary", vbWhite)
    Call SetControlBackColorIfExists("txtCurrentStatus", vbWhite)
    Call SetControlBackColorIfExists("txtBestCandidateNumber", vbWhite)
    Call SetControlBackColorIfExists("txtBestCandidateDate", vbWhite)
    Call SetControlBackColorIfExists("txtBestCandidateComment", vbWhite)
    Call SetControlBackColorIfExists("txtSelectedOperationNumber", vbWhite)
    Call SetControlBackColorIfExists("txtSelectedOperationDate", vbWhite)
    Call SetControlBackColorIfExists("txtCandidateComment", vbWhite)
End Sub

Private Sub SetControlBackColorIfExists(ByVal controlName As String, ByVal backColorValue As Long)
    Dim ctrl As Object

    On Error Resume Next
    Set ctrl = Me.Controls(controlName)
    If Not ctrl Is Nothing Then ctrl.BackColor = backColorValue
    On Error GoTo 0
End Sub

Private Sub MakeReadOnly(ByVal targetControl As MSForms.control)
    On Error Resume Next
    targetControl.Locked = True
    targetControl.Enabled = True
    targetControl.TabStop = False
    On Error GoTo 0
End Sub

Private Sub ApplyEntryTabOrder()
    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Call SetTabIndexIfExists("chkUseBestCandidate", 0, False)
        Call SetTabIndexIfExists("lstCandidates", 1, True)
        Call SetTabIndexIfExists("btnApply", 2, True)
        Call SetTabIndexIfExists("btnApplyNext", 3, True)
        Call SetTabIndexIfExists("btnNextPending", 4, True)
        Call SetTabIndexIfExists("btnOpenIncOutRow", 5, True)
        Call SetTabIndexIfExists("btnClose", 6, True)
    Else
        Call SetTabIndexIfExists("chkUseBestCandidate", 0, True)
        Call SetTabIndexIfExists("txtSelectedOperationNumber", 1, True)
        Call SetTabIndexIfExists("txtSelectedOperationDate", 2, True)
        Call SetTabIndexIfExists("lstCandidates", 3, True)
        Call SetTabIndexIfExists("btnApply", 4, True)
        Call SetTabIndexIfExists("btnApplyNext", 5, True)
        Call SetTabIndexIfExists("btnNextPending", 6, True)
        Call SetTabIndexIfExists("btnClose", 7, True)
    End If
End Sub

Private Sub SetTabIndexIfExists(ByVal controlName As String, ByVal tabIndexValue As Integer, ByVal tabStopValue As Boolean)
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(controlName)
    If ctrl Is Nothing Then Exit Sub
    ctrl.TabIndex = tabIndexValue
    ctrl.TabStop = tabStopValue
    On Error GoTo 0
End Sub

Private Sub ApplyBackfillCaptions()
    On Error Resume Next
    Me.Caption = LocalizationManager.GetText("Legacy package backfill review")
    Me.lblQueueSummary.Caption = LocalizationManager.GetText("Backfill queue")
    Me.lblReviewId.Caption = LocalizationManager.GetText("Backfill ID")
    Me.lblIncOutRow.Caption = LocalizationManager.GetText("IncOut row")
    Me.lblCurrentStatus.Caption = LocalizationManager.GetText("Review status")
    Me.lblRecordNumber.Caption = LocalizationManager.GetText("Record number")
    Me.lblService.Caption = LocalizationManager.GetText("Package ID")
    Me.lblAmount.Caption = LocalizationManager.GetText("Parent amount")
    Me.lblDocumentType.Caption = LocalizationManager.GetText("Parent document type")
    Me.lblDocumentNumber.Caption = LocalizationManager.GetText("Parent document number")
    Me.lblDocumentDate.Caption = LocalizationManager.GetText("Parent document date")
    Me.lblCounterparty.Caption = LocalizationManager.GetText("Counterparty")
    Me.lblCandidates.Caption = LocalizationManager.GetText("Package proposals")
    Me.lblBestCandidateNumber.Caption = LocalizationManager.GetText("1C operation number")
    Me.lblBestCandidateDate.Caption = LocalizationManager.GetText("1C operation date")
    Me.chkUseBestCandidate.Caption = LocalizationManager.GetText("Use current proposal")
    Me.lblSelectedOperationNumber.Caption = LocalizationManager.GetText("Proposed child number")
    Me.lblSelectedOperationDate.Caption = LocalizationManager.GetText("Proposed child date")
    Me.lblBestCandidateComment.Caption = LocalizationManager.GetText("1C comment")
    Me.lblCandidateComment.Caption = LocalizationManager.GetText("Proposed description")
    Me.btnApply.Caption = LocalizationManager.GetText("Apply")
    Me.btnApplyNext.Caption = LocalizationManager.GetText("Apply and next")
    Me.btnNextPending.Caption = LocalizationManager.GetText("Next pending")
    If Not mOpenSourceButton Is Nothing Then mOpenSourceButton.Caption = LocalizationManager.GetText("Open IncOut row")
    Me.btnClose.Caption = LocalizationManager.GetText("Close")
    On Error GoTo 0
End Sub

Private Sub ApplyFormLayout()
    Dim marginX As Single
    Dim contentWidth As Single
    Dim currentTop As Single
    Dim labelOffset As Single
    Dim leftColumnWidth As Single
    Dim rightColumnLeft As Single
    Dim rightColumnWidth As Single
    Dim fieldWidth As Single
    Dim buttonTop As Single
    Dim parentTypeWidth As Single
    Dim parentNumberWidth As Single
    Dim parentDateWidth As Single

    marginX = 12
    labelOffset = 14
    contentWidth = DESIGN_FORM_WIDTH - (marginX * 2) - 18
    leftColumnWidth = 430
    rightColumnLeft = marginX + leftColumnWidth + 16
    rightColumnWidth = contentWidth - leftColumnWidth - 16
    fieldWidth = 104

    Me.lblQueueSummary.Left = marginX
    Me.lblQueueSummary.Top = 10
    Me.lblQueueSummary.Width = contentWidth
    Me.lblQueueSummary.Height = 18
    Me.txtQueueSummary.Left = marginX
    Me.txtQueueSummary.Top = 28
    Me.txtQueueSummary.Width = contentWidth
    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Me.txtQueueSummary.Height = 52
        currentTop = 84
    Else
        Me.txtQueueSummary.Height = 28
        currentTop = 64
    End If

    PlaceField Me.lblReviewId, Me.txtReviewId, marginX, currentTop, 108
    PlaceField Me.lblIncOutRow, Me.txtIncOutRow, marginX + 118, currentTop, 72
    PlaceField Me.lblCurrentStatus, Me.txtCurrentStatus, marginX + 200, currentTop, 136

    currentTop = currentTop + 40
    PlaceField Me.lblRecordNumber, Me.txtRecordNumber, marginX, currentTop, 96
    PlaceField Me.lblService, Me.txtService, marginX + 108, currentTop, 92
    PlaceField Me.lblAmount, Me.txtAmount, marginX + 212, currentTop, 112

    currentTop = currentTop + 40
    parentTypeWidth = 200
    parentNumberWidth = 108
    parentDateWidth = 106
    PlaceField Me.lblDocumentType, Me.txtDocumentType, marginX, currentTop, parentTypeWidth
    PlaceField Me.lblDocumentNumber, Me.txtDocumentNumber, marginX + parentTypeWidth + 12, currentTop, parentNumberWidth
    PlaceField Me.lblDocumentDate, Me.txtDocumentDate, marginX + parentTypeWidth + parentNumberWidth + 24, currentTop, parentDateWidth

    currentTop = currentTop + 40
    Me.lblCounterparty.Left = marginX
    Me.lblCounterparty.Top = currentTop
    Me.lblCounterparty.Width = 110
    Me.txtCounterparty.Left = marginX
    Me.txtCounterparty.Top = currentTop + labelOffset
    Me.txtCounterparty.Width = leftColumnWidth
    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Me.txtCounterparty.Height = 34
    Else
        Me.txtCounterparty.Height = 54
    End If

    Me.lblCandidates.Left = rightColumnLeft
    Me.lblCandidates.Top = 64
    Me.lblCandidates.Width = 150
    Me.lstCandidates.Left = rightColumnLeft
    Me.lstCandidates.Top = 82
    Me.lstCandidates.Width = rightColumnWidth
    Me.lstCandidates.Height = 166

    currentTop = Me.txtCounterparty.Top + Me.txtCounterparty.Height + 8
    PlaceField Me.lblBestCandidateNumber, Me.txtBestCandidateNumber, marginX, currentTop, 126
    PlaceField Me.lblBestCandidateDate, Me.txtBestCandidateDate, marginX + 138, currentTop, 110
    If mReviewMode <> REVIEW_MODE_BACKFILL Then
        Me.chkUseBestCandidate.Left = marginX + 260
        Me.chkUseBestCandidate.Top = currentTop + 16
        Me.chkUseBestCandidate.Width = 180
        Me.chkUseBestCandidate.Visible = True
    Else
        Me.chkUseBestCandidate.Visible = False
    End If

    currentTop = currentTop + 38
    PlaceField Me.lblSelectedOperationNumber, Me.txtSelectedOperationNumber, marginX, currentTop, 148
    PlaceField Me.lblSelectedOperationDate, Me.txtSelectedOperationDate, marginX + 160, currentTop, 118

    currentTop = currentTop + 38
    Me.lblBestCandidateComment.Left = marginX
    Me.lblBestCandidateComment.Top = currentTop
    Me.lblBestCandidateComment.Width = leftColumnWidth
    Me.txtBestCandidateComment.Left = marginX
    Me.txtBestCandidateComment.Top = currentTop + labelOffset
    Me.txtBestCandidateComment.Width = leftColumnWidth
    Me.txtBestCandidateComment.Height = 94

    Me.lblCandidateComment.Left = rightColumnLeft
    Me.lblCandidateComment.Top = currentTop
    Me.lblCandidateComment.Width = rightColumnWidth
    Me.txtCandidateComment.Left = rightColumnLeft
    Me.txtCandidateComment.Top = currentTop + labelOffset
    Me.txtCandidateComment.Width = rightColumnWidth
    Me.txtCandidateComment.Height = 94

    buttonTop = currentTop + 118
    Me.btnApply.Left = marginX
    Me.btnApply.Top = buttonTop
    Me.btnApply.Width = 98
    Me.btnApplyNext.Left = Me.btnApply.Left + Me.btnApply.Width + 8
    Me.btnApplyNext.Top = buttonTop
    Me.btnApplyNext.Width = 132
    Me.btnNextPending.Left = Me.btnApplyNext.Left + Me.btnApplyNext.Width + 8
    Me.btnNextPending.Top = buttonTop
    Me.btnNextPending.Width = 132
    If mReviewMode = REVIEW_MODE_BACKFILL Then
        Call EnsureBackfillOpenRowButton
        mOpenSourceButton.Left = Me.btnNextPending.Left + Me.btnNextPending.Width + 8
        mOpenSourceButton.Top = buttonTop
        mOpenSourceButton.Width = 132
        mOpenSourceButton.Height = Me.btnNextPending.Height
        mOpenSourceButton.Visible = True
    ElseIf Not mOpenSourceButton Is Nothing Then
        mOpenSourceButton.Visible = False
    End If
    Me.btnClose.Width = 82
    Me.btnClose.Left = marginX + contentWidth - Me.btnClose.Width
    Me.btnClose.Top = buttonTop

    Me.txtReviewRowIndex.Left = marginX + contentWidth - 60
    Me.txtReviewRowIndex.Top = buttonTop + 32
    Me.txtReviewRowIndex.Width = 54
End Sub

Private Sub PlaceField(ByVal lbl As MSForms.Label, ByVal txt As MSForms.control, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthValue As Single)
    lbl.Left = leftPos
    lbl.Top = topPos
    lbl.Width = widthValue
    txt.Left = leftPos
    txt.Top = topPos + 14
    txt.Width = widthValue
End Sub

Private Sub EnsureBackfillOpenRowButton()
    Dim existingControl As Object

    On Error Resume Next
    Set existingControl = Me.Controls("btnOpenIncOutRow")
    On Error GoTo 0

    If mReviewMode <> REVIEW_MODE_BACKFILL Then
        If Not existingControl Is Nothing Then existingControl.Visible = False
        Exit Sub
    End If

    If existingControl Is Nothing Then
        Set mOpenSourceButton = Me.Controls.Add("Forms.CommandButton.1", "btnOpenIncOutRow", True)
    Else
        Set mOpenSourceButton = existingControl
    End If

    mOpenSourceButton.Visible = True
    mOpenSourceButton.Caption = LocalizationManager.GetText("Open IncOut row")
    mOpenSourceButton.TakeFocusOnClick = False
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
    Call EnsureBackfillOpenRowButton

    contentWidth = DESIGN_FORM_WIDTH
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
End Sub

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
    GetFormContentHeight = bottomEdge + 34
End Function
