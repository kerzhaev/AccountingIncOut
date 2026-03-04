VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormProgress 
   Caption         =   "UserForm1"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "UserFormProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
' FORM CODE - UserFormProgress
'==============================================

Option Explicit

Private Sub btnCancel_Click()
    Call ProgressManager.CancelOperation
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Prevent closing the form with the X button
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Call ProgressManager.CancelOperation
    End If
End Sub
