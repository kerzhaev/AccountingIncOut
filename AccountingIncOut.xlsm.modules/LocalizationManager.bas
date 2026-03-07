Attribute VB_Name = "LocalizationManager"
'==============================================
' LOCALIZATION MODULE - LocalizationManager
' Purpose: Dynamic translation of UI and messages from dictionary table
' Version: 1.1.0
' Date: 08.03.2026
'==============================================

Option Explicit

Private translationDict As Object
Private isInitialized As Boolean

' Initialize the dictionary from the Excel table
Public Sub InitializeLocalization()
    On Error GoTo InitError

    Set translationDict = CreateObject("Scripting.Dictionary")
    translationDict.CompareMode = 1
    isInitialized = False

    ' Preferred source after sheet rename.
    If TryLoadLocalizationFromSource("Dictionaries", "TableLocalization") Then Exit Sub

    ' Backward-compatible fallback.
    If TryLoadLocalizationFromSource("Localization", "TableLocalization") Then Exit Sub

    Call LogLocalizationIssue("Localization dictionary/table not found. Falling back to key text.")
    Exit Sub

InitError:
    isInitialized = False
    Set translationDict = Nothing
    Call LogLocalizationIssue("Error initializing localization: " & Err.description)
End Sub

Private Function TryLoadLocalizationFromSource(ByVal sheetName As String, ByVal tableName As String) As Boolean
    Dim wsLoc As Worksheet
    Dim tblLoc As ListObject
    Dim i As Long
    Dim key As String
    Dim val As String

    Set wsLoc = CommonUtilities.GetWorksheetSafe(sheetName)
    If wsLoc Is Nothing Then Exit Function

    Set tblLoc = CommonUtilities.GetListObjectSafe(wsLoc, tableName)
    If tblLoc Is Nothing Then Exit Function

    If tblLoc.ListRows.Count > 0 And Not tblLoc.DataBodyRange Is Nothing Then
        For i = 1 To tblLoc.ListRows.Count
            key = Trim(CStr(tblLoc.DataBodyRange.Cells(i, 1).value))
            val = Trim(CStr(tblLoc.DataBodyRange.Cells(i, 2).value))

            If key <> "" Then
                If translationDict.Exists(key) Then
                    translationDict(key) = val
                Else
                    translationDict.Add key, val
                End If
            End If
        Next i
    End If

    isInitialized = True
    TryLoadLocalizationFromSource = True
    Debug.Print "Localization initialized from " & sheetName & ". Loaded " & translationDict.Count & " keys."
End Function

Private Sub LogLocalizationIssue(ByVal message As String)
    On Error GoTo SilentFail

    Debug.Print message
    Call SystemLogger.LogOperation("Localization", message, "WARNING", 0)
    Exit Sub

SilentFail:
    Debug.Print "Localization log fallback: " & message
End Sub

' Get translated text by key
Public Function GetText(key As String, Optional defaultText As String = "") As String
    If Not isInitialized Then Call InitializeLocalization

    If translationDict Is Nothing Or Not isInitialized Then
        GetText = IIf(defaultText <> "", defaultText, key)
        Exit Function
    End If

    If translationDict.Exists(key) Then
        GetText = translationDict(key)
    Else
        GetText = IIf(defaultText <> "", defaultText, key)
    End If
End Function

' Automatically translate all controls on a UserForm
Public Sub TranslateForm(frm As Object)
    Dim ctrl As Object
    Dim translatedText As String

    If Not isInitialized Then Call InitializeLocalization
    If translationDict Is Nothing Then Exit Sub

    On Error GoTo TranslateError

    translatedText = GetText(frm.Caption)
    If translatedText <> frm.Caption And translatedText <> "" Then
        frm.Caption = translatedText
    End If

    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "Label" Or TypeName(ctrl) = "CommandButton" Or _
           TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "OptionButton" Or _
           TypeName(ctrl) = "Page" Then

            translatedText = GetText(ctrl.Caption)
            If translatedText <> ctrl.Caption And translatedText <> "" Then
                ctrl.Caption = translatedText
            End If

            If ctrl.ControlTipText <> "" Then
                translatedText = GetText(ctrl.ControlTipText)
                If translatedText <> ctrl.ControlTipText And translatedText <> "" Then
                    ctrl.ControlTipText = translatedText
                End If
            End If
        End If
    Next ctrl

    Exit Sub

TranslateError:
    Call LogLocalizationIssue("TranslateForm error: " & Err.description)
End Sub

' Force reload dictionary (useful during development/adding new terms)
Public Sub ReloadDictionary()
    isInitialized = False
    Set translationDict = Nothing
    Call InitializeLocalization
End Sub
