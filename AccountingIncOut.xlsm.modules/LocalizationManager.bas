Attribute VB_Name = "LocalizationManager"
'==============================================
' LOCALIZATION MODULE - LocalizationManager
' Purpose: Dynamic translation of UI and messages from dictionary table
' Version: 1.0.0
' Date: 05.03.2026
'==============================================

Option Explicit

Private translationDict As Object
Private isInitialized As Boolean

' Initialize the dictionary from the Excel table
Public Sub InitializeLocalization()
    Dim wsLoc As Worksheet
    Dim tblLoc As ListObject
    Dim i As Long
    Dim key As String
    Dim val As String
    
    On Error GoTo InitError
    
    ' Create dictionary with text compare (case-insensitive)
    Set translationDict = CreateObject("Scripting.Dictionary")
    translationDict.CompareMode = 1
    
    ' Get worksheet and table safely
    Set wsLoc = CommonUtilities.GetWorksheetSafe("Localization")
    If wsLoc Is Nothing Then
        Debug.Print "Localization sheet not found!"
        isInitialized = True ' Prevent infinite loops
        Exit Sub
    End If
    
    Set tblLoc = CommonUtilities.GetListObjectSafe(wsLoc, "TableLocalization")
    If tblLoc Is Nothing Then
        Debug.Print "Localization table not found!"
        isInitialized = True
        Exit Sub
    End If
    
    ' Load data into dictionary
    If tblLoc.ListRows.Count > 0 And Not tblLoc.DataBodyRange Is Nothing Then
        For i = 1 To tblLoc.ListRows.Count
            key = Trim(CStr(tblLoc.DataBodyRange.Cells(i, 1).value))
            val = Trim(CStr(tblLoc.DataBodyRange.Cells(i, 2).value))
            
            If key <> "" Then
                If Not translationDict.Exists(key) Then
                    translationDict.Add key, val
                Else
                    translationDict(key) = val ' Update if duplicate
                End If
            End If
        Next i
    End If
    
    isInitialized = True
    Debug.Print "Localization initialized. Loaded " & translationDict.Count & " keys."
    Exit Sub
    
InitError:
    Debug.Print "Error initializing localization: " & Err.description
    isInitialized = False
End Sub

' Get translated text by key
Public Function GetText(key As String, Optional defaultText As String = "") As String
    If Not isInitialized Then Call InitializeLocalization
    
    If translationDict Is Nothing Then
        GetText = IIf(defaultText <> "", defaultText, key)
        Exit Function
    End If
    
    If translationDict.Exists(key) Then
        GetText = translationDict(key)
    Else
        ' If key not found, return the key itself or default text
        GetText = IIf(defaultText <> "", defaultText, key)
    End If
End Function

' Automatically translate all controls on a UserForm
Public Sub TranslateForm(frm As Object)
    Dim ctrl As Object
    Dim translatedText As String
    
    If Not isInitialized Then Call InitializeLocalization
    
    On Error Resume Next
    
    ' Translate Form Caption
    translatedText = GetText(frm.Caption)
    If translatedText <> frm.Caption And translatedText <> "" Then
        frm.Caption = translatedText
    End If
    
    ' Translate Controls (Labels, Buttons, CheckBoxes)
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "Label" Or TypeName(ctrl) = "CommandButton" Or _
           TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "OptionButton" Or _
           TypeName(ctrl) = "Page" Then
           
            translatedText = GetText(ctrl.Caption)
            If translatedText <> ctrl.Caption And translatedText <> "" Then
                ctrl.Caption = translatedText
            End If
            
            ' Translate ControlTipText if exists
            If ctrl.ControlTipText <> "" Then
                translatedText = GetText(ctrl.ControlTipText)
                If translatedText <> ctrl.ControlTipText And translatedText <> "" Then
                    ctrl.ControlTipText = translatedText
                End If
            End If
        End If
    Next ctrl
    
    On Error GoTo 0
End Sub

' Force reload dictionary (useful during development/adding new terms)
Public Sub ReloadDictionary()
    isInitialized = False
    Call InitializeLocalization
End Sub

