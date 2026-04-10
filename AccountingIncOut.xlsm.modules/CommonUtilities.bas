Attribute VB_Name = "CommonUtilities"
'==============================================
' COMMON UTILITY FUNCTIONS MODULE - CommonUtilities
' Purpose: Centralized storage of common utility functions
' State: CREATED AS A RESULT OF REFACTORING
' Version: 1.0.0
' Date: 10.01.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' =============================================
' SAFE WORKSHEET RETRIEVAL
' =============================================
' @description Gets worksheet by name with error handling
' @param   sheetName [String] Worksheet name
' @return  [Worksheet] Worksheet object or Nothing if not found
' =============================================
Public Function GetWorksheetSafe(sheetName As String) As Worksheet
    Dim Ws As Worksheet
    
    On Error Resume Next
    Set Ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    Set GetWorksheetSafe = Ws
End Function

' =============================================
' SAFE TABLE RETRIEVAL
' =============================================
' @description Gets table (ListObject) by name with error handling
' @param   Ws [Worksheet] Worksheet containing the table
' @param   tableName [String] Table name
' @return  [ListObject] Table object or Nothing if not found
' =============================================
Public Function GetListObjectSafe(Ws As Worksheet, tableName As String) As ListObject
    Dim tbl As ListObject
    
    On Error Resume Next
    If Not Ws Is Nothing Then
        Set tbl = Ws.ListObjects(tableName)
    End If
    On Error GoTo 0
    
    Set GetListObjectSafe = tbl
End Function

' =============================================
' SAFE TABLE COLUMN RETRIEVAL
' =============================================
Public Function GetListColumnSafe(ByVal targetTable As ListObject, ByVal columnName As String) As ListColumn
    Dim listColumn As ListColumn

    On Error Resume Next
    If Not targetTable Is Nothing Then
        Set listColumn = targetTable.ListColumns(columnName)
    End If
    On Error GoTo 0

    Set GetListColumnSafe = listColumn
End Function


' =============================================
' DATE FORMAT VALIDATION
' =============================================
' @description Checks correctness of date format DD.MM.YY
' @param   DateText [String] Date text in DD.MM.YY format
' @return  [Boolean] True if format is correct and date is not in the future
' =============================================
Public Function IsValidDateFormat(DateText As String) As Boolean
    On Error GoTo DateError
    
    ' Check DD.MM.YY format
    If Len(DateText) = 8 And Mid(DateText, 3, 1) = "." And Mid(DateText, 6, 1) = "." Then
        Dim TestDate As Date
        ' Convert YY to full year (20YY)
        Dim fullDateText As String
        fullDateText = Left(DateText, 6) & "20" & Right(DateText, 2)
        TestDate = CDate(fullDateText)
        
        ' Check that date is not in the future
        If TestDate > Date Then
            IsValidDateFormat = False
        Else
            IsValidDateFormat = True
        End If
    Else
        IsValidDateFormat = False
    End If
    
    Exit Function
    
DateError:
    IsValidDateFormat = False
End Function

' =============================================
' FORMAT DATE FROM CELL
' =============================================
' @description Formats date from cell to DD.MM.YY string
' @param   cell [Range] Cell with date
' @return  [String] Formatted date or empty string
' =============================================
Public Function FormatDateCell(cell As Range) As String
    On Error GoTo FormatError
    
    If Not IsEmpty(cell.value) And IsDate(cell.value) Then
        FormatDateCell = Format(cell.value, "dd.mm.yy")
    Else
        FormatDateCell = ""
    End If
    
    Exit Function
    
FormatError:
    FormatDateCell = ""
End Function

' =============================================
' WRITE DATE TO CELL
' =============================================
' @description Writes date in DD.MM.YY format to cell
' @param   cell [Range] Cell to write to
' @param   DateText [String] Date text in DD.MM.YY format
' =============================================
Public Sub WriteDateToCell(cell As Range, DateText As String)
    If Trim(DateText) <> "" And IsValidDateFormat(DateText) Then
        ' Convert YY to full year for writing
        Dim fullDateText As String
        fullDateText = Left(DateText, 6) & "20" & Right(DateText, 2)
        cell.value = CDate(fullDateText)
    Else
        cell.value = ""
    End If
End Sub

' =============================================
' CORRESPONDENT NORMALIZATION
' =============================================
Public Function CorrespondentTextsMatch(ByVal leftText As String, ByVal rightText As String) As Boolean
    Dim leftKey As String
    Dim rightKey As String
    Dim leftNormalized As String
    Dim rightNormalized As String

    leftKey = BuildCorrespondentMatchKey(leftText)
    rightKey = BuildCorrespondentMatchKey(rightText)

    If Len(leftKey) = 0 Or Len(rightKey) = 0 Then Exit Function

    If Left$(leftKey, 8) = "MILUNIT:" And Left$(rightKey, 8) = "MILUNIT:" Then
        CorrespondentTextsMatch = (StrComp(leftKey, rightKey, vbTextCompare) = 0)
        Exit Function
    End If

    leftNormalized = NormalizeCorrespondentText(leftText)
    rightNormalized = NormalizeCorrespondentText(rightText)

    If Len(leftNormalized) = 0 Or Len(rightNormalized) = 0 Then Exit Function

    If StrComp(leftNormalized, rightNormalized, vbTextCompare) = 0 Then
        CorrespondentTextsMatch = True
    ElseIf InStr(1, leftNormalized, rightNormalized, vbTextCompare) > 0 Then
        CorrespondentTextsMatch = True
    ElseIf InStr(1, rightNormalized, leftNormalized, vbTextCompare) > 0 Then
        CorrespondentTextsMatch = True
    End If
End Function

Public Function BuildCorrespondentMatchKey(ByVal sourceText As String) As String
    Dim militaryToken As String
    Dim rpbsCode As String
    Dim normalizedText As String

    militaryToken = ExtractCorrespondentMilitaryToken(sourceText)
    If Len(militaryToken) > 0 Then
        BuildCorrespondentMatchKey = "MILUNIT:" & militaryToken

        If ShouldUseRpbsInCorrespondentKey(sourceText) Then
            rpbsCode = ExtractCorrespondentRpbsCode(sourceText)
            If Len(rpbsCode) > 0 Then
                BuildCorrespondentMatchKey = BuildCorrespondentMatchKey & "|RPBS:" & rpbsCode
            End If
        End If
        Exit Function
    End If

    normalizedText = NormalizeCorrespondentText(sourceText)
    If Len(normalizedText) > 0 Then
        BuildCorrespondentMatchKey = "TEXT:" & normalizedText
    End If
End Function

Public Function NormalizeCorrespondentText(ByVal sourceText As String) As String
    NormalizeCorrespondentText = CollapseCorrespondentSpaces(PrepareCorrespondentSource(sourceText, False))
End Function

Public Function ExtractCorrespondentMilitaryToken(ByVal sourceText As String) As String
    Dim preparedText As String
    Dim markerTail As String

    preparedText = PrepareCorrespondentSource(sourceText, True)
    If Len(preparedText) = 0 Then Exit Function

    If IsMilitaryToken(Trim$(preparedText)) Then
        ExtractCorrespondentMilitaryToken = Trim$(preparedText)
        Exit Function
    End If

    markerTail = GetTextAfterMilitaryMarker(preparedText)
    If Len(markerTail) > 0 Then
        ExtractCorrespondentMilitaryToken = ExtractFirstMilitaryToken(markerTail)
    End If
End Function

Public Function ExtractCorrespondentRpbsCode(ByVal sourceText As String) As String
    Dim preparedText As String
    Dim markerPos As Long
    Dim currentChar As String
    Dim i As Long

    preparedText = PrepareCorrespondentSource(sourceText, True)
    If Len(preparedText) = 0 Then Exit Function

    markerPos = InStr(1, preparedText, GetRpbsMarker(), vbTextCompare)
    If markerPos = 0 Then Exit Function

    i = markerPos + Len(GetRpbsMarker())
    Do While i <= Len(preparedText)
        currentChar = Mid$(preparedText, i, 1)
        If currentChar <> " " And currentChar <> ":" Then Exit Do
        i = i + 1
    Loop

    Do While i <= Len(preparedText)
        currentChar = Mid$(preparedText, i, 1)
        If IsAsciiLetter(currentChar) Or IsDigitCharacter(currentChar) Or IsCyrillicLetter(currentChar) Then
            ExtractCorrespondentRpbsCode = ExtractCorrespondentRpbsCode & currentChar
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    ExtractCorrespondentRpbsCode = NormalizeCodeLookalikes(ExtractCorrespondentRpbsCode)
End Function

Public Function IsBranchCorrespondent(ByVal sourceText As String) As Boolean
    Dim normalizedText As String
    normalizedText = PrepareCorrespondentSource(sourceText, False)
    IsBranchCorrespondent = (InStr(1, normalizedText, GetBranchMarker(), vbTextCompare) > 0)
End Function

Private Function PrepareCorrespondentSource(ByVal sourceText As String, ByVal keepHyphen As Boolean) As String
    Dim preparedText As String
    Dim i As Long
    Dim currentChar As String

    preparedText = UCase$(Trim$(sourceText))
    preparedText = Replace(preparedText, ChrW$(1025), ChrW$(1045))
    preparedText = Replace(preparedText, ChrW$(1105), ChrW$(1045))
    preparedText = Replace(preparedText, vbCr, " ")
    preparedText = Replace(preparedText, vbLf, " ")
    preparedText = Replace(preparedText, vbTab, " ")
    preparedText = Replace(preparedText, Chr$(34), " ")

    For i = 1 To Len(preparedText)
        currentChar = Mid$(preparedText, i, 1)

        If IsAsciiLetter(currentChar) Or IsCyrillicLetter(currentChar) Or IsDigitCharacter(currentChar) Then
            PrepareCorrespondentSource = PrepareCorrespondentSource & currentChar
        ElseIf keepHyphen And currentChar = "-" Then
            PrepareCorrespondentSource = PrepareCorrespondentSource & currentChar
        ElseIf currentChar = "/" Then
            PrepareCorrespondentSource = PrepareCorrespondentSource & currentChar
        Else
            PrepareCorrespondentSource = PrepareCorrespondentSource & " "
        End If
    Next i
End Function

Private Function CollapseCorrespondentSpaces(ByVal sourceText As String) As String
    Dim result As String

    result = Trim$(sourceText)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    CollapseCorrespondentSpaces = result
End Function

Private Function ShouldUseRpbsInCorrespondentKey(ByVal sourceText As String) As Boolean
    ShouldUseRpbsInCorrespondentKey = IsBranchCorrespondent(sourceText)
End Function

Private Function GetTextAfterMilitaryMarker(ByVal sourceText As String) As String
    Dim markers As Variant
    Dim marker As Variant
    Dim markerPos As Long
    Dim startPos As Long

    markers = Array(GetMilitaryMarkerFull(), GetMilitaryMarkerAlternate(), GetMilitaryMarkerTypo(), GetMilitaryMarkerSlash(), GetMilitaryMarkerShort())

    For Each marker In markers
        markerPos = InStr(1, sourceText, CStr(marker), vbTextCompare)
        If markerPos > 0 Then
            startPos = markerPos + Len(CStr(marker))
            GetTextAfterMilitaryMarker = Mid$(sourceText, startPos)
            Exit Function
        End If
    Next marker
End Function

Private Function ExtractFirstMilitaryToken(ByVal sourceText As String) As String
    Dim parts() As String
    Dim i As Long
    Dim currentPart As String

    parts = Split(sourceText, " ")
    For i = LBound(parts) To UBound(parts)
        currentPart = Trim$(parts(i))
        If IsMilitaryToken(currentPart) Then
            ExtractFirstMilitaryToken = currentPart
            Exit Function
        End If
    Next i
End Function

Private Function IsMilitaryToken(ByVal tokenText As String) As Boolean
    Dim dashPos As Long
    Dim leftPart As String
    Dim rightPart As String

    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then Exit Function

    dashPos = InStr(tokenText, "-")
    If dashPos = 0 Then
        IsMilitaryToken = IsDigitsOnly(tokenText) And Len(tokenText) >= 4 And Len(tokenText) <= 6
        Exit Function
    End If

    If InStr(dashPos + 1, tokenText, "-") > 0 Then Exit Function

    leftPart = Left$(tokenText, dashPos - 1)
    rightPart = Mid$(tokenText, dashPos + 1)

    If Not (IsDigitsOnly(leftPart) And Len(leftPart) >= 4 And Len(leftPart) <= 6) Then Exit Function
    If Len(rightPart) = 0 Or Len(rightPart) > 4 Then Exit Function
    If Not IsAsciiCyrillicOrDigitsOnly(rightPart) Then Exit Function

    IsMilitaryToken = True
End Function

Private Function IsDigitsOnly(ByVal sourceText As String) As Boolean
    Dim i As Long
    Dim currentChar As String

    If Len(sourceText) = 0 Then Exit Function

    For i = 1 To Len(sourceText)
        currentChar = Mid$(sourceText, i, 1)
        If Not IsDigitCharacter(currentChar) Then Exit Function
    Next i

    IsDigitsOnly = True
End Function

Private Function IsAsciiCyrillicOrDigitsOnly(ByVal sourceText As String) As Boolean
    Dim i As Long
    Dim currentChar As String

    If Len(sourceText) = 0 Then Exit Function

    For i = 1 To Len(sourceText)
        currentChar = Mid$(sourceText, i, 1)
        If Not (IsAsciiLetter(currentChar) Or IsCyrillicLetter(currentChar) Or IsDigitCharacter(currentChar)) Then Exit Function
    Next i

    IsAsciiCyrillicOrDigitsOnly = True
End Function

Private Function IsAsciiLetter(ByVal sourceChar As String) As Boolean
    Dim codePoint As Long

    If Len(sourceChar) = 0 Then Exit Function
    codePoint = AscW(sourceChar)
    IsAsciiLetter = ((codePoint >= 65 And codePoint <= 90) Or (codePoint >= 97 And codePoint <= 122))
End Function

Private Function IsCyrillicLetter(ByVal sourceChar As String) As Boolean
    Dim codePoint As Long

    If Len(sourceChar) = 0 Then Exit Function
    codePoint = AscW(sourceChar)
    IsCyrillicLetter = ((codePoint >= 1040 And codePoint <= 1103) Or codePoint = 1025 Or codePoint = 1105)
End Function

Private Function IsDigitCharacter(ByVal sourceChar As String) As Boolean
    Dim codePoint As Long

    If Len(sourceChar) = 0 Then Exit Function
    codePoint = AscW(sourceChar)
    IsDigitCharacter = (codePoint >= 48 And codePoint <= 57)
End Function

Private Function GetMilitaryMarkerFull() As String
    GetMilitaryMarkerFull = ChrW$(1042) & ChrW$(1054) & ChrW$(1049) & ChrW$(1057) & ChrW$(1050) & ChrW$(1054) & ChrW$(1042) & ChrW$(1040) & ChrW$(1071) & " " & ChrW$(1063) & ChrW$(1040) & ChrW$(1057) & ChrW$(1058) & ChrW$(1068)
End Function

Private Function GetMilitaryMarkerAlternate() As String
    GetMilitaryMarkerAlternate = ChrW$(1042) & ChrW$(1054) & ChrW$(1048) & ChrW$(1053) & ChrW$(1057) & ChrW$(1050) & ChrW$(1040) & ChrW$(1071) & " " & ChrW$(1063) & ChrW$(1040) & ChrW$(1057) & ChrW$(1058) & ChrW$(1068)
End Function

Private Function GetMilitaryMarkerTypo() As String
    GetMilitaryMarkerTypo = ChrW$(1042) & ChrW$(1054) & ChrW$(1057) & ChrW$(1050) & ChrW$(1054) & ChrW$(1042) & ChrW$(1040) & ChrW$(1071) & " " & ChrW$(1063) & ChrW$(1040) & ChrW$(1057) & ChrW$(1058) & ChrW$(1068)
End Function

Private Function GetMilitaryMarkerSlash() As String
    GetMilitaryMarkerSlash = ChrW$(1042) & "/" & ChrW$(1063)
End Function

Private Function GetMilitaryMarkerShort() As String
    GetMilitaryMarkerShort = ChrW$(1042) & ChrW$(1063)
End Function

Private Function GetRpbsMarker() As String
    GetRpbsMarker = ChrW$(1056) & ChrW$(1055) & ChrW$(1041) & ChrW$(1057)
End Function

Private Function GetBranchMarker() As String
    GetBranchMarker = ChrW$(1060) & ChrW$(1048) & ChrW$(1051) & ChrW$(1048) & ChrW$(1040) & ChrW$(1051)
End Function

Private Function NormalizeCodeLookalikes(ByVal sourceText As String) As String
    Dim result As String

    result = UCase$(sourceText)
    result = Replace(result, ChrW$(1040), "A")
    result = Replace(result, ChrW$(1042), "B")
    result = Replace(result, ChrW$(1045), "E")
    result = Replace(result, ChrW$(1050), "K")
    result = Replace(result, ChrW$(1052), "M")
    result = Replace(result, ChrW$(1053), "H")
    result = Replace(result, ChrW$(1054), "O")
    result = Replace(result, ChrW$(1056), "P")
    result = Replace(result, ChrW$(1057), "C")
    result = Replace(result, ChrW$(1058), "T")
    result = Replace(result, ChrW$(1059), "Y")
    result = Replace(result, ChrW$(1061), "X")

    NormalizeCodeLookalikes = result
End Function
