Attribute VB_Name = "SearchModule"
'==============================================
' TABLE SEARCH MODULE - SearchModule
' Purpose: Implementation of instant search across all table columns
' State: INTEGRATED WITH LOCALIZATION MANAGER
' Version: 1.7.3
' Date: 09.08.2025
' Author: Evgeniy Kerzhaev, FKU "95 FES" MO RF
'==============================================

Option Explicit

' Global variables to save search state
Private LastSearchText As String
Private SearchResultsVisible As Boolean

Public Sub PerformSearch()
    Dim SearchText As String
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim i As Long, j As Long
    Dim cellValue As String
    Dim FoundCount As Integer
    
    SearchText = Trim(UCase(UserFormVhIsh.txtSearch.Text))
    
    ' Save search text for possible restoration
    LastSearchText = SearchText
    
    ' Clear results if search is empty
    If Len(SearchText) = 0 Then
        With UserFormVhIsh
            .lstSearchResults.Clear
            .lstSearchResults.Visible = False
            .lblStatusBar.Caption = LocalizationManager.GetText("Enter text to search")
        End With
        SearchResultsVisible = False
        Exit Sub
    End If
    
    On Error GoTo SearchError
    
    Set wsData = ThisWorkbook.Worksheets("IncOut")
    Set tblData = wsData.ListObjects("TableIncOut")
    
    UserFormVhIsh.lstSearchResults.Clear
    
    ' Search across all rows and columns of the table
    If tblData.ListRows.Count > 0 Then
        For i = 1 To tblData.ListRows.Count
            Dim hasMatch As Boolean
            hasMatch = False
            
            For j = 1 To tblData.ListColumns.Count
                cellValue = UCase(CStr(tblData.DataBodyRange.Cells(i, j).value))
                
                If InStr(cellValue, SearchText) > 0 Then
                    hasMatch = True
                    Exit For
                End If
            Next j
            
            If hasMatch Then
                Call AddImprovedSearchResult(i, tblData)
                FoundCount = FoundCount + 1
            End If
            
            ' Limit number of results
            If FoundCount >= 25 Then Exit For
        Next i
    End If
    
    ' Display results
    Call DisplaySearchResults(FoundCount)
    
    Exit Sub
    
SearchError:
    UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Search error: ") & Err.description
    UserFormVhIsh.lstSearchResults.Visible = False
    SearchResultsVisible = False
End Sub

' IMPROVED PROCEDURE: Adding search result with numbering
' Added fields: Received from (9), Executor (11), Confirmation status (19)
' Removed duplication of row number at the end (hidden key stored separately)
Private Sub AddImprovedSearchResult(rowNum As Long, tbl As ListObject)
    Dim ResultText As String
    Dim serviceText As String
    Dim docTypeInOut As String
    Dim docTypeText As String
    Dim docNumberText As String
    Dim amountText As String
    Dim frpText As String
    Dim DateText As String
    Dim fromWhomText As String
    Dim executorText As String
    Dim statusText As String
    
    On Error GoTo AddError
    
    ' Core data
    serviceText = CStr(tbl.DataBodyRange.Cells(rowNum, 2).value)      ' Service
    docTypeInOut = CStr(tbl.DataBodyRange.Cells(rowNum, 3).value)     ' Doc Group (Inc./Out.)
    docTypeText = CStr(tbl.DataBodyRange.Cells(rowNum, 4).value)      ' Document Type
    docNumberText = CStr(tbl.DataBodyRange.Cells(rowNum, 5).value)    ' Document Number
    amountText = CStr(tbl.DataBodyRange.Cells(rowNum, 6).value)       ' Amount
    frpText = CStr(tbl.DataBodyRange.Cells(rowNum, 7).value)          ' FRP Number
    
    ' Date (8)
    If IsDate(tbl.DataBodyRange.Cells(rowNum, 8).value) Then
        DateText = Format(tbl.DataBodyRange.Cells(rowNum, 8).value, "dd.mm.yy")
    Else
        DateText = CStr(tbl.DataBodyRange.Cells(rowNum, 8).value)
    End If
    
    ' New fields
    fromWhomText = CStr(tbl.DataBodyRange.Cells(rowNum, 9).value)     ' Received from
    executorText = CStr(tbl.DataBodyRange.Cells(rowNum, 11).value)    ' Executor
    statusText = CStr(tbl.DataBodyRange.Cells(rowNum, 19).value)      ' Confirmation status
    
    ' Format result string
    ' Start with ">" and row number
    ResultText = ">" & rowNum & ": "
    
    If Trim(serviceText) <> "" Then
        ResultText = ResultText & "[" & serviceText & "] "
    End If
    
    If Trim(docTypeInOut) <> "" Then
        ResultText = ResultText & docTypeInOut & " "
    End If
    
    If Trim(docTypeText) <> "" Then
        ResultText = ResultText & docTypeText & " "
    End If
    
    If Trim(docNumberText) <> "" Then
        ResultText = ResultText & LocalizationManager.GetText("No.") & docNumberText & " "
    End If
    
    If Trim(amountText) <> "" And amountText <> "0" Then
        ResultText = ResultText & "(" & amountText & LocalizationManager.GetText(" rub.") & ") "
    End If
    
    If Trim(frpText) <> "" Then
        ResultText = ResultText & LocalizationManager.GetText("FRP:") & frpText & " "
    End If
    
    If Trim(DateText) <> "" Then
        ResultText = ResultText & LocalizationManager.GetText("from ") & DateText & " "
    End If
    
    ' Add fields to the tail of the result:
    If Trim(fromWhomText) <> "" Then
        ResultText = ResultText & "| " & LocalizationManager.GetText("From: ") & fromWhomText & " "
    End If
    
    If Trim(executorText) <> "" Then
        ResultText = ResultText & "| " & LocalizationManager.GetText("Exec.: ") & executorText & " "
    End If
    
    If Trim(statusText) <> "" Then
        ResultText = ResultText & "| " & LocalizationManager.GetText("Status: ") & statusText & " "
    End If
    
    ' CHANGED: Do not trim string for correct width calculation
    
    ' Add to list
    With UserFormVhIsh.lstSearchResults
        .AddItem ResultText
        ' Store hidden record identifier via separator,
        ' but DO NOT add it visually to the text
        If .ListCount > 0 Then
            .List(.ListCount - 1, 0) = ResultText & Chr(1) & CStr(rowNum)
        End If
    End With
    
    Exit Sub
    
AddError:
    ' Skip adding result errors
End Sub

' UPDATED PROCEDURE: Display results with automatic width expansion
Private Sub DisplaySearchResults(FoundCount As Integer)
    With UserFormVhIsh
        If .lstSearchResults.ListCount > 0 Then
            .lstSearchResults.Visible = True
            SearchResultsVisible = True
            
            ' NEW: Automatic width calculation based on content
            Call AutoResizeSearchResultsWidth
            
            ' Fixed height remains
            .lstSearchResults.Height = 120
            
            ' Status bar
            .lblStatusBar.Caption = LocalizationManager.GetText("Found: ") & FoundCount & LocalizationManager.GetText(" records | ") & _
                                    LocalizationManager.GetText("Navigation: ^v or click to jump | ") & _
                                    LocalizationManager.GetText("Search remains active")
        Else
            .lstSearchResults.Visible = False
            SearchResultsVisible = False
            .lblStatusBar.Caption = LocalizationManager.GetText("For query '") & .txtSearch.Text & LocalizationManager.GetText("' nothing found")
        End If
    End With
End Sub

' NEW PROCEDURE: Automatic field resizing based on content
Private Sub AutoResizeSearchResultsWidth()
    Dim maxWidth As Single
    Dim textWidth As Single
    Dim i As Long
    Dim testText As String
    Dim minWidth As Single
    Dim maxAllowedWidth As Single
    
    On Error GoTo ResizeError
    
    ' Min and max width
    minWidth = 420 ' Minimum width
    maxAllowedWidth = 800 ' Maximum width (to stay within form boundaries)
    
    maxWidth = minWidth
    
    With UserFormVhIsh.lstSearchResults
        ' Loop through all list items to find the longest one
        For i = 0 To .ListCount - 1
            ' Get text without hidden part (before Chr(1) separator)
            testText = .List(i, 0)
            If InStr(testText, Chr(1)) > 0 Then
                testText = Left(testText, InStr(testText, Chr(1)) - 1)
            End If
            
            ' NEW: Calculate text width considering font
            textWidth = CalculateTextWidth(testText, .Font.Name, .Font.Size)
            
            ' Update max width with a small padding
            If textWidth + 20 > maxWidth Then
                maxWidth = textWidth + 20
            End If
        Next i
        
        ' Constrain width within reasonable limits
        If maxWidth < minWidth Then
            maxWidth = minWidth
        ElseIf maxWidth > maxAllowedWidth Then
            maxWidth = maxAllowedWidth
        End If
        
        ' Apply new width
        .Width = maxWidth
        
        ' NEW: Update ColumnWidths for correct display
        .ColumnWidths = CStr(maxWidth - 10)
    End With
    
    Exit Sub
    
ResizeError:
    ' On error, use standard width
    UserFormVhIsh.lstSearchResults.Width = minWidth
End Sub

' NEW FUNCTION: Calculate text width
Private Function CalculateTextWidth(Text As String, fontName As String, fontSize As Single) As Single
    Dim textLength As Long
    Dim avgCharWidth As Single
    
    On Error GoTo WidthError
    
    textLength = Len(Text)
    
    ' Approximate character width in pixels depending on font size
    ' For Segoe UI, average char width is about 0.6 of font height
    avgCharWidth = fontSize * 0.6
    
    ' Calculate total width with a small margin
    CalculateTextWidth = textLength * avgCharWidth
    
    Exit Function
    
WidthError:
    ' On error, return base width
    CalculateTextWidth = 420
End Function

' RADICALLY REWORKED PROCEDURE: SelectSearchResult
Public Sub SelectSearchResult()
    Dim selectedRow As Long
    Dim resultData As String
    Dim parts As Variant
    
    On Error GoTo SelectError
    
    With UserFormVhIsh.lstSearchResults
        If .ListIndex >= 0 Then
            ' Extract row number from hidden item data
            resultData = .List(.ListIndex, 0)
            parts = Split(resultData, Chr(1))
            
            If UBound(parts) >= 1 Then
                selectedRow = CLng(parts(1))
                
                ' Jump to the found record
                Call NavigationModule.NavigateToRecord(selectedRow)
                
                ' KEY CHANGE: DO NOT clear search and DO NOT hide results!
                ' Instead, highlight the selected item and update status
                
                ' Update status bar with selected record info
                UserFormVhIsh.lblStatusBar.Caption = LocalizationManager.GetText("Jump to record No.") & selectedRow & " | " & _
                                                     LocalizationManager.GetText("Found: ") & .ListCount & LocalizationManager.GetText(" records | ") & _
                                                     LocalizationManager.GetText("Search active: """) & UserFormVhIsh.txtSearch.Text & """"
                
                ' Visually highlight selected item (remains highlighted)
                .BackColor = RGB(240, 248, 255) ' Light blue background for active search
            End If
        End If
    End With
    
    Exit Sub
    
SelectError:
    MsgBox LocalizationManager.GetText("Error selecting search result: ") & Err.description, vbExclamation, LocalizationManager.GetText("Error")
End Sub

' NEW PROCEDURE: Clear search on user demand
Public Sub ClearSearch()
    With UserFormVhIsh
        .txtSearch.Text = ""
        .lstSearchResults.Clear
        .lstSearchResults.Visible = False
        .lblStatusBar.Caption = LocalizationManager.GetText("Search cleared")
    End With
    LastSearchText = ""
    SearchResultsVisible = False
End Sub

' NEW PROCEDURE: Restore last search
Public Sub RestoreLastSearch()
    If LastSearchText <> "" Then
        UserFormVhIsh.txtSearch.Text = LastSearchText
        Call PerformSearch
    End If
End Sub

' NEW PROCEDURE: Navigate search results with keyboard
Public Sub NavigateSearchResults(direction As String)
    With UserFormVhIsh.lstSearchResults
        If .Visible And .ListCount > 0 Then
            Select Case UCase(direction)
                Case "UP"
                    If .ListIndex > 0 Then
                        .ListIndex = .ListIndex - 1
                    Else
                        .ListIndex = .ListCount - 1 ' Jump to last item
                    End If
                Case "DOWN"
                    If .ListIndex < .ListCount - 1 Then
                        .ListIndex = .ListIndex + 1
                    Else
                        .ListIndex = 0 ' Jump to first item
                    End If
                Case "FIRST"
                    .ListIndex = 0
                Case "LAST"
                    .ListIndex = .ListCount - 1
            End Select
            
            ' Automatically jump to selected record
            Call SelectSearchResult
        End If
    End With
End Sub

' NEW PROCEDURE: Check if search is active
Public Function IsSearchActive() As Boolean
    IsSearchActive = SearchResultsVisible And (UserFormVhIsh.lstSearchResults.ListCount > 0)
End Function

' NEW PROCEDURE: Get current search info
Public Function GetSearchInfo() As String
    If IsSearchActive() Then
        With UserFormVhIsh.lstSearchResults
            GetSearchInfo = LocalizationManager.GetText("Active search: """) & UserFormVhIsh.txtSearch.Text & """ | " & _
                            LocalizationManager.GetText("Found: ") & .ListCount & LocalizationManager.GetText(" records | ") & _
                            LocalizationManager.GetText("Selected: ") & (.ListIndex + 1) & LocalizationManager.GetText(" of ") & .ListCount
        End With
    Else
        GetSearchInfo = LocalizationManager.GetText("Search inactive")
    End If
End Function

' NEW PROCEDURE: Highlight found text in form
Public Sub HighlightSearchTermInForm()
    ' This procedure can be expanded to highlight found text in form fields
    ' Leaving as a stub for future development
    Dim searchTerm As String
    searchTerm = UserFormVhIsh.txtSearch.Text
    
    If Len(searchTerm) > 0 Then
        ' Add logic here to highlight found text in form fields
        ' E.g., change background color of fields containing the search text
    End If
End Sub

