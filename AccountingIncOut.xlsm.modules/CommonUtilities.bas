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
