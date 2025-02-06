Attribute VB_Name = "Ranking_Points"
Option Explicit

' Constants for better maintainability
Private Const DATA_FILE_NAME As String = "data.xlsx"
Private Const ALPHA_FILE_PATTERN As String = "Alpha*.xlsx"
Private Const FONT_NAME As String = "Ottawa"
Private Const FONT_SIZE As Integer = 8
Private Const TITLE_FONT_SIZE As Integer = 8

' Variables
Public AlphaListFilePath As String
Public alphaList As Workbook
Public cancelPressed As Boolean

Sub PlayerRankings()
    Dim DataFilePath As String
    Dim wb As Workbook
    Dim entrants As Worksheet
    Dim currentColumn As Integer
    Dim fso As Object
    Dim numberOfEntries As Integer
    Dim category As String
    
    ' Initialize file paths and check for files
    DataFilePath = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.path) & DATA_FILE_NAME
    AlphaListFilePath = GetAlphaListFilePath(ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.path), ALPHA_FILE_PATTERN)
    
    If Not UtilityFunctions.FileExists(DataFilePath) Then Exit Sub
    If Not UtilityFunctions.FileExists(AlphaListFilePath) Then Exit Sub
    
    ' Request password if necessary
    cancelPressed = False
    AlphaListPasswordInputForm.Show
    If cancelPressed Then Exit Sub
    
    ' Open the data workbook
    Set wb = Workbooks.Open(DataFilePath)
    Set entrants = wb.Worksheets("MASTER")
    numberOfEntries = GetNumberOfEntries(entrants)

    currentColumn = 10
    Do While Not IsEmpty(entrants.Cells(1, currentColumn)) And entrants.Cells(1, currentColumn).Value <> "Entry"
        category = GetCategory(entrants.Cells(1, currentColumn).Value)
        
        ' Create and populate new sheet
        CreateNewSheet entrants, currentColumn, category
        
        currentColumn = currentColumn + 1
    Loop

    wb.Save
    alphaList.Close SaveChanges:=False
End Sub

' Get category based on event data
Private Function GetCategory(eventValue As String) As String
    Dim subCategory As Integer
    
    subCategory = UtilityFunctions.ExtractNumberFromString(eventValue)
    
    Select Case eventValue
        Case "CB", "CG": GetCategory = "Cadet"
        Case "JB", "JG": GetCategory = "Junior"
        Case Else
            If subCategory > 5 And subCategory <= 15 Then
                GetCategory = "Cadet"
            ElseIf subCategory <= 19 Then
                GetCategory = "Junior"
            ElseIf subCategory = 21 Then
                GetCategory = "Senior"
            ElseIf subCategory >= 40 Then
                GetCategory = "Veteran"
            Else
                GetCategory = "Senior"
            End If
    End Select
End Function

' Create a new sheet and populate data
Private Sub CreateNewSheet(entrants As Worksheet, currentColumn As Integer, category As String)
    Dim newSheet As Worksheet
    Dim titles As Variant
    Dim row As Integer
    Dim points As Integer
    Dim i As Integer
    Dim lastRow As Long

    ' Create new sheet
    Set newSheet = Worksheets.Add
    newSheet.Name = entrants.Cells(1, currentColumn).Value
    If Err.number <> 0 Then Exit Sub
    newSheet.Move After:=Worksheets(Worksheets.Count)

    ' Set Titles
    titles = Array("Licence No", "Name", "County", "Points")
    For i = 1 To UBound(titles) + 1
        With newSheet.Cells(1, i)
            .Value = titles(i - 1)
            .Font.Name = FONT_NAME
            .Font.Size = TITLE_FONT_SIZE
            .Font.Bold = True
        End With
    Next i
    
    ' Populate Player Details
    row = 2
    For i = 2 To GetNumberOfEntries(entrants) + 1
        If Not IsEmpty(entrants.Cells(i, currentColumn)) Then
            ' Licence Number
            newSheet.Cells(row, 1).Formula = "=MASTER!C" & i
            newSheet.Cells(row, 1).Font.Name = FONT_NAME
            newSheet.Cells(row, 1).Font.Size = FONT_SIZE

            ' Name
            newSheet.Cells(row, 2).Formula = "=MASTER!D" & i & " & "" "" & MASTER!E" & i
            newSheet.Cells(row, 2).Font.Name = FONT_NAME
            newSheet.Cells(row, 2).Font.Size = FONT_SIZE

            ' County
            newSheet.Cells(row, 3).Formula = "=MASTER!G" & i
            newSheet.Cells(row, 3).Font.Name = FONT_NAME
            newSheet.Cells(row, 3).Font.Size = FONT_SIZE
            
            ' Points
            points = GetRankingPoints(category, newSheet.Cells(row, 1).Value)
            newSheet.Cells(row, 4).Value = points
            newSheet.Cells(row, 4).Font.Name = FONT_NAME
            newSheet.Cells(row, 4).Font.Size = FONT_SIZE

            row = row + 1
        End If
    Next i

    newSheet.Columns("A:D").AutoFit
    ' Find last used row in Column D
    lastRow = newSheet.Cells(newSheet.Rows.Count, "D").End(xlUp).Row

    ' Sort only columns A to D based on values in Column D
    newSheet.Range("A1:D" & lastRow).Sort Key1:=newSheet.Range("D1"), Order1:=xlDescending, Header:=xlYes
End Sub

' Get ranking points based on category and licence number
Private Function GetRankingPoints(category As String, licenceNumber As Double) As Integer
    Dim rankingList As Worksheet
    Dim lastRow As Long
    Dim points As Integer
    Dim row As Long

    Set rankingList = alphaList.Worksheets(1)
    lastRow = rankingList.Cells(rankingList.Rows.Count, "H").End(xlUp).row
    points = 0
    
    For row = 2 To lastRow
        If rankingList.Cells(row, "H").Value = category And rankingList.Cells(row, "J").Value = licenceNumber Then
            points = rankingList.Cells(row, "E").Value
            Exit For
        End If
    Next row
    
    GetRankingPoints = points
End Function

' Get the number of entries in the "Entries" sheet
Private Function GetNumberOfEntries(entrants As Worksheet) As Integer
    Dim entries As Integer
    Dim row As Long

    entries = 0
    row = 2

    Do While Not IsEmpty(entrants.Cells(row, 4))
        entries = entries + 1
        row = row + 1
    Loop

    GetNumberOfEntries = entries
End Function

' Find the Alpha List file based on the pattern
Private Function GetAlphaListFilePath(folderPath As String, filePattern As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        If file.Name Like filePattern Then
            GetAlphaListFilePath = file.path
            Exit Function
        End If
    Next file
    
    GetAlphaListFilePath = "" ' Return empty string if no file is found
End Function
