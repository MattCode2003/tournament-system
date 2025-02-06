Attribute VB_Name = "Group_Creation"

Public dataWorkbook As Workbook
Public currentEvent As String
Private Const DATA_FILE_NAME As String = "data.xlsx"
Private Const DRAW_FILE_NAME As String = "draw.xlsx"

Sub CreateGroups()
    Dim DataFilePath As String
    Dim currentWorksheet As Worksheet
    Dim numberOfEntries As Integer
    Dim recommendedSeeds As Integer
    Dim numberOfSeeds As Integer
    Dim straightToKO As Integer
    Dim snake As Boolean
    Dim groupSize As Integer
    Dim smallerGroupSize As Boolean
    Dim seeds As Integer
    Dim numberOfGroups As Integer
    Dim msg As String

    ' Sanity Check for data file
    DataFilePath = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.path) & DATA_FILE_NAME
    If Not UtilityFunctions.FileExists(DataFilePath) Then Exit Sub
    
    Set dataWorkbook = Workbooks.Open(DataFilePath)
    
    Do While True
        ' Choose event to create groups for
        GroupCreationForm.Show

        ' Setup important info
        Set currentWorksheet = dataWorkbook.Worksheets(currentEvent)
        numberOfEntries = Application.WorksheetFunction.CountA(currentWorksheet.Range("A:A")) - 1
        recommendedSeeds = RecommendedSeedNumbers(numberOfEntries)
        
        ' Gets User inputs
        msg = "You have " & Trim(Str(numberOfEntries)) & " entries." & Chr(13)
        msg = msg & "How many Go Straight to the knockouts?"
        straightToKO = Int(InputBox(Trim(msg), , Str(0)))
        currentWorksheet.Cells(numberOfEntries + 2, 2).Value = straightToKO

        snake = (MsgBox("Do you wish to use the snake system to form groups?", _
                vbYesNo + vbDefaultButton2, "Snake System") = vbYes)
        
        If Not snake Then
            seeds = Int(InputBox("How many seeds are there?", "Number of Seeds", Str(recommendedSeeds - straightToKO)))
        End If
        groupSize = Int(InputBox("How many players is in a normal group", "Group Size", Str(4)))
        
        If numberOfEntries Mod groupSize <> 0 Then
            smallerGroupSize = (MsgBox("Do spare players result in smaller groups?", _
                vbYesNo + vbDefaultButton2, "Smaller Group Sizes") = vbYes)
        End If

        ' Calculates the number of groups needed
        numberOfGroups = numberOfEntries \ groupSize
        If smallerGroupSize Then numberOfGroups = numberOfGroups + 1
        If Not smallerGroupSize Then groupSize = groupSize + 1

        ' Creates draw using either random or snaked groups
        If snake Then
            SnakeDraw numberOfEntries, groupSize, numberOfGroups
        Else
            ' RandomDraw()
            End
        End If
    Loop

End Sub

' Calculates the number of recommended seeds
Private Function RecommendedSeedNumbers(numberOfEntries As Integer) As Integer
    Dim recommendedSeeds as Integer
    Dim ix As Integer

    recommendedSeeds = 1 + Int(numberOfEntries / 24)
    ix = Int(Log(recommendedSeeds) / Log(2#))
    If recommendedSeeds > Exp(ix * Log(2#)) Then ix = ix + 1
    recommendedSeeds = Exp(ix * Log(2#)) * 2

    RecommendedSeedNumbers = recommendedSeeds
End Function


Sub SnakeDraw(numberOfEntries As Integer, groupSize As Integer, numberOfGroups As Integer)
    Dim straightToKO As Integer
    Dim groupNumber As Integer
    Dim firstEmptyColumn As Long
    Dim i As Integer
    Dim direction As Integer
    Dim licenceNumberGroup() As Variant
    Dim playerGroup() As String
    Dim countyGroup() As String

    ' Initialize variables
    straightToKO = Cells(numberOfEntries + 2, 2)
    firstEmptyColumn = Cells(1, Columns.Count).End(xlToLeft).Column + 1
    groupNumber = 1
    direction = 1
    ReDim licenceNumberGroup(1 To numberOfGroups, 1 To groupSize)
    ReDim playerGroup(1 To numberOfGroups, 1 To groupSize)
    ReDim countyGroup(1 To numberOfGroups, 1 To groupSize)
    
    ' Adds the top seeds in each group
    For i = 2 To numberOfGroups + 1
        licenceNumberGroup(i - 1, 1) = Cells(i, 1)
        playerGroup(i - 1, 1) = Cells(i, 2)
        countyGroup(i - 1, 1) = Cells(i, 3)

        groupNumber = groupNumber + direction
    Next i

    direction = -1

    ' Does the remaining players
    For i = i + 1 To numberOfEntries + 1
    Next i

End Sub

Sub RandomDraw()
End Sub