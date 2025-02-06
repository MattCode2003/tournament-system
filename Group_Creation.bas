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
    Dim positionInGroup As Integer
    Dim direction As Integer
    Dim licenceNumberGroup() As Variant
    Dim playerGroup() As String
    Dim countyGroup() As String
    Dim newGroupInfomation(1 To 3) As Integer

    ' Initialize variables
    straightToKO = Cells(numberOfEntries + 2, 2)
    firstEmptyColumn = Cells(1, Columns.Count).End(xlToLeft).Column + 1
    groupNumber = 1
    groupPosition = 1
    direction = 1
    ReDim licenceNumberGroup(1 To numberOfGroups, 1 To groupSize)
    ReDim playerGroup(1 To numberOfGroups, 1 To groupSize)
    ReDim countyGroup(1 To numberOfGroups, 1 To groupSize)
    
    ' Adds the top seeds in each group
    For i = 2 To numberOfGroups + 1
        licenceNumberGroup(i - 1, groupPosition) = Cells(i, 1)
        playerGroup(i - 1, groupPosition) = Cells(i, 2)
        countyGroup(i - 1, groupPosition) = Cells(i, 3)

        groupNumber = groupNumber + direction
    Next i

    groupPosition = groupPosition + 2
    direction = -1

    ' Does the remaining players
    For i = i + 1 To numberOfEntries
        ' check if the current space is taken
        ' this will be because of county clashes
        Do Until playerGroup(groupNumber, groupPosition) = ""
            newGroupInfomation = ChangeGroup(groupNumber, groupPosition, numberOfGroups, direction)
            groupNumber = newGroupInfomation(1)
            groupPosition = newGroupInfomation(2)
            direction = newGroupInfomation(3)
        Loop

        ' Checks if there is a county clash
    Next i

End Sub

' Calculates the next group and the position in that group
Function ChangeGroup(groupNumber As Integer, groupPosition As Integer, numberOfGroups As Integer, direction As Integer) As Variant
    Dim results(1 To 3) As Integer
 
    If groupNumber = 1 And direction = -1 Then
        results(1) = groupNumber
        results(2) = groupPosition + 1
        results(3) = 1
    ElseIf groupNumber = numberOfGroups And direction = 1 Then
        results(1) = groupNumber
        results(2) = groupPosition + 1
        results(3) = -1
    Else
        results(1) = groupNumber + direction
        results(2) = groupPosition
        results(3) = direction
    End If
    
    ChangeGroup = results
        
End Function


Sub RandomDraw()
End Sub