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
            SnakeDraw numberOfEntries, numberOfGroups
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


' Adjusted Snake System with Even Group Distribution
Sub SnakeDraw(numberOfEntries As Integer, numberOfGroups As Integer)
    Dim players As Collection
    Dim groups() As Collection
    Dim i As Integer
    Dim groupNumber As Integer, groupPosition As Integer
    Dim direction As Integer
    Dim player As Object
    Dim results As Variant

    ' Get Players
    Set players = CreatePlayers(numberOfEntries)
    
    ' Initialize groups (each group is a collection)
    ReDim groups(1 To numberOfGroups)
    For i = 1 To numberOfGroups
        Set groups(i) = New Collection
    Next i

    ' Initialize variables for group and position
    groupNumber = 1
    groupPosition = 1
    direction = 1 ' Start by filling groups in a forward direction (1)

    ' Distribute players into groups using the snake system
    For Each player In players
        ' Add player to the current group
        groups(groupNumber).Add player
        
        ' Determine the next group and position using the snake system
        results = ChangeGroup(groupNumber, groupPosition, numberOfGroups, direction)
        groupNumber = results(1)
        groupPosition = results(2)
        direction = results(3)
    Next player

    ' For debugging or verification purposes, print group assignments
    Dim column As Integer
    For i = 1 To numberOfGroups
        column = Cells(1, Columns.Count).End(xlToLeft).Column + 1
        For Each player In groups(i)
            Cells(i + 1, column).Value = player.LicenceNumber
            column = column + 1
            Cells(i + 1, column).Value = player.Name
            column = column + 1
            Cells(i + 1, Column).Value = player.Association
            column = column + 1
        Next Player
    Next i
    AssociationClash groups, numberOfEntries, numberOfGroups
End Sub




Private Function CreatePlayers(numberOfEntries As Integer) As Collection
    Dim players As New Collection
    Dim p As Player
    Dim i As Integer

    For i = 2 + Cells(numberOfEntries + 2, 2).Value To numberOfEntries + 1
        Set p = New Player
        p.LicenceNumber = Cells(i, 1)
        p.Name = Cells(i, 2)
        p.Association = Cells(i, 3)
        
        players.Add p
    Next i

    Set CreatePlayers = players
End Function


' Calculates the next group and the position in that group
Private Function ChangeGroup(groupNumber As Integer, groupPosition As Integer, numberOfGroups As Integer, direction As Integer) As Variant
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


Sub AssociationClash(groups() As Collection, numberOfEntries As Integer, numberOfGroups As Integer)
    Dim groupNumber As Integer
    Dim groupPosition As Integer
    Dim direction As Integer
    Dim i As Integer
    Dim j As Integer
    Dim association As String
    Dim clash As Boolean
    Dim newGroup As Integer
    Dim newPosition As Integer
    Dim newDirection As Integer

    ' Initialize variables
    groupNumber = numberOfGroups
    groupPosition = 2
    direction = -1

    ' goes through the groups in the same way as snake system
    For i = numberOfGroups + 1 To numberOfEntries
        ' gets the specific players association
        association = groups(groupNumber).Item(groupPosition).Association

        ' Goes through all the previous players in the same group
        ' to find any clashes
        clash = False
        For j = 1 To groupPosition - 1
            If association = groups(groupNumber).Item(j).Association Then
                clash = True
                Exit For
            End If
        Next j

        ' if there is a clash then search through the groups using the changegroup sub
        ' to find a group where there is no clash
        ' once a group is found put the player in that group and position
        ' this moves everyone else down a group/position
    Next i
End Sub

Sub RandomDraw()
End Sub