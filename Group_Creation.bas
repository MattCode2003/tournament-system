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

        ' Calculates the number of groups needed        
        If numberOfEntries Mod groupSize = 0 Then
            numberOfGroups = numberOfEntries \ groupSize
        Else
            smallerGroupSize = (MsgBox("Do spare players result in smaller groups?", _
                vbYesNo + vbDefaultButton2, "Smaller Group Sizes") = vbYes)
            
            If smallerGroupSize Then
                numberOfGroups = (numberOfEntries \ groupSize) + 1
            Else
                numberOfGroups = numberOfEntries \ groupSize
                groupSize = (numberOfEntries + numberOfGroups - 1) \ numberOfGroups
            End If
        End If

        ' Creates draw using either random or snaked groups
        If snake Then
            SnakeDraw numberOfEntries, numberOfGroups, groupSize
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
Sub SnakeDraw(number_of_entries As Integer, number_of_groups As Integer, max_group_size As Integer)
    Dim players As Collection
    Dim groups() As Variant
    Dim i As Integer
    Dim j As Integer
    Dim snake_path() As Variant
    Dim player As Player
    Dim location As Integer
    Dim group As Integer
    Dim position As Integer
    Dim stored_location As Integer
    Dim clash As Boolean
    
    ' === Get Players ===
    Set players = CreatePlayers(number_of_entries)
    
    ' === Initialize groups (2D Array) ===
    ReDim groups(1 To number_of_groups)
    For i = 1 To number_of_groups
        Dim inner_array() As Variant
        ReDim inner_array(1 To max_group_size)

        ' Initalise each position in the inner array with a new player object
        For J = 1 To max_group_size
            Set inner_array(j) = New Player
        Next J

        groups(i) = inner_array
    Next i



    ' === Build Snake Path ===
    snake_path = BuildSnakePath(number_of_groups, max_group_size)

    ' === Assign players to groups ===
    location = 1
    For i = 1 To number_of_entries

        ' If location occupied then move to the next avaiable location
        While groups(snake_path(location, 1))(snake_path(location, 2)).Name <> ""
            location = location + 1
        Wend

        ' Gets the player
        Set player = players(i)

        group = snake_path(location, 1)
        position = snake_path(location, 2)

        ' Checks for Association Clash
        If HasAssociationClash(groups(group), player.Association) > 0 Then
            clash = True
            stored_location = location
            Do
                location = location + 1
                group = snake_path(location, 1)
                position = snake_path(location, 2)
                

                If HasAssociationClash(groups(group), player.Association) = 0 Then
                    Set groups(group)(position) = player
                    location = stored_location
                    clash = False
                
                ElseIf location = stored_location + 10 Or location = number_of_groups * max_group_size Then
                    location = stored_location
                    group = snake_path(location, 1)
                    position = snake_path(location, 2)
                    Set groups(group)(position) = player
                    location = location + 1
                    clash = False
                End If
                    
            Loop While clash = True
        
        ' what to do when there isnt an association clash
        Else
            Set groups(group)(position) = player
            location = location + 1
        End If

    Next i


    ' puts the groups on excel
    Dim column As Integer
    For i = 1 To number_of_groups
        column = Cells(1, Columns.Count).End(xlToLeft).Column + 1
        Dim p As Integer
        For p = LBound(groups(i)) To UBound(groups(i))
            Set player = groups(i)(p)
            Cells(i + 1, column).Value = player.LicenceNumber
            column = column + 1
            Cells(i + 1, column).Value = player.Name
            column = column + 1
            Cells(i + 1, column).Value = player.Association
            column = column + 1
        Next p
    Next i

End Sub



' THIS CAN BE MADE MORE EFFICIENT
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


' Builds the path the snake will follow
Private Function BuildSnakePath(number_of_groups As Integer, max_group_size As Integer)
    Dim snake_path() As Variant
    Dim i As Integer
    Dim group As Integer
    Dim position As Integer
    Dim direction As Integer
    Dim places As Integer

    places = number_of_groups * max_group_size

    ReDim snake_path(1 To places, 1 To 2)

    group = 1
    position = 1
    direction = 1

    For i = 1 To places
        snake_path(i, 1) = group
        snake_path(i, 2) = position

        'Change direction at ends
        If direction = 1 Then
            If group = number_of_groups Then
                direction = -1
                position = position + 1
            Else
                group = group + 1
            End If
        Else
            If group = 1 Then
                direction = 1
                position = position + 1
            Else
                group = group - 1
            End If
        End If
    Next i

    BuildSnakePath = snake_path
End Function


' Counts the number of association clashes in a given group
Private Function HasAssociationClash(group As Variant, association As String) As Integer
    Dim player As Variant
    Dim number_of_clashes As Integer

    number_of_clashes = 0

    For Each player In group
        If Not IsEmpty(player) And IsObject(player) Then
            If TypeName(player) = "Player" Then
                If player.Association = association Then
                    number_of_clashes = number_of_clashes + 1
                End If
            End If
        End If
    Next player

    HasAssociationClash = number_of_clashes
End Function

Sub RandomDraw()
End Sub