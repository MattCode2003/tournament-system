Attribute VB_Name = "Draw_Sheet"

Private DataFilePath As String
Private wb As Workbook
Private Const DATA_FILE_NAME As String = "data.xlsx"

Sub DrawAndGroupSheetForm()

    DataFilePath = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.path) & DATA_FILE_NAME
    If Not UtilityFunctions.FileExists(DataFilePath) Then Exit Sub

    Set wb = Workbooks.Open(DataFilePath)
    GroupSheetsForm.Show
End Sub



Public Sub CreateDrawSheet()
    Dim pass_checks As Boolean
    Dim draw_ws As Worksheet
    Dim comp_name As String
    Dim event_sheets As Collection
    Const title As String = "(Players in each group go across)"
    Dim row As Integer
    Dim event_ws As Worksheet
    Dim max_players As Integer
    Dim col As Integer
    Dim start_cell As Range
    Dim number_of_groups As Integer
    Dim event_col As Integer
    Dim event_row As Integer
    Dim event_cell As Range

    ' Does all the sanity checks
    pass_checks = SanityChecks
    If Not pass_checks Then Exit Sub

    ' Creates a new worksheet
    Set draw_ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    draw_ws.Name = "Draw"

    ' Creates the titles
    comp_name = ThisWorkbook.Worksheets("General Settings").Cells(3, 2)
    With draw_ws.Range("A1:I1")
        .Merge
        .Value = comp_name & "   " & title
        .Characters(0, Len(comp_name)).Font.Bold = True
        .Characters(Len(comp_name) + 4, Len(title)).Font.Color = RGB(255, 0, 0)
    End With

    ' Goes through each event to see if there is any groups creates
    ' Stores them in event_sheets
    Set event_sheets = GetEvents()

    row = 3
    For Each event_ws In event_sheets
        ' Get the max number of players in an group
        max_players = GetMaxNumberPlayersInGroup(event_ws)
        
        ' Does the title section
        draw_ws.Cells(row, 1).Value = "Date"
        draw_ws.Cells(row, 2).Value = "Event"
        draw_ws.Cells(row, 3).Value = "Time"
        draw_ws.Cells(row, 4).Value = "Group"

        col = 5
        For i = 1 To max_players
            draw_ws.Cells(row, col).Value = "Cod" & Chr(64 + i)
            col = col + 1
            draw_ws.Cells(row, col).Value = "Player" & Chr(64 + i)
            col = col + 1
            draw_ws.Cells(row, col).Value = "c" & Chr(64 + i)
            col = col + 1
        Next i

        Rows(row).RowHeight = 19.5
        
        With draw_ws.Range(draw_ws.Cells(row, 1), draw_ws.Cells(row, col - 1))
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
            .Borders(xlRight).LineStyle = xlContinuous
            .Borders(xlRight).Weight = xlThin
            .Borders(xlRight).Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With draw_ws.Cells(row, draw_ws.Columns.Count).End(xlToLeft).Borders(xlRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With

        With draw_ws.Cells(row, 1).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With

        row = row + 1

        ' Calculates the number of groups
        Set start_cell = event_ws.Cells(1, Columns.Count).End(xlToLeft).Offset(1, 1)
        number_of_groups = event_ws.Cells(event_ws.Rows.Count, start_cell.column).End(xlUp).row - 1

        For i = 1 To number_of_groups
            ' Date
            With draw_ws.Cells(row, 1)
                .Font.Bold = True
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeTop).Color = RGB(0, 0, 0)

                If i = number_of_groups Then
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                End If
            End With

            ' Event
            With draw_ws.Cells(row, 2)
                .Value = event_ws.Name
                .Font.Bold = True
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeTop).Color = RGB(0, 0, 0)

                If i = number_of_groups Then
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                End If
            End With

            ' Time
            With draw_ws.Cells(row, 3)
                .Font.Bold = True
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeTop).Color = RGB(0, 0, 0)

                If i = number_of_groups Then
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                End If
            End With
            
            ' Group Number
            With draw_ws.Cells(row, 4)
                .Value = i
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeTop).Color = RGB(0, 0, 0)

                If i = number_of_groups Then
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                End If
            End With

            ' Player Details
            event_col = start_cell.column
            event_row = start_cell.row + (i - 1)
            Set event_cell = event_ws.Cells(event_row, event_col)
            
            col = 5
            Do While event_cell.Value <> ""
                With draw_ws.Cells(row, col)
                    .Value = event_cell.Value
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).Weight = xlThin
                    .Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlThin
                    .Borders(xlEdgeTop).Color = RGB(0, 0, 0)

                    If i = number_of_groups Then
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).Weight = xlMedium
                        .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                    End If
                End With

                col = col + 1
                event_col = event_col + 1

                Set event_cell = event_ws.Cells(event_row, event_col)
            Loop

            With draw_ws.Cells(row, col - 1)
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlEdgeRight).Color = RGB(0, 0, 0)
            End With

            row = row + 1
        Next i
        row = row + 1
    Next event_ws
End Sub



Private Function SanityChecks() As Boolean
    Dim has_form_sheet As Boolean
    Dim has_draw_dheet As Boolean
    Dim required_sheet_count As Integer
    Dim user_response As VbMsgBoxResult
    Dim ws As Worksheet

    ' Check if form sheet exists
    has_form_sheet = UtilityFunctions.CheckSheetNames(wb, "Form")

    ' Determine required minimum sheet count
    required_sheet_count = IIf(has_form_sheet, 3, 2)

    ' Check if workbook meets minimum sheet requirement
    If wb.Worksheets.Count >= required_sheet_count Then

        ' Check if draw sheet already exists
        has_draw_sheet = UtilityFunctions.CheckSheetNames(wb, "Draw")

        If has_draw_sheet Then
            user_response = MsgBox( _
                "A draw already exists. Continuing will delete and recreate it." & vbCrLf & _
                "Are you sure you want to continue?", _
                vbYesNo + vbExclamation, "Warning")

            If user_response = vbYes Then
                Set ws = wb.Sheets("Draw")
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
                SanityChecks = True
            Else
                SanityChecks = False
            End If

        Else
            SanityChecks = True
        End If

    Else
        MsgBox "Ranking points and groups not created", vbCritical, "Error"
        SanityChecks = False
    End If
End Function



' Goes through each event to see if there is any groups creates
Private Function GetEvents() As Collection
    Dim wss As New Collection
    Dim i As Integer
    Dim first_empty_cell As Range
    Dim ws As Worksheet

    For i = IIf(UtilityFunctions.CheckSheetNames(wb, "Form"), 1, 2) To wb.Worksheets.Count - 1
        Set ws = wb.Worksheets(i)
        If ws.Cells(1, Columns.Count).End(xlToLeft).Offset(1, 1) <> "" Then
            wss.Add ws
        End If
    Next i

    Set GetEvents = wss
End Function


' THIS DOES NOT FUNCTION CORRECTLY
Private Function GetMaxNumberPlayersInGroup(ws As Worksheet) As Integer
    Dim start_cell As Range
    Dim max_count As Integer
    Dim current_count As Integer
    Dim current_row As Integer
    Dim current_cell As Range
    Dim group_start_col As Integer

    ' Identify the starting point for group data
    Set start_cell = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Offset(1, 1)
    group_start_col = start_cell.column
    current_row = start_cell.row

    max_count = 0

    ' Loop through rows until there's no group data in the starting column
    Do While ws.Cells(current_row, group_start_col).Value <> ""
        current_count = 0
        Set current_cell = ws.Cells(current_row, group_start_col)

        ' Count non-empty cells across the row starting from the group start column
        Do While current_cell.Value <> ""
            current_count = current_count + 1
            Set current_cell = current_cell.Offset(0, 1)
        Loop

        If current_count > max_count Then
            max_count = current_count
        End If

        current_row = current_row + 1
    Loop

    ' Each player takes 3 columns
    GetMaxNumberPlayersInGroup = max_count \ 3
End Function

