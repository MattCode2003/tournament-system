Attribute VB_Name = "Group_Sheets"

Private DataFilePath As String
Private wb As Workbook
Private ws As Worksheet
Private Const DATA_FILE_NAME As String = "data.xlsx"
Public events As Variant


Sub AllEventGroupSheets()
    ' Set data.xlsx file location
    DataFilePath = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.path) & DATA_FILE_NAME
    Set wb = Workbooks.Open(DataFilePath)
    Set ws = wb.Sheets("Draw")
End Sub


Sub SingleEventGroupSheet()
    Dim form As New EventSelectionForm
    Dim unique_values_index As Variant
    Dim i As Long

    ' Set data.xlsx file location
    DataFilePath = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.path) & DATA_FILE_NAME
    Set wb = Workbooks.Open(DataFilePath)

    ' Gets all the events
    unique_values_index = GetUniqueValues()

    'Creates an array of just the events
    ReDim events(0 To UBound(unique_values_index))
    For i = 0 To UBound(unique_values_index)
        events(i) = unique_values_index(i, 0)
    Next i

    ' Gets the event from the user
    form.Show

    ' Gets the starting row and creates the group sheet
    For i = 0 To UBound(unique_values_index)
        If unique_values_index(i, 0) = form.selected_event_Value Then
            Call CreateGroupSheet(GetGroupsFromSheet(CLng(unique_values_index(i, 1))))
            Exit For
        End If
    Next i

End Sub


' Gets all the different events in the draw sheet
Private Function GetUniqueValues() As Variant
    Dim dict As Object
    Dim last_row As Long
    Dim i As Long
    Dim cell_value As Variant
    Dim data_array As Variant
    Dim result_array As Variant

    Set ws = wb.Sheets("Draw")
    Set dict = CreateObject("Scripting.Dictionary")
    last_row = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Get all values in column B
    data_array = ws.Range("B2:B" & last_row).Value

    ' Collect unique values and their first row number
    For i = 1 To UBound(data_array, 1)
        cell_value = data_array(i, 1)
        If Not dict.exists(cell_value) And cell_value <> "" And cell_value <> "Event" Then
            dict.Add cell_value, i
        End If
    Next i

    ' Convert dictionary to 2D array
    If dict.Count > 0 Then
        ReDim result_array(0 To dict.Count - 1, 0 To 1)
        i = 0
        Dim key As Variant
        For Each key In dict.Keys
            result_array(i, 0) = key              ' Unique value
            result_array(i, 1) = dict(key) + 1    ' First occurrence row number
            i = i + 1
        Next key

        GetUniqueValues = result_array
    Else
        GetUniqueValues = Array()
    End If
End Function


' Gets the groups for the event
Private Function GetGroupsFromSheet(start_row As Long) As Collection
    Dim groups As New Collection
    Dim current_group As Collection
    Dim player As Player
    Dim data_array As Variant
    Dim row As Long
    Dim col As Long
    Dim i As Long
    Dim last_row As Long
    Dim last_col As Long

    ' Gets the number of groups in the event
    last_row = start_row
    Do While Not IsEmpty(ws.Cells(last_row + 1, "B").Value)
        last_row = last_row + 1
    Loop

    ' Gets all the players in the groups
    last_col = ws.Cells(start_row - 1, ws.Columns.Count).End(xlToLeft).Column
    data_array = ws.Range(ws.Cells(start_row, "E"), ws.Cells(last_row, last_col)).Value

    ' Loops through each group
    For row = LBound(data_array, 1) To UBound(data_array, 1)
        Set current_group = New Collection

        ' Adds the players in that group to the collection
        For i = 1 To UBound(data_array, 2) / 3
            col = (i * 3) - 2
            If data_array(row, col) = "" Then Exit For
            Set player = New Player
            player.LicenceNumber = data_array(row, col)
            player.Name = data_array(row, col + 1)
            player.Association = data_array(row, col + 2)
            current_group.Add player
        Next i
        groups.Add current_group
    Next row

    Set GetGroupsFromSheet = groups
End Function


' Creates the group sheets for an event
Private Sub CreateGroupSheet(groups As Collection)
    Debug.Print wb.Sheets("Draw").Cells(start_row, 6).Value
End Sub