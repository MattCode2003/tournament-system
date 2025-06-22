Attribute VB_Name = "Group_Sheets"

Private DataFilePath As String
Private wb As Workbook
Private Const DATA_FILE_NAME As String = "data.xlsx"

Sub AllEventGroupSheets()
    ' Set data.xlsx file location
    DataFilePath = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.path) & DATA_FILE_NAME
    Set wb = Workbooks.Open(DataFilePath)
End Sub

Sub SingleEventGroupSheet()
    Dim events As Variant

    ' Set data.xlsx file location
    DataFilePath = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.path) & DATA_FILE_NAME
    Set wb = Workbooks.Open(DataFilePath)

    events = GetUniqueValues()
    Debug.Print events(0)
End Sub


' Gets all the different events in the draw sheet
Private Function GetUniqueValues() As Variant
    Dim ws As Worksheet
    Dim dict As Object
    Dim last_row as Long
    Dim i As Long
    Dim cell_value As Variant
    Dim data_array As Variant
    Dim result_array() As Variant

    Set ws = wb.Sheets("Draw")
    Set dict = CreateObject("Scripting.Dictionary")
    last_row = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Gets Reads all the values and stores them in an array
    data_array = ws.Range("B2:B" & last_row).Value

    ' Gets all the unique values and stores them in a dictionary
    For i = 1 To UBound(data_array, 1)
        cell_value = data_array(i, 1)
        If Not dict.exists(cell_value) And cell_value <> "" And cell_value <> "Event" Then
            dict. Add cell_value, Nothing
        End If
    Next i

    ' Clears the memory
    data_array = Empty

    ' Converts the dictionary into an array
    If dict.Count > 0 Then
        ReDim result_array(0 To dict.Count - 1)
        i = 0
        For Each cell_value In dict.Keys
            result_array(i) = cell_value
            i = i + 1
        Next cell_value

        GetUniqueValues = result_array
    Else
        GetUniqueValues = Array()
    End If
End Function