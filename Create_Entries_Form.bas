Attribute VB_Name = "Create_Entries_Form"

Sub Create_Entries_Form()
    Dim filePath As String
    Dim i As Integer, x As Integer, j As Integer
    Dim event_included As Variant, event_name As Variant
    Dim EventSettings As Worksheet
    Dim titles As Variant, column_width As Variant
    Dim wb As Workbook
    Dim EntryLocation As String, column_letter As String
    Dim rng As Range
    Dim user_response As VbMsgBoxResult
    Dim entrant_number As Integer
    Dim Delimiter As String
    Dim fso As Object
    Dim comp_date As String
    Dim age_category As Integer
    Dim general_settings As Worksheet
    Dim comp_month As Integer
    Dim comp_year As Integer
    Dim birth_year As Integer
    Dim row As Integer
    Dim formula_string As String
    Dim or_condition As String
    
    
    entrant_number = 301
    Set general_settings = Workbooks("Main.xlsm").Worksheets("General Settings")

    filePath = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.Path) & "data.xlsx"

    ' Check if file exists and confirm overwrite
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        user_response = MsgBox("Creating a new entry list will delete your current one." & vbCrLf & _
                               "Would you like to continue?", vbYesNo + vbExclamation, "Warning")
        If user_response = vbNo Then Exit Sub
        Kill filePath
    End If

    ' Checks if competition date has been entered
    If general_settings.Range("B4").Value = "" Then
        MsgBox ("No Competition Date Provided")
        Exit Sub
    End If

    ' Create new workbook and save
    Set wb = Workbooks.Add
    EntryLocation = ThisWorkbook.path & UtilityFunctions.GetDelimiter(ThisWorkbook.Path) & "data.xlsx"
    Application.Wait Now + TimeValue("00:00:02")
    wb.SaveAs EntryLocation
    wb.Sheets(1).Name = "MASTER"

    ' Activate the new workbook
    wb.Activate

    '============================ Titles ===============================

    ' General title settings
    With Rows(1)
        .RowHeight = 10.5
        .Font.Bold = True
        .Font.Size = 8
        .Font.Name = "Arial"
    End With

    ' Initial set of titles
    titles = Array("Entry No", "Player No", "Licence No", "First Name", "Surname", "DOB", "County", "Sex", "Email")
    column_width = Array(6.64, 9.91, 9.18, 11.09, 18.64, 9.36, 6.36, 3.09, 37.91)

    For i = 1 To UBound(titles) + 1
        With Cells(1, i)
            .Value = titles(i - 1)
            .HorizontalAlignment = xlCenter
        End With
        Columns(i).ColumnWidth = column_width(i - 1)
    Next i

    ' Event titles from "Event Settings"
    Set EventSettings = Workbooks("Main.xlsm").Worksheets("Event Settings")

    x = 3
    Do
        event_name = EventSettings.Range("J" & x).Value
        event_included = EventSettings.Range("B" & x).Value
        If IsEmpty(event_name) Then Exit Do
        If Not IsEmpty(event_included) Then
            With Cells(1, i)
                .Value = event_name
                .HorizontalAlignment = xlCenter
            End With
            Columns(i).ColumnWidth = 3.64
            i = i + 1
        End If

        x = x + 1
    Loop

    ' Final set of titles
    titles = Array("Entry", "Paid", "Owes", "Comments")
    column_width = Array(11.46, 11.46, 11.46, 27.09)

    For x = LBound(titles) To UBound(titles)
        With Cells(1, i)
            .Value = titles(x)
            .HorizontalAlignment = xlCenter
        End With
        Columns(i).ColumnWidth = column_width(x)

        If titles(x) <> "Comments" Then
            column_letter = Split(Cells(1, i).Address, "$")(1)
            Range(column_letter & "2:" & column_letter & entrant_number).NumberFormat = _
                "_(£* #,##0.00_);_(£* (#,##0.00);_(£* ""-""??_);_(@_)"
        End If

        ' Conditional formatting and formula for "Owes"
        If titles(x) = "Owes" Then
            column_letter = Split(Cells(1, i).Address, "$")(1)
            For j = 2 To entrant_number
                Cells(j, i).Formula = "=" & Cells(j, i - 2).Address & "-" & Cells(j, i - 1).Address
            Next j

            Set rng = Range(column_letter & "2:" & column_letter & entrant_number)
            With rng.FormatConditions
                .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                .Item(1).Interior.Color = RGB(255, 255, 0)
            End With
        End If
        i = i + 1
    Next x

    '============================= Formatting ===============================

    ' Entry Number
    With Range("A2:A" & entrant_number)
        .Formula = "=ROW()-1"
        .HorizontalAlignment = xlCenter
    End With

    ' General formatting
    With Range("A2:AF" & entrant_number + 3)
        .Font.Size = 8
        .Font.Name = "Arial"
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With

    ' DOB column formatting
    Range("F2:F" & entrant_number).NumberFormat = "dd-mmm-yy"

    ' Freeze panes
    Range("G2").Select
    ActiveWindow.FreezePanes = True

    ' Birthday formatting for age group events
    For x = 10 To i - 5

        ' Gets category age
        If Cells(1, x).Value = "JB" Or Cells(1, x).Value = "JG" Then
            age_category = 19
        ElseIf Cells(1, x).Value = "CB" Or Cells(1, x).Value = "CG" Then
            age_category = 15
        Else
            age_category = UtilityFunctions.ExtractNumberFromString(Cells(1, x).Value)
        End If
        
        ' Goes to next iteration if there is no age Requirement
        If age_category <> 0 Then
            ' Calculates the max year to be born
            comp_date = FullDateToShortDate(general_settings.Range("B4").Value)
            comp_month = CInt(Mid(comp_date, 4, 2))
            comp_year = CInt(Right(comp_date, 4))

            birth_year = comp_year - age_category
            
            ' Compensates for new season
            If comp_month >= 8 Then
                birth_year = birth_year + 1
            End If

            column_letter = Split(Cells(1, x).Address, "$")(1)

            ' Highlight yellow if black condition is met AND cell has a value
            Set rng = Range(column_letter & "2:" & column_letter & entrant_number + 3)
            With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(F2<>"""", YEAR(F2)<" & birth_year & ", " & Cells(2, x).Address(False, False) & "<>"""")")
                .Interior.Color = RGB(255, 255, 0) ' Yellow
                .StopIfTrue = True ' Prevent lower rules from applying
            End With

            
            With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(F2<>"""", YEAR(F2)<" & birth_year & ")")
                .Interior.Color = RGB(0, 0, 0)
                .StopIfTrue = False
            End With
        End If

        column_letter = Split(Cells(1, x).Address, "$")(1)


        ' Gender Formatting
        ' If player is male
        If InStr(1, Cells(1, x).Value, "B", vbTextCompare) or InStr(1, Cells(1, x).Value, "M", vbTextCompare) Then
            Set rng = Range(column_letter & "2:" & column_letter & entrant_number + 3)

            ' Sets background to yellow if invalid entry
            With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(H2<>"""", H2<>""M"", " & column_letter & "2<>"""")")
                .Interior.Color = RGB(255, 255, 0)
                .StopIfTrue = True
            End With

            ' Sets background to black if not eligable
            With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(H2<>"""", H2<>""M"")")
                .Interior.Color = RGB(0, 0, 0)
                .StopIfTrue = False
            End With
        End If

        ' If player is female
        If InStr(1, Cells(1, x).Value, "G", vbTextCompare) or InStr(1, Cells(1, x).Value, "W", vbTextCompare) Then
            Set rng = Range(column_letter & "2:" & column_letter & entrant_number + 3)

            ' Sets background to yellow if invalid entry
            With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(H2<>"""", H2<>""F"", " & column_letter & "2<>"""")")
                .Interior.Color = RGB(255, 255, 0)
                .StopIfTrue = True
            End With

            ' Sets background to black if not eligable
            With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(H2<>"""", H2<>""F"")")
                .Interior.Color = RGB(0, 0, 0)
                .StopIfTrue = False
            End With
        End If
    Next x


    ' Adds the price of each event to the worksheet
    With Cells(1, i + 10)
        .Value = "Category"
        .HorizontalAlignment = xlCenter
    End With
    With Cells(1, i + 11)
        .Value = "Price"
        .HorizontalAlignment = xlCenter
    End With

    x = 3
    row = 2
    Do
        event_name = EventSettings.Range("J" & x).Value
        event_included = EventSettings.Range("B" & x).Value
        If IsEmpty(event_name) Then Exit Do
        If Not IsEmpty(event_included) Then
            With Cells(row, i + 10)
                .Value = event_name
                .HorizontalAlignment = xlCenter
            End With
            With Cells(row, i + 11)
                .Value = event_included
                .HorizontalAlignment = xlCenter
                .Font.Size = 8
                .Font.Name = "Arial"
            End With
            row = row + 1
        End If

        x = x + 1
    Loop
    
    ' Adds the price of the admin fee
    with Cells(row, i + 10)
        .Value = "Admin"
        .HorizontalAlignment = xlCenter
    End With
    with Cells(row, i + 11)
        .Value = EventSettings.Cells(x + 2, 2)
        .HorizontalAlignment = xlCenter
        .Font.Size = 8
        .Font.Name = "Arial"
    End With

    Columns(i + 10).ColumnWidth = 7
    Columns(i + 11).ColumnWidth = 7

    ' Creates the formula for the entries
    formula_string = "=SUM("
    or_condition = ""
    For x = 10 to i - 5
        formula_string = formula_string & "IF(" & Cells(2, x).Address(False, False) & "<>"""", " & _
                        Cells(x - 8, i + 11).Address(False, False) & ", 0), "
        or_condition = or_condition + Cells(2, x).Address(False, False) & "<>"""", "
    Next x
    
    
    or_condition = Left(or_condition, Len(or_condition) - 2)  ' Remove trailing comma and space
    formula_string = formula_string & "IF(OR(" & or_condition & "), " & Cells(row, i + 11) & ", 0))"


    ' Creates the formula for the entries
    column_letter = Split(Cells(1, i - 4).Address, "$")(1)
    Set rng = Range(column_letter & "2:" & column_letter & entrant_number + 3)
    rng.Formula = formula_string

    '=========================== Footer ================================

    ' Border
    column_letter = Split(Cells(1, i - 1).Address, "$")(1)
    Set rng = Range("A" & entrant_number + 4 & ":" & column_letter & entrant_number + 4)
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    With rng.Font
        .Name = "Arial"
        .Size = 8
    End With

    ' Total Entrants
    Cells(entrant_number + 4, 4).Value = "Total"
    Cells(entrant_number + 4, 5).Formula = "=COUNTA(E2:E" & entrant_number + 3 & ")"
    Cells(entrant_number + 4, 5).HorizontalAlignment = xlCenter
    Cells(entrant_number + 4, 6).Value = "Total Number"

    ' Event Totals
    For x = 10 To i - 5
        column_letter = Split(Cells(1, x).Address, "$")(1)
        With Cells(entrant_number + 4, x)
            .Formula = "=COUNTA(" & column_letter & "2:" & column_letter & entrant_number + 3 & ")"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next x

    ' Money Totals
    For x = i - 4 To i - 2
        column_letter = Split(Cells(1, x).Address, "$")(1)
        With Cells(entrant_number + 4, x)
            .Formula = "=SUM(" & column_letter & "1:" & column_letter & entrant_number + 3 & ")"
            .NumberFormat = "_(£* #,##0.00_);_(£* (#,##0.00);_(£* ""-""??_);_(@_)"
        End With
    Next x

    wb.Save
End Sub

'======================== Date Conversion =============================

Function FullDateToShortDate(inputDate As String) As String
    Dim formattedDate As String
    Dim dateValue As Date

    ' Remove ordinal suffix (e.g., 'st', 'nd', 'rd', 'th') using Replace function
    inputDate = Replace(inputDate, "st", "")
    inputDate = Replace(inputDate, "nd", "")
    inputDate = Replace(inputDate, "rd", "")
    inputDate = Replace(inputDate, "th", "")
    
    ' Convert string to date
    dateValue = CDate(Mid(inputDate, InStr(inputDate, " ")))
    
    ' Format the date to dd/mm/yyyy
    formattedDate = Format(dateValue, "dd/mm/yyyy")
    
    ' Return the formatted date
    FullDateToShortDate = formattedDate
End Function



