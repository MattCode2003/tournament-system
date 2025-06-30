Attribute VB_Name = "Group_Sheet_Design"

Private Const FONT As String = "Calibri"



' IMPORTANT INFO
' Table 1 = very top info such as event and group number
' Table 2 = Group Seeding, Licence Number, Player Name, County, Points, position
' Table 3 = Count Back stuff
' Table 4 = Match scores
' Table 5 = dc/sc/wc



'================================================================================ Group of 3 ================================================================================


Sub Group3(ws As worksheet, tournament_name As String, event_name As String, group_number As Long, start_time As String, table As String, dates As String, group As Variant)
    Dim i As Integer ' General For loop index
    Dim row As Integer
    Dim player_letter_locations As Variant
    Dim umpire_letter_locations As Variant

    ' Column Widths
    ws.Range("A1:BO1").EntireColumn.ColumnWidth = 1

    ' Row heights
    ws.Rows("1:3").RowHeight = 18.5         ' Table 1
    ws.Rows(4).RowHeight = 29               ' Gap between Table 1 and 2/3
    ws.Rows(5).RowHeight = 26               ' Table 2/3 Header
    ws.Rows("6:8").RowHeight = 30           ' Table 2/3 Info
    ws.Rows(9).RowHeight = 29               ' Gap between Table 2/3 and 4
    ws.Rows(10).RowHeight = 18.5            ' Table 4 Header
    ws.Rows("11:16").RowHeight = 31         ' Table 4 Info
    ws.Rows(31).RowHeight = 22              ' Table 5

    ' Table 1
    Call Table1Format(ws.Range("A1:AX1"), "Tournament: " & tournament_name, 14)
    Call Table1Format(ws.Range("A2:AX2"), "Event: " & event_name, 14)
    Call Table1Format(ws.Range("A3:AX3"), "Group: " & group_number, 14)
    Call Table1Format(ws.Range("AY1:BO1"), "Time: " & start_time, 14)
    Call Table1Format(ws.Range("AY2:BO2"), "Table: " & table, 14)
    Call Table1Format(ws.Range("AY3:BO3"), "Date: " & dates, 14)

    ' Creates the "For Referee's use in case of a tie" text box
    With ws.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=433.3125, _
            Top:=70, _
            Width:=108.692, _
            Height:=10)

        .TextFrame.Characters.Text = "For Referee's use in case of a tie"
        .TextFrame2.TextRange.Font.Size = 7
        .TextFrame2.TextRange.Font.Name = "Calibri (Body)"
        .Line.Visible = msoFalse
    End With

    ' Table 2 headers
    Call TableHeader(ws.Range("B5:D5"), "", "left", 11, False)
    Call TableHeader(ws.Range("E5:J5"), "Licence No", "centre", 11, False)
    Call TableHeader(ws.Range("K5:AG5"), "Full Name", "centre", 11, True)
    Call TableHeader(ws.Range("AH5:AK5"), "County", "centre", 8, False)
    Call TableHeader(ws.Range("AL5:AO5"), "Points", "centre", 8, False)
    Call TableHeader(ws.Range("AP5:AS5"), "Position", "right", 8, False)

    ' Table 2 Info
    For i = 1 To 3
        row = 5 + i
        
        Call TableInfoFormat(ws.Range("B" & row & ":D" & row), Chr(64 + i), True, i = 3, 14, 21, xlHAlignCenter, xlDouble, xlContinuous)
        Call TableInfoFormat(ws.Range("E" & row & ":J" & row), group(i).LicenceNumber, True, i = 3, 14, 21)
        Call TableInfoFormat(ws.Range("K" & row & ":AG" & row), group(i).Name, True, i = 3, 14, 21)
        Call TableInfoFormat(ws.Range("AH" & row & ":AK" & row), group(i).Association, False, i = 3, 14, 21)
        Call TableInfoFormat(ws.Range("AL" & row & ":AO" & row), "", False, i = 3, 14, 21)
        Call TableInfoFormat(ws.Range("AP" & row & ":AS" & row), "", False, i = 3, 14, 21, xlHAlignCenter, xlContinuous, xlDouble)
    Next i

    ' Table 3 Headers
    Call TableHeader(ws.Range("AW5:BB5"), "Sets", "left", 11, False)
    Call TableHeader(ws.Range("BC5:BH5"), "Games", "centre", 11, False)
    Call TableHeader(ws.Range("BI5:BN5"), "Points", "right", 11, False)

    ' Table 3 Info
    For i = 1 To 3
        row = 5 + i
        Call TableInfoFormat(ws.Range("AW" & row & ":AY" & row), "", False, i = 3, 14, 21, xlHAlignCenter, xlDouble, xlDot)
        Call TableInfoFormat(ws.Range("AZ" & row & ":BB" & row), "", False, i = 3, 14, 21, xlHAlignCenter, xlDot, xlContinuous)
        Call TableInfoFormat(ws.Range("BC" & row & ":BE" & row), "", False, i = 3, 14, 21, xlHAlignCenter, xlContinuous, xlDot)
        Call TableInfoFormat(ws.Range("BF" & row & ":BH" & row), "", False, i = 3, 14, 21, xlHAlignCenter, xlDot, xlContinuous)
        Call TableInfoFormat(ws.Range("BI" & row & ":BK" & row), "", False, i = 3, 14, 21, xlHAlignCenter, xlContinuous, xlDot)
        Call TableInfoFormat(ws.Range("BL" & row & ":BN" & row), "", False, i = 3, 14, 21, xlHAlignCenter, xlDot, xlDouble)
    Next i

    ' Table 4 Headers
    Call TableHeader(ws.Range("A10:D10"), "", "left", 11, False)
    Call TableHeader(ws.Range("E10:I10"), "Licence No", "centre", 9, False)
    Call TableHeader(ws.Range("J10:Y10"), "Player", "centre", 12, True)
    Call TableHeader(ws.Range("Z10:AN10"), "Coach", "centre", 12, True)
    Call TableHeader(ws.Range("AO10:AQ10"), "Ump", "centre", 8, False)
    Call TableHeader(ws.Range("AR10:AU10"), "Game 1", "centre", 8, False)
    Call TableHeader(ws.Range("AV10:AY10"), "Game 2", "centre", 8, False)
    Call TableHeader(ws.Range("AZ10:BC10"), "Game 3", "centre", 8, False)
    Call TableHeader(ws.Range("BD10:BG10"), "Game 4", "centre", 8, False)
    Call TableHeader(ws.Range("BH10:BK10"), "Game 5", "centre", 8, False)
    Call TableHeader(ws.Range("BL10:BO10"), "Winner", "right", 8, False)

    ' Table 4 Match Numbers
    row = 1
    For i = 11 To 15 Step 2
        Call TableInfoFormat(ws.Range("A" & i & ":B" & i + 1), row, False, i = 15, 12, 21, xlHAlignCenter, xlDouble, xlContinuous)
        row = row + 1
    Next i

    ' Create player orders
    player_letter_locations = Array("A", "C", "B", "C", "A", "B")
    umpire_letter_locations = Array("B", "A", "C")

    ' Table 4 Info
    For i = 0 To 5 Step 2
        Call TableInfoFormat(ws.Range("C" & i + 11 & ":D" & i + 11), player_letter_locations(i), True, False, 12, 21)                                                                                              ' Top Player Letter
        Call TableInfoFormat(ws.Range("E" & i + 11 & ":I" & i + 11), group(Asc(player_letter_locations(i)) - 64).LicenceNumber, False, False, 13, 21)                                                              ' Top Player Licence Number
        Call TableInfoFormat(ws.Range("J" & i + 11 & ":Y" & i + 11), group(Asc(player_letter_locations(i)) - 64).Name, False, False, 13, 21, xlHAlignLeft)                                                         ' Top Player Name
        Call TableInfoFormat(ws.Range("Z" & i + 11 & ":AN" & i + 11), "", False, False, 13, 21, xlHAlignLeft)                                                                                                      ' Top Coach
        Call TableInfoFormat(ws.Range("AR" & i + 11 & ":AU" & i + 11), "", False, False, 13, 21)                                                                                                                   ' Top Game 1
        Call TableInfoFormat(ws.Range("AV" & i + 11 & ":AY" & i + 11), "", False, False, 13, 21)                                                                                                                   ' Top Game 2
        Call TableInfoFormat(ws.Range("AZ" & i + 11 & ":BC" & i + 11), "", False, False, 13, 21)                                                                                                                   ' Top Game 3
        Call TableInfoFormat(ws.Range("BD" & i + 11 & ":BG" & i + 11), "", False, False, 13, 21)                                                                                                                   ' Top Game 4
        Call TableInfoFormat(ws.Range("BH" & i + 11 & ":BK" & i + 11), "", False, False, 13, 21)                                                                                                                   ' Top Game 5
        
        Call TableInfoFormat(ws.Range("AO" & i + 11 & ":AQ" & i + 12), umpire_letter_locations(i / 2), False, i = 4, 13, 21)                                                                                      ' Umpire Letter
        Call TableInfoFormat(ws.Range("BL" & i + 11 & ":BO" & i + 12), "", False, i = 4, 13, 21, xlHAlignCenter, xlContinuous, xlDouble)                                                                          ' Winner  

        Call TableInfoFormat(ws.Range("C" & i + 12 & ":D" & i + 12), player_letter_locations(i + 1), True, i = 4, 12, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                      ' Bottom Player Letter
        Call TableInfoFormat(ws.Range("E" & i + 12 & ":I" & i + 12), group(Asc(player_letter_locations(i + 1)) - 64).LicenceNumber, False, i = 4, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)      ' Botttom Player Licence Number
        Call TableInfoFormat(ws.Range("J" & i + 12 & ":Y" & i + 12), group(Asc(player_letter_locations(i + 1)) - 64).Name, False, i = 4, 13, 21, xlHAlignLeft, xlContinuous, xlContinuous, xlDot)                 ' Bottom Player Name
        Call TableInfoFormat(ws.Range("Z" & i + 12 & ":AN" & i + 12), "", False, i = 4, 13, 21, xlHAlignLeft, xlContinuous, xlContinuous, xlDot)                                                                  ' Bottom Coach
        Call TableInfoFormat(ws.Range("AR" & i + 12 & ":AU" & i + 12), "", False, i = 4, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 1
        Call TableInfoFormat(ws.Range("AV" & i + 12 & ":AY" & i + 12), "", False, i = 4, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 2
        Call TableInfoFormat(ws.Range("AZ" & i + 12 & ":BC" & i + 12), "", False, i = 4, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 3
        Call TableInfoFormat(ws.Range("BD" & i + 12 & ":BG" & i + 12), "", False, i = 4, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 4
        Call TableInfoFormat(ws.Range("BH" & i + 12 & ":BK" & i + 12), "", False, i = 4, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 5
    Next i

    ' Table 5
    Call TableInfoFormat(ws.Range("AQ31:AV31"), "dc", False, True, 11, 21, xlHAlignCenter, xlDouble, xlContinuous, xlDouble)
    Call TableInfoFormat(ws.Range("AW31:BB31"), "sc", False, True, 11, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDouble)
    Call TableInfoFormat(ws.Range("BC31:BH31"), "wc", False, True, 11, 21, xlHAlignCenter, xlContinuous, xlDouble, xlDouble)

    ws.Range("A1:BO31").Interior.Color = RGB(255, 255, 255)
    With ws.PageSetup
        .PrintArea = "A1:BO31"
        .LeftMargin = 28.35
        .RightMargin = 28.35
    End With
End Sub


'================================================================================ Group of 4 ================================================================================


Sub Group4(ws As worksheet, tournament_name As String, event_name As String, group_number As Long, start_time As String, table As String, dates As String, group As Variant)
    Dim i As Integer ' General For loop index
    Dim row As Integer
    Dim player_letter_locations As Variant
    Dim umpire_letter_locations As Variant
    
    ' Column Widths
    ws.Range("A1:BO1").EntireColumn.ColumnWidth = 1

    ' Row heights
    ws.Rows("1:3").RowHeight = 18.5         ' Table 1
    ws.Rows(4).RowHeight = 29               ' Gap between Table 1 and 2/3
    ws.Rows(5).RowHeight = 26               ' Table 2/3 Header
    ws.Rows("6:9").RowHeight = 30           ' Table 2/3 Info
    ws.Rows(10).RowHeight = 29              ' Gap between Table 2/3 and 4
    ws.Rows(11).RowHeight = 16              ' Table 4 Header
    ws.Rows("12:23").RowHeight = 31         ' Table 4 Info
    ws.Rows(31).RowHeight = 22              ' Table 5

    ' Table 1
    Call Table1Format(ws.Range("A1:AX1"), "Tournament: " & tournament_name, 14)
    Call Table1Format(ws.Range("A2:AX2"), "Event: " & event_name, 14)
    Call Table1Format(ws.Range("A3:AX3"), "Group: " & group_number, 14)
    Call Table1Format(ws.Range("AY1:BO1"), "Time: " & start_time, 14)
    Call Table1Format(ws.Range("AY2:BO2"), "Table: " & table, 14)
    Call Table1Format(ws.Range("AY3:BO3"), "Date: " & dates, 14)
    

    ' Creates the "For Referee's use in case of a tie" text box
    With ws.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=433.3125, _
            Top:=70, _
            Width:=108.692, _
            Height:=10)

        .TextFrame.Characters.Text = "For Referee's use in case of a tie"
        .TextFrame2.TextRange.Font.Size = 7
        .TextFrame2.TextRange.Font.Name = "Calibri (Body)"
        .Line.Visible = msoFalse
    End With

    ' Table 2 headers
    Call TableHeader(ws.Range("B5:D5"), "", "left", 11, False)
    Call TableHeader(ws.Range("E5:J5"), "Licence No", "centre", 11, False)
    Call TableHeader(ws.Range("K5:AG5"), "Full Name", "centre", 11, True)
    Call TableHeader(ws.Range("AH5:AK5"), "County", "centre", 8, False)
    Call TableHeader(ws.Range("AL5:AO5"), "Points", "centre", 8, False)
    Call TableHeader(ws.Range("AP5:AS5"), "Position", "right", 8, False)

    ' Table 2 Info
    For i = 1 To 4
        row = 5 + i
        
        Call TableInfoFormat(ws.Range("B" & row & ":D" & row), Chr(64 + i), True, i = 4, 14, 21, xlHAlignCenter, xlDouble, xlContinuous)
        Call TableInfoFormat(ws.Range("E" & row & ":J" & row), group(i).LicenceNumber, True, i = 4, 14, 21)
        Call TableInfoFormat(ws.Range("K" & row & ":AG" & row), group(i).Name, True, i = 4, 14, 21)
        Call TableInfoFormat(ws.Range("AH" & row & ":AK" & row), group(i).Association, False, i = 4, 14, 21)
        Call TableInfoFormat(ws.Range("AL" & row & ":AO" & row), "", False, i = 4, 14, 21)
        Call TableInfoFormat(ws.Range("AP" & row & ":AS" & row), "", False, i = 4, 14, 21, xlHAlignCenter, xlContinuous, xlDouble)
    Next i

    ' Table 3 Headers
    Call TableHeader(ws.Range("AW5:BB5"), "Sets", "left", 11, False)
    Call TableHeader(ws.Range("BC5:BH5"), "Games", "centre", 11, False)
    Call TableHeader(ws.Range("BI5:BN5"), "Points", "right", 11, False)

    ' Table 3 Info
    For i = 1 To 4
        row = 5 + i
        Call TableInfoFormat(ws.Range("AW" & row & ":AY" & row), "", False, i = 4, 14, 21, xlHAlignCenter, xlDouble, xlDot)
        Call TableInfoFormat(ws.Range("AZ" & row & ":BB" & row), "", False, i = 4, 14, 21, xlHAlignCenter, xlDot, xlContinuous)
        Call TableInfoFormat(ws.Range("BC" & row & ":BE" & row), "", False, i = 4, 14, 21, xlHAlignCenter, xlContinuous, xlDot)
        Call TableInfoFormat(ws.Range("BF" & row & ":BH" & row), "", False, i = 4, 14, 21, xlHAlignCenter, xlDot, xlContinuous)
        Call TableInfoFormat(ws.Range("BI" & row & ":BK" & row), "", False, i = 4, 14, 21, xlHAlignCenter, xlContinuous, xlDot)
        Call TableInfoFormat(ws.Range("BL" & row & ":BN" & row), "", False, i = 4, 14, 21, xlHAlignCenter, xlDot, xlDouble)
    Next i

    ' Table 4 Headers
    Call TableHeader(ws.Range("A11:D11"), "", "left", 11, False)
    Call TableHeader(ws.Range("E11:I11"), "Licence No", "centre", 9, False)
    Call TableHeader(ws.Range("J11:Y11"), "Player", "centre", 12, True)
    Call TableHeader(ws.Range("Z11:AN11"), "Coach", "centre", 12, True)
    Call TableHeader(ws.Range("AO11:AQ11"), "Ump", "centre", 8, False)
    Call TableHeader(ws.Range("AR11:AU11"), "Game 1", "centre", 8, False)
    Call TableHeader(ws.Range("AV11:AY11"), "Game 2", "centre", 8, False)
    Call TableHeader(ws.Range("AZ11:BC11"), "Game 3", "centre", 8, False)
    Call TableHeader(ws.Range("BD11:BG11"), "Game 4", "centre", 8, False)
    Call TableHeader(ws.Range("BH11:BK11"), "Game 5", "centre", 8, False)
    Call TableHeader(ws.Range("BL11:BO11"), "Winner", "right", 8, False)

    ' Table 4 Match Numbers
    row = 1
    For i = 12 To 22 Step 2
        Call TableInfoFormat(ws.Range("A" & i & ":B" & i + 1), row, False, i = 22, 12, 21, xlHAlignCenter, xlDouble, xlContinuous)
        row = row + 1
    Next i

    ' Create player orders
    player_letter_locations = Array("B", "C", "A", "D", "B", "D", "A", "C", "D", "C", "A", "B")
    umpire_letter_locations = Array("A", "B", "C", "D", "A", "C")

    ' Table 4 Info
    For i = 0 To 11 Step 2
        Call TableInfoFormat(ws.Range("C" & i + 12 & ":D" & i + 12), player_letter_locations(i), True, False, 12, 21)                                                                                              ' Top Player Letter
        Call TableInfoFormat(ws.Range("E" & i + 12 & ":I" & i + 12), group(Asc(player_letter_locations(i)) - 64).LicenceNumber, False, False, 13, 21)                                                              ' Top Player Licence Number
        Call TableInfoFormat(ws.Range("J" & i + 12 & ":Y" & i + 12), group(Asc(player_letter_locations(i)) - 64).Name, False, False, 13, 21, xlHAlignLeft)                                                         ' Top Player Name
        Call TableInfoFormat(ws.Range("Z" & i + 12 & ":AN" & i + 12), "", False, False, 13, 21, xlHAlignLeft)                                                                                                      ' Top Coach
        Call TableInfoFormat(ws.Range("AR" & i + 12 & ":AU" & i + 12), "", False, False, 13, 21)                                                                                                                   ' Top Game 1
        Call TableInfoFormat(ws.Range("AV" & i + 12 & ":AY" & i + 12), "", False, False, 13, 21)                                                                                                                   ' Top Game 2
        Call TableInfoFormat(ws.Range("AZ" & i + 12 & ":BC" & i + 12), "", False, False, 13, 21)                                                                                                                   ' Top Game 3
        Call TableInfoFormat(ws.Range("BD" & i + 12 & ":BG" & i + 12), "", False, False, 13, 21)                                                                                                                   ' Top Game 4
        Call TableInfoFormat(ws.Range("BH" & i + 12 & ":BK" & i + 12), "", False, False, 13, 21)                                                                                                                   ' Top Game 5
        
        Call TableInfoFormat(ws.Range("AO" & i + 12 & ":AQ" & i + 13), umpire_letter_locations(i / 2), False, i = 10, 13, 21)                                                                                      ' Umpire Letter
        Call TableInfoFormat(ws.Range("BL" & i + 12 & ":BO" & i + 13), "", False, i = 10, 13, 21, xlHAlignCenter, xlContinuous, xlDouble)                                                                          ' Winner  

        Call TableInfoFormat(ws.Range("C" & i + 13 & ":D" & i + 13), player_letter_locations(i + 1), True, i = 10, 12, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                      ' Bottom Player Letter
        Call TableInfoFormat(ws.Range("E" & i + 13 & ":I" & i + 13), group(Asc(player_letter_locations(i + 1)) - 64).LicenceNumber, False, i = 10, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)      ' Botttom Player Licence Number
        Call TableInfoFormat(ws.Range("J" & i + 13 & ":Y" & i + 13), group(Asc(player_letter_locations(i + 1)) - 64).Name, False, i = 10, 13, 21, xlHAlignLeft, xlContinuous, xlContinuous, xlDot)                 ' Bottom Player Name
        Call TableInfoFormat(ws.Range("Z" & i + 13 & ":AN" & i + 13), "", False, i = 10, 13, 21, xlHAlignLeft, xlContinuous, xlContinuous, xlDot)                                                                  ' Bottom Coach
        Call TableInfoFormat(ws.Range("AR" & i + 13 & ":AU" & i + 13), "", False, i = 10, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 1
        Call TableInfoFormat(ws.Range("AV" & i + 13 & ":AY" & i + 13), "", False, i = 10, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 2
        Call TableInfoFormat(ws.Range("AZ" & i + 13 & ":BC" & i + 13), "", False, i = 10, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 3
        Call TableInfoFormat(ws.Range("BD" & i + 13 & ":BG" & i + 13), "", False, i = 10, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 4
        Call TableInfoFormat(ws.Range("BH" & i + 13 & ":BK" & i + 13), "", False, i = 10, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 5
    Next i

    ' Table 5
    Call TableInfoFormat(ws.Range("AQ31:AV31"), "dc", False, True, 11, 21, xlHAlignCenter, xlDouble, xlContinuous, xlDouble)
    Call TableInfoFormat(ws.Range("AW31:BB31"), "sc", False, True, 11, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDouble)
    Call TableInfoFormat(ws.Range("BC31:BH31"), "wc", False, True, 11, 21, xlHAlignCenter, xlContinuous, xlDouble, xlDouble)

    ws.Range("A1:BO31").Interior.Color = RGB(255, 255, 255)
    With ws.PageSetup
        .PrintArea = "A1:BO31"
        .LeftMargin = 28.35
        .RightMargin = 28.35
    End With
End Sub


'================================================================================ Group of 5 ================================================================================


Sub Group5(ws As worksheet, tournament_name As String, event_name As String, group_number As Long, start_time As String, table As String, dates As String, group As Variant)
    Dim i As Integer ' General For loop index
    Dim row As Integer
    Dim player_letter_locations As Variant
    Dim umpire_letter_locations As Variant
    
    ' Column Widths
    ws.Range("A1:BO1").EntireColumn.ColumnWidth = 1

    ' Row heights
    ws.Rows("1:3").RowHeight = 18.5         ' Table 1
    ws.Rows(4).RowHeight = 29               ' Gap between Table 1 and 2/3
    ws.Rows(5).RowHeight = 24.8             ' Table 2/3 Header
    ws.Rows("6:10").RowHeight = 28          ' Table 2/3 Info
    ws.Rows(11).RowHeight = 29              ' Gap between Table 2/3 and 4
    ws.Rows(12).RowHeight = 16              ' Table 4 Header
    ws.Rows("13:32").RowHeight = 30         ' Table 4 Info
    ws.Rows(34).RowHeight = 22              ' Table 5

    ' Table 1
    Call Table1Format(ws.Range("A1:AX1"), "Tournament: " & tournament_name, 14)
    Call Table1Format(ws.Range("A2:AX2"), "Event: " & event_name, 14)
    Call Table1Format(ws.Range("A3:AX3"), "Group: " & group_number, 14)
    Call Table1Format(ws.Range("AY1:BO1"), "Time: " & start_time, 14)
    Call Table1Format(ws.Range("AY2:BO2"), "Table: " & table, 14)
    Call Table1Format(ws.Range("AY3:BO3"), "Date: " & dates, 14)
    

    ' Creates the "For Referee's use in case of a tie" text box
    With ws.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=433.3125, _
            Top:=70, _
            Width:=108.692, _
            Height:=10)

        .TextFrame.Characters.Text = "For Referee's use in case of a tie"
        .TextFrame2.TextRange.Font.Size = 7
        .TextFrame2.TextRange.Font.Name = "Calibri (Body)"
        .Line.Visible = msoFalse
    End With

    ' Table 2 headers
    Call TableHeader(ws.Range("B5:D5"), "", "left", 11, False)
    Call TableHeader(ws.Range("E5:J5"), "Licence No", "centre", 11, False)
    Call TableHeader(ws.Range("K5:AG5"), "Full Name", "centre", 11, True)
    Call TableHeader(ws.Range("AH5:AK5"), "County", "centre", 8, False)
    Call TableHeader(ws.Range("AL5:AO5"), "Points", "centre", 8, False)
    Call TableHeader(ws.Range("AP5:AS5"), "Position", "right", 8, False)

    ' Table 2 Info
    For i = 1 To 5
        row = 5 + i
        
        Call TableInfoFormat(ws.Range("B" & row & ":D" & row), Chr(64 + i), True, i = 5, 14, 21, xlHAlignCenter, xlDouble, xlContinuous)
        Call TableInfoFormat(ws.Range("E" & row & ":J" & row), group(i).LicenceNumber, True, i = 5, 14, 21)
        Call TableInfoFormat(ws.Range("K" & row & ":AG" & row), group(i).Name, True, i = 5, 14, 21)
        Call TableInfoFormat(ws.Range("AH" & row & ":AK" & row), group(i).Association, False, i = 5, 14, 21)
        Call TableInfoFormat(ws.Range("AL" & row & ":AO" & row), "", False, i = 5, 14, 21)
        Call TableInfoFormat(ws.Range("AP" & row & ":AS" & row), "", False, i = 5, 14, 21, xlHAlignCenter, xlContinuous, xlDouble)
    Next i

    ' Table 3 Headers
    Call TableHeader(ws.Range("AW5:BB5"), "Sets", "left", 11, False)
    Call TableHeader(ws.Range("BC5:BH5"), "Games", "centre", 11, False)
    Call TableHeader(ws.Range("BI5:BN5"), "Points", "right", 11, False)

    ' Table 3 Info
    For i = 1 To 5
        row = 5 + i
        Call TableInfoFormat(ws.Range("AW" & row & ":AY" & row), "", False, i = 5, 14, 21, xlHAlignCenter, xlDouble, xlDot)
        Call TableInfoFormat(ws.Range("AZ" & row & ":BB" & row), "", False, i = 5, 14, 21, xlHAlignCenter, xlDot, xlContinuous)
        Call TableInfoFormat(ws.Range("BC" & row & ":BE" & row), "", False, i = 5, 14, 21, xlHAlignCenter, xlContinuous, xlDot)
        Call TableInfoFormat(ws.Range("BF" & row & ":BH" & row), "", False, i = 5, 14, 21, xlHAlignCenter, xlDot, xlContinuous)
        Call TableInfoFormat(ws.Range("BI" & row & ":BK" & row), "", False, i = 5, 14, 21, xlHAlignCenter, xlContinuous, xlDot)
        Call TableInfoFormat(ws.Range("BL" & row & ":BN" & row), "", False, i = 5, 14, 21, xlHAlignCenter, xlDot, xlDouble)
    Next i

    ' Table 4 Headers
    Call TableHeader(ws.Range("A12:D12"), "", "left", 11, False)
    Call TableHeader(ws.Range("E12:I12"), "Licence No", "centre", 9, False)
    Call TableHeader(ws.Range("J12:Y12"), "Player", "centre", 12, True)
    Call TableHeader(ws.Range("Z12:AN12"), "Coach", "centre", 12, True)
    Call TableHeader(ws.Range("AO12:AQ12"), "Ump", "centre", 8, False)
    Call TableHeader(ws.Range("AR12:AU12"), "Game 1", "centre", 8, False)
    Call TableHeader(ws.Range("AV12:AY12"), "Game 2", "centre", 8, False)
    Call TableHeader(ws.Range("AZ12:BC12"), "Game 3", "centre", 8, False)
    Call TableHeader(ws.Range("BD12:BG12"), "Game 4", "centre", 8, False)
    Call TableHeader(ws.Range("BH12:BK12"), "Game 5", "centre", 8, False)
    Call TableHeader(ws.Range("BL12:BO12"), "Winner", "right", 8, False)

    ' Table 4 Match Numbers
    row = 1
    For i = 13 To 31 Step 2
        Call TableInfoFormat(ws.Range("A" & i & ":B" & i + 1), row, False, i = 31, 12, 21, xlHAlignCenter, xlDouble, xlContinuous)
        row = row + 1
    Next i

    ' Create player orders
    player_letter_locations = Array("B", "E", "C", "D", "A", "E", "B", "C", "A", "D", "E", "C", "A", "C", "B", "D", "A", "B", "D", "E")
    umpire_letter_locations = Array("A", "E", "C", "D", "B", "A", "E", "C", "D", "B")

        ' Table 4 Info
    For i = 0 To 19 Step 2
        Call TableInfoFormat(ws.Range("C" & i + 13 & ":D" & i + 13), player_letter_locations(i), True, False, 12, 21)                                                                                              ' Top Player Letter
        Call TableInfoFormat(ws.Range("E" & i + 13 & ":I" & i + 13), group(Asc(player_letter_locations(i)) - 64).LicenceNumber, False, False, 13, 21)                                                              ' Top Player Licence Number
        Call TableInfoFormat(ws.Range("J" & i + 13 & ":Y" & i + 13), group(Asc(player_letter_locations(i)) - 64).Name, False, False, 13, 21, xlHAlignLeft)                                                         ' Top Player Name
        Call TableInfoFormat(ws.Range("Z" & i + 13 & ":AN" & i + 13), "", False, False, 13, 21, xlHAlignLeft)                                                                                                      ' Top Coach
        Call TableInfoFormat(ws.Range("AR" & i + 13 & ":AU" & i + 13), "", False, False, 13, 21)                                                                                                                   ' Top Game 1
        Call TableInfoFormat(ws.Range("AV" & i + 13 & ":AY" & i + 13), "", False, False, 13, 21)                                                                                                                   ' Top Game 2
        Call TableInfoFormat(ws.Range("AZ" & i + 13 & ":BC" & i + 13), "", False, False, 13, 21)                                                                                                                   ' Top Game 3
        Call TableInfoFormat(ws.Range("BD" & i + 13 & ":BG" & i + 13), "", False, False, 13, 21)                                                                                                                   ' Top Game 4
        Call TableInfoFormat(ws.Range("BH" & i + 13 & ":BK" & i + 13), "", False, False, 13, 21)                                                                                                                   ' Top Game 5
        
        Call TableInfoFormat(ws.Range("AO" & i + 13 & ":AQ" & i + 14), umpire_letter_locations(i / 2), False, i = 18, 13, 21)                                                                                      ' Umpire Letter
        Call TableInfoFormat(ws.Range("BL" & i + 13 & ":BO" & i + 14), "", False, i = 18, 13, 21, xlHAlignCenter, xlContinuous, xlDouble)                                                                          ' Winner  

        Call TableInfoFormat(ws.Range("C" & i + 14 & ":D" & i + 14), player_letter_locations(i + 1), True, i = 18, 12, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                      ' Bottom Player Letter
        Call TableInfoFormat(ws.Range("E" & i + 14 & ":I" & i + 14), group(Asc(player_letter_locations(i + 1)) - 64).LicenceNumber, False, i = 18, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)      ' Botttom Player Licence Number
        Call TableInfoFormat(ws.Range("J" & i + 14 & ":Y" & i + 14), group(Asc(player_letter_locations(i + 1)) - 64).Name, False, i = 18, 13, 21, xlHAlignLeft, xlContinuous, xlContinuous, xlDot)                 ' Bottom Player Name
        Call TableInfoFormat(ws.Range("Z" & i + 14 & ":AN" & i + 14), "", False, i = 18, 13, 21, xlHAlignLeft, xlContinuous, xlContinuous, xlDot)                                                                  ' Bottom Coach
        Call TableInfoFormat(ws.Range("AR" & i + 14 & ":AU" & i + 14), "", False, i = 18, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 1
        Call TableInfoFormat(ws.Range("AV" & i + 14 & ":AY" & i + 14), "", False, i = 18, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 2
        Call TableInfoFormat(ws.Range("AZ" & i + 14 & ":BC" & i + 14), "", False, i = 18, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 3
        Call TableInfoFormat(ws.Range("BD" & i + 14 & ":BG" & i + 14), "", False, i = 18, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 4
        Call TableInfoFormat(ws.Range("BH" & i + 14 & ":BK" & i + 14), "", False, i = 18, 13, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDot)                                                               ' Bottom Game 5
    Next i

    ' Table 5
    Call TableInfoFormat(ws.Range("AQ34:AV34"), "dc", False, True, 11, 21, xlHAlignCenter, xlDouble, xlContinuous, xlDouble)
    Call TableInfoFormat(ws.Range("AW34:BB34"), "sc", False, True, 11, 21, xlHAlignCenter, xlContinuous, xlContinuous, xlDouble)
    Call TableInfoFormat(ws.Range("BC34:BH34"), "wc", False, True, 11, 21, xlHAlignCenter, xlContinuous, xlDouble, xlDouble)

    ws.Range("A1:BO34").Interior.Color = RGB(255, 255, 255)
    With ws.PageSetup
        .PrintArea = "A1:BO34"
        .LeftMargin = 28.35
        .RightMargin = 28.35
    End With
End Sub


'================================================================================ Group of 5 ================================================================================


Sub Group6(ws As worksheet, tournament_name As String, event_name As String, group_number As Long, start_time As String, table As String, dates As String, group As Variant)
End Sub




' General Info Section
Private Sub Table1Format(rng As Range, text As String, font_size As Integer)
    Dim b As XlBordersIndex

    With rng
        .Merge
        .Font.Name = FONT
        .Font.Size = font_size
        .Value = text

        For b = xlEdgeLeft To xlInsideHorizontal
            With .Borders(b)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        Next b
    End With
End Sub


' Creates all the headers
Private Sub TableHeader(rng As Range, text As String, location As String, font_size As Integer, bold As Boolean)
    With rng
        .Merge
        .Font.Name = FONT
        .Font.Size = font_size
        .Font.Bold = bold
        .Value = text
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        ' Top border (always double)
        With .Borders(xlEdgeTop)
            .LineStyle = xlDouble
            .Color = RGB(0, 0, 0)
        End With

        ' Bottom border (always continuous)
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With

        ' Right border
        With .Borders(xlEdgeRight)
            .LineStyle = IIf(LCase(location) = "right", xlDouble, xlContinuous)
            If .LineStyle = xlContinuous Then .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With

        ' Left border
        With .Borders(xlEdgeLeft)
            .LineStyle = IIf(LCase(location) = "left", xlDouble, xlContinuous)
            If .LineStyle = xlContinuous Then .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End With
End Sub


Private Sub TableInfoFormat(rng As Range, cell_value As Variant, is_bold As Boolean, _
                        is_last_row As Boolean, font_size As Integer, max_length As Integer, _
                        Optional horizontal_alignment As XlHAlign = xlHAlignCenter, _
                        Optional left_border_style As xlLineStyle = xlContinuous, _
                        Optional right_border_style As xlLineStyle = xlContinuous, _
                        Optional top_border_style As xlLineStyle = xlContinuous)

    If Len(cell_value) > max_length Then fontSize = font_size * max_length / Len(cell_value)
    
    With rng
        .Merge
        .Font.Name = FONT
        .Font.Size = font_size
        .Font.Bold = is_bold
        .Value = cell_value
        .HorizontalAlignment = horizontal_alignment
        .VerticalAlignment = xlCenter
        
        ' Top Border
        With .Borders(xlEdgeTop)
            .LineStyle = top_border_style
            .Color = RGB(0, 0, 0)
        End With

        ' Bottom Border
        With .Borders(xlEdgeBottom)
            If is_last_row Then
                .LineStyle = xlDouble
                .Color = RGB(0, 0, 0)
            Else
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End If
        End With

        ' Left Border
        With .Borders(xlEdgeLeft)
            .LineStyle = left_border_style
            .Color = RGB(0, 0, 0)
        End With

        ' Right Border
        With .Borders(xlEdgeRight)
            .LineStyle = right_border_style
            .Color = RGB(0, 0, 0)
        End With
    End With
End Sub