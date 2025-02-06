VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GroupCreationForm 
   Caption         =   "Group Creation"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "GroupCreationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GroupCreationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim index As Integer

Private Sub UserForm_Initialize()

   ' Back button setup
   With Me.Back1
      ' .Font.Name = "Wingdings"
      .Caption = ChrW(8592)
      .Font.Size = 16
   End With

   ' Forward button setup
   With Me.Forward1
      ' .Font.Name = "Wingdings"
      .Caption = ChrW(8594)
      .Font.Size = 16
   End With

   ' Sheet index Setup
   index = 2
   dataWorkbook.Sheets(2).Activate
   Me.EventLabel.Caption = dataWorkbook.Sheets(2).Name

End Sub

Private Sub Forward1_Click()
   ' Works out the sheet index of the next event
   If index < dataWorkbook.Sheets.count Then
      index = index + 1
   Else
      index = 2
   End If

   ' Displays the next event
   dataWorkbook.Sheets(index).Activate
   Me.EventLabel.Caption = dataWorkbook.Sheets(index).Name
End Sub

Private Sub Back1_Click()
   ' Works out the sheet index of the next event
   If index > 2 Then
      index = index - 1
   Else
      index = dataWorkbook.sheets.count
   End If

   ' Displays the next event
   dataWorkbook.Sheets(index).Activate
   Me.EventLabel.Caption = dataWorkbook.Sheets(index).Name
End Sub

Private Sub ExitButton_Click()
   Me.Hide
   End
End Sub

Private Sub OkButton_Click()
   currentEvent = dataWorkbook.sheets(index).Name
   Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    End
End Sub