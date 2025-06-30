VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GroupSheetsForm 
   Caption         =   "Group Menu"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "GroupSheetsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GroupSheetsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub DrawSheet_Click()
    Call Draw_Sheet.CreateDrawSheet()
End Sub

Private Sub SingleGroupSheets_Click()
    Call Group_Sheets.EventGroupSheet("SINGLE")
End Sub

Private Sub MutipleGroupSheets_Click()
    Call Group_Sheets.EventGroupSheet("ALL")
End Sub

Private Sub ExitButton_Click()
    End
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    End
End Sub
