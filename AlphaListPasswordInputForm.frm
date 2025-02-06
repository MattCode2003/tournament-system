VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AlphaListPasswordInputForm 
   Caption         =   "Alpha List Password"
   ClientHeight    =   3012
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4716
   OleObjectBlob   =   "AlphaListPasswordInputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AlphaListPasswordInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancelButton_Click()
    cancelPressed = True
    Me.Hide
End Sub

Private Sub Label1_Click()

End Sub

Private Sub okButton_Click()
    ' Enables Error Handling
    On Error Resume Next

    Application.ScreenUpdating = False
    Set alphaList = Workbooks.Open(Filename:=AlphaListFilePath, Password:=Me.inputBox.Text, ReadOnly:=True)
    
    ' Incorrect Password
    If Err.number <> 0 Then
        Me.errorMessage.Visible = True
    Else
        ' Correct Password
        alphaList.Windows(1).Visible = False
        Me.Hide
    End If

    Application.ScreenUpdating = True

End Sub

Private Sub UserForm_Click()
    Exit Sub
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    cancelPressed = True
    Me.Hide
End Sub
