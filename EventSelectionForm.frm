VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EventSelectionForm 
   Caption         =   "Event Selection"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "EventSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EventSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private selected_event As String

Public Property Get selected_event_Value() As String
   selected_event_Value = selected_event
End Property

Private Sub UserForm_Initialize
   Dim item As Variant
   For Each item In events
      Me.ComboBox.AddItem item
   Next item
End Sub


' NEED TO ADD SANITY CHECKS HERE
Private Sub SubmitButton_Click()
   selected_event = Me.ComboBox.Value
   Me.Hide
End Sub
