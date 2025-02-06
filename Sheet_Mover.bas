Attribute VB_Name = "Sheet_Mover"
Sub GoToGeneralSettings()
    Dim sheetName As String
    ' Specify the sheet name here
    sheetName = "General Settings"
    
    ' Check if the sheet exists
    On Error Resume Next
    If Not Worksheets(sheetName) Is Nothing Then
        Worksheets(sheetName).Activate
    Else
        MsgBox "Sheet " & sheetName & " does not exist.", vbExclamation
    End If
    On Error GoTo 0
End Sub

Sub GoToEventSettings()
    Dim sheetName As String
    ' Specify the sheet name here
    sheetName = "Event Settings"
    
    ' Check if the sheet exists
    On Error Resume Next
    If Not Worksheets(sheetName) Is Nothing Then
        Worksheets(sheetName).Activate
    Else
        MsgBox "Sheet " & sheetName & " does not exist.", vbExclamation
    End If
    On Error GoTo 0
End Sub

Sub GoToMainMenu()
    Dim sheetName As String
    ' Specify the sheet name here
    sheetName = "Main Menu"
    
    ' Check if the sheet exists
    On Error Resume Next
    If Not Worksheets(sheetName) Is Nothing Then
        Worksheets(sheetName).Activate
    Else
        MsgBox "Sheet " & sheetName & " does not exist.", vbExclamation
    End If
    On Error GoTo 0
End Sub

