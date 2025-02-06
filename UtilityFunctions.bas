Attribute VB_Name = "UtilitiyFunctions"

'========================== Extract Number ============================

Public Function ExtractNumberFromString(inputString As String) As Integer
    Dim i As Long
    Dim numStart As Long
    Dim numEnd As Long
    Dim tempNum As String

    ' Initialize variables
    tempNum = ""
    numStart = 0
    numEnd = 0

    ' Loop through each character in the string
    For i = 1 To Len(inputString)
        If IsNumeric(Mid(inputString, i, 1)) Then
            If numStart = 0 Then numStart = i ' Mark the start of the number
            tempNum = tempNum & Mid(inputString, i, 1) ' Build the number string
        Else
            If numStart > 0 Then Exit For ' Stop once the number ends
        End If
    Next i

    ' Return the extracted number if found
    If tempNum <> "" Then
        ExtractNumberFromString = CInt(tempNum)
    Else
        ExtractNumberFromString = CInt(0) ' Return an error if no number is found
    End If
End Function


'================================ Get Delimiter ================================


' Get file delimiter based on OneDrive or local
Public Function GetDelimiter(path As String) As String
    GetDelimiter = IIf(Left(path, 5) = "https", "/", "\")
End Function


'============================== File Exists ====================================


' Check if file exists
Public Function FileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbCritical
        FileExists = False
    Else
        FileExists = True
    End If
End Function