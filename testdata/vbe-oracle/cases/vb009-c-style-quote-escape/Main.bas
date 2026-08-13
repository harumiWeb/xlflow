Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim text As String
    Dim index As Long
    If Mid$(text, index, 1) <> "\"" Then
        Debug.Print text
    End If
End Sub
