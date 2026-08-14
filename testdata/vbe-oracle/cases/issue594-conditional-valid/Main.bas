Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim ready As Boolean
    Dim fallback As Boolean
    If ready Then
        Debug.Print "ready"
    ElseIf fallback Then
        Debug.Print "fallback"
    Else
        Debug.Print "default"
    End If
End Sub
