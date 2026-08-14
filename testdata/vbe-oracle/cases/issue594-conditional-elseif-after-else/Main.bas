Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim ready As Boolean
    If ready Then
        Debug.Print "ready"
    Else
        Debug.Print "fallback"
    ElseIf ready Then
        Debug.Print "too late"
    End If
End Sub
