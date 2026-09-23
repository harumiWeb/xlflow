Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim value As Long
    Select Case value
    Case 1
        Debug.Print "one"
    Case 1
        Debug.Print "unreachable duplicate"
    Case 2 To 9
        Debug.Print "range"
    Case 5
        Debug.Print "unreachable covered"
    End Select
End Sub
