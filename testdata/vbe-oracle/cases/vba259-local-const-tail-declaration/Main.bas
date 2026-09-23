Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim value As Long
    Select Case value
    Case 1
        Debug.Print "one"
    Case Else
        Debug.Print "other"
    End Select
    Const Tail As Long = 5
    Debug.Print Tail
End Sub
