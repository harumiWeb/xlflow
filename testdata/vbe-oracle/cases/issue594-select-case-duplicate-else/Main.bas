Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim value As Long
    Select Case value
    Case 1
        Debug.Print "one"
    Case Else
        Debug.Print "first default"
    Case Else
        Debug.Print "second default"
    End Select
End Sub
