Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim value As Long
    Select Case value
    Case 1
        Debug.Print "one"
    Case Else
        Debug.Print "default"
    End Select
End Sub
