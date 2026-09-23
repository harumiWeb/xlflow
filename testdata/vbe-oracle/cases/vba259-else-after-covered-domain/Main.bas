Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim value As Byte
    Select Case value
    Case 0 To 127
        Debug.Print "low"
    Case 128 To 255
        Debug.Print "high"
    Case Else
        Debug.Print "unreachable else"
    End Select
End Sub
