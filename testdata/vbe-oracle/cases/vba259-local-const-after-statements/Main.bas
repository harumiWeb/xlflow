Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim value As Long
    Select Case value
    Case Limit
        Debug.Print "limit"
    Case Else
        Debug.Print "other"
    End Select
    Const Limit As Long = 5
End Sub
