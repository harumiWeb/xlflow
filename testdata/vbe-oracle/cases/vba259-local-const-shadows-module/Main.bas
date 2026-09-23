Attribute VB_Name = "Main"
Option Explicit

Private Const Limit As Long = 10

Public Sub Main()
    Const Limit As Long = 5
    Dim value As Long
    Select Case value
    Case Limit
        Debug.Print "limit"
    Case 5
        Debug.Print "five"
    End Select
End Sub
