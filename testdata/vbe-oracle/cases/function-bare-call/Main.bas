Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeFunctionBareCall()
    Dim result As Long
    result = AddOne(1)
End Sub

Private Function AddOne(ByVal value As Long) As Long
    AddOne = value + 1
End Function
