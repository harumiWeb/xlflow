Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeMissingRequiredArgument()
    Dim result As Long
    result = AddTwo(1)
End Sub

Private Function AddTwo(ByVal first As Long, ByVal second As Long) As Long
    AddTwo = first + second
End Function
