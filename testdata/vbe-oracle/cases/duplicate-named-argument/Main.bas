Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeDuplicateNamedArgument()
    Dim result As Long
    result = AddTwo(first:=1, first:=2)
End Sub

Private Function AddTwo(ByVal first As Long, ByVal second As Long) As Long
    AddTwo = first + second
End Function
