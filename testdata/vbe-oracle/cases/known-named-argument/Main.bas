Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeKnownNamedArgument()
    Dim result As Long
    result = AddTwo(second:=2, first:=1)
End Sub

Private Function AddTwo(ByVal first As Long, ByVal second As Long) As Long
    AddTwo = first + second
End Function
