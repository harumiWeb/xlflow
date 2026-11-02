Attribute VB_Name = "Main"
Option Explicit
#Const flags = 3
#If (flags And 1) = 1 Then
Public Function Answer() As Long
    Answer = 42
End Function
#Else

This is plain prose, not VBA code.
## Heading
1. item
#End If
