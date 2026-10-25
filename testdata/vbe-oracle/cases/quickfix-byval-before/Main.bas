Attribute VB_Name = "Main"
Option Explicit

Private Function Increment(ByRef value As Long) As Long ' preserve this comment
    Increment = value + 1
End Function

Private Function Echo(ByRef value As String) As String
    Echo = value
End Function

Private Function Flip(ByRef value As Boolean) As Boolean
    Flip = Not value
End Function

Private Function SameDate(ByRef value As Date) As Date
    SameDate = value
End Function
