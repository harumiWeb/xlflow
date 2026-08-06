Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeByRefParenthesizedVariable()
    Dim value As Long
    value = 1
    MutateLong (value)
End Sub

Private Sub MutateLong(ByRef value As Long)
    value = value + 1
End Sub
