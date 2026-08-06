Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeByRefBareVariable()
    Dim value As Long
    value = 1
    MutateLong value
End Sub

Private Sub MutateLong(ByRef value As Long)
    value = value + 1
End Sub
