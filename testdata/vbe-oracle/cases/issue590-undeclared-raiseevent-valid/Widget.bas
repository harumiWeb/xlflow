Attribute VB_Name = "Widget"
Option Explicit
Public Event Changed(ByVal value As Long)
Public Sub ProbeIssue590UndeclaredRaiseEventValid()
    RaiseEvent Changed(1)
End Sub
