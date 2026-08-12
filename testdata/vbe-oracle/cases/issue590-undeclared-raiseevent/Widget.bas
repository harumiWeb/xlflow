Attribute VB_Name = "Widget"
Option Explicit
Public Event Changed(ByVal value As Long)
Public Sub ProbeIssue590UndeclaredRaiseEvent()
    RaiseEvent MissingEvent(1)
End Sub
