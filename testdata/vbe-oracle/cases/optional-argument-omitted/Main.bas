Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeOptionalArgumentOmitted()
    Dim result As String
    result = JoinText("left")
End Sub

Private Function JoinText(ByVal first As String, Optional ByVal second As String = "right") As String
    JoinText = first & second
End Function
