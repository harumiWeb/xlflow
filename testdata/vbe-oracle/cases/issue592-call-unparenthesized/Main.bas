Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeIssue592CallUnparenthesized()
    Call WritePair 1, 2
End Sub

Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
