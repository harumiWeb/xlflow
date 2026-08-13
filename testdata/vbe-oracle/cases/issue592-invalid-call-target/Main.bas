Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeIssue592InvalidCallTarget()
    Call (1 + 2)
End Sub
