Attribute VB_Name = "Main"
Option Explicit
Public Const ScalarValue As Long = 1
Public Sub ProbeIssue590NonCallable()
    Debug.Print ScalarValue()
End Sub
