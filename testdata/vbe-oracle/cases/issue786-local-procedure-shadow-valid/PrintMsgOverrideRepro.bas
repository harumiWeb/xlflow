Attribute VB_Name = "PrintMsgOverrideRepro"
Option Explicit

Private Sub printMsg(ByVal level As Long, ByVal text As String, ByVal fromProcedure As String, Optional ByVal isHeader As Boolean = False)
End Sub

Public Sub CallIt()
    printMsg 1, "hello", "CallIt"
End Sub
