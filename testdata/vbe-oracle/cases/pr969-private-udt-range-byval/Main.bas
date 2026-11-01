Attribute VB_Name = "Main"
Option Explicit
Private Sub Observe(ByVal value As Range, ByVal qualified As Excel.Range)
    Debug.Print value.Row, qualified.Row
End Sub
