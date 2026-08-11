Attribute VB_Name = "ThisWorkbook"
Option Explicit
Private Sub Workbook_Open()
    Debug.Print TypeName(Me)
End Sub
