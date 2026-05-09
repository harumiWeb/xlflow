Attribute VB_Name = "ThisWorkbook"
Option Explicit

Private Sub Workbook_Open()
  Main.Run
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  App.HandleWorkbookBeforeClose
End Sub
