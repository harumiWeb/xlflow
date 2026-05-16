Attribute VB_Name = "App"
Option Explicit

Public Sub RunCore(ByVal wb As Workbook)
  If wb Is Nothing Then Exit Sub
  StartInvaderGame
End Sub
