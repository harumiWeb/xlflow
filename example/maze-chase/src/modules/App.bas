Attribute VB_Name = "App"
Option Explicit

Public Sub RunCore(ByVal wb As Workbook)
    If wb Is Nothing Then
        Exit Sub
    End If

    If XlflowRuntime.IsHeadless() Then
        XlflowDebug.Log "Maze Chase requires interactive Excel. Use xlflow test for automated verification."
        Exit Sub
    End If

    Load MazeChaseForm
    MazeChaseForm.Show vbModeless
End Sub
