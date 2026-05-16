Attribute VB_Name = "App"
Option Explicit

Private Const LAUNCH_BUTTON_NAME As String = "CalendarPickerLaunchButton"

Public Sub RunCore(ByVal wb As Workbook)
  Dim targetCell As Range
  Dim picker As CalendarPicker
  Dim wasConfirmed As Boolean
  Dim SelectedDate As Date

  Set targetCell = ResolveTargetCell(wb)

  Set picker = New CalendarPicker
  SelectedDate = picker.PickDate(wasConfirmed)

  If wasConfirmed Then
    WriteDateToCell targetCell, SelectedDate
  End If

  Unload picker
  Set picker = Nothing
End Sub

Private Function ResolveTargetCell(ByVal wb As Workbook) As Range
  Dim currentSelection As Object
  Dim selectedRange As Range

  Set currentSelection = wb.Application.Selection

  If TypeName(currentSelection) <> "Range" Then
    Err.Raise vbObjectError + 1000, "App.ResolveTargetCell", "単一セルを選択してから実行してください。"
  End If

  Set selectedRange = currentSelection

  If selectedRange.Cells.CountLarge <> 1 Then
    Err.Raise vbObjectError + 1001, "App.ResolveTargetCell", "複数セルではなく単一セルを選択してから実行してください。"
  End If

  If Not selectedRange.Parent.Parent Is wb Then
    Err.Raise vbObjectError + 1002, "App.ResolveTargetCell", "このブック上のセルを選択してから実行してください。"
  End If

  Set ResolveTargetCell = selectedRange.Cells(1, 1)
End Function

Private Sub WriteDateToCell(ByVal targetCell As Range, ByVal SelectedDate As Date)
  targetCell.value = SelectedDate
  targetCell.NumberFormat = "yyyy/mm/dd"
End Sub

Public Sub InstallLaunchButton(ByVal wb As Workbook)
  Dim ws As Worksheet
  Dim targetArea As Range
  Dim launchShape As Shape

  Set ws = wb.Worksheets(1)
  Set targetArea = ws.Range("B2:E4")

  Set launchShape = FindLaunchButton(ws)
  If launchShape Is Nothing Then
    Set launchShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetArea.Left, targetArea.Top, targetArea.Width, targetArea.Height)
    launchShape.Name = LAUNCH_BUTTON_NAME
  Else
    launchShape.Left = targetArea.Left
    launchShape.Top = targetArea.Top
    launchShape.Width = targetArea.Width
    launchShape.Height = targetArea.Height
  End If

  With launchShape
    .OnAction = "'" & Replace$(wb.Name, "'", "''") & "'!Ui.RunFromButton"
    .TextFrame.Characters.Text = "カレンダーピッカーを開く"
    .TextFrame.HorizontalAlignment = xlHAlignCenter
    .TextFrame.VerticalAlignment = xlVAlignCenter
    .TextFrame.Characters.Font.Color = RGB(31, 78, 121)
    .TextFrame.Characters.Font.Bold = True
    .Placement = xlMoveAndSize
    .Fill.ForeColor.RGB = RGB(233, 242, 255)
    .Line.ForeColor.RGB = RGB(91, 155, 213)
  End With
End Sub

Private Function FindLaunchButton(ByVal ws As Worksheet) As Shape
  Dim candidate As Shape

  For Each candidate In ws.Shapes
    If candidate.Name = LAUNCH_BUTTON_NAME Then
      Set FindLaunchButton = candidate
      Exit Function
    End If
  Next candidate
End Function
