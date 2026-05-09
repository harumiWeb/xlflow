Attribute VB_Name = "App"
Option Explicit

Private Const INPUT_LABEL_CELL As String = "A2"
Private Const INPUT_VALUE_CELL As String = "B2"
Private Const STATUS_CELL As String = "A3"
Private Const OUTPUT_TOP_LEFT_CELL As String = "B5"
Private Const OUTPUT_MAX_MODULES As Long = 41

Public Sub RunCore(ByVal wb As Workbook)
  Dim targetSheet As Worksheet
  Dim sourceText As String
  Dim matrix As Variant

  On Error GoTo ErrHandler

  Set targetSheet = wb.Worksheets(1)
  PrepareSheet targetSheet

  sourceText = Trim$(CStr(targetSheet.Range(INPUT_VALUE_CELL).Value2))
  If Len(sourceText) = 0 Then
    ClearOutput targetSheet, targetSheet.Range(OUTPUT_TOP_LEFT_CELL)
    targetSheet.Range(STATUS_CELL).Value2 = "B2 に QR 化したい文字列を入力してください"
    Exit Sub
  End If

  matrix = QrCode.BuildMatrix(sourceText)
  PaintMatrix targetSheet, targetSheet.Range(OUTPUT_TOP_LEFT_CELL), matrix
  targetSheet.Range(STATUS_CELL).Value2 = "QR generated for " & sourceText
  Exit Sub

ErrHandler:
  targetSheet.Range(STATUS_CELL).Value2 = "Error: " & Err.Description
  Err.Raise Err.Number, Err.source, Err.Description
End Sub

Private Sub PrepareSheet(ByVal targetSheet As Worksheet)
  targetSheet.Range("A1").Value2 = "QR Code Generator"
  targetSheet.Range(INPUT_LABEL_CELL).Value2 = "Input"
  targetSheet.Range(STATUS_CELL).Value2 = ""
  targetSheet.Range("A1:B3").Font.Bold = True
  targetSheet.Range(INPUT_VALUE_CELL).HorizontalAlignment = xlLeft
End Sub

Public Sub HandleInputChanged(ByVal targetSheet As Worksheet)
  Dim eventsWereEnabled As Boolean

  eventsWereEnabled = Application.EnableEvents
  On Error GoTo CleanFail

  Application.EnableEvents = False
  RunCore targetSheet.Parent

CleanExit:
  Application.EnableEvents = eventsWereEnabled
  Exit Sub

CleanFail:
  Application.EnableEvents = eventsWereEnabled
  Err.Raise Err.Number, Err.source, Err.Description
End Sub

Private Sub PaintMatrix(ByVal targetSheet As Worksheet, ByVal topLeft As Range, ByVal matrix As Variant)
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim moduleCount As Long
  Dim quietZone As Long
  Dim renderSize As Long
  Dim outputRange As Range
  Dim cell As Range

  quietZone = 4
  moduleCount = UBound(matrix, 1) - LBound(matrix, 1) + 1
  renderSize = moduleCount + quietZone * 2

  Set outputRange = topLeft.Resize(OUTPUT_MAX_MODULES, OUTPUT_MAX_MODULES)
  outputRange.Clear
  outputRange.Interior.Color = RGB(255, 255, 255)
  outputRange.Borders.LineStyle = xlLineStyleNone

  For rowIndex = 0 To renderSize - 1
    targetSheet.Rows(topLeft.Row + rowIndex).RowHeight = 18
  Next rowIndex

  For columnIndex = 0 To renderSize - 1
    targetSheet.Columns(topLeft.Column + columnIndex).ColumnWidth = 2.57
  Next columnIndex

  For Each cell In topLeft.Resize(renderSize, renderSize)
    cell.Interior.Color = RGB(255, 255, 255)
    cell.Borders.LineStyle = xlLineStyleNone
  Next cell

  For rowIndex = 0 To moduleCount - 1
    For columnIndex = 0 To moduleCount - 1
      Set cell = topLeft.offset(rowIndex + quietZone, columnIndex + quietZone)
      If matrix(rowIndex, columnIndex) = 1 Then
        cell.Interior.Color = RGB(0, 0, 0)
      Else
        cell.Interior.Color = RGB(255, 255, 255)
      End If
    Next columnIndex
  Next rowIndex
End Sub

Private Sub ClearOutput(ByVal targetSheet As Worksheet, ByVal topLeft As Range)
  With topLeft.Resize(OUTPUT_MAX_MODULES, OUTPUT_MAX_MODULES)
    .Clear
    .Interior.Color = RGB(255, 255, 255)
    .Borders.LineStyle = xlLineStyleNone
  End With
End Sub
