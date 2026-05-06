Attribute VB_Name = "WeatherSheet"
Option Explicit

Private Const HEADER_ROW As Long = 7
Private Const FORECAST_FIRST_ROW As Long = 8
Private Const MAX_FORECAST_COUNT As Long = 3

Public Sub RenderForecast(ByVal wb As Workbook, ByVal forecastRoot As Object)
  Dim ws As Worksheet
  Dim forecasts As Collection
  Dim descriptionNode As Object
  Dim previousScreenUpdating As Boolean

  previousScreenUpdating = Application.ScreenUpdating
  On Error GoTo ErrHandler

  Application.ScreenUpdating = False
  Set ws = wb.Worksheets(1)
  Set descriptionNode = GetChildObject(forecastRoot, "description")
  Set forecasts = GetChildArray(forecastRoot, "forecasts")

  PrepareCanvas ws
  WriteSummary ws, forecastRoot, descriptionNode
  WriteForecastHeaders ws
  WriteForecastRows ws, forecasts

CleanExit:
  Application.ScreenUpdating = previousScreenUpdating
  Exit Sub

ErrHandler:
  Application.ScreenUpdating = previousScreenUpdating
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub PrepareCanvas(ByVal ws As Worksheet)
  ws.Cells.UnMerge
  ws.Cells.Clear

  ws.Range("A1:G1").Merge
  ws.Range("A2:G2").Merge
  ws.Range("A3:G5").Merge

  ws.Columns("A").ColumnWidth = 14
  ws.Columns("B").ColumnWidth = 10
  ws.Columns("C").ColumnWidth = 14
  ws.Columns("D").ColumnWidth = 30
  ws.Columns("E").ColumnWidth = 39
  ws.Columns("F").ColumnWidth = 8
  ws.Columns("G").ColumnWidth = 8

  ws.Rows("1:2").RowHeight = 24
  ws.Rows("3:5").RowHeight = 62
  ws.Rows(HEADER_ROW).RowHeight = 26
  ws.Rows("8:10").RowHeight = 54

  With ws.Range("A1:G10").Font
    .Name = "Meiryo"
    .Size = 11
  End With
End Sub

Private Sub WriteSummary(ByVal ws As Worksheet, ByVal forecastRoot As Object, ByVal descriptionNode As Object)
  With ws.Range("A1")
    .value = GetChildString(forecastRoot, "title")
    .Font.Bold = True
    .Font.Size = 18
    .HorizontalAlignment = xlLeft
  End With

  With ws.Range("A2")
    .value = "更新: " & GetChildString(forecastRoot, "publicTimeFormatted")
    .Font.Size = 11
    .HorizontalAlignment = xlLeft
  End With

  With ws.Range("A3:G5")
    .value = NormalizeDescriptionText(GetChildString(descriptionNode, "bodyText"))
    .WrapText = True
    .VerticalAlignment = xlTop
    .HorizontalAlignment = xlLeft
    .Interior.Color = RGB(248, 248, 248)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
  End With
End Sub

Private Sub WriteForecastHeaders(ByVal ws As Worksheet)
  With ws.Range("A7:G7")
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = RGB(221, 235, 247)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
  End With

  ws.Range("A7").value = "日付"
  ws.Range("B7").value = "画像"
  ws.Range("C7").value = "天気"
  ws.Range("D7").value = "詳細"
  ws.Range("E7").value = "降水確率(00-06/06-12/12-18/18-24)"
  ws.Range("F7").value = "最低"
  ws.Range("G7").value = "最高"
End Sub

Private Sub WriteForecastRows(ByVal ws As Worksheet, ByVal forecasts As Collection)
  Dim rowIndex As Long
  Dim forecastIndex As Long
  Dim forecastItem As Object
  Dim detailNode As Object
  Dim temperatureNode As Object
  Dim chanceNode As Object
  Dim telop As String

  For forecastIndex = 1 To MinLong(MAX_FORECAST_COUNT, forecasts.Count)
    rowIndex = FORECAST_FIRST_ROW + forecastIndex - 1
    Set forecastItem = forecasts.Item(forecastIndex)
    Set detailNode = GetChildObject(forecastItem, "detail")
    Set temperatureNode = GetChildObject(forecastItem, "temperature")
    Set chanceNode = GetChildObject(forecastItem, "chanceOfRain")
    telop = GetChildString(forecastItem, "telop")

    ws.Cells(rowIndex, 1).value = GetChildString(forecastItem, "dateLabel")
    ws.Cells(rowIndex, 2).value = WeatherIconText(telop)
    ws.Cells(rowIndex, 3).value = telop
    ws.Cells(rowIndex, 4).value = DetailText(detailNode, telop)
    ws.Cells(rowIndex, 5).value = ChanceOfRainText(chanceNode)
    ws.Cells(rowIndex, 6).value = TemperatureValueText(temperatureNode, "min")
    ws.Cells(rowIndex, 7).value = TemperatureValueText(temperatureNode, "max")

    With ws.Range("A" & CStr(rowIndex) & ":G" & CStr(rowIndex))
      .Borders.LineStyle = xlContinuous
      .Borders.Weight = xlThin
      .VerticalAlignment = xlCenter
      .WrapText = True
    End With

    With ws.Cells(rowIndex, 2)
      .Font.Size = 28
      .Font.Bold = True
      .Font.Color = WeatherIconColor(telop)
      .HorizontalAlignment = xlCenter
    End With

    ws.Cells(rowIndex, 1).HorizontalAlignment = xlCenter
    ws.Cells(rowIndex, 3).HorizontalAlignment = xlCenter
    ws.Cells(rowIndex, 6).HorizontalAlignment = xlCenter
    ws.Cells(rowIndex, 7).HorizontalAlignment = xlCenter
    ws.Cells(rowIndex, 4).HorizontalAlignment = xlLeft
    ws.Cells(rowIndex, 5).HorizontalAlignment = xlCenter
  Next forecastIndex
End Sub

Private Function GetChildObject(ByVal parent As Object, ByVal key As String) As Object
  If Not parent.Exists(key) Then
    Err.Raise vbObjectError + 2200, "WeatherSheet.GetChildObject", "Missing JSON property '" & key & "'."
  End If

  If Not IsObject(parent(key)) Then
    Err.Raise vbObjectError + 2201, "WeatherSheet.GetChildObject", "JSON property '" & key & "' is not an object."
  End If

  Set GetChildObject = parent(key)
End Function

Private Function GetChildArray(ByVal parent As Object, ByVal key As String) As Collection
  If Not parent.Exists(key) Then
    Err.Raise vbObjectError + 2202, "WeatherSheet.GetChildArray", "Missing JSON property '" & key & "'."
  End If

  If Not IsObject(parent(key)) Then
    Err.Raise vbObjectError + 2203, "WeatherSheet.GetChildArray", "JSON property '" & key & "' is not an array."
  End If

  Set GetChildArray = parent(key)
End Function

Private Function GetChildString(ByVal parent As Object, ByVal key As String) As String
  If Not parent.Exists(key) Then
    Err.Raise vbObjectError + 2204, "WeatherSheet.GetChildString", "Missing JSON property '" & key & "'."
  End If

  If IsNull(parent(key)) Then
    Err.Raise vbObjectError + 2205, "WeatherSheet.GetChildString", "JSON property '" & key & "' is null."
  End If

  GetChildString = CStr(parent(key))
End Function

Private Function GetChildStringOrDefault(ByVal parent As Object, ByVal key As String, ByVal defaultValue As String) As String
  If Not parent.Exists(key) Then
    GetChildStringOrDefault = defaultValue
  ElseIf IsNull(parent(key)) Then
    GetChildStringOrDefault = defaultValue
  Else
    GetChildStringOrDefault = CStr(parent(key))
  End If
End Function

Private Function TemperatureValueText(ByVal temperatureNode As Object, ByVal key As String) As String
  Dim edgeNode As Object
  Set edgeNode = GetChildObject(temperatureNode, key)
  TemperatureValueText = GetChildStringOrDefault(edgeNode, "celsius", "--")
End Function

Private Function DetailText(ByVal detailNode As Object, ByVal telop As String) As String
  DetailText = GetChildStringOrDefault(detailNode, "weather", telop)
  DetailText = Replace(DetailText, vbCr, " ")
  DetailText = Replace(DetailText, vbLf, " ")
End Function

Private Function ChanceOfRainText(ByVal chanceNode As Object) As String
  ChanceOfRainText = _
    "00-06: " & GetChildStringOrDefault(chanceNode, "T00_06", "--%") & " / 06-12: " & GetChildStringOrDefault(chanceNode, "T06_12", "--%") & " /" & vbLf & _
    "12-18: " & GetChildStringOrDefault(chanceNode, "T12_18", "--%") & " / 18-24: " & GetChildStringOrDefault(chanceNode, "T18_24", "--%")
End Function

Private Function NormalizeDescriptionText(ByVal bodyText As String) As String
  NormalizeDescriptionText = Replace(bodyText, vbCrLf, vbLf)
  NormalizeDescriptionText = Replace(NormalizeDescriptionText, vbCr, vbLf)
End Function

Private Function WeatherIconText(ByVal telop As String) As String
  Dim hasSunny As Boolean
  Dim hasCloud As Boolean
  Dim hasRain As Boolean
  Dim hasSnow As Boolean

  hasSunny = InStr(1, telop, "晴", vbTextCompare) > 0
  hasCloud = InStr(1, telop, "曇", vbTextCompare) > 0 Or InStr(1, telop, "くもり", vbTextCompare) > 0
  hasRain = InStr(1, telop, "雨", vbTextCompare) > 0
  hasSnow = InStr(1, telop, "雪", vbTextCompare) > 0

  If hasSunny And hasCloud Then
    WeatherIconText = ChrW$(&H2600) & ChrW$(&H2601)
  ElseIf hasCloud And hasRain Then
    WeatherIconText = ChrW$(&H2601) & ChrW$(&H2614)
  ElseIf hasSunny And hasRain Then
    WeatherIconText = ChrW$(&H2600) & ChrW$(&H2614)
  ElseIf hasCloud And hasSnow Then
    WeatherIconText = ChrW$(&H2601) & ChrW$(&H2744)
  ElseIf hasSunny Then
    WeatherIconText = ChrW$(&H2600)
  ElseIf hasCloud Then
    WeatherIconText = ChrW$(&H2601)
  ElseIf hasRain Then
    WeatherIconText = ChrW$(&H2614)
  ElseIf hasSnow Then
    WeatherIconText = ChrW$(&H2744)
  Else
    WeatherIconText = "?"
  End If
End Function

Private Function WeatherIconColor(ByVal telop As String) As Long
  If InStr(1, telop, "晴", vbTextCompare) > 0 Then
    WeatherIconColor = RGB(237, 125, 49)
  ElseIf InStr(1, telop, "雨", vbTextCompare) > 0 Then
    WeatherIconColor = RGB(91, 155, 213)
  ElseIf InStr(1, telop, "雪", vbTextCompare) > 0 Then
    WeatherIconColor = RGB(112, 173, 71)
  Else
    WeatherIconColor = RGB(165, 165, 165)
  End If
End Function

Private Function MinLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
  If leftValue < rightValue Then
    MinLong = leftValue
  Else
    MinLong = rightValue
  End If
End Function
