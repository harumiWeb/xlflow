Attribute VB_Name = "MarketDashboard"
Option Explicit

Private Const DashboardSheetName As String = "Dashboard"
Private Const DataSheetName As String = "MarketData"
Private Const SummaryHeaderRow As Long = 8
Private Const SummaryFirstRow As Long = 9
Private Const DetailHeaderRow As Long = 15
Private Const DetailFirstRow As Long = 16
Private Const ChartHeaderRow As Long = 21
Private Const ChartTopRow As Long = 22
Private Const ChartBottomRow As Long = 38

Public Sub RefreshDashboard(ByVal wb As Workbook, ByVal apiKey As String)
  Dim dashboardSheet As Worksheet
  Dim dataSheet As Worksheet
  Dim watchlist As Variant
  Dim watchIndex As Long
  Dim summaryRow As Long
  Dim dataRow As Long
  Dim blockStartRow As Long
  Dim symbol As String
  Dim displayName As String
  Dim interval As String
  Dim outputSize As Long
  Dim accentColor As Long
  Dim seriesTable As Variant
  Dim nextDataRow As Long
  Dim quoteTable As Variant
  Dim percentChangeValue As Double
  Dim weeklyChangeValue As Double
  Dim openMarketCount As Long
  Dim totalPercentChange As Double
  Dim bestSymbol As String
  Dim bestPercentChange As Double
  Dim worstSymbol As String
  Dim worstPercentChange As Double
  Dim chartCount As Long
  Dim chartSymbols() As String
  Dim chartDisplayNames() As String
  Dim chartAccentColors() As Long
  Dim chartBlockStartRows() As Long
  Dim chartNextRows() As Long
  Dim chartDateColumns() As Long
  Dim chartCloseColumns() As Long

  Set dashboardSheet = GetPrimaryDashboardSheet(wb)
  Set dataSheet = EnsureSheet(wb, DataSheetName)
  watchlist = BuildWatchlist()
  chartCount = UBound(watchlist, 1)

  ReDim chartSymbols(1 To chartCount)
  ReDim chartDisplayNames(1 To chartCount)
  ReDim chartAccentColors(1 To chartCount)
  ReDim chartBlockStartRows(1 To chartCount)
  ReDim chartNextRows(1 To chartCount)
  ReDim chartDateColumns(1 To chartCount)
  ReDim chartCloseColumns(1 To chartCount)

  PrepareDashboardSheet dashboardSheet
  PrepareDataSheet dataSheet

  WriteDashboardHeader dashboardSheet
  WriteSummaryHeader dashboardSheet, SummaryHeaderRow
  WriteDetailHeader dashboardSheet, DetailHeaderRow

  bestPercentChange = -1E+30
  worstPercentChange = 1E+30

  summaryRow = SummaryFirstRow
  dataRow = 1

  For watchIndex = LBound(watchlist, 1) To UBound(watchlist, 1)
    symbol = CStr(watchlist(watchIndex, 1))
    displayName = CStr(watchlist(watchIndex, 2))
    interval = CStr(watchlist(watchIndex, 3))
    outputSize = CLng(watchlist(watchIndex, 4))
    accentColor = CLng(watchlist(watchIndex, 5))

    quoteTable = CsvUtil.ParseCsvText(TwelveDataClient.FetchQuoteCsv(symbol, apiKey))
    seriesTable = CsvUtil.ParseCsvText(TwelveDataClient.FetchTimeSeriesCsv(symbol, apiKey, interval, outputSize))
    percentChangeValue = QuoteNumber(quoteTable, "percent_change") / 100#
    weeklyChangeValue = CalculateSeriesChange(seriesTable, 7)
    totalPercentChange = totalPercentChange + percentChangeValue

    If QuoteIsOpen(quoteTable) Then
      openMarketCount = openMarketCount + 1
    End If

    If percentChangeValue > bestPercentChange Then
      bestPercentChange = percentChangeValue
      bestSymbol = symbol
    End If

    If percentChangeValue < worstPercentChange Then
      worstPercentChange = percentChangeValue
      worstSymbol = symbol
    End If

    WriteSummaryRow dashboardSheet, summaryRow, symbol, displayName, quoteTable, weeklyChangeValue
    WriteDetailRow dashboardSheet, DetailFirstRow + (watchIndex - 1), symbol, quoteTable, weeklyChangeValue

    blockStartRow = dataRow
    WriteSeriesBlock dataSheet, dataRow, symbol, displayName, seriesTable, nextDataRow
    chartSymbols(watchIndex) = symbol
    chartDisplayNames(watchIndex) = displayName
    chartAccentColors(watchIndex) = accentColor
    chartBlockStartRows(watchIndex) = blockStartRow
    chartNextRows(watchIndex) = nextDataRow
    chartDateColumns(watchIndex) = CsvUtil.CsvHeaderIndex(seriesTable, "datetime")
    chartCloseColumns(watchIndex) = CsvUtil.CsvHeaderIndex(seriesTable, "close")
    dataRow = nextDataRow
    WriteSparkline dashboardSheet, dataSheet, summaryRow, 13, blockStartRow, nextDataRow, seriesTable, accentColor

    summaryRow = summaryRow + 1
  Next watchIndex

  WriteHeroCards dashboardSheet, UBound(watchlist, 1), totalPercentChange / UBound(watchlist, 1), bestSymbol, bestPercentChange, worstSymbol, worstPercentChange, openMarketCount
  WriteChartSectionHeader dashboardSheet, ChartHeaderRow
  WriteOverviewCharts dashboardSheet, dataSheet, chartSymbols, chartDisplayNames, chartAccentColors, chartBlockStartRows, chartNextRows, chartDateColumns, chartCloseColumns, chartCount
  FormatDashboardSheet dashboardSheet
  dataSheet.Visible = xlSheetHidden
End Sub

Public Function BuildWatchlist() As Variant
  Dim items(1 To 4, 1 To 5) As Variant

  items(1, 1) = "AAPL"
  items(1, 2) = "Apple"
  items(1, 3) = "1day"
  items(1, 4) = 15
  items(1, 5) = RGB(37, 99, 235)

  items(2, 1) = "MSFT"
  items(2, 2) = "Microsoft"
  items(2, 3) = "1day"
  items(2, 4) = 15
  items(2, 5) = RGB(16, 185, 129)

  items(3, 1) = "BTC/USD"
  items(3, 2) = "Bitcoin"
  items(3, 3) = "1day"
  items(3, 4) = 15
  items(3, 5) = RGB(245, 158, 11)

  items(4, 1) = "ETH/USD"
  items(4, 2) = "Ethereum"
  items(4, 3) = "1day"
  items(4, 4) = 15
  items(4, 5) = RGB(139, 92, 246)

  BuildWatchlist = items
End Function

Private Sub WriteDashboardHeader(ByVal ws As Worksheet)
  With ws.Range("A1:M1")
    .Merge
    .Value2 = "Twelve Data Market Dashboard"
    .Font.Size = 18
    .Font.Bold = True
    .Font.Color = vbWhite
    .Interior.Color = RGB(15, 23, 42)
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
  End With

  With ws.Range("A2:M2")
    .Merge
    .Value2 = "Equities and crypto snapshots refreshed " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    .Font.Size = 10
    .Font.Color = RGB(71, 85, 105)
    .Interior.Color = RGB(248, 250, 252)
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
  End With

  ws.Rows(1).RowHeight = 28
  ws.Rows(2).RowHeight = 20
  ws.Rows(3).RowHeight = 8
End Sub

Private Sub WriteHeroCards(ByVal ws As Worksheet, ByVal assetCount As Long, ByVal averageChange As Double, ByVal bestSymbol As String, ByVal bestChange As Double, ByVal worstSymbol As String, ByVal worstChange As Double, ByVal openMarketCount As Long)
  WriteHeroCard ws, ws.Range("A4:C6"), "Watchlist", CStr(assetCount), "tracked assets", RGB(37, 99, 235)
  WriteHeroCard ws, ws.Range("D4:F6"), "Average Move", Format$(averageChange, "+0.00%;-0.00%;0.00%"), "daily change", RGB(16, 185, 129)
  WriteHeroCard ws, ws.Range("G4:I6"), "Top Mover", bestSymbol & " " & Format$(bestChange, "+0.00%;-0.00%;0.00%"), "strongest today", RGB(245, 158, 11)
  WriteHeroCard ws, ws.Range("J4:M6"), "Open Markets", CStr(openMarketCount), worstSymbol & " " & Format$(worstChange, "+0.00%;-0.00%;0.00%"), RGB(139, 92, 246)
End Sub

Private Sub WriteHeroCard(ByVal ws As Worksheet, ByVal targetRange As Range, ByVal title As String, ByVal valueText As String, ByVal detailText As String, ByVal accentColor As Long)
  targetRange.Merge
  targetRange.Value2 = title & vbLf & valueText & vbLf & detailText
  targetRange.WrapText = True
  targetRange.HorizontalAlignment = xlLeft
  targetRange.VerticalAlignment = xlCenter
  targetRange.Font.Color = vbWhite
  targetRange.Interior.Color = accentColor
  targetRange.Borders.LineStyle = xlContinuous
  targetRange.Borders.Color = RGB(226, 232, 240)
  targetRange.Font.Bold = True
  targetRange.Font.Size = 11
End Sub

Private Sub WriteSummaryHeader(ByVal ws As Worksheet, ByVal headerRow As Long)
  Dim headers As Variant
  Dim headerIndex As Long

  headers = Array("Symbol", "Name", "Last", "Change", "% Change", "1W %", "Open", "Day Range", "52W Range", "Market", "Updated", "Exchange", "Trend")
  For headerIndex = LBound(headers) To UBound(headers)
    ws.Cells(headerRow, headerIndex + 1).Value2 = headers(headerIndex)
  Next headerIndex

  With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, UBound(headers) + 1))
    .Font.Bold = True
    .Font.Color = vbWhite
    .Interior.Color = RGB(30, 41, 59)
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(51, 65, 85)
  End With

  ws.Columns("A").ColumnWidth = 11
  ws.Columns("B").ColumnWidth = 22
  ws.Columns("C").ColumnWidth = 12
  ws.Columns("D").ColumnWidth = 12
  ws.Columns("E").ColumnWidth = 12
  ws.Columns("F").ColumnWidth = 12
  ws.Columns("G").ColumnWidth = 12
  ws.Columns("H").ColumnWidth = 18
  ws.Columns("I").ColumnWidth = 20
  ws.Columns("J").ColumnWidth = 11
  ws.Columns("K").ColumnWidth = 14
  ws.Columns("L").ColumnWidth = 14
  ws.Columns("M").ColumnWidth = 24
End Sub

Private Sub WriteSummaryRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal symbol As String, ByVal displayName As String, ByVal quoteTable As Variant, ByVal weeklyChangeValue As Double)
  Dim nameCol As Long
  Dim exchangeCol As Long
  Dim datetimeCol As Long
  Dim closeCol As Long
  Dim changeCol As Long
  Dim percentCol As Long
  Dim openCol As Long
  Dim highCol As Long
  Dim lowCol As Long
  Dim fiftyTwoWeekLowCol As Long
  Dim fiftyTwoWeekHighCol As Long
  Dim marketText As String
  Dim dataRow As Long
  Dim lastPrice As Double
  Dim changeValue As Double
  Dim percentValue As Double
  Dim rowRange As Range

  nameCol = CsvUtil.CsvHeaderIndex(quoteTable, "name")
  exchangeCol = CsvUtil.CsvHeaderIndex(quoteTable, "exchange")
  datetimeCol = CsvUtil.CsvHeaderIndex(quoteTable, "datetime")
  closeCol = CsvUtil.CsvHeaderIndex(quoteTable, "close")
  changeCol = CsvUtil.CsvHeaderIndex(quoteTable, "change")
  percentCol = CsvUtil.CsvHeaderIndex(quoteTable, "percent_change")
  openCol = CsvUtil.CsvHeaderIndex(quoteTable, "open")
  highCol = CsvUtil.CsvHeaderIndex(quoteTable, "high")
  lowCol = CsvUtil.CsvHeaderIndex(quoteTable, "low")
  fiftyTwoWeekLowCol = CsvUtil.CsvHeaderIndex(quoteTable, "fifty_two_week_low")
  fiftyTwoWeekHighCol = CsvUtil.CsvHeaderIndex(quoteTable, "fifty_two_week_high")
  dataRow = 2

  lastPrice = Val(CStr(quoteTable(dataRow, closeCol)))
  changeValue = Val(CStr(quoteTable(dataRow, changeCol)))
  percentValue = Val(CStr(quoteTable(dataRow, percentCol))) / 100#
  If QuoteIsOpen(quoteTable) Then
    marketText = "OPEN"
  Else
    marketText = "CLOSED"
  End If

  ws.Cells(rowNumber, 1).Value2 = symbol
  ws.Cells(rowNumber, 2).Value2 = CStr(quoteTable(dataRow, nameCol))
  ws.Cells(rowNumber, 3).Value2 = lastPrice
  ws.Cells(rowNumber, 3).NumberFormat = "0.00"
  ws.Cells(rowNumber, 4).Value2 = changeValue
  ws.Cells(rowNumber, 4).NumberFormat = "+0.00;-0.00;0.00"
  ws.Cells(rowNumber, 5).Value2 = percentValue
  ws.Cells(rowNumber, 5).NumberFormat = "+0.00%;-0.00%;0.00%"
  ws.Cells(rowNumber, 6).Value2 = weeklyChangeValue
  ws.Cells(rowNumber, 6).NumberFormat = "+0.00%;-0.00%;0.00%"
  ws.Cells(rowNumber, 7).Value2 = Val(CStr(quoteTable(dataRow, openCol)))
  ws.Cells(rowNumber, 7).NumberFormat = "0.00"
  ws.Cells(rowNumber, 8).Value2 = FormatRangeText(Val(CStr(quoteTable(dataRow, lowCol))), Val(CStr(quoteTable(dataRow, highCol))))
  ws.Cells(rowNumber, 9).Value2 = FormatRangeText(Val(CStr(quoteTable(dataRow, fiftyTwoWeekLowCol))), Val(CStr(quoteTable(dataRow, fiftyTwoWeekHighCol))))
  ws.Cells(rowNumber, 10).Value2 = marketText
  ws.Cells(rowNumber, 11).Value2 = CStr(quoteTable(dataRow, datetimeCol))
  ws.Cells(rowNumber, 12).Value2 = CStr(quoteTable(dataRow, exchangeCol))

  Set rowRange = ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 13))
  StylePerformanceRow rowRange, changeValue
  ws.Cells(rowNumber, 1).Font.Bold = True
  ws.Cells(rowNumber, 1).HorizontalAlignment = xlCenter
  ws.Rows(rowNumber).RowHeight = 24
End Sub

Private Sub WriteDetailHeader(ByVal ws As Worksheet, ByVal headerRow As Long)
  Dim headers As Variant
  Dim headerIndex As Long

  headers = Array("Deep Dive", "Prev Close", "52W Low", "52W High", "Dist. High", "Dist. Low", "Volume", "Avg / 7D", "Currency")
  For headerIndex = LBound(headers) To UBound(headers)
    ws.Cells(headerRow, headerIndex + 1).Value2 = headers(headerIndex)
  Next headerIndex

  With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, UBound(headers) + 1))
    .Font.Bold = True
    .Font.Color = vbWhite
    .Interior.Color = RGB(15, 23, 42)
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(51, 65, 85)
  End With
End Sub

Private Sub WriteDetailRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal symbol As String, ByVal quoteTable As Variant, ByVal weeklyChangeValue As Double)
  Dim prevClose As Double
  Dim low52 As Double
  Dim high52 As Double
  Dim currentClose As Double
  Dim volumeText As String
  Dim avgText As String

  prevClose = QuoteNumber(quoteTable, "previous_close")
  low52 = QuoteNumber(quoteTable, "fifty_two_week_low")
  high52 = QuoteNumber(quoteTable, "fifty_two_week_high")
  currentClose = QuoteNumber(quoteTable, "close")
  volumeText = QuoteTextOrDefault(quoteTable, "volume", "-")
  avgText = QuoteTextOrDefault(quoteTable, "average_volume", Format$(weeklyChangeValue, "+0.00%;-0.00%;0.00%"))

  ws.Cells(rowNumber, 1).Value2 = symbol
  ws.Cells(rowNumber, 2).Value2 = prevClose
  ws.Cells(rowNumber, 2).NumberFormat = "0.00"
  ws.Cells(rowNumber, 3).Value2 = low52
  ws.Cells(rowNumber, 3).NumberFormat = "0.00"
  ws.Cells(rowNumber, 4).Value2 = high52
  ws.Cells(rowNumber, 4).NumberFormat = "0.00"
  ws.Cells(rowNumber, 5).Value2 = FormatDistanceFrom(currentClose, high52)
  ws.Cells(rowNumber, 6).Value2 = FormatDistanceFrom(currentClose, low52)
  ws.Cells(rowNumber, 7).Value2 = volumeText
  ws.Cells(rowNumber, 8).Value2 = avgText
  ws.Cells(rowNumber, 9).Value2 = QuoteTextOrDefault(quoteTable, "currency", "USD")

  With ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 9))
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(226, 232, 240)
    .Interior.Color = RGB(248, 250, 252)
  End With
End Sub

Private Sub WriteSeriesBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal symbol As String, ByVal displayName As String, ByVal seriesTable As Variant, ByRef nextRow As Long)
  Dim headerCol As Long
  Dim sourceRow As Long
  Dim targetRow As Long
  Dim columnIndex As Long
  Dim columnCount As Long
  Dim rowCount As Long
  Dim lastRow As Long
  Dim titleRange As Range
  Dim headerRange As Range

  rowCount = UBound(seriesTable, 1)
  columnCount = UBound(seriesTable, 2)

  Set titleRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, columnCount))
  titleRange.Merge
  titleRange.Value2 = symbol & "  " & displayName
  titleRange.Font.Bold = True
  titleRange.Font.Color = vbWhite
  titleRange.Interior.Color = RGB(15, 23, 42)

  Set headerRange = ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + 1, columnCount))
  For columnIndex = 1 To columnCount
    ws.Cells(startRow + 1, columnIndex).Value2 = seriesTable(1, columnIndex)
  Next columnIndex
  headerRange.Font.Bold = True
  headerRange.Font.Color = vbWhite
  headerRange.Interior.Color = RGB(51, 65, 85)

  lastRow = startRow + rowCount
  For sourceRow = 2 To rowCount
    targetRow = startRow + sourceRow
    For columnIndex = 1 To columnCount
      ws.Cells(targetRow, columnIndex).Value2 = seriesTable(rowCount - sourceRow + 2, columnIndex)
    Next columnIndex
  Next sourceRow

  With ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, columnCount))
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(203, 213, 225)
  End With

  nextRow = lastRow + 2
End Sub

Private Sub WriteSparkline(ByVal dashboardSheet As Worksheet, ByVal dataSheet As Worksheet, ByVal summaryRow As Long, ByVal targetColumn As Long, ByVal blockStartRow As Long, ByVal nextDataRow As Long, ByVal seriesTable As Variant, ByVal accentColor As Long)
  Dim sparkTarget As Range
  Dim trendText As String

  trendText = BuildTrendText(seriesTable)

  Set sparkTarget = dashboardSheet.Cells(summaryRow, targetColumn)
  sparkTarget.Value2 = trendText
  sparkTarget.Font.Name = "Consolas"
  sparkTarget.Font.Color = accentColor
  sparkTarget.Font.Bold = True
  sparkTarget.HorizontalAlignment = xlCenter
  sparkTarget.VerticalAlignment = xlCenter
End Sub

Private Function BuildTrendText(ByVal seriesTable As Variant) As String
  Const MaxPoints As Long = 12
  Dim closeCol As Long
  Dim totalPoints As Long
  Dim sampleCount As Long
  Dim values() As Double
  Dim pointIndex As Long
  Dim sourceRow As Long
  Dim minValue As Double
  Dim maxValue As Double
  Dim normalized As Double
  Dim bucket As Long
  Dim trendText As String

  closeCol = CsvUtil.CsvHeaderIndex(seriesTable, "close")
  totalPoints = UBound(seriesTable, 1) - 1
  If totalPoints <= 0 Then
    BuildTrendText = "-"
    Exit Function
  End If

  sampleCount = totalPoints
  If sampleCount > MaxPoints Then
    sampleCount = MaxPoints
  End If

  ReDim values(1 To sampleCount)
  minValue = 1E+30
  maxValue = -1E+30

  For pointIndex = 1 To sampleCount
    sourceRow = 2 + sampleCount - pointIndex
    values(pointIndex) = Val(CStr(seriesTable(sourceRow, closeCol)))
    If values(pointIndex) < minValue Then
      minValue = values(pointIndex)
    End If
    If values(pointIndex) > maxValue Then
      maxValue = values(pointIndex)
    End If
  Next pointIndex

  For pointIndex = 1 To sampleCount
    If maxValue <= minValue Then
      bucket = 3
    Else
      normalized = (values(pointIndex) - minValue) / (maxValue - minValue)
      bucket = CLng(Int(normalized * 7))
    End If
    trendText = trendText & TrendGlyph(bucket)
  Next pointIndex

  BuildTrendText = trendText
End Function

Private Function TrendGlyph(ByVal bucket As Long) As String
  Select Case bucket
    Case Is <= 0
      TrendGlyph = ChrW$(&H2581)
    Case 1
      TrendGlyph = ChrW$(&H2582)
    Case 2
      TrendGlyph = ChrW$(&H2583)
    Case 3
      TrendGlyph = ChrW$(&H2584)
    Case 4
      TrendGlyph = ChrW$(&H2585)
    Case 5
      TrendGlyph = ChrW$(&H2586)
    Case 6
      TrendGlyph = ChrW$(&H2587)
    Case Else
      TrendGlyph = ChrW$(&H2588)
  End Select
End Function

Private Sub WriteChartSectionHeader(ByVal ws As Worksheet, ByVal rowNumber As Long)
  With ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 13))
    .Merge
    .Value2 = "Market Charts"
    .Font.Bold = True
    .Font.Color = vbWhite
    .Interior.Color = RGB(15, 23, 42)
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(51, 65, 85)
  End With

  ws.Rows(rowNumber).RowHeight = 22
End Sub

Private Sub WriteOverviewCharts(ByVal dashboardSheet As Worksheet, ByVal dataSheet As Worksheet, ByRef chartSymbols() As String, ByRef chartDisplayNames() As String, ByRef chartAccentColors() As Long, ByRef chartBlockStartRows() As Long, ByRef chartNextRows() As Long, ByRef chartDateColumns() As Long, ByRef chartCloseColumns() As Long, ByVal chartCount As Long)
  Dim previousScreenUpdating As Boolean

  previousScreenUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = True

  CreatePriceChart dashboardSheet, dataSheet, "A22:G38", "Equities Price Trend", chartSymbols, chartDisplayNames, chartAccentColors, chartBlockStartRows, chartNextRows, chartDateColumns, chartCloseColumns, chartCount, False
  CreatePriceChart dashboardSheet, dataSheet, "H22:M38", "Crypto Price Trend", chartSymbols, chartDisplayNames, chartAccentColors, chartBlockStartRows, chartNextRows, chartDateColumns, chartCloseColumns, chartCount, True

  Application.ScreenUpdating = previousScreenUpdating
End Sub

Private Sub CreatePriceChart(ByVal dashboardSheet As Worksheet, ByVal dataSheet As Worksheet, ByVal targetAddress As String, ByVal chartTitle As String, ByRef chartSymbols() As String, ByRef chartDisplayNames() As String, ByRef chartAccentColors() As Long, ByRef chartBlockStartRows() As Long, ByRef chartNextRows() As Long, ByRef chartDateColumns() As Long, ByRef chartCloseColumns() As Long, ByVal chartCount As Long, ByVal cryptoOnly As Boolean)
  Dim targetRange As Range
  Dim chartObject As ChartObject
  Dim chart As Chart
  Dim seriesIndex As Long

  Set targetRange = dashboardSheet.Range(targetAddress)
  Set chartObject = dashboardSheet.ChartObjects.Add(targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
  chartObject.Placement = xlMoveAndSize

  Set chart = chartObject.Chart
  chart.ChartType = xlLine
  chart.HasTitle = True
  chart.ChartTitle.Text = chartTitle
  chart.HasLegend = True
  chart.Legend.Position = xlLegendPositionBottom

  For seriesIndex = 1 To chartCount
    If SymbolMatchesGroup(chartSymbols(seriesIndex), cryptoOnly) Then
      AddChartSeries chart, dataSheet, chartDisplayNames(seriesIndex), chartAccentColors(seriesIndex), chartBlockStartRows(seriesIndex), chartNextRows(seriesIndex), chartDateColumns(seriesIndex), chartCloseColumns(seriesIndex)
    End If
  Next seriesIndex

  chart.Axes(xlCategory).TickLabelSpacing = 3
End Sub

Private Sub AddChartSeries(ByVal chart As Chart, ByVal dataSheet As Worksheet, ByVal seriesName As String, ByVal accentColor As Long, ByVal blockStartRow As Long, ByVal nextDataRow As Long, ByVal dateColumn As Long, ByVal closeColumn As Long)
  Dim chartSeries As Series
  Dim firstDataRow As Long
  Dim lastDataRow As Long

  firstDataRow = blockStartRow + 2
  lastDataRow = nextDataRow - 2

  Set chartSeries = chart.SeriesCollection.NewSeries
  chartSeries.Name = seriesName
  chartSeries.XValues = dataSheet.Range(dataSheet.Cells(firstDataRow, dateColumn), dataSheet.Cells(lastDataRow, dateColumn))
  chartSeries.Values = dataSheet.Range(dataSheet.Cells(firstDataRow, closeColumn), dataSheet.Cells(lastDataRow, closeColumn))
  chartSeries.Border.Color = accentColor
End Sub

Private Function SymbolMatchesGroup(ByVal symbol As String, ByVal cryptoOnly As Boolean) As Boolean
  Dim isCrypto As Boolean

  isCrypto = InStr(1, symbol, "/", vbTextCompare) > 0
  SymbolMatchesGroup = (isCrypto = cryptoOnly)
End Function

Private Sub PrepareDashboardSheet(ByVal ws As Worksheet)
  ClearWorksheet ws
  ws.Cells.Font.Name = "Calibri"
  ws.Cells.Font.Size = 10
  ws.Range("A:M").ColumnWidth = 12
  ws.Rows("4:6").RowHeight = 42
  ws.Rows(ChartTopRow & ":" & ChartBottomRow).RowHeight = 18
End Sub

Private Sub PrepareDataSheet(ByVal ws As Worksheet)
  ClearWorksheet ws
  ws.Cells.Font.Name = "Calibri"
  ws.Cells.Font.Size = 9
  ws.Visible = xlSheetVisible
End Sub

Private Sub FormatDashboardSheet(ByVal ws As Worksheet)
  With ws.Range("A4:M6")
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(226, 232, 240)
  End With

  With ws.Range("A8:M" & SummaryFirstRow + 3)
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(226, 232, 240)
  End With

  With ws.Range("A15:I19")
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(226, 232, 240)
  End With

  With ws.Range(ws.Cells(ChartHeaderRow, 1), ws.Cells(ChartBottomRow, 13))
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(226, 232, 240)
  End With
End Sub

Private Sub StylePerformanceRow(ByVal targetRange As Range, ByVal changeValue As Double)
  targetRange.Borders.LineStyle = xlContinuous
  targetRange.Borders.Color = RGB(226, 232, 240)

  If changeValue > 0 Then
    targetRange.Interior.Color = RGB(236, 253, 245)
    targetRange.Font.Color = RGB(4, 120, 87)
  ElseIf changeValue < 0 Then
    targetRange.Interior.Color = RGB(254, 242, 242)
    targetRange.Font.Color = RGB(153, 27, 27)
  Else
    targetRange.Interior.Color = RGB(248, 250, 252)
    targetRange.Font.Color = RGB(71, 85, 105)
  End If
End Sub

Private Function QuoteTextOrDefault(ByVal quoteTable As Variant, ByVal headerName As String, ByVal defaultValue As String) As String
  Dim columnIndex As Long

  columnIndex = OptionalCsvHeaderIndex(quoteTable, headerName)
  If columnIndex = 0 Then
    QuoteTextOrDefault = defaultValue
  Else
    QuoteTextOrDefault = CStr(quoteTable(2, columnIndex))
  End If
End Function

Private Function QuoteNumber(ByVal quoteTable As Variant, ByVal headerName As String) As Double
  QuoteNumber = Val(QuoteTextOrDefault(quoteTable, headerName, "0"))
End Function

Private Function QuoteIsOpen(ByVal quoteTable As Variant) As Boolean
  QuoteIsOpen = StrComp(QuoteTextOrDefault(quoteTable, "is_market_open", "false"), "true", vbTextCompare) = 0
End Function

Private Function OptionalCsvHeaderIndex(ByVal csvTable As Variant, ByVal headerName As String) As Long
  Dim columnIndex As Long

  For columnIndex = LBound(csvTable, 2) To UBound(csvTable, 2)
    If StrComp(CStr(csvTable(1, columnIndex)), headerName, vbTextCompare) = 0 Then
      OptionalCsvHeaderIndex = columnIndex
      Exit Function
    End If
  Next columnIndex
End Function

Private Function CalculateSeriesChange(ByVal seriesTable As Variant, ByVal lookbackPeriods As Long) As Double
  Dim closeCol As Long
  Dim latestClose As Double
  Dim baselineRow As Long
  Dim baselineClose As Double

  closeCol = CsvUtil.CsvHeaderIndex(seriesTable, "close")
  latestClose = Val(CStr(seriesTable(2, closeCol)))
  baselineRow = 2 + lookbackPeriods
  If baselineRow > UBound(seriesTable, 1) Then
    baselineRow = UBound(seriesTable, 1)
  End If
  baselineClose = Val(CStr(seriesTable(baselineRow, closeCol)))

  If baselineClose = 0 Then
    CalculateSeriesChange = 0
  Else
    CalculateSeriesChange = (latestClose / baselineClose) - 1#
  End If
End Function

Private Function FormatRangeText(ByVal lowValue As Double, ByVal highValue As Double) As String
  FormatRangeText = Format$(lowValue, "0.00") & " - " & Format$(highValue, "0.00")
End Function

Private Function FormatDistanceFrom(ByVal currentValue As Double, ByVal referenceValue As Double) As String
  Dim distance As Double

  If referenceValue = 0 Then
    FormatDistanceFrom = "-"
    Exit Function
  End If

  distance = (currentValue / referenceValue) - 1#
  FormatDistanceFrom = Format$(distance, "+0.00%;-0.00%;0.00%")
End Function

Private Sub ClearWorksheet(ByVal ws As Worksheet)
  Dim chartIndex As Long
  Dim chartCount As Long
  Dim previousScreenUpdating As Boolean

  ws.Cells.UnMerge
  ws.Cells.Clear

  chartCount = ws.ChartObjects.Count
  If chartCount > 0 Then
    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
  End If

  For chartIndex = chartCount To 1 Step -1
    ws.ChartObjects(chartIndex).Delete
  Next chartIndex

  If chartCount > 0 Then
    Application.ScreenUpdating = previousScreenUpdating
  End If
End Sub

Private Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
  Dim ws As Worksheet

  Set ws = FindSheet(wb, sheetName)
  If ws Is Nothing Then
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
  End If

  Set EnsureSheet = ws
End Function

Private Function FindSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
  Dim ws As Worksheet

  For Each ws In wb.Worksheets
    If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
      Set FindSheet = ws
      Exit Function
    End If
  Next ws
End Function

Private Function GetPrimaryDashboardSheet(ByVal wb As Workbook) As Worksheet
  Dim ws As Worksheet

  Set ws = FindSheet(wb, DashboardSheetName)
  If ws Is Nothing Then
    Set ws = wb.Worksheets(1)
    If StrComp(ws.Name, DashboardSheetName, vbTextCompare) <> 0 Then
      ws.Name = DashboardSheetName
    End If
  ElseIf ws.Index <> 1 Then
    ws.Move Before:=wb.Worksheets(1)
  End If

  Set GetPrimaryDashboardSheet = ws
End Function
