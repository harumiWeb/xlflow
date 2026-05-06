Attribute VB_Name = "TwelveDataClient"
Option Explicit

Private Const BaseUrl As String = "https://api.twelvedata.com/"
Private Const CacheSheetName As String = "ApiCache"
Private Const CacheTtlSeconds As Long = 90

Public Function FetchQuoteCsv(ByVal symbol As String, ByVal apiKey As String) As String
  FetchQuoteCsv = FetchCachedText("quote|" & symbol, BuildUrl("quote", symbol, apiKey, vbNullString, 0))
End Function

Public Function FetchTimeSeriesCsv(ByVal symbol As String, ByVal apiKey As String, ByVal interval As String, ByVal outputSize As Long) As String
  FetchTimeSeriesCsv = FetchCachedText("time_series|" & symbol & "|" & interval & "|" & CStr(outputSize), BuildUrl("time_series", symbol, apiKey, interval, outputSize))
End Function

Private Function BuildUrl(ByVal endpoint As String, ByVal symbol As String, ByVal apiKey As String, ByVal interval As String, ByVal outputSize As Long) As String
  Dim url As String

  url = BaseUrl & endpoint & "?symbol=" & UrlEncode(symbol) & "&apikey=" & UrlEncode(apiKey) & "&format=csv"
  If Len(interval) > 0 Then
    url = url & "&interval=" & UrlEncode(interval)
  End If
  If outputSize > 0 Then
    url = url & "&outputsize=" & CStr(outputSize)
  End If

  BuildUrl = url
End Function

Private Function FetchCachedText(ByVal cacheKey As String, ByVal url As String) As String
  Dim cacheSheet As Worksheet
  Dim cacheRow As Long
  Dim payload As String

  Set cacheSheet = EnsureCacheSheet(ThisWorkbook)
  ResolveCacheRowIndex cacheSheet, cacheKey, cacheRow

  If cacheRow > 0 Then
    If DateDiff("s", CDate(cacheSheet.Cells(cacheRow, 2).Value), Now) < CacheTtlSeconds Then
      FetchCachedText = CStr(cacheSheet.Cells(cacheRow, 3).Value)
      Exit Function
    End If
  Else
    cacheRow = cacheSheet.Cells(cacheSheet.Rows.Count, 1).End(xlUp).Row + 1
    If cacheRow < 2 Then
      cacheRow = 2
    End If
  End If

  payload = HttpGet(url)
  cacheSheet.Cells(cacheRow, 1).Value2 = cacheKey
  cacheSheet.Cells(cacheRow, 2).Value = Now
  cacheSheet.Cells(cacheRow, 3).Value = payload
  FetchCachedText = payload
End Function

Private Function HttpGet(ByVal url As String) As String
  Dim request As Object

  Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
  request.SetTimeouts 5000, 5000, 10000, 10000
  request.Open "GET", url, False
  request.SetRequestHeader "User-Agent", "xlflow/market-dashboard"
  request.Send

  If request.Status < 200 Or request.Status >= 300 Then
    Err.Raise vbObjectError + 701, "TwelveDataClient.HttpGet", "HTTP " & request.Status & " " & request.StatusText & ": " & Left$(request.ResponseText, 500)
  End If

  HttpGet = request.ResponseText
End Function

Private Function EnsureCacheSheet(ByVal wb As Workbook) As Worksheet
  Dim ws As Worksheet

  Set ws = FindSheet(wb, CacheSheetName)
  If ws Is Nothing Then
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = CacheSheetName
    ws.Cells(1, 1).Value2 = "cache_key"
    ws.Cells(1, 2).Value2 = "fetched_at"
    ws.Cells(1, 3).Value2 = "payload"
    ws.Visible = xlSheetVeryHidden
  End If

  Set EnsureCacheSheet = ws
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

Private Sub ResolveCacheRowIndex(ByVal cacheSheet As Worksheet, ByVal cacheKey As String, ByRef cacheRow As Long)
  Dim lastRow As Long
  Dim rowIndex As Long

  lastRow = cacheSheet.Cells(cacheSheet.Rows.Count, 1).End(xlUp).Row
  For rowIndex = 2 To lastRow
    If StrComp(CStr(cacheSheet.Cells(rowIndex, 1).Value2), cacheKey, vbTextCompare) = 0 Then
      cacheRow = rowIndex
      Exit Sub
    End If
  Next rowIndex
End Sub

Private Function UrlEncode(ByVal value As String) As String
  Dim encoded As String
  Dim charIndex As Long
  Dim currentChar As String
  Dim codePoint As Long

  For charIndex = 1 To Len(value)
    currentChar = Mid$(value, charIndex, 1)
    codePoint = AscW(currentChar)
    If codePoint < 0 Then
      codePoint = codePoint + 65536
    End If

    Select Case codePoint
      Case 45, 46, 95, 126, 48 To 57, 65 To 90, 97 To 122
        encoded = encoded & currentChar
      Case 32
        encoded = encoded & "%20"
      Case Else
        encoded = encoded & "%" & Right$("0" & Hex$(codePoint), 2)
    End Select
  Next charIndex

  UrlEncode = encoded
End Function
