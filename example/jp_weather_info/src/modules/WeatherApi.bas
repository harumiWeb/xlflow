Attribute VB_Name = "WeatherApi"
Option Explicit

Private Const WEATHER_API_BASE_URL As String = "https://weather.tsukumijima.net/api/forecast/city/"
Private Const DEFAULT_CITY_CODE_VALUE As String = "130010"

Public Function DefaultCityCode() As String
  DefaultCityCode = DEFAULT_CITY_CODE_VALUE
End Function

Public Function FetchForecast(ByVal cityCode As String) As Object
  Dim responseText As String
  responseText = GetJsonResponse(BuildForecastUrl(cityCode))
  Set FetchForecast = JsonParser.ParseJson(responseText)
End Function

Private Function BuildForecastUrl(ByVal cityCode As String) As String
  If Len(Trim$(cityCode)) = 0 Then
    Err.Raise vbObjectError + 2000, "WeatherApi.BuildForecastUrl", "cityCode must not be empty."
  End If

  BuildForecastUrl = WEATHER_API_BASE_URL & Trim$(cityCode)
End Function

Private Function GetJsonResponse(ByVal url As String) As String
  Dim request As Object
  Set request = CreateObject("WinHttp.WinHttpRequest.5.1")

  request.SetTimeouts 5000, 5000, 15000, 15000
  request.Open "GET", url, False
  request.SetRequestHeader "Accept", "application/json"
  request.Send

  If CLng(request.Status) <> 200 Then
    Err.Raise vbObjectError + 2001, "WeatherApi.GetJsonResponse", _
      "Weather API request failed with status " & CStr(request.Status) & " " & CStr(request.StatusText)
  End If

  GetJsonResponse = DecodeUtf8Response(request.responseBody)
End Function

Private Function DecodeUtf8Response(ByVal responseBody As Variant) As String
  Dim stream As Object
  Dim streamOpened As Boolean
  Dim errNumber As Long
  Dim errSource As String
  Dim errDescription As String

  On Error GoTo ErrHandler

  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 1
  stream.Open
  streamOpened = True
  stream.Write responseBody
  stream.Position = 0
  stream.Type = 2
  stream.Charset = "utf-8"

  DecodeUtf8Response = stream.ReadText

CleanExit:
  If streamOpened Then
    stream.Close
  End If
  Exit Function

ErrHandler:
  errNumber = Err.Number
  errSource = Err.Source
  errDescription = Err.Description
  If Not streamOpened Then
    Err.Raise errNumber, errSource, errDescription
  End If

  On Error GoTo CloseFailed
  stream.Close
  On Error GoTo 0
  Err.Raise errNumber, errSource, errDescription

CloseFailed:
    On Error GoTo 0
  Err.Clear
  Err.Raise errNumber, errSource, errDescription
End Function
