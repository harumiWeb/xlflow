Attribute VB_Name = "NewsApi"
Option Explicit

Private Const NEWS_API_BASE_URL As String = "https://newsapi.org/v2/everything"
Private Const DEFAULT_QUERY As String = "world"
Private Const DEFAULT_LANGUAGE As String = "en"
Private Const DEFAULT_SORT_ORDER As String = "publishedAt"
Private Const DEFAULT_PAGE_SIZE As Long = 15

Public Function FetchWorldNews() As Object
  Dim responseText As String
  Dim parsedRoot As Object

  responseText = GetJsonResponse(BuildWorldNewsUrl(), GetRequiredApiKey())
  Set parsedRoot = JsonParser.ParseJson(responseText)
  ValidateSuccessResponse parsedRoot

  Set FetchWorldNews = parsedRoot
End Function

Public Function BuildWorldNewsUrl() As String
  BuildWorldNewsUrl = NEWS_API_BASE_URL & _
    "?q=" & UrlEncode(DEFAULT_QUERY) & _
    "&language=" & UrlEncode(DEFAULT_LANGUAGE) & _
    "&sortBy=" & UrlEncode(DEFAULT_SORT_ORDER) & _
    "&pageSize=" & CStr(DEFAULT_PAGE_SIZE)
End Function

Private Function GetRequiredApiKey() As String
  GetRequiredApiKey = Trim$(Environ$("NEWSAPI_KEY"))
  If Len(GetRequiredApiKey) = 0 Then
    Err.Raise vbObjectError + 2300, "NewsApi.GetRequiredApiKey", "NEWSAPI_KEY environment variable is not set."
  End If
End Function

Private Function GetJsonResponse(ByVal url As String, ByVal apiKey As String) As String
  Dim request As Object
  Dim responseText As String

  Set request = CreateObject("WinHttp.WinHttpRequest.5.1")

  request.SetTimeouts 5000, 5000, 15000, 15000
  request.Open "GET", url, False
  request.SetRequestHeader "Accept", "application/json"
  request.SetRequestHeader "X-Api-Key", apiKey
  request.Send

  responseText = DecodeUtf8Response(request.responseBody)

  If CLng(request.Status) <> 200 Then
    Err.Raise vbObjectError + 2301, "NewsApi.GetJsonResponse", _
      "NewsAPI request failed with status " & CStr(request.Status) & " " & CStr(request.StatusText) & ". " & LimitText(responseText, 240)
  End If

  GetJsonResponse = responseText
End Function

Private Sub ValidateSuccessResponse(ByVal newsRoot As Object)
  Dim statusValue As String
  Dim messageValue As String

  statusValue = GetStringValue(newsRoot, "status", "missing")
  If LCase$(statusValue) <> "ok" Then
    messageValue = GetStringValue(newsRoot, "message", "NewsAPI returned a non-ok status.")
    Err.Raise vbObjectError + 2302, "NewsApi.ValidateSuccessResponse", messageValue
  End If
End Sub

Private Function GetStringValue(ByVal parent As Object, ByVal key As String, ByVal defaultValue As String) As String
  If Not parent.Exists(key) Then
    GetStringValue = defaultValue
  ElseIf IsNull(parent(key)) Then
    GetStringValue = defaultValue
  Else
    GetStringValue = CStr(parent(key))
  End If
End Function

Private Function LimitText(ByVal value As String, ByVal maxLength As Long) As String
  If Len(value) <= maxLength Then
    LimitText = value
  Else
    LimitText = Left$(value, maxLength - 3) & "..."
  End If
End Function

Private Function UrlEncode(ByVal value As String) As String
  Dim index As Long
  Dim ch As String
  Dim encoded As String
  Dim codePoint As Long

  For index = 1 To Len(value)
    ch = Mid$(value, index, 1)
    codePoint = AscW(ch)

    Select Case codePoint
      Case 48 To 57, 65 To 90, 97 To 122
        encoded = encoded & ch
      Case 45, 46, 95, 126
        encoded = encoded & ch
      Case 32
        encoded = encoded & "%20"
      Case Else
        encoded = encoded & "%" & Right$("0" & Hex$(codePoint And &HFF), 2)
    End Select
  Next index

  UrlEncode = encoded
End Function

Private Function DecodeUtf8Response(ByVal responseBody As Variant) As String
  Dim stream As Object
  Dim streamOpened As Boolean

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
  If streamOpened Then
    stream.Close
  End If
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
