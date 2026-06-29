Attribute VB_Name = "modApiClient"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const DEFAULT_API_BASE_URL As String = "https://laws.e-gov.go.jp/api/2"
Private Const HTTP_METHOD_GET As String = "GET"
Private Const AD_TYPE_BINARY As Long = 1
Private Const AD_TYPE_TEXT As Long = 2
Private Const AD_READ_ALL As Long = -1
Private m_LastApiCallAt As Double

Public Function ApiGetLaws( _
    Optional ByVal limit As Long = 0, _
    Optional ByVal offset As Long = 0, _
    Optional ByVal lawTitle As String = "", _
    Optional ByVal lawNum As String = "", _
    Optional ByVal lawId As String = "") As String

    Dim path As String
    path = "/laws"
    If limit > 0 Then
        path = AppendQueryParameter(path, "limit", CStr(limit))
    End If
    If offset > 0 Then
        path = AppendQueryParameter(path, "offset", CStr(offset))
    End If
    If Len(Trim$(lawTitle)) > 0 Then
        path = AppendQueryParameter(path, "law_title", Trim$(lawTitle))
    End If
    If Len(Trim$(lawNum)) > 0 Then
        path = AppendQueryParameter(path, "law_num", Trim$(lawNum))
    End If
    If Len(Trim$(lawId)) > 0 Then
        path = AppendQueryParameter(path, "law_id", Trim$(lawId))
    End If

    ApiGetLaws = ApiGet(path, "法令名一覧取得", "", "", limit)
End Function

Public Function ApiGetLawById(ByVal lawId As String) As String
    lawId = Trim$(lawId)
    If Len(lawId) = 0 Then
        Err.Raise vbObjectError + 6002, "modApiClient.ApiGetLawById", "LawId is required."
    End If

    ApiGetLawById = ApiGet(AppendQueryParameter("/laws", "law_id", lawId), "法令名一覧取得", lawId, "", 1)
End Function

Public Function ApiGetLawRevisions(ByVal lawId As String) As String
    lawId = Trim$(lawId)
    If Len(lawId) = 0 Then
        Err.Raise vbObjectError + 6003, "modApiClient.ApiGetLawRevisions", "LawId is required."
    End If

    ApiGetLawRevisions = ApiGet("/law_revisions/" & lawId, "施行日候補取得", lawId, "", -1)
End Function

Public Function ApiGetLawData(ByVal lawIdentifier As String, Optional ByVal lawId As String = "", Optional ByVal enforcementDate As String = "") As String
    lawIdentifier = Trim$(lawIdentifier)
    If Len(lawIdentifier) = 0 Then
        Err.Raise vbObjectError + 6004, "modApiClient.ApiGetLawData", "Law identifier is required."
    End If

    ApiGetLawData = ApiGet("/law_data/" & lawIdentifier, "法令本文取得", lawId, enforcementDate, -1)
End Function

Public Function ApiGet( _
    ByVal pathOrUrl As String, _
    Optional ByVal processKind As String = "API通信", _
    Optional ByVal lawId As String = "", _
    Optional ByVal enforcementDate As String = "", _
    Optional ByVal expectedItemCount As Long = -1) As String

    On Error GoTo ErrHandler

    Dim endpointUrl As String
    endpointUrl = BuildApiUrl(pathOrUrl)

    Dim retryLimit As Long
    retryLimit = modSettings.GetSettingLong("ApiRetryCount", 1)

    Dim attempt As Long
    Dim statusCode As Long
    Dim responseText As String
    Dim elapsedMs As Long
    Dim lastErrorMessage As String

    For attempt = 0 To retryLimit
        WaitForRateLimit

        Dim startedAt As Double
        startedAt = Timer

        Dim request As Object
        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
        request.Open HTTP_METHOD_GET, endpointUrl, False
        request.SetTimeouts _
            modSettings.GetSettingLong("TimeoutResolveMs", 10000), _
            modSettings.GetSettingLong("TimeoutConnectMs", 10000), _
            modSettings.GetSettingLong("TimeoutSendMs", 10000), _
            modSettings.GetSettingLong("TimeoutReceiveMs", 30000)
        request.SetRequestHeader "Accept", "application/json"
        request.SetRequestHeader "Accept-Charset", "utf-8"
        request.SetRequestHeader "User-Agent", "xlflow-law-viewer/0.1"
        request.Send

        statusCode = CLng(request.Status)
        responseText = ResponseBodyAsUtf8Text(request)
        elapsedMs = ElapsedMilliseconds(startedAt)

        If statusCode >= 200 And statusCode < 300 Then
            modLogger.LogApiCall processKind, endpointUrl, statusCode, True, lawId, enforcementDate, expectedItemCount, elapsedMs, ""
            ApiGet = responseText
            Exit Function
        End If

        lastErrorMessage = Left$(responseText, 500)
        modLogger.LogApiCall processKind, endpointUrl, statusCode, False, lawId, enforcementDate, expectedItemCount, elapsedMs, lastErrorMessage

        If Not ShouldRetryStatus(statusCode) Then
            Exit For
        End If
    Next attempt

    Err.Raise vbObjectError + 6001, "modApiClient.ApiGet", "API request failed. HTTP status=" & CStr(statusCode) & " " & lastErrorMessage
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "API", "modApiClient.ApiGet", Err.description, lawId, enforcementDate, pathOrUrl
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function BuildApiUrl(ByVal pathOrUrl As String) As String
    If IsAbsoluteUrl(pathOrUrl) Then
        BuildApiUrl = pathOrUrl
        Exit Function
    End If

    Dim baseUrl As String
    baseUrl = modSettings.GetSettingText("ApiBaseUrl", DEFAULT_API_BASE_URL)
    Do While Right$(baseUrl, 1) = "/"
        baseUrl = Left$(baseUrl, Len(baseUrl) - 1)
    Loop

    If Left$(pathOrUrl, 1) = "/" Then
        BuildApiUrl = baseUrl & pathOrUrl
    Else
        BuildApiUrl = baseUrl & "/" & pathOrUrl
    End If
End Function

Public Function AppendQueryParameter(ByVal queryString As String, ByVal parameterName As String, ByVal parameterValue As String) As String
    Dim separator As String
    If Len(queryString) = 0 Or InStr(1, queryString, "?", vbBinaryCompare) = 0 Then
        separator = "?"
    Else
        separator = "&"
    End If

    AppendQueryParameter = queryString & separator & UrlEncodeAscii(parameterName) & "=" & UrlEncodeAscii(parameterValue)
End Function

Private Function IsAbsoluteUrl(ByVal value As String) As Boolean
    IsAbsoluteUrl = (LCase$(Left$(value, 7)) = "http://") Or (LCase$(Left$(value, 8)) = "https://")
End Function

Private Function ShouldRetryStatus(ByVal statusCode As Long) As Boolean
    ShouldRetryStatus = (statusCode = 408 Or statusCode = 429 Or statusCode >= 500)
End Function

Private Sub WaitForRateLimit()
    Dim minimumIntervalSeconds As Double
    minimumIntervalSeconds = modSettings.GetSettingDouble("ApiMinIntervalSeconds", 0.5)

    If minimumIntervalSeconds <= 0 Then
        m_LastApiCallAt = Timer
        Exit Sub
    End If

    If m_LastApiCallAt > 0 Then
        Dim elapsedSeconds As Double
        elapsedSeconds = Timer - m_LastApiCallAt
        If elapsedSeconds < 0 Then
            elapsedSeconds = elapsedSeconds + 86400#
        End If

        If elapsedSeconds < minimumIntervalSeconds Then
            Sleep CLng((minimumIntervalSeconds - elapsedSeconds) * 1000#)
        End If
    End If

    m_LastApiCallAt = Timer
End Sub

Private Function ElapsedMilliseconds(ByVal startedAt As Double) As Long
    Dim elapsedSeconds As Double
    elapsedSeconds = Timer - startedAt
    If elapsedSeconds < 0 Then
        elapsedSeconds = elapsedSeconds + 86400#
    End If
    ElapsedMilliseconds = CLng(elapsedSeconds * 1000#)
End Function

Private Function ResponseBodyAsUtf8Text(ByVal request As Object) As String
    Dim bodyBytes As Variant
    bodyBytes = request.ResponseBody

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = AD_TYPE_BINARY
    stream.Open
    stream.Write bodyBytes
    stream.position = 0
    stream.Type = AD_TYPE_TEXT
    stream.Charset = "utf-8"
    ResponseBodyAsUtf8Text = stream.ReadText(AD_READ_ALL)
    stream.Close
End Function

Private Function UrlEncodeAscii(ByVal value As String) As String
    Dim valueBytes As Variant
    valueBytes = Utf8Bytes(value)

    Dim startIndex As Long
    startIndex = LBound(valueBytes)
    If UBound(valueBytes) - startIndex + 1 >= 3 Then
        If CLng(valueBytes(startIndex)) = &HEF _
            And CLng(valueBytes(startIndex + 1)) = &HBB _
            And CLng(valueBytes(startIndex + 2)) = &HBF Then
            startIndex = startIndex + 3
        End If
    End If

    Dim result As String
    Dim index As Long
    For index = startIndex To UBound(valueBytes)
        Dim byteValue As Long
        byteValue = CLng(valueBytes(index))

        Dim character As String
        If byteValue >= 0 And byteValue <= 127 Then
            character = Chr$(byteValue)
        Else
            character = ""
        End If

        If (byteValue >= 48 And byteValue <= 57) _
            Or (byteValue >= 65 And byteValue <= 90) _
            Or (byteValue >= 97 And byteValue <= 122) _
            Or character = "-" Or character = "_" Or character = "." Or character = "~" Then
            result = result & character
        Else
            result = result & "%" & Right$("0" & Hex$(byteValue), 2)
        End If
    Next index

    UrlEncodeAscii = result
End Function

Private Function Utf8Bytes(ByVal value As String) As Variant
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = AD_TYPE_TEXT
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.position = 0
    stream.Type = AD_TYPE_BINARY
    Utf8Bytes = stream.Read(AD_READ_ALL)
    stream.Close
End Function
