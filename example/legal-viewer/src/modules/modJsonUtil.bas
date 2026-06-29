Attribute VB_Name = "modJsonUtil"
Option Explicit

Public Function JsonIsLikelyJson(ByVal JsonText As String) As Boolean
    Dim trimmedText As String
    trimmedText = Trim$(StripUtf8Bom(JsonText))

    JsonIsLikelyJson = (Left$(trimmedText, 1) = "{" And Right$(trimmedText, 1) = "}") _
        Or (Left$(trimmedText, 1) = "[" And Right$(trimmedText, 1) = "]")
End Function

Public Function JsonRequireObject(ByVal JsonText As String, ByVal sourceName As String) As String
    Dim trimmedText As String
    trimmedText = Trim$(StripUtf8Bom(JsonText))

    If Left$(trimmedText, 1) <> "{" Or Right$(trimmedText, 1) <> "}" Then
        Err.Raise vbObjectError + 6101, "modJsonUtil.JsonRequireObject", sourceName & " is not a JSON object."
    End If

    JsonRequireObject = trimmedText
End Function

Public Function JsonRequireArray(ByVal JsonText As String, ByVal sourceName As String) As String
    Dim trimmedText As String
    trimmedText = Trim$(StripUtf8Bom(JsonText))

    If Left$(trimmedText, 1) <> "[" Or Right$(trimmedText, 1) <> "]" Then
        Err.Raise vbObjectError + 6102, "modJsonUtil.JsonRequireArray", sourceName & " is not a JSON array."
    End If

    JsonRequireArray = trimmedText
End Function

Public Function JsonExtractString(ByVal JsonText As String, ByVal propertyName As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrHandler

    Dim needle As String
    needle = """" & propertyName & """"

    Dim namePosition As Long
    namePosition = InStr(1, JsonText, needle, vbTextCompare)
    If namePosition = 0 Then
        JsonExtractString = defaultValue
        Exit Function
    End If

    Dim colonPosition As Long
    colonPosition = InStr(namePosition + Len(needle), JsonText, ":", vbBinaryCompare)
    If colonPosition = 0 Then
        JsonExtractString = defaultValue
        Exit Function
    End If

    Dim valuePosition As Long
    valuePosition = colonPosition + 1
    Do While valuePosition <= Len(JsonText) And InStr(1, " " & vbTab & vbCr & vbLf, Mid$(JsonText, valuePosition, 1), vbBinaryCompare) > 0
        valuePosition = valuePosition + 1
    Loop

    If Mid$(JsonText, valuePosition, 1) <> """" Then
        JsonExtractString = defaultValue
        Exit Function
    End If

    JsonExtractString = ReadJsonString(JsonText, valuePosition)
    Exit Function

ErrHandler:
    JsonExtractString = defaultValue
End Function

Public Function JsonEscape(ByVal value As String) As String
    Dim result As String
    Dim index As Long

    For index = 1 To Len(value)
        Dim character As String
        character = Mid$(value, index, 1)

        Select Case character
            Case """"
                result = result & "\" & """"
            Case "\"
                result = result & "\\"
            Case vbCr
                result = result & "\r"
            Case vbLf
                result = result & "\n"
            Case vbTab
                result = result & "\t"
            Case Else
                result = result & character
        End Select
    Next index

    JsonEscape = result
End Function

Public Function ParseJsonObject(ByVal JsonText As String) As Object
    Set ParseJsonObject = JsonConverter.ParseJson(JsonRequireObject(JsonText, "JSON"))
End Function

Public Function ParseJsonArray(ByVal JsonText As String) As Object
    Set ParseJsonArray = JsonConverter.ParseJson(JsonRequireArray(JsonText, "JSON"))
End Function

Public Function JsonHasKey(ByVal jsonObject As Object, ByVal propertyName As String) As Boolean
    On Error GoTo ErrHandler

    JsonHasKey = CBool(jsonObject.Exists(propertyName))
    Exit Function

ErrHandler:
    JsonHasKey = False
End Function

Public Function JsonObjectProperty(ByVal jsonObject As Object, ByVal propertyName As String) As Object
    On Error GoTo ErrHandler

    If jsonObject Is Nothing Then
        Exit Function
    End If
    If Not JsonHasKey(jsonObject, propertyName) Then
        Exit Function
    End If
    If IsObject(jsonObject(propertyName)) Then
        Set JsonObjectProperty = jsonObject(propertyName)
    End If
    Exit Function

ErrHandler:
    Set JsonObjectProperty = Nothing
End Function

Public Function JsonTextProperty(ByVal jsonObject As Object, ByVal propertyName As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrHandler

    If jsonObject Is Nothing Then
        JsonTextProperty = defaultValue
        Exit Function
    End If
    If Not JsonHasKey(jsonObject, propertyName) Then
        JsonTextProperty = defaultValue
        Exit Function
    End If

    Dim value As Variant
    value = jsonObject(propertyName)
    If IsNull(value) Or IsEmpty(value) Then
        JsonTextProperty = defaultValue
    Else
        JsonTextProperty = CStr(value)
    End If
    Exit Function

ErrHandler:
    JsonTextProperty = defaultValue
End Function

Public Function StripUtf8Bom(ByVal value As String) As String
    If Len(value) > 0 And AscW(Left$(value, 1)) = &HFEFF Then
        StripUtf8Bom = Mid$(value, 2)
    Else
        StripUtf8Bom = value
    End If
End Function

Private Function ReadJsonString(ByVal JsonText As String, ByVal quotePosition As Long) As String
    Dim result As String
    Dim index As Long
    Dim escaping As Boolean

    For index = quotePosition + 1 To Len(JsonText)
        Dim character As String
        character = Mid$(JsonText, index, 1)

        If escaping Then
            Select Case character
                Case """"
                    result = result & """"
                Case "\"
                    result = result & "\"
                Case "/"
                    result = result & "/"
                Case "b"
                    result = result & Chr$(8)
                Case "f"
                    result = result & Chr$(12)
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case Else
                    result = result & character
            End Select
            escaping = False
        ElseIf character = "\" Then
            escaping = True
        ElseIf character = """" Then
            Exit For
        Else
            result = result & character
        End If
    Next index

    ReadJsonString = result
End Function
