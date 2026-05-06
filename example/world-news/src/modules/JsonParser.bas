Attribute VB_Name = "JsonParser"
Option Explicit

Private m_jsonText As String
Private m_index As Long

Public Function ParseJson(ByVal jsonText As String) As Object
  m_jsonText = jsonText
  m_index = 1

  SkipWhitespace
  If PeekChar() <> "{" Then
    RaiseParseError "Top-level JSON value must be an object."
  End If

  Set ParseJson = ParseObjectNode()

  SkipWhitespace
  If m_index <= Len(m_jsonText) Then
    RaiseParseError "Unexpected trailing characters."
  End If
End Function

Private Function ParseObjectNode() As Object
  Dim dict As Object
  Dim key As String

  Set dict = CreateObject("Scripting.Dictionary")
  ExpectChar "{"
  SkipWhitespace

  If PeekChar() = "}" Then
    m_index = m_index + 1
    Set ParseObjectNode = dict
    Exit Function
  End If

  Do
    key = ParseStringLiteral()
    SkipWhitespace
    ExpectChar ":"
    SkipWhitespace
    AddParsedValueToDictionary dict, key
    SkipWhitespace

    Select Case PeekChar()
      Case "}"
        m_index = m_index + 1
        Exit Do
      Case ","
        m_index = m_index + 1
        SkipWhitespace
      Case Else
        RaiseParseError "Expected ',' or '}' in object."
    End Select
  Loop

  Set ParseObjectNode = dict
End Function

Private Function ParseArrayNode() As Collection
  Dim items As Collection

  Set items = New Collection
  ExpectChar "["
  SkipWhitespace

  If PeekChar() = "]" Then
    m_index = m_index + 1
    Set ParseArrayNode = items
    Exit Function
  End If

  Do
    AddParsedValueToCollection items
    SkipWhitespace

    Select Case PeekChar()
      Case "]"
        m_index = m_index + 1
        Exit Do
      Case ","
        m_index = m_index + 1
        SkipWhitespace
      Case Else
        RaiseParseError "Expected ',' or ']' in array."
    End Select
  Loop

  Set ParseArrayNode = items
End Function

Private Sub AddParsedValueToDictionary(ByVal dict As Object, ByVal key As String)
  Select Case PeekChar()
    Case "{"
      dict.Add key, ParseObjectNode()
    Case "["
      dict.Add key, ParseArrayNode()
    Case Else
      dict.Add key, ParseScalarValue()
  End Select
End Sub

Private Sub AddParsedValueToCollection(ByVal items As Collection)
  Select Case PeekChar()
    Case "{"
      items.Add ParseObjectNode()
    Case "["
      items.Add ParseArrayNode()
    Case Else
      items.Add ParseScalarValue()
  End Select
End Sub

Private Function ParseScalarValue() As Variant
  Select Case PeekChar()
    Case """"
      ParseScalarValue = ParseStringLiteral()
    Case "t"
      ParseLiteral "true"
      ParseScalarValue = True
    Case "f"
      ParseLiteral "false"
      ParseScalarValue = False
    Case "n"
      ParseLiteral "null"
      ParseScalarValue = Null
    Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
      ParseScalarValue = ParseNumberLiteral()
    Case Else
      RaiseParseError "Unexpected value token."
  End Select
End Function

Private Function ParseStringLiteral() As String
  Dim result As String
  Dim ch As String
  Dim hexValue As String

  ExpectChar """"

  Do While m_index <= Len(m_jsonText)
    ch = Mid$(m_jsonText, m_index, 1)
    m_index = m_index + 1

    Select Case ch
      Case """"
        ParseStringLiteral = result
        Exit Function
      Case "\"
        If m_index > Len(m_jsonText) Then
          RaiseParseError "Unexpected end of string escape sequence."
        End If

        ch = Mid$(m_jsonText, m_index, 1)
        m_index = m_index + 1

        Select Case ch
          Case """", "\", "/"
            result = result & ch
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
          Case "u"
            hexValue = Mid$(m_jsonText, m_index, 4)
            If Len(hexValue) <> 4 Or Not hexValue Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
              RaiseParseError "Invalid unicode escape sequence."
            End If
            result = result & ChrW$(CLng("&H" & hexValue))
            m_index = m_index + 4
          Case Else
            RaiseParseError "Invalid escape sequence '\\" & ch & "'."
        End Select
      Case Else
        result = result & ch
    End Select
  Loop

  RaiseParseError "Unterminated string literal."
End Function

Private Function ParseNumberLiteral() As Variant
  Dim startIndex As Long
  Dim token As String

  startIndex = m_index

  If PeekChar() = "-" Then
    m_index = m_index + 1
  End If

  ConsumeDigits

  If PeekChar() = "." Then
    m_index = m_index + 1
    ConsumeDigits
  End If

  If PeekChar() = "e" Or PeekChar() = "E" Then
    m_index = m_index + 1
    If PeekChar() = "+" Or PeekChar() = "-" Then
      m_index = m_index + 1
    End If
    ConsumeDigits
  End If

  token = Mid$(m_jsonText, startIndex, m_index - startIndex)

  If InStr(1, token, ".", vbBinaryCompare) > 0 Or InStr(1, token, "e", vbTextCompare) > 0 Then
    ParseNumberLiteral = CDbl(token)
    Exit Function
  End If

  On Error GoTo IntegerOverflow
  ParseNumberLiteral = CLng(token)
  On Error GoTo 0
  Exit Function

IntegerOverflow:
  If Err.Number = 6 Then
    Err.Clear
    ParseNumberLiteral = CDbl(token)
    On Error GoTo 0
    Exit Function
  End If
  RaiseParseError "Invalid number literal '" & token & "'."
End Function

Private Sub ConsumeDigits()
  Dim startIndex As Long

  startIndex = m_index
  Do While m_index <= Len(m_jsonText)
    Select Case Mid$(m_jsonText, m_index, 1)
      Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
        m_index = m_index + 1
      Case Else
        Exit Do
    End Select
  Loop

  If startIndex = m_index Then
    RaiseParseError "Expected digit."
  End If
End Sub

Private Sub ParseLiteral(ByVal literalText As String)
  If Mid$(m_jsonText, m_index, Len(literalText)) <> literalText Then
    RaiseParseError "Expected literal '" & literalText & "'."
  End If

  m_index = m_index + Len(literalText)
End Sub

Private Sub ExpectChar(ByVal expected As String)
  If PeekChar() <> expected Then
    RaiseParseError "Expected '" & expected & "'."
  End If

  m_index = m_index + 1
End Sub

Private Sub SkipWhitespace()
  Do While m_index <= Len(m_jsonText)
    Select Case Mid$(m_jsonText, m_index, 1)
      Case " ", vbTab, vbCr, vbLf
        m_index = m_index + 1
      Case Else
        Exit Do
    End Select
  Loop
End Sub

Private Function PeekChar() As String
  If m_index > Len(m_jsonText) Then
    PeekChar = vbNullString
  Else
    PeekChar = Mid$(m_jsonText, m_index, 1)
  End If
End Function

Private Sub RaiseParseError(ByVal message As String)
  Err.Raise vbObjectError + 2100, "JsonParser.ParseJson", message & " (position " & CStr(m_index) & ")"
End Sub
