Attribute VB_Name = "modTableFormatter"
Option Explicit

Public Function MarkdownTableFromArray(ByRef tableValues As Variant, Optional ByVal hasHeader As Boolean = True) As String
    On Error GoTo ErrHandler

    Dim rowCount As Long
    Dim columnCount As Long
    GetTableBounds tableValues, rowCount, columnCount
    If rowCount = 0 Or columnCount = 0 Then
        MarkdownTableFromArray = ""
        Exit Function
    End If

    Dim result As String
    Dim rowIndex As Long
    For rowIndex = 1 To rowCount
        result = AppendLine(result, MarkdownRow(tableValues, rowIndex, columnCount))
        If rowIndex = 1 And hasHeader Then
            result = AppendLine(result, MarkdownSeparatorRow(columnCount))
        End If
    Next rowIndex

    MarkdownTableFromArray = result
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "TableFormatter", "modTableFormatter.MarkdownTableFromArray", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function AsciiTableFromArray(ByRef tableValues As Variant, Optional ByVal hasHeader As Boolean = True) As String
    On Error GoTo ErrHandler

    Dim rowCount As Long
    Dim columnCount As Long
    GetTableBounds tableValues, rowCount, columnCount
    If rowCount = 0 Or columnCount = 0 Then
        AsciiTableFromArray = ""
        Exit Function
    End If

    Dim columnWidths() As Long
    columnWidths = TableColumnWidths(tableValues, rowCount, columnCount)

    Dim borderLine As String
    borderLine = AsciiBorderLine(columnWidths, columnCount)

    Dim result As String
    result = borderLine

    Dim rowIndex As Long
    For rowIndex = 1 To rowCount
        result = AppendLine(result, AsciiRow(tableValues, rowIndex, columnWidths, columnCount))
        If rowIndex = 1 And hasHeader Then
            result = AppendLine(result, borderLine)
        End If
    Next rowIndex
    result = AppendLine(result, borderLine)

    AsciiTableFromArray = result
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "TableFormatter", "modTableFormatter.AsciiTableFromArray", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function TextDisplayWidth(ByVal value As String) As Long
    Dim index As Long
    Dim codeUnit As Long
    Dim nextCodeUnit As Long
    Dim codePoint As Long
    Dim widthValue As Long

    index = 1
    Do While index <= Len(value)
        codeUnit = Utf16CodeUnitAt(value, index)
        If IsHighSurrogate(codeUnit) And index < Len(value) Then
            nextCodeUnit = Utf16CodeUnitAt(value, index + 1)
            If IsLowSurrogate(nextCodeUnit) Then
                codePoint = 65536 + ((codeUnit - 55296) * 1024) + (nextCodeUnit - 56320)
                index = index + 2
            Else
                codePoint = codeUnit
                index = index + 1
            End If
        Else
            codePoint = codeUnit
            index = index + 1
        End If
        widthValue = widthValue + CharacterDisplayWidth(codePoint)
    Loop

    TextDisplayWidth = widthValue
End Function

Public Function TableFormatterSmoke() As String
    On Error GoTo ErrHandler

    Dim values(1 To 3, 1 To 2) As String
    values(1, 1) = "項目"
    values(1, 2) = "値"
    values(2, 1) = "労働時間"
    values(2, 2) = "8"
    values(3, 1) = "ABC"
    values(3, 2) = "日本語"

    Dim markdownText As String
    markdownText = MarkdownTableFromArray(values, True)
    If InStr(1, markdownText, "| 項目 | 値 |", vbBinaryCompare) = 0 _
        Or InStr(1, markdownText, "| --- | --- |", vbBinaryCompare) = 0 _
        Or InStr(1, markdownText, "| ABC | 日本語 |", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7101, "modTableFormatter.TableFormatterSmoke", "Markdown table output mismatch."
    End If

    Dim asciiText As String
    asciiText = AsciiTableFromArray(values, True)
    If InStr(1, asciiText, "+----------+--------+", vbBinaryCompare) = 0 _
        Or InStr(1, asciiText, "| 項目     | 値     |", vbBinaryCompare) = 0 _
        Or InStr(1, asciiText, "| 労働時間 | 8      |", vbBinaryCompare) = 0 _
        Or InStr(1, asciiText, "| ABC      | 日本語 |", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7102, "modTableFormatter.TableFormatterSmoke", "ASCII table output mismatch: " & Replace(asciiText, vbCrLf, "\n")
    End If

    If TextDisplayWidth("ABC") <> 3 _
        Or TextDisplayWidth("項目") <> 4 _
        Or TextDisplayWidth("労働時間") <> 8 _
        Or TextDisplayWidth("ｱｲｳ") <> 3 _
        Or TextDisplayWidth("１２") <> 4 _
        Or TextDisplayWidth(ChrW$(55360) & ChrW$(56331)) <> 2 _
        Or TextDisplayWidth("葛" & ChrW$(56128) & ChrW$(56576)) <> 2 Then
        Err.Raise vbObjectError + 7103, "modTableFormatter.TableFormatterSmoke", "Display width calculation mismatch."
    End If

    TableFormatterSmoke = "ok:" & CStr(Len(asciiText))
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "TableFormatter", "modTableFormatter.TableFormatterSmoke", Err.description, "", "", ""
    TableFormatterSmoke = "error:" & Err.description
End Function

Private Sub GetTableBounds(ByRef tableValues As Variant, ByRef rowCount As Long, ByRef columnCount As Long)
    If Not IsArray(tableValues) Then
        Err.Raise vbObjectError + 7104, "modTableFormatter.GetTableBounds", "Table values must be a two-dimensional array."
    End If

    On Error GoTo ErrHandler
    rowCount = UBound(tableValues, 1) - LBound(tableValues, 1) + 1
    columnCount = UBound(tableValues, 2) - LBound(tableValues, 2) + 1
    Exit Sub

ErrHandler:
    Err.Raise vbObjectError + 7105, "modTableFormatter.GetTableBounds", "Table values must be a two-dimensional array."
End Sub

Private Function MarkdownRow(ByRef tableValues As Variant, ByVal rowOffset As Long, ByVal columnCount As Long) As String
    Dim rowIndex As Long
    rowIndex = LBound(tableValues, 1) + rowOffset - 1

    Dim result As String
    result = "|"

    Dim columnOffset As Long
    Dim columnIndex As Long
    For columnOffset = 1 To columnCount
        columnIndex = LBound(tableValues, 2) + columnOffset - 1
        result = result & " " & MarkdownCellText(CellText(tableValues(rowIndex, columnIndex))) & " |"
    Next columnOffset

    MarkdownRow = result
End Function

Private Function MarkdownSeparatorRow(ByVal columnCount As Long) As String
    Dim result As String
    result = "|"

    Dim columnIndex As Long
    For columnIndex = 1 To columnCount
        result = result & " --- |"
    Next columnIndex

    MarkdownSeparatorRow = result
End Function

Private Function TableColumnWidths(ByRef tableValues As Variant, ByVal rowCount As Long, ByVal columnCount As Long) As Long()
    Dim result() As Long
    ReDim result(1 To columnCount)

    Dim rowOffset As Long
    Dim columnOffset As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim widthValue As Long
    For rowOffset = 1 To rowCount
        rowIndex = LBound(tableValues, 1) + rowOffset - 1
        For columnOffset = 1 To columnCount
            columnIndex = LBound(tableValues, 2) + columnOffset - 1
            widthValue = TextDisplayWidth(CellText(tableValues(rowIndex, columnIndex)))
            If widthValue > result(columnOffset) Then
                result(columnOffset) = widthValue
            End If
        Next columnOffset
    Next rowOffset

    For columnOffset = 1 To columnCount
        If result(columnOffset) < 1 Then
            result(columnOffset) = 1
        End If
    Next columnOffset

    TableColumnWidths = result
End Function

Private Function AsciiBorderLine(ByRef columnWidths() As Long, ByVal columnCount As Long) As String
    Dim result As String
    result = "+"

    Dim columnIndex As Long
    For columnIndex = 1 To columnCount
        result = result & String$(columnWidths(columnIndex) + 2, "-") & "+"
    Next columnIndex

    AsciiBorderLine = result
End Function

Private Function AsciiRow(ByRef tableValues As Variant, ByVal rowOffset As Long, ByRef columnWidths() As Long, ByVal columnCount As Long) As String
    Dim rowIndex As Long
    rowIndex = LBound(tableValues, 1) + rowOffset - 1

    Dim result As String
    result = "|"

    Dim columnOffset As Long
    Dim columnIndex As Long
    Dim valueText As String
    For columnOffset = 1 To columnCount
        columnIndex = LBound(tableValues, 2) + columnOffset - 1
        valueText = CellText(tableValues(rowIndex, columnIndex))
        result = result & " " & PadDisplayRight(valueText, columnWidths(columnOffset)) & " |"
    Next columnOffset

    AsciiRow = result
End Function

Private Function PadDisplayRight(ByVal value As String, ByVal targetWidth As Long) As String
    Dim padCount As Long
    padCount = targetWidth - TextDisplayWidth(value)
    If padCount < 0 Then
        padCount = 0
    End If

    PadDisplayRight = value & String$(padCount, " ")
End Function

Private Function CharacterDisplayWidth(ByVal codePoint As Long) As Long
    If IsZeroWidthCodePoint(codePoint) Then
        CharacterDisplayWidth = 0
    ElseIf IsWideCodePoint(codePoint) Then
        CharacterDisplayWidth = 2
    Else
        CharacterDisplayWidth = 1
    End If
End Function

Private Function Utf16CodeUnitAt(ByVal value As String, ByVal index As Long) As Long
    Dim result As Long
    result = AscW(Mid$(value, index, 1))
    If result < 0 Then
        result = result + 65536
    End If
    Utf16CodeUnitAt = result
End Function

Private Function IsHighSurrogate(ByVal codeUnit As Long) As Boolean
    IsHighSurrogate = (codeUnit >= 55296 And codeUnit <= 56319)
End Function

Private Function IsLowSurrogate(ByVal codeUnit As Long) As Boolean
    IsLowSurrogate = (codeUnit >= 56320 And codeUnit <= 57343)
End Function

Private Function IsZeroWidthCodePoint(ByVal codePoint As Long) As Boolean
    IsZeroWidthCodePoint = ((codePoint >= 768 And codePoint <= 879) _
        Or (codePoint >= 6832 And codePoint <= 6911) _
        Or (codePoint >= 7616 And codePoint <= 7679) _
        Or (codePoint >= 65024 And codePoint <= 65039) _
        Or (codePoint >= 917760 And codePoint <= 917999) _
        Or (codePoint >= 12441 And codePoint <= 12442) _
        Or codePoint = 8203 _
        Or codePoint = 8204 _
        Or codePoint = 8205)
End Function

Private Function IsWideCodePoint(ByVal codePoint As Long) As Boolean
    IsWideCodePoint = ((codePoint >= 4352 And codePoint <= 4447) _
        Or codePoint = 9001 _
        Or codePoint = 9002 _
        Or (codePoint >= 11904 And codePoint <= 42191) _
        Or (codePoint >= 131072 And codePoint <= 196605) _
        Or (codePoint >= 196608 And codePoint <= 262141) _
        Or (codePoint >= 44032 And codePoint <= 55203) _
        Or (codePoint >= 63744 And codePoint <= 64255) _
        Or (codePoint >= 65040 And codePoint <= 65049) _
        Or (codePoint >= 65072 And codePoint <= 65135) _
        Or (codePoint >= 65281 And codePoint <= 65376) _
        Or (codePoint >= 65504 And codePoint <= 65510))
End Function

Private Function MarkdownCellText(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, "|", "\|")
    value = Replace(value, vbCrLf, "<br>")
    value = Replace(value, vbCr, "<br>")
    value = Replace(value, vbLf, "<br>")
    MarkdownCellText = value
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function

Private Function AppendLine(ByVal baseText As String, ByVal partText As String) As String
    If Len(baseText) = 0 Then
        AppendLine = partText
    Else
        AppendLine = baseText & vbCrLf & partText
    End If
End Function
