Attribute VB_Name = "modCompare"
Option Explicit

Public Function CompareTextBlocks( _
    ByVal leftText As String, _
    ByVal rightText As String, _
    Optional ByVal leftLabel As String = "左", _
    Optional ByVal rightLabel As String = "右") As String

    On Error GoTo ErrHandler

    Dim leftLines As Variant
    Dim rightLines As Variant
    leftLines = SplitLinesForCompare(leftText)
    rightLines = SplitLinesForCompare(rightText)

    Dim leftCount As Long
    Dim rightCount As Long
    leftCount = CountCompareLines(leftLines)
    rightCount = CountCompareLines(rightLines)

    Dim result As String
    result = leftLabel & vbCrLf & rightLabel

    Dim leftIndex As Long
    Dim rightIndex As Long
    leftIndex = 1
    rightIndex = 1

    Do While leftIndex <= leftCount Or rightIndex <= rightCount
        If leftIndex > leftCount Then
            result = AppendLine(result, "+ " & CompareLineAt(rightLines, rightIndex))
            rightIndex = rightIndex + 1
        ElseIf rightIndex > rightCount Then
            result = AppendLine(result, "- " & CompareLineAt(leftLines, leftIndex))
            leftIndex = leftIndex + 1
        ElseIf StrComp(CompareLineAt(leftLines, leftIndex), CompareLineAt(rightLines, rightIndex), vbBinaryCompare) = 0 Then
            result = AppendLine(result, "  " & CompareLineAt(leftLines, leftIndex))
            leftIndex = leftIndex + 1
            rightIndex = rightIndex + 1
        ElseIf leftIndex < leftCount And StrComp(CompareLineAt(leftLines, leftIndex + 1), CompareLineAt(rightLines, rightIndex), vbBinaryCompare) = 0 Then
            result = AppendLine(result, "- " & CompareLineAt(leftLines, leftIndex))
            leftIndex = leftIndex + 1
        ElseIf rightIndex < rightCount And StrComp(CompareLineAt(leftLines, leftIndex), CompareLineAt(rightLines, rightIndex + 1), vbBinaryCompare) = 0 Then
            result = AppendLine(result, "+ " & CompareLineAt(rightLines, rightIndex))
            rightIndex = rightIndex + 1
        Else
            result = AppendCharDiffLines(result, CompareLineAt(leftLines, leftIndex), CompareLineAt(rightLines, rightIndex))
            leftIndex = leftIndex + 1
            rightIndex = rightIndex + 1
        End If
    Loop

    CompareTextBlocks = result
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Compare", "modCompare.CompareTextBlocks", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function CompareBodyUnitRows( _
    ByVal leftBodyUnitRow As Long, _
    ByVal rightBodyUnitRow As Long, _
    Optional ByVal tableDisplayMode As String = "") As String

    On Error GoTo ErrHandler

    If Len(Trim$(tableDisplayMode)) = 0 Then
        tableDisplayMode = modLawNavigator.TableDisplayModeMarkdown()
    End If

    CompareBodyUnitRows = CompareTextBlocks( _
        modLawNavigator.BodyContextPreviewText(leftBodyUnitRow, tableDisplayMode), _
        modLawNavigator.BodyContextPreviewText(rightBodyUnitRow, tableDisplayMode), _
        "左", _
        "右")
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Compare", "modCompare.CompareBodyUnitRows", Err.description, "", "", CStr(leftBodyUnitRow) & ":" & CStr(rightBodyUnitRow)
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function CompareSmoke() As String
    On Error GoTo ErrHandler

    Dim diffText As String
    diffText = CompareTextBlocks("ABCDE", "ABXDE")

    If InStr(1, diffText, "  AB", vbBinaryCompare) = 0 _
        Or InStr(1, diffText, "- C", vbBinaryCompare) = 0 _
        Or InStr(1, diffText, "+ X", vbBinaryCompare) = 0 _
        Or InStr(1, diffText, "  DE", vbBinaryCompare) = 0 _
        Or InStr(1, diffText, "- ABCDE", vbBinaryCompare) > 0 _
        Or InStr(1, diffText, "+ ABXDE", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7401, "modCompare.CompareSmoke", "Diff output mismatch."
    End If

    CompareSmoke = "ok:" & CStr(Len(diffText))
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Compare", "modCompare.CompareSmoke", Err.description, "", "", ""
    CompareSmoke = "error:" & Err.description
End Function

Private Function SplitLinesForCompare(ByVal value As String) As Variant
    value = Replace(value, vbCrLf, vbLf)
    value = Replace(value, vbCr, vbLf)
    SplitLinesForCompare = Split(value, vbLf)
End Function

Private Function CountCompareLines(ByVal lines As Variant) As Long
    On Error GoTo FailSafe
    If Not IsArray(lines) Then Exit Function
    CountCompareLines = UBound(lines) - LBound(lines) + 1
    Exit Function

FailSafe:
    CountCompareLines = 0
End Function

Private Function CompareLineAt(ByVal lines As Variant, ByVal lineNumber As Long) As String
    On Error GoTo FailSafe

    If Not IsArray(lines) Then Exit Function
    If lineNumber < 1 Then Exit Function
    If lineNumber > UBound(lines) - LBound(lines) + 1 Then Exit Function

    CompareLineAt = CStr(lines(LBound(lines) + lineNumber - 1))
    Exit Function

FailSafe:
    CompareLineAt = ""
End Function

Private Function AppendLine(ByVal baseText As String, ByVal partText As String) As String
    If Len(baseText) = 0 Then
        AppendLine = partText
    ElseIf Len(partText) = 0 Then
        AppendLine = baseText
    Else
        AppendLine = baseText & vbCrLf & partText
    End If
End Function

Private Function AppendCharDiffLines(ByVal baseText As String, ByVal leftText As String, ByVal rightText As String) As String
    Dim prefixLength As Long
    Dim suffixLength As Long
    Dim prefixText As String
    Dim leftMiddle As String
    Dim rightMiddle As String
    Dim suffixText As String

    prefixLength = CommonPrefixLength(leftText, rightText)
    suffixLength = CommonSuffixLength(leftText, rightText, prefixLength)

    If prefixLength > 0 Then
        prefixText = Left$(leftText, prefixLength)
        baseText = AppendLine(baseText, "  " & prefixText)
    End If

    leftMiddle = Mid$(leftText, prefixLength + 1, Len(leftText) - prefixLength - suffixLength)
    rightMiddle = Mid$(rightText, prefixLength + 1, Len(rightText) - prefixLength - suffixLength)

    If Len(leftMiddle) > 0 Then
        baseText = AppendLine(baseText, "- " & leftMiddle)
    End If
    If Len(rightMiddle) > 0 Then
        baseText = AppendLine(baseText, "+ " & rightMiddle)
    End If

    If suffixLength > 0 Then
        suffixText = Right$(leftText, suffixLength)
        baseText = AppendLine(baseText, "  " & suffixText)
    End If

    AppendCharDiffLines = baseText
End Function

Private Function CommonPrefixLength(ByVal leftText As String, ByVal rightText As String) As Long
    Dim limitLength As Long
    Dim index As Long

    limitLength = Len(leftText)
    If Len(rightText) < limitLength Then
        limitLength = Len(rightText)
    End If

    For index = 1 To limitLength
        If Mid$(leftText, index, 1) <> Mid$(rightText, index, 1) Then Exit For
    Next index

    CommonPrefixLength = index - 1
End Function

Private Function CommonSuffixLength(ByVal leftText As String, ByVal rightText As String, ByVal prefixLength As Long) As Long
    Dim leftRemaining As Long
    Dim rightRemaining As Long
    Dim limitLength As Long
    Dim index As Long

    leftRemaining = Len(leftText) - prefixLength
    rightRemaining = Len(rightText) - prefixLength
    If leftRemaining < 0 Then leftRemaining = 0
    If rightRemaining < 0 Then rightRemaining = 0
    limitLength = leftRemaining
    If rightRemaining < limitLength Then
        limitLength = rightRemaining
    End If

    For index = 1 To limitLength
        If Mid$(leftText, Len(leftText) - index + 1, 1) <> Mid$(rightText, Len(rightText) - index + 1, 1) Then Exit For
    Next index

    CommonSuffixLength = index - 1
End Function
