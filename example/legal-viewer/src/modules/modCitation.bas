Attribute VB_Name = "modCitation"
Option Explicit

Private Const BODY_UNIT_COL_LAW_TITLE As Long = 5
Private Const BODY_UNIT_COL_LAW_NUM As Long = 6
Private Const BODY_UNIT_COL_UNIT_KIND As Long = 7
Private Const BODY_UNIT_COL_ARTICLE_TITLE As Long = 8
Private Const BODY_UNIT_COL_PARAGRAPH_NUM As Long = 10
Private Const BODY_UNIT_COL_ITEM_TITLE As Long = 11
Private Const BODY_UNIT_COL_COUNT As Long = 13

Private Const FORMAT_SIMPLE As String = "シンプル引用"
Private Const FORMAT_WITH_SOURCE As String = "出典情報つき引用"
Private Const FORMAT_MARKDOWN As String = "Markdown引用"
Private Const FORMAT_TABLE_MARKDOWN As String = "表引用（Markdownテーブル）"
Private Const FORMAT_TABLE_ASCII As String = "表引用（ASCII罫線テーブル）"
Private Const FORMAT_EXCEL_TSV As String = "Excel貼付用タブ区切り"

Public Function CitationFormatSimple() As String
    CitationFormatSimple = FORMAT_SIMPLE
End Function

Public Function CitationFormatWithSource() As String
    CitationFormatWithSource = FORMAT_WITH_SOURCE
End Function

Public Function CitationFormatMarkdown() As String
    CitationFormatMarkdown = FORMAT_MARKDOWN
End Function

Public Function CitationFormatTableMarkdown() As String
    CitationFormatTableMarkdown = FORMAT_TABLE_MARKDOWN
End Function

Public Function CitationFormatTableAscii() As String
    CitationFormatTableAscii = FORMAT_TABLE_ASCII
End Function

Public Function CitationFormatExcelTsv() As String
    CitationFormatExcelTsv = FORMAT_EXCEL_TSV
End Function

Public Sub ConfigureCitationFormatComboBox(ByVal targetComboBox As Object, Optional ByVal defaultFormat As String = "")
    If Len(defaultFormat) = 0 Then
        defaultFormat = FORMAT_WITH_SOURCE
    End If

    targetComboBox.Clear
    targetComboBox.AddItem FORMAT_SIMPLE
    targetComboBox.AddItem FORMAT_WITH_SOURCE
    targetComboBox.AddItem FORMAT_MARKDOWN
    targetComboBox.AddItem FORMAT_TABLE_MARKDOWN
    targetComboBox.AddItem FORMAT_TABLE_ASCII
    targetComboBox.AddItem FORMAT_EXCEL_TSV

    If CitationFormatIsSupported(defaultFormat) Then
        targetComboBox.value = defaultFormat
    Else
        targetComboBox.value = FORMAT_WITH_SOURCE
    End If
End Sub

Public Function CitationTextFromBodyUnitRow(ByVal bodyUnitStoreRow As Long, ByVal citationFormat As String, Optional ByVal tableDisplayMode As String = "") As String
    On Error GoTo ErrHandler

    citationFormat = NormalizeCitationFormat(citationFormat)
    Dim effectiveTableMode As String
    effectiveTableMode = EffectiveTableDisplayMode(citationFormat, tableDisplayMode)

    Dim bodyText As String
    bodyText = modLawNavigator.BodyUnitPreviewText(bodyUnitStoreRow, effectiveTableMode)
    If Len(Trim$(bodyText)) = 0 Then
        CitationTextFromBodyUnitRow = ""
        Exit Function
    End If

    Dim sourceText As String
    sourceText = CitationSourceText(bodyUnitStoreRow)

    Select Case citationFormat
        Case FORMAT_SIMPLE
            CitationTextFromBodyUnitRow = bodyText
        Case FORMAT_MARKDOWN
            CitationTextFromBodyUnitRow = MarkdownQuoteText(bodyText, sourceText)
        Case FORMAT_EXCEL_TSV
            CitationTextFromBodyUnitRow = ExcelTabDelimitedText(bodyText, sourceText)
        Case FORMAT_TABLE_MARKDOWN, FORMAT_TABLE_ASCII, FORMAT_WITH_SOURCE
            CitationTextFromBodyUnitRow = bodyText & vbCrLf & vbCrLf & "出典: " & sourceText
        Case Else
            CitationTextFromBodyUnitRow = bodyText & vbCrLf & vbCrLf & "出典: " & sourceText
    End Select
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Citation", "modCitation.CitationTextFromBodyUnitRow", Err.description, "", "", CStr(bodyUnitStoreRow)
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function CitationSmoke() As String
    On Error GoTo ErrHandler

    modLawNavigator.PrepareBodyNavigatorSmokeData

    Dim sourcedText As String
    sourcedText = CitationTextFromBodyUnitRow(3, FORMAT_WITH_SOURCE, modLawNavigator.TableDisplayModeMarkdown())
    If InStr(1, sourcedText, "受給権は差し押えることができない。", vbBinaryCompare) = 0 _
        Or InStr(1, sourcedText, "出典: テスト法（令和七年法律第一号） 第十一条 第1項", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7201, "modCitation.CitationSmoke", "Sourced citation text mismatch."
    End If

    Dim markdownText As String
    markdownText = CitationTextFromBodyUnitRow(7, FORMAT_TABLE_MARKDOWN)
    If InStr(1, markdownText, "| 項目 | 値 |", vbBinaryCompare) = 0 _
        Or InStr(1, markdownText, "+------+-----+", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7202, "modCitation.CitationSmoke", "Markdown table citation mismatch."
    End If

    Dim asciiText As String
    asciiText = CitationTextFromBodyUnitRow(7, FORMAT_TABLE_ASCII)
    If InStr(1, asciiText, "+------+-----+", vbBinaryCompare) = 0 _
        Or InStr(1, asciiText, "| 項目 | 値 |", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7203, "modCitation.CitationSmoke", "ASCII table citation mismatch."
    End If

    Dim tsvText As String
    tsvText = CitationTextFromBodyUnitRow(7, FORMAT_EXCEL_TSV)
    If InStr(1, tsvText, "項目" & vbTab & "値", vbBinaryCompare) = 0 _
        Or InStr(1, tsvText, "| 項目 | 値 |", vbBinaryCompare) > 0 _
        Or InStr(1, tsvText, "+------+-----+", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7204, "modCitation.CitationSmoke", "Excel TSV citation mismatch."
    End If

    CitationSmoke = "ok:" & CStr(Len(sourcedText) + Len(markdownText) + Len(asciiText) + Len(tsvText))
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Citation", "modCitation.CitationSmoke", Err.description, "", "", ""
    CitationSmoke = "error:" & Err.description
End Function

Private Function CitationFormatIsSupported(ByVal citationFormat As String) As Boolean
    citationFormat = NormalizeCitationFormat(citationFormat)
    CitationFormatIsSupported = (Len(citationFormat) > 0)
End Function

Private Function NormalizeCitationFormat(ByVal citationFormat As String) As String
    Select Case Trim$(citationFormat)
        Case FORMAT_SIMPLE, FORMAT_WITH_SOURCE, FORMAT_MARKDOWN, FORMAT_TABLE_MARKDOWN, FORMAT_TABLE_ASCII, FORMAT_EXCEL_TSV
            NormalizeCitationFormat = Trim$(citationFormat)
        Case Else
            NormalizeCitationFormat = FORMAT_WITH_SOURCE
    End Select
End Function

Private Function EffectiveTableDisplayMode(ByVal citationFormat As String, ByVal tableDisplayMode As String) As String
    If StrComp(citationFormat, FORMAT_TABLE_ASCII, vbBinaryCompare) = 0 Then
        EffectiveTableDisplayMode = modLawNavigator.TableDisplayModeAscii()
    ElseIf StrComp(citationFormat, FORMAT_TABLE_MARKDOWN, vbBinaryCompare) = 0 _
        Or StrComp(citationFormat, FORMAT_EXCEL_TSV, vbBinaryCompare) = 0 Then
        EffectiveTableDisplayMode = modLawNavigator.TableDisplayModeMarkdown()
    ElseIf Len(tableDisplayMode) > 0 Then
        EffectiveTableDisplayMode = tableDisplayMode
    Else
        EffectiveTableDisplayMode = modLawNavigator.TableDisplayModeMarkdown()
    End If
End Function

Private Function ExcelTabDelimitedText(ByVal bodyText As String, ByVal sourceText As String) As String
    Dim tableText As String
    tableText = MarkdownTableToTabDelimited(bodyText)
    If Len(tableText) = 0 Then
        tableText = bodyText
    End If

    If Len(sourceText) > 0 Then
        ExcelTabDelimitedText = tableText & vbCrLf & vbCrLf & "出典" & vbTab & sourceText
    Else
        ExcelTabDelimitedText = tableText
    End If
End Function

Private Function MarkdownTableToTabDelimited(ByVal bodyText As String) As String
    Dim normalizedText As String
    normalizedText = Replace(bodyText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)

    Dim lines As Variant
    lines = Split(normalizedText, vbLf)

    Dim result As String
    Dim index As Long
    For index = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(CStr(lines(index)))
        If IsMarkdownTableRow(lineText) And Not IsMarkdownSeparatorRow(lineText) Then
            result = AppendLine(result, MarkdownTableRowToTabDelimited(lineText))
        End If
    Next index

    MarkdownTableToTabDelimited = result
End Function

Private Function IsMarkdownTableRow(ByVal lineText As String) As Boolean
    IsMarkdownTableRow = (Len(lineText) >= 2 And Left$(lineText, 1) = "|" And Right$(lineText, 1) = "|")
End Function

Private Function IsMarkdownSeparatorRow(ByVal lineText As String) As Boolean
    If Not IsMarkdownTableRow(lineText) Then
        IsMarkdownSeparatorRow = False
        Exit Function
    End If

    Dim innerText As String
    innerText = Mid$(lineText, 2, Len(lineText) - 2)
    innerText = Replace(innerText, "|", "")
    innerText = Replace(innerText, "-", "")
    innerText = Replace(innerText, ":", "")
    innerText = Replace(innerText, " ", "")
    IsMarkdownSeparatorRow = (Len(innerText) = 0)
End Function

Private Function MarkdownTableRowToTabDelimited(ByVal lineText As String) As String
    Dim innerText As String
    innerText = Mid$(lineText, 2, Len(lineText) - 2)
    innerText = Replace(innerText, "\|", ChrW$(30))

    Dim cells As Variant
    cells = Split(innerText, "|")

    Dim result As String
    Dim index As Long
    For index = LBound(cells) To UBound(cells)
        If Len(result) > 0 Then
            result = result & vbTab
        End If
        result = result & NormalizeMarkdownCellText(CStr(cells(index)))
    Next index

    MarkdownTableRowToTabDelimited = result
End Function

Private Function NormalizeMarkdownCellText(ByVal CellText As String) As String
    CellText = Trim$(CellText)
    CellText = Replace(CellText, ChrW$(30), "|")
    CellText = Replace(CellText, "\\", "\")
    CellText = Replace(CellText, "<br>", " ")
    CellText = Replace(CellText, vbTab, " ")
    NormalizeMarkdownCellText = CellText
End Function

Private Function CitationSourceText(ByVal bodyUnitStoreRow As Long) As String
    If Not modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        CitationSourceText = ""
        Exit Function
    End If

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())
    If bodyUnitStoreRow < 2 Or bodyUnitStoreRow > modSheetStore.LastUsedRow(bodySheet) Then
        CitationSourceText = ""
        Exit Function
    End If

    Dim bodyValues As Variant
    bodyValues = bodySheet.Range(bodySheet.cells(1, 1), bodySheet.cells(modSheetStore.LastUsedRow(bodySheet), BODY_UNIT_COL_COUNT)).Value2

    Dim result As String
    result = BodyValueText(bodyValues, bodyUnitStoreRow, BODY_UNIT_COL_LAW_TITLE)

    Dim lawNum As String
    lawNum = BodyValueText(bodyValues, bodyUnitStoreRow, BODY_UNIT_COL_LAW_NUM)
    If Len(lawNum) > 0 Then
        result = result & "（" & lawNum & "）"
    End If

    Dim heading As String
    heading = CitationUnitHeading(bodyValues, bodyUnitStoreRow)
    If Len(heading) > 0 Then
        result = result & " " & heading
    End If

    CitationSourceText = result
End Function

Private Function CitationUnitHeading(ByRef bodyValues As Variant, ByVal rowIndex As Long) As String
    Dim unitKind As String
    unitKind = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_UNIT_KIND)

    Dim result As String
    result = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_ARTICLE_TITLE)

    If StrComp(unitKind, "Paragraph", vbBinaryCompare) = 0 Then
        Dim paragraphNum As String
        paragraphNum = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_PARAGRAPH_NUM)
        If Len(paragraphNum) > 0 Then
            result = AppendSpace(result, "第" & paragraphNum & "項")
        End If
    End If

    Dim itemTitle As String
    itemTitle = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_ITEM_TITLE)
    If Len(itemTitle) > 0 Then
        result = AppendSpace(result, itemTitle)
    End If

    CitationUnitHeading = result
End Function

Private Function MarkdownQuoteText(ByVal bodyText As String, ByVal sourceText As String) As String
    Dim normalizedText As String
    normalizedText = Replace(bodyText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)

    Dim lines As Variant
    lines = Split(normalizedText, vbLf)

    Dim result As String
    Dim index As Long
    For index = LBound(lines) To UBound(lines)
        result = AppendLine(result, "> " & CStr(lines(index)))
    Next index
    result = AppendLine(result, "")
    result = AppendLine(result, "> 出典: " & sourceText)

    MarkdownQuoteText = result
End Function

Private Function AppendSpace(ByVal baseText As String, ByVal partText As String) As String
    If Len(baseText) = 0 Then
        AppendSpace = partText
    ElseIf Len(partText) = 0 Then
        AppendSpace = baseText
    Else
        AppendSpace = baseText & " " & partText
    End If
End Function

Private Function AppendLine(ByVal baseText As String, ByVal lineText As String) As String
    If Len(baseText) = 0 Then
        AppendLine = lineText
    Else
        AppendLine = baseText & vbCrLf & lineText
    End If
End Function

Private Function BodyValueText(ByRef bodyValues As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If IsError(bodyValues(rowIndex, columnIndex)) Or IsNull(bodyValues(rowIndex, columnIndex)) Or IsEmpty(bodyValues(rowIndex, columnIndex)) Then
        BodyValueText = ""
    Else
        BodyValueText = CStr(bodyValues(rowIndex, columnIndex))
    End If
End Function
