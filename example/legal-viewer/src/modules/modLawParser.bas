Attribute VB_Name = "modLawParser"
Option Explicit

Private Const SHEET_BODY_UNITS As String = "_lv_tmp_law_body_units"
Private Const BODY_UNIT_COLUMN_COUNT As Long = 13
Private Const BODY_UNIT_COL_LAW_REVISION_ID As Long = 3
Private Const EXCEL_CELL_TEXT_LIMIT As Long = 30000

Private Type LawBodyMeta
    lawId As String
    lawRevisionId As String
    enforcementDate As String
    lawTitle As String
    lawNum As String
End Type

Public Function BodyUnitSheetName() As String
    BodyUnitSheetName = SHEET_BODY_UNITS
End Function

Public Function BodyUnitCacheKey(ByVal lawRevisionId As String) As String
    BodyUnitCacheKey = SHEET_BODY_UNITS & "|" & Trim$(lawRevisionId)
End Function

Public Function RefreshLawBodyUnitCache(ByVal lawRevisionId As String, Optional ByVal lawId As String = "", Optional ByVal enforcementDate As String = "") As Long
    On Error GoTo ErrHandler

    RefreshLawBodyUnitCache = RefreshLawBodyUnitCacheTimed(lawRevisionId, lawId, enforcementDate)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawParser", "modLawParser.RefreshLawBodyUnitCache", Err.description, lawId, enforcementDate, lawRevisionId
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function BodyUnitCacheCount(Optional ByVal lawRevisionId As String = "") As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_BODY_UNITS) Then
        BodyUnitCacheCount = 0
        Exit Function
    End If

    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Worksheets(SHEET_BODY_UNITS)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(targetSheet)

    Dim rowIndex As Long
    Dim countValue As Long
    For rowIndex = 2 To lastRow
        If Len(Trim$(lawRevisionId)) = 0 _
            Or StrComp(CellText(targetSheet.cells(rowIndex, 3).Value2), lawRevisionId, vbTextCompare) = 0 Then
            countValue = countValue + 1
        End If
    Next rowIndex

    BodyUnitCacheCount = countValue
End Function

Public Function LawBodyParseSmoke(ByVal lawRevisionId As String) As String
    On Error GoTo ErrHandler

    Dim unitCount As Long
    unitCount = RefreshLawBodyUnitCache(lawRevisionId)
    LawBodyParseSmoke = "ok:" & CStr(unitCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawParser", "modLawParser.LawBodyParseSmoke", Err.description, "", "", lawRevisionId
    LawBodyParseSmoke = "error:" & Err.description
End Function

Public Function LawBodyParsePerformanceSmoke(ByVal lawRevisionId As String, Optional ByVal lawId As String = "", Optional ByVal enforcementDate As String = "") As String
    On Error GoTo ErrHandler

    Dim apiMs As Long
    Dim jsonMs As Long
    Dim buildMs As Long
    Dim writeMs As Long
    Dim writeMode As String
    Dim unitCount As Long
    unitCount = RefreshLawBodyUnitCacheTimed(lawRevisionId, lawId, enforcementDate, apiMs, jsonMs, buildMs, writeMs, writeMode)

    LawBodyParsePerformanceSmoke = "ok:units=" & CStr(unitCount) & _
        ";apiMs=" & CStr(apiMs) & _
        ";jsonMs=" & CStr(jsonMs) & _
        ";buildMs=" & CStr(buildMs) & _
        ";writeMs=" & CStr(writeMs) & _
        ";writeMode=" & writeMode
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawParser", "modLawParser.LawBodyParsePerformanceSmoke", Err.description, lawId, enforcementDate, lawRevisionId
    LawBodyParsePerformanceSmoke = "error:" & Err.description
End Function

Public Function BodyUnitBulkWriteSmoke() As String
    On Error GoTo ErrHandler

    Dim targetSheet As Worksheet
    Set targetSheet = EnsureTempSheet(ThisWorkbook, SHEET_BODY_UNITS, BodyUnitHeaders())
    modSheetStore.ClearDataRows targetSheet

    Dim meta As LawBodyMeta
    meta.lawId = "TEST001"
    meta.lawRevisionId = "TESTREV001"
    meta.enforcementDate = "2025-04-01"
    meta.lawTitle = "テスト法"
    meta.lawNum = "令和七年法律第一号"

    Dim bodyUnitRows As Collection
    Set bodyUnitRows = New Collection
    AddBodyUnitRow bodyUnitRows, meta, "Article", "第一条", "目的", "", "", "000001", "第一条テスト本文。"
    AddBodyUnitRow bodyUnitRows, meta, "Paragraph", "第一条", "目的", "1", "", "000001.0001", "テスト本文。"
    Dim writeMode As String
    writeMode = ReplaceRowsForRevision(targetSheet, meta.lawRevisionId, bodyUnitRows)
    If StrComp(writeMode, "append", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 6710, "modLawParser.BodyUnitBulkWriteSmoke", "Initial write did not use append mode."
    End If

    If BodyUnitCacheCount("TESTREV001") <> 2 Then
        Err.Raise vbObjectError + 6711, "modLawParser.BodyUnitBulkWriteSmoke", "Initial bulk write count mismatch."
    End If

    Dim secondMeta As LawBodyMeta
    secondMeta = meta
    secondMeta.lawRevisionId = "TESTREV002"

    Set bodyUnitRows = New Collection
    AddBodyUnitRow bodyUnitRows, secondMeta, "Article", "第一条", "別revision", "", "", "000001", "別revision本文。"
    writeMode = ReplaceRowsForRevision(targetSheet, secondMeta.lawRevisionId, bodyUnitRows)
    If StrComp(writeMode, "append", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 6713, "modLawParser.BodyUnitBulkWriteSmoke", "New revision write did not use append mode."
    End If
    If BodyUnitCacheCount("TESTREV001") <> 2 Or BodyUnitCacheCount("TESTREV002") <> 1 Then
        Err.Raise vbObjectError + 6714, "modLawParser.BodyUnitBulkWriteSmoke", "Append fast path did not preserve existing revision rows."
    End If

    Set bodyUnitRows = New Collection
    AddBodyUnitRow bodyUnitRows, meta, "Article", "第二条", "", "", "", "000002", "第二条差し替え本文。"
    writeMode = ReplaceRowsForRevision(targetSheet, meta.lawRevisionId, bodyUnitRows)
    If StrComp(writeMode, "replace", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 6715, "modLawParser.BodyUnitBulkWriteSmoke", "Existing revision write did not use replace mode."
    End If

    If BodyUnitCacheCount("TESTREV001") <> 1 Or BodyUnitCacheCount("TESTREV002") <> 1 Then
        Err.Raise vbObjectError + 6712, "modLawParser.BodyUnitBulkWriteSmoke", "Replacement bulk write count mismatch."
    End If

    BodyUnitBulkWriteSmoke = "ok:" & CStr(BodyUnitCacheCount("TESTREV001"))
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawParser", "modLawParser.BodyUnitBulkWriteSmoke", Err.description, "", "", ""
    BodyUnitBulkWriteSmoke = "error:" & Err.description
End Function

Public Function BodyUnitLongTextSmoke() As String
    On Error GoTo ErrHandler

    Dim targetSheet As Worksheet
    Set targetSheet = EnsureTempSheet(ThisWorkbook, SHEET_BODY_UNITS, BodyUnitHeaders())
    modSheetStore.ClearDataRows targetSheet

    Dim meta As LawBodyMeta
    meta.lawId = "TEST001"
    meta.lawRevisionId = "TESTREV_LONG"
    meta.enforcementDate = "2025-04-01"
    meta.lawTitle = "テスト法"
    meta.lawNum = "令和七年法律第一号"

    Dim longText As String
    longText = String$(EXCEL_CELL_TEXT_LIMIT + 1000, "A")

    Dim bodyUnitRows As Collection
    Set bodyUnitRows = New Collection
    AddBodyUnitRow bodyUnitRows, meta, "Paragraph", "第一条", "", "1", "", "000001.0001", longText
    ReplaceRowsForRevision targetSheet, meta.lawRevisionId, bodyUnitRows

    Dim writtenLength As Long
    writtenLength = Len(CStr(targetSheet.cells(2, BODY_UNIT_COLUMN_COUNT).Value2))
    If writtenLength <> EXCEL_CELL_TEXT_LIMIT Then
        Err.Raise vbObjectError + 6716, "modLawParser.BodyUnitLongTextSmoke", "Long text was not constrained before worksheet write."
    End If

    BodyUnitLongTextSmoke = "ok:" & CStr(writtenLength)
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawParser", "modLawParser.BodyUnitLongTextSmoke", Err.description, "", "", ""
    BodyUnitLongTextSmoke = "error:" & Err.description
End Function

Private Function RefreshLawBodyUnitCacheTimed(ByVal lawRevisionId As String, Optional ByVal lawId As String = "", Optional ByVal enforcementDate As String = "", Optional ByRef apiMs As Long = 0, Optional ByRef jsonMs As Long = 0, Optional ByRef buildMs As Long = 0, Optional ByRef writeMs As Long = 0, Optional ByRef writeMode As String = "") As Long
    lawRevisionId = Trim$(lawRevisionId)
    If Len(lawRevisionId) = 0 Then
        Err.Raise vbObjectError + 6701, "modLawParser.RefreshLawBodyUnitCacheTimed", "LawRevisionId is required."
    End If

    Dim startedAt As Single
    startedAt = Timer
    Dim responseJson As String
    responseJson = modLawBody.GetLawDataJson(lawRevisionId, lawId, enforcementDate)
    apiMs = MillisecondsSince(startedAt)

    startedAt = Timer
    Dim root As Object
    Set root = modJsonUtil.ParseJsonObject(responseJson)

    Dim fullText As Object
    Set fullText = modJsonUtil.JsonObjectProperty(root, "law_full_text")
    If fullText Is Nothing Then
        Err.Raise vbObjectError + 6702, "modLawParser.RefreshLawBodyUnitCacheTimed", "law_full_text is missing."
    End If

    Dim meta As LawBodyMeta
    meta = MetadataFromResponse(root, lawId, enforcementDate)
    jsonMs = MillisecondsSince(startedAt)

    startedAt = Timer
    Dim bodyUnitRows As Collection
    Set bodyUnitRows = New Collection
    Dim unitCount As Long
    Dim bodyOrdinal As Long
    AppendBodyNodes fullText, bodyUnitRows, meta, unitCount, bodyOrdinal

    If unitCount = 0 Then
        AddBodyUnitRow bodyUnitRows, meta, "LawText", "", "", "", "", "000000", NormalizeText(NodeText(fullText))
        unitCount = 1
    End If
    buildMs = MillisecondsSince(startedAt)

    startedAt = Timer
    Dim targetSheet As Worksheet
    Set targetSheet = EnsureTempSheet(ThisWorkbook, SHEET_BODY_UNITS, BodyUnitHeaders())
    writeMode = ReplaceRowsForRevision(targetSheet, meta.lawRevisionId, bodyUnitRows)
    writeMs = MillisecondsSince(startedAt)

    RefreshLawBodyUnitCacheTimed = unitCount
End Function

Private Function MetadataFromResponse(ByVal root As Object, ByVal fallbackLawId As String, ByVal fallbackEnforcementDate As String) As LawBodyMeta
    Dim meta As LawBodyMeta

    Dim lawInfo As Object
    Set lawInfo = modJsonUtil.JsonObjectProperty(root, "law_info")

    Dim revisionInfo As Object
    Set revisionInfo = modJsonUtil.JsonObjectProperty(root, "revision_info")

    meta.lawId = modJsonUtil.JsonTextProperty(lawInfo, "law_id", fallbackLawId)
    meta.lawRevisionId = modJsonUtil.JsonTextProperty(revisionInfo, "law_revision_id")
    meta.enforcementDate = modLawRevision.NormalizeEnforcementDate(modJsonUtil.JsonTextProperty(revisionInfo, "amendment_enforcement_date", fallbackEnforcementDate))
    If Len(meta.enforcementDate) = 0 Then
        meta.enforcementDate = modLawRevision.NormalizeEnforcementDate(fallbackEnforcementDate)
    End If
    meta.lawTitle = modJsonUtil.JsonTextProperty(revisionInfo, "law_title")
    meta.lawNum = modJsonUtil.JsonTextProperty(lawInfo, "law_num")

    If Len(meta.lawRevisionId) = 0 Then
        Err.Raise vbObjectError + 6703, "modLawParser.MetadataFromResponse", "law_revision_id is missing."
    End If

    MetadataFromResponse = meta
End Function

Public Function BodyUnitCoverageSmoke() As String
    On Error GoTo ErrHandler

    Dim meta As LawBodyMeta
    meta.lawId = "TEST001"
    meta.lawRevisionId = "TESTREV_COVERAGE"
    meta.enforcementDate = "2025-04-01"
    meta.lawTitle = "テスト法"
    meta.lawNum = "令和七年法律第一号"

    Dim root As Object
    Set root = MakeJsonNode("Law")
    AddJsonChild root, MakeArticleSmokeNode("第一条", "目的", "本則本文の検索語。")
    AddJsonChild root, MakeSupplementarySmokeNode()
    AddJsonChild root, MakeAppdxTableSmokeNode()
    AddJsonChild root, MakeAppdxListSmokeNode()

    Dim bodyUnitRows As Collection
    Set bodyUnitRows = New Collection

    Dim unitCount As Long
    Dim bodyOrdinal As Long
    AppendBodyNodes root, bodyUnitRows, meta, unitCount, bodyOrdinal

    If Not BodyUnitRowsContain(bodyUnitRows, "Paragraph", "本則本文の検索語") Then
        Err.Raise vbObjectError + 6717, "modLawParser.BodyUnitCoverageSmoke", "Main provision paragraph was not parsed."
    End If
    If Not BodyUnitRowsContain(bodyUnitRows, "Paragraph", "附則側の検索語") Then
        Err.Raise vbObjectError + 6718, "modLawParser.BodyUnitCoverageSmoke", "Supplementary provision paragraph was not parsed."
    End If
    If Not BodyUnitRowsContain(bodyUnitRows, "AppdxTable", "別表検索語") Then
        Err.Raise vbObjectError + 6719, "modLawParser.BodyUnitCoverageSmoke", "Appended table text was not parsed."
    End If
    If Not BodyUnitRowsContain(bodyUnitRows, "AppdxTable", "Markdownテーブル") _
        Or Not BodyUnitRowsContain(bodyUnitRows, "AppdxTable", "ASCII罫線テーブル") Then
        Err.Raise vbObjectError + 6720, "modLawParser.BodyUnitCoverageSmoke", "Appended table formatted text was not generated."
    End If
    If Not BodyUnitRowsContain(bodyUnitRows, "AppdxTable", vbCrLf & "一　物の製造") _
        Or Not BodyUnitRowsContain(bodyUnitRows, "AppdxTable", vbCrLf & "十五　焼却") Then
        Err.Raise vbObjectError + 6721, "modLawParser.BodyUnitCoverageSmoke", "Appended numbered list text was not line-broken."
    End If

    BodyUnitCoverageSmoke = "ok:" & CStr(unitCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawParser", "modLawParser.BodyUnitCoverageSmoke", Err.description, "", "", ""
    BodyUnitCoverageSmoke = "error:" & Err.description
End Function

Private Sub AppendBodyNodes(ByVal node As Object, ByVal bodyUnitRows As Collection, ByRef meta As LawBodyMeta, ByRef unitCount As Long, ByRef bodyOrdinal As Long)
    If node Is Nothing Then
        Exit Sub
    End If

    Dim tagName As String
    tagName = NodeTag(node)

    If StrComp(tagName, "Article", vbBinaryCompare) = 0 Then
        bodyOrdinal = bodyOrdinal + 1
        AppendArticleUnit node, bodyUnitRows, meta, unitCount, bodyOrdinal
        Exit Sub
    End If

    If IsStandaloneBodyUnitTag(tagName) Then
        bodyOrdinal = bodyOrdinal + 1
        AppendStandaloneBodyUnit node, bodyUnitRows, meta, unitCount, bodyOrdinal
        Exit Sub
    End If

    Dim beforeCount As Long
    beforeCount = unitCount

    Dim children As Object
    Set children = NodeChildren(node)
    If children Is Nothing Then
        If IsSupplementaryFallbackTag(tagName) Then
            bodyOrdinal = bodyOrdinal + 1
            AppendStandaloneBodyUnit node, bodyUnitRows, meta, unitCount, bodyOrdinal
        End If
        Exit Sub
    End If

    Dim child As Variant
    For Each child In children
        If IsObject(child) Then
            AppendBodyNodes child, bodyUnitRows, meta, unitCount, bodyOrdinal
        End If
    Next child

    If unitCount = beforeCount And IsSupplementaryFallbackTag(tagName) Then
        bodyOrdinal = bodyOrdinal + 1
        AppendStandaloneBodyUnit node, bodyUnitRows, meta, unitCount, bodyOrdinal
    End If
End Sub

Private Sub AppendArticleUnit(ByVal articleNode As Object, ByVal bodyUnitRows As Collection, ByRef meta As LawBodyMeta, ByRef unitCount As Long, ByVal articleOrdinal As Long)
    Dim articleTitle As String
    articleTitle = ChildTextByTag(articleNode, "ArticleTitle")

    Dim articleCaption As String
    articleCaption = ChildTextByTag(articleNode, "ArticleCaption")

    Dim articleKey As String
    articleKey = Format$(articleOrdinal, "000000")

    AddBodyUnitRow bodyUnitRows, meta, "Article", articleTitle, articleCaption, "", "", articleKey, ""
    unitCount = unitCount + 1

    Dim children As Object
    Set children = NodeChildren(articleNode)
    If children Is Nothing Then
        Exit Sub
    End If

    Dim paragraphOrdinal As Long
    Dim child As Variant
    For Each child In children
        If IsObject(child) And StrComp(NodeTag(child), "Paragraph", vbBinaryCompare) = 0 Then
            paragraphOrdinal = paragraphOrdinal + 1
            AppendParagraphUnit child, bodyUnitRows, meta, unitCount, articleTitle, articleCaption, articleKey, paragraphOrdinal
        End If
    Next child
End Sub

Private Sub AppendParagraphUnit(ByVal paragraphNode As Object, ByVal bodyUnitRows As Collection, ByRef meta As LawBodyMeta, ByRef unitCount As Long, ByVal articleTitle As String, ByVal articleCaption As String, ByVal articleKey As String, ByVal paragraphOrdinal As Long)
    Dim paragraphNum As String
    paragraphNum = AttributeText(paragraphNode, "Num")
    If Len(paragraphNum) = 0 Then
        paragraphNum = ChildTextByTag(paragraphNode, "ParagraphNum")
    End If

    Dim paragraphText As String
    paragraphText = ChildTextByTag(paragraphNode, "ParagraphSentence")
    If Len(paragraphText) = 0 Then
        paragraphText = NodeText(paragraphNode)
    End If

    Dim paragraphKey As String
    paragraphKey = articleKey & "." & Format$(paragraphOrdinal, "0000")

    AddBodyUnitRow bodyUnitRows, meta, "Paragraph", articleTitle, articleCaption, paragraphNum, "", paragraphKey, NormalizeText(paragraphText)
    unitCount = unitCount + 1

    Dim itemOrdinal As Long
    AppendItemUnits paragraphNode, bodyUnitRows, meta, unitCount, articleTitle, articleCaption, paragraphNum, paragraphKey, itemOrdinal
End Sub

Private Sub AppendItemUnits(ByVal node As Object, ByVal bodyUnitRows As Collection, ByRef meta As LawBodyMeta, ByRef unitCount As Long, ByVal articleTitle As String, ByVal articleCaption As String, ByVal paragraphNum As String, ByVal paragraphKey As String, ByRef itemOrdinal As Long)
    Dim children As Object
    Set children = NodeChildren(node)
    If children Is Nothing Then
        Exit Sub
    End If

    Dim child As Variant
    For Each child In children
        If IsObject(child) Then
            If StrComp(NodeTag(child), "Item", vbBinaryCompare) = 0 Then
                itemOrdinal = itemOrdinal + 1
                AppendSingleItemUnit child, bodyUnitRows, meta, unitCount, articleTitle, articleCaption, paragraphNum, paragraphKey, itemOrdinal
            Else
                AppendItemUnits child, bodyUnitRows, meta, unitCount, articleTitle, articleCaption, paragraphNum, paragraphKey, itemOrdinal
            End If
        End If
    Next child
End Sub

Private Sub AppendSingleItemUnit(ByVal itemNode As Object, ByVal bodyUnitRows As Collection, ByRef meta As LawBodyMeta, ByRef unitCount As Long, ByVal articleTitle As String, ByVal articleCaption As String, ByVal paragraphNum As String, ByVal paragraphKey As String, ByVal itemOrdinal As Long)
    Dim itemTitle As String
    itemTitle = ChildTextByTag(itemNode, "ItemTitle")

    Dim itemText As String
    itemText = ChildTextByTag(itemNode, "ItemSentence")
    If Len(itemText) = 0 Then
        itemText = NodeText(itemNode)
    End If

    AddBodyUnitRow bodyUnitRows, meta, "Item", articleTitle, articleCaption, paragraphNum, itemTitle, paragraphKey & "." & Format$(itemOrdinal, "0000"), NormalizeText(itemText)
    unitCount = unitCount + 1
End Sub

Private Sub AppendStandaloneBodyUnit(ByVal node As Object, ByVal bodyUnitRows As Collection, ByRef meta As LawBodyMeta, ByRef unitCount As Long, ByVal bodyOrdinal As Long)
    Dim unitText As String
    unitText = StandaloneBodyUnitText(node)
    If Len(unitText) = 0 Then
        Exit Sub
    End If

    AddBodyUnitRow bodyUnitRows, meta, StandaloneBodyUnitKind(NodeTag(node)), StandaloneBodyUnitTitle(node), "", "", "", Format$(bodyOrdinal, "000000"), unitText
    unitCount = unitCount + 1
End Sub

Private Function StandaloneBodyUnitText(ByVal node As Object) As String
    Dim unitText As String
    unitText = NormalizeText(NodeText(node))

    If StrComp(StandaloneBodyUnitKind(NodeTag(node)), "AppdxTable", vbBinaryCompare) = 0 Then
        unitText = RemoveLeadingText(unitText, StandaloneBodyUnitTitle(node))
        unitText = BreakJapaneseNumberedList(unitText)

        Dim formattedTables As String
        formattedTables = modLawTableParser.FormattedTablesTextFromNode(node)
        If Len(formattedTables) > 0 Then
            If Len(unitText) > 0 Then
                unitText = unitText & vbCrLf & formattedTables
            Else
                unitText = formattedTables
            End If
        End If
    End If

    StandaloneBodyUnitText = unitText
End Function

Private Function RemoveLeadingText(ByVal baseText As String, ByVal leadingText As String) As String
    leadingText = Trim$(leadingText)
    If Len(leadingText) > 0 And Left$(baseText, Len(leadingText)) = leadingText Then
        RemoveLeadingText = Trim$(Mid$(baseText, Len(leadingText) + 1))
    Else
        RemoveLeadingText = baseText
    End If
End Function

Private Function BreakJapaneseNumberedList(ByVal value As String) As String
    Dim listStart As Long
    listStart = JapaneseNumberedListStartPosition(value)
    If listStart = 0 Then
        BreakJapaneseNumberedList = value
        Exit Function
    End If

    Dim listText As String
    listText = Mid$(value, listStart)

    Dim segments As Collection
    Set segments = JapaneseNumberedListSegments(listText)
    If segments.Count < 2 Then
        BreakJapaneseNumberedList = value
        Exit Function
    End If

    Dim result As String
    result = Trim$(Left$(value, listStart - 1))

    Dim segmentIndex As Long
    For segmentIndex = 1 To segments.Count
        result = AppendLine(result, CStr(segments(segmentIndex)))
    Next segmentIndex

    BreakJapaneseNumberedList = result
End Function

Private Function JapaneseNumberedListStartPosition(ByVal value As String) As Long
    Dim closeParenPosition As Long
    closeParenPosition = InStr(1, value, "）", vbBinaryCompare)
    Do While closeParenPosition > 0
        Dim nextPosition As Long
        nextPosition = NextNonSpacePosition(value, closeParenPosition + 1)
        If nextPosition > 0 And Mid$(value, nextPosition, 1) = "一" Then
            JapaneseNumberedListStartPosition = nextPosition
            Exit Function
        End If
        closeParenPosition = InStr(closeParenPosition + 1, value, "）", vbBinaryCompare)
    Loop

    If Left$(value, 1) = "一" Then
        JapaneseNumberedListStartPosition = 1
    End If
End Function

Private Function NextNonSpacePosition(ByVal value As String, ByVal startPosition As Long) As Long
    Dim position As Long
    For position = startPosition To Len(value)
        If Mid$(value, position, 1) <> " " Then
            NextNonSpacePosition = position
            Exit Function
        End If
    Next position
End Function

Private Function JapaneseNumberedListSegments(ByVal listText As String) As Collection
    Dim markers As Variant
    markers = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")

    Dim segments As Collection
    Set segments = New Collection

    Dim previousStart As Long
    Dim previousMarker As String
    Dim searchFrom As Long
    searchFrom = 1

    Dim markerIndex As Long
    For markerIndex = LBound(markers) To UBound(markers)
        Dim marker As String
        marker = CStr(markers(markerIndex))

        Dim markerPosition As Long
        markerPosition = FindNumberMarkerPosition(listText, marker, searchFrom)
        If markerPosition = 0 Then
            Exit For
        End If

        If previousStart > 0 Then
            segments.Add FormatJapaneseNumberedSegment(previousMarker, Mid$(listText, previousStart + Len(previousMarker), markerPosition - previousStart - Len(previousMarker)))
        End If

        previousStart = markerPosition
        previousMarker = marker
        searchFrom = markerPosition + Len(marker)
    Next markerIndex

    If previousStart > 0 Then
        segments.Add FormatJapaneseNumberedSegment(previousMarker, Mid$(listText, previousStart + Len(previousMarker)))
    End If

    Set JapaneseNumberedListSegments = segments
End Function

Private Function FindNumberMarkerPosition(ByVal value As String, ByVal marker As String, ByVal startPosition As Long) As Long
    Dim foundPosition As Long
    foundPosition = InStr(startPosition, value, marker, vbBinaryCompare)

    Do While foundPosition > 0
        If foundPosition = 1 Then
            FindNumberMarkerPosition = foundPosition
            Exit Function
        ElseIf Mid$(value, foundPosition - 1, 1) <> "第" Then
            FindNumberMarkerPosition = foundPosition
            Exit Function
        End If
        foundPosition = InStr(foundPosition + Len(marker), value, marker, vbBinaryCompare)
    Loop
End Function

Private Function FormatJapaneseNumberedSegment(ByVal marker As String, ByVal segmentText As String) As String
    FormatJapaneseNumberedSegment = marker & "　" & Trim$(segmentText)
End Function

Private Function AppendLine(ByVal baseText As String, ByVal lineText As String) As String
    lineText = Trim$(lineText)
    If Len(lineText) = 0 Then
        AppendLine = baseText
    ElseIf Len(baseText) = 0 Then
        AppendLine = lineText
    Else
        AppendLine = baseText & vbCrLf & lineText
    End If
End Function

Private Sub AddBodyUnitRow(ByVal bodyUnitRows As Collection, ByRef meta As LawBodyMeta, ByVal unitKind As String, ByVal articleTitle As String, ByVal articleCaption As String, ByVal paragraphNum As String, ByVal itemTitle As String, ByVal sortKey As String, ByVal unitText As String)
    Dim rowValues As Variant
    rowValues = Array( _
        Format$(Now, "yyyy-mm-dd hh:nn:ss"), _
        CellTextForWrite(meta.lawId), _
        CellTextForWrite(meta.lawRevisionId), _
        CellTextForWrite(meta.enforcementDate), _
        CellTextForWrite(meta.lawTitle), _
        CellTextForWrite(meta.lawNum), _
        CellTextForWrite(unitKind), _
        CellTextForWrite(articleTitle), _
        CellTextForWrite(articleCaption), _
        CellTextForWrite(paragraphNum), _
        CellTextForWrite(itemTitle), _
        CellTextForWrite(sortKey), _
        CellTextForWrite(unitText))

    bodyUnitRows.Add rowValues
End Sub

Private Function ReplaceRowsForRevision(ByVal targetSheet As Worksheet, ByVal lawRevisionId As String, ByVal newRows As Collection) As String
    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(targetSheet)

    If newRows.Count = 0 Then
        ReplaceRowsForRevision = "empty"
        Exit Function
    End If

    If lastRow < 2 Or Not RevisionExists(targetSheet, lastRow, lawRevisionId) Then
        AppendBodyUnitRows targetSheet, Application.Max(2, lastRow + 1), newRows
        ReplaceRowsForRevision = "append"
        Exit Function
    End If

    Dim existingValues As Variant
    existingValues = targetSheet.Range(targetSheet.cells(2, 1), targetSheet.cells(lastRow, BODY_UNIT_COLUMN_COUNT)).Value2

    Dim retainedCount As Long
    Dim rowIndex As Long
    If IsArray(existingValues) Then
        For rowIndex = 1 To UBound(existingValues, 1)
            If StrComp(CellText(existingValues(rowIndex, BODY_UNIT_COL_LAW_REVISION_ID)), lawRevisionId, vbTextCompare) <> 0 Then
                retainedCount = retainedCount + 1
            End If
        Next rowIndex
    End If

    Dim outputCount As Long
    outputCount = retainedCount + newRows.Count

    ClearBodyUnitData targetSheet, lastRow
    If outputCount = 0 Then
        ReplaceRowsForRevision = "replace"
        Exit Function
    End If

    Dim outputValues As Variant
    ReDim outputValues(1 To outputCount, 1 To BODY_UNIT_COLUMN_COUNT)

    Dim outputRow As Long
    Dim columnIndex As Long
    If IsArray(existingValues) Then
        For rowIndex = 1 To UBound(existingValues, 1)
            If StrComp(CellText(existingValues(rowIndex, BODY_UNIT_COL_LAW_REVISION_ID)), lawRevisionId, vbTextCompare) <> 0 Then
                outputRow = outputRow + 1
                For columnIndex = 1 To BODY_UNIT_COLUMN_COUNT
                    outputValues(outputRow, columnIndex) = existingValues(rowIndex, columnIndex)
                Next columnIndex
            End If
        Next rowIndex
    End If

    Dim rowValues As Variant
    Dim newRowIndex As Long
    For newRowIndex = 1 To newRows.Count
        rowValues = newRows(newRowIndex)
        outputRow = outputRow + 1
        For columnIndex = LBound(rowValues) To UBound(rowValues)
            outputValues(outputRow, columnIndex - LBound(rowValues) + 1) = rowValues(columnIndex)
        Next columnIndex
    Next newRowIndex

    targetSheet.Range(targetSheet.cells(2, 1), targetSheet.cells(outputCount + 1, BODY_UNIT_COLUMN_COUNT)).Value2 = outputValues
    ReplaceRowsForRevision = "replace"
End Function

Private Function RevisionExists(ByVal targetSheet As Worksheet, ByVal lastRow As Long, ByVal lawRevisionId As String) As Boolean
    If lastRow < 2 Then
        Exit Function
    End If

    Dim revisionValues As Variant
    revisionValues = targetSheet.Range(targetSheet.cells(2, BODY_UNIT_COL_LAW_REVISION_ID), targetSheet.cells(lastRow, BODY_UNIT_COL_LAW_REVISION_ID)).Value2

    If Not IsArray(revisionValues) Then
        RevisionExists = (StrComp(CellText(revisionValues), lawRevisionId, vbTextCompare) = 0)
        Exit Function
    End If

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(revisionValues, 1)
        If StrComp(CellText(revisionValues(rowIndex, 1)), lawRevisionId, vbTextCompare) = 0 Then
            RevisionExists = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub AppendBodyUnitRows(ByVal targetSheet As Worksheet, ByVal startRow As Long, ByVal newRows As Collection)
    Dim outputValues As Variant
    ReDim outputValues(1 To newRows.Count, 1 To BODY_UNIT_COLUMN_COUNT)

    Dim outputRow As Long
    Dim columnIndex As Long
    Dim rowValues As Variant
    For outputRow = 1 To newRows.Count
        rowValues = newRows(outputRow)
        For columnIndex = LBound(rowValues) To UBound(rowValues)
            outputValues(outputRow, columnIndex - LBound(rowValues) + 1) = rowValues(columnIndex)
        Next columnIndex
    Next outputRow

    targetSheet.Range(targetSheet.cells(startRow, 1), targetSheet.cells(startRow + newRows.Count - 1, BODY_UNIT_COLUMN_COUNT)).Value2 = outputValues
End Sub

Private Sub ClearBodyUnitData(ByVal targetSheet As Worksheet, ByVal lastRow As Long)
    If lastRow > 1 Then
        targetSheet.Range(targetSheet.cells(2, 1), targetSheet.cells(lastRow, BODY_UNIT_COLUMN_COUNT)).ClearContents
    End If
End Sub

Private Function EnsureTempSheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String, ByVal headers As Variant) As Worksheet
    Dim targetSheet As Worksheet
    If modSheetStore.SheetExists(targetWorkbook, sheetName) Then
        Set targetSheet = targetWorkbook.Worksheets(sheetName)
    Else
        Set targetSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count))
        targetSheet.Name = sheetName
    End If

    targetSheet.cells.NumberFormat = "@"
    If Application.CountA(targetSheet.rows(1)) = 0 Then
        WriteHeaders targetSheet, headers
    End If
    targetSheet.Visible = xlSheetVeryHidden
    Set EnsureTempSheet = targetSheet
End Function

Private Sub WriteHeaders(ByVal targetSheet As Worksheet, ByVal headers As Variant)
    Dim index As Long
    For index = LBound(headers) To UBound(headers)
        targetSheet.cells(1, index - LBound(headers) + 1).Value2 = headers(index)
    Next index
    targetSheet.rows(1).Font.Bold = True
End Sub

Private Function BodyUnitHeaders() As Variant
    BodyUnitHeaders = Array( _
        "CachedAt", _
        "LawId", _
        "LawRevisionId", _
        "EnforcementDate", _
        "LawTitle", _
        "LawNum", _
        "UnitKind", _
        "ArticleTitle", _
        "ArticleCaption", _
        "ParagraphNum", _
        "ItemTitle", _
        "SortKey", _
        "Text")
End Function

Private Function NodeTag(ByVal node As Object) As String
    NodeTag = modJsonUtil.JsonTextProperty(node, "tag")
End Function

Private Function NodeChildren(ByVal node As Object) As Object
    Set NodeChildren = modJsonUtil.JsonObjectProperty(node, "children")
End Function

Private Function AttributeText(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributes As Object
    Set attributes = modJsonUtil.JsonObjectProperty(node, "attr")
    AttributeText = modJsonUtil.JsonTextProperty(attributes, attributeName)
End Function

Private Function ChildTextByTag(ByVal node As Object, ByVal childTag As String) As String
    Dim children As Object
    Set children = NodeChildren(node)
    If children Is Nothing Then
        Exit Function
    End If

    Dim result As String
    Dim child As Variant
    For Each child In children
        If IsObject(child) And StrComp(NodeTag(child), childTag, vbBinaryCompare) = 0 Then
            result = result & NodeText(child)
        End If
    Next child
    ChildTextByTag = NormalizeText(result)
End Function

Private Function NodeText(ByVal node As Variant) As String
    If IsObject(node) Then
        Dim children As Object
        Set children = NodeChildren(node)
        If children Is Nothing Then
            NodeText = ""
            Exit Function
        End If

        Dim result As String
        Dim child As Variant
        For Each child In children
            result = result & NodeText(child)
        Next child
        NodeText = result
    ElseIf IsNull(node) Or IsEmpty(node) Then
        NodeText = ""
    Else
        NodeText = CStr(node)
    End If
End Function

Private Function NormalizeText(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")
    value = Trim$(value)

    Do While InStr(1, value, "  ", vbBinaryCompare) > 0
        value = Replace(value, "  ", " ")
    Loop

    NormalizeText = value
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function

Private Function IsStandaloneBodyUnitTag(ByVal tagName As String) As Boolean
    If Len(tagName) = 0 Then
        Exit Function
    End If

    Select Case tagName
        Case "AppdxTable", "SupplProvisionAppdxTable", "TableStruct", "Table", _
             "AppdxNote", "AppdxStyle", "AppdxFormat", "AppdxFig", _
             "SupplProvisionAppdxStyle", "SupplProvisionAppdxFormat", "SupplProvisionAppdxFig"
            IsStandaloneBodyUnitTag = True
        Case Else
            IsStandaloneBodyUnitTag = (Left$(tagName, 5) = "Appdx" Or InStr(1, tagName, "Appdx", vbBinaryCompare) > 0)
    End Select
End Function

Private Function IsSupplementaryFallbackTag(ByVal tagName As String) As Boolean
    IsSupplementaryFallbackTag = (StrComp(tagName, "SupplProvision", vbBinaryCompare) = 0)
End Function

Private Function StandaloneBodyUnitKind(ByVal tagName As String) As String
    If StrComp(tagName, "SupplProvision", vbBinaryCompare) = 0 Then
        StandaloneBodyUnitKind = "SupplementaryProvision"
    ElseIf StrComp(tagName, "Table", vbBinaryCompare) = 0 _
        Or StrComp(tagName, "TableStruct", vbBinaryCompare) = 0 _
        Or InStr(1, tagName, "AppdxTable", vbBinaryCompare) > 0 Then
        StandaloneBodyUnitKind = "AppdxTable"
    ElseIf Left$(tagName, 5) = "Appdx" Or InStr(1, tagName, "Appdx", vbBinaryCompare) > 0 Then
        StandaloneBodyUnitKind = "Appdx"
    Else
        StandaloneBodyUnitKind = "LawText"
    End If
End Function

Private Function StandaloneBodyUnitTitle(ByVal node As Object) As String
    Dim titleTags As Variant
    titleTags = Array("AppdxTableTitle", "SupplProvisionAppdxTableTitle", "AppdxStyleTitle", "AppdxFormatTitle", "AppdxFigTitle", "SupplProvisionLabel", "TableStructTitle")

    Dim index As Long
    Dim titleText As String
    For index = LBound(titleTags) To UBound(titleTags)
        titleText = ChildTextByTag(node, CStr(titleTags(index)))
        If Len(titleText) > 0 Then
            StandaloneBodyUnitTitle = titleText
            Exit Function
        End If
    Next index

    StandaloneBodyUnitTitle = UnitKindFallbackTitle(StandaloneBodyUnitKind(NodeTag(node)))
End Function

Private Function UnitKindFallbackTitle(ByVal unitKind As String) As String
    Select Case unitKind
        Case "AppdxTable"
            UnitKindFallbackTitle = "別表"
        Case "Appdx"
            UnitKindFallbackTitle = "別表等"
        Case "SupplementaryProvision"
            UnitKindFallbackTitle = "附則"
        Case Else
            UnitKindFallbackTitle = "本文"
    End Select
End Function

Private Function CellTextForWrite(ByVal value As String) As String
    If Len(value) <= EXCEL_CELL_TEXT_LIMIT Then
        CellTextForWrite = value
    Else
        CellTextForWrite = Left$(value, EXCEL_CELL_TEXT_LIMIT)
    End If
End Function

Private Function MillisecondsSince(ByVal startedAt As Single) As Long
    Dim elapsedSeconds As Single
    elapsedSeconds = Timer - startedAt
    If elapsedSeconds < 0 Then
        elapsedSeconds = elapsedSeconds + 86400!
    End If

    MillisecondsSince = CLng(elapsedSeconds * 1000!)
End Function

Private Function MakeJsonNode(ByVal tagName As String, Optional ByVal textValue As String = "") As Object
    Dim node As Object
    Set node = CreateObject("Scripting.Dictionary")

    Dim children As Collection
    Set children = New Collection

    node.Add "tag", tagName
    node.Add "children", children
    If Len(textValue) > 0 Then
        children.Add textValue
    End If

    Set MakeJsonNode = node
End Function

Private Sub AddJsonChild(ByVal parentNode As Object, ByVal childNode As Object)
    Dim children As Object
    Set children = NodeChildren(parentNode)
    children.Add childNode
End Sub

Private Function MakeArticleSmokeNode(ByVal articleTitle As String, ByVal articleCaption As String, ByVal paragraphText As String) As Object
    Dim articleNode As Object
    Set articleNode = MakeJsonNode("Article")
    AddJsonChild articleNode, MakeJsonNode("ArticleTitle", articleTitle)
    If Len(articleCaption) > 0 Then
        AddJsonChild articleNode, MakeJsonNode("ArticleCaption", articleCaption)
    End If

    Dim paragraphNode As Object
    Set paragraphNode = MakeJsonNode("Paragraph")
    AddJsonChild paragraphNode, MakeJsonNode("ParagraphNum", "1")
    AddJsonChild paragraphNode, MakeJsonNode("ParagraphSentence", paragraphText)
    AddJsonChild articleNode, paragraphNode

    Set MakeArticleSmokeNode = articleNode
End Function

Private Function MakeSupplementarySmokeNode() As Object
    Dim supplementaryNode As Object
    Set supplementaryNode = MakeJsonNode("SupplProvision")
    AddJsonChild supplementaryNode, MakeJsonNode("SupplProvisionLabel", "附則")
    AddJsonChild supplementaryNode, MakeArticleSmokeNode("第一条", "", "附則側の検索語。")
    Set MakeSupplementarySmokeNode = supplementaryNode
End Function

Private Function MakeAppdxTableSmokeNode() As Object
    Dim tableNode As Object
    Set tableNode = MakeJsonNode("AppdxTable")
    AddJsonChild tableNode, MakeJsonNode("AppdxTableTitle", "別表第一")

    Dim tableStruct As Object
    Set tableStruct = MakeJsonNode("TableStruct")

    Dim headerRow As Object
    Set headerRow = MakeJsonNode("TableHeaderRow")
    AddJsonChild headerRow, MakeJsonNode("TableHeaderColumn", "項目")
    AddJsonChild headerRow, MakeJsonNode("TableHeaderColumn", "値")
    AddJsonChild tableStruct, headerRow

    Dim dataRow As Object
    Set dataRow = MakeJsonNode("TableRow")
    AddJsonChild dataRow, MakeJsonNode("TableColumn", "別表検索語")
    AddJsonChild dataRow, MakeJsonNode("TableColumn", "8")
    AddJsonChild tableStruct, dataRow

    AddJsonChild tableNode, tableStruct
    Set MakeAppdxTableSmokeNode = tableNode
End Function

Private Function MakeAppdxListSmokeNode() As Object
    Dim tableNode As Object
    Set tableNode = MakeJsonNode("AppdxTable")
    AddJsonChild tableNode, MakeJsonNode("AppdxTableTitle", "別表列挙")
    AddJsonChild tableNode, MakeJsonNode("RelatedArticleNum", "（第三十三条、第四十条、第四十一条関係）")
    AddJsonChild tableNode, MakeJsonNode("AppdxTableText", "一物の製造の事業二鉱業の事業三土木の事業四運送の事業五貨物取扱いの事業六農林の事業七畜産の事業八販売の事業九金融の事業十映画の事業十一郵便の事業十二教育の事業十三保健衛生の事業十四旅館の事業十五焼却の事業")
    Set MakeAppdxListSmokeNode = tableNode
End Function

Private Function BodyUnitRowsContain(ByVal bodyUnitRows As Collection, ByVal unitKind As String, ByVal expectedText As String) As Boolean
    Dim rowValues As Variant
    Dim rowIndex As Long
    For rowIndex = 1 To bodyUnitRows.Count
        rowValues = bodyUnitRows(rowIndex)
        If StrComp(CStr(rowValues(6)), unitKind, vbBinaryCompare) = 0 _
            And InStr(1, CStr(rowValues(12)), expectedText, vbBinaryCompare) > 0 Then
            BodyUnitRowsContain = True
            Exit Function
        End If
    Next rowIndex
End Function
