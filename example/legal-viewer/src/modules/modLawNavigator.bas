Attribute VB_Name = "modLawNavigator"
Option Explicit

Private Const BODY_UNIT_COL_LAW_REVISION_ID As Long = 3
Private Const BODY_UNIT_COL_UNIT_KIND As Long = 7
Private Const BODY_UNIT_COL_ARTICLE_TITLE As Long = 8
Private Const BODY_UNIT_COL_ARTICLE_CAPTION As Long = 9
Private Const BODY_UNIT_COL_PARAGRAPH_NUM As Long = 10
Private Const BODY_UNIT_COL_ITEM_TITLE As Long = 11
Private Const BODY_UNIT_COL_SORT_KEY As Long = 12
Private Const BODY_UNIT_COL_TEXT As Long = 13
Private Const BODY_UNIT_COL_COUNT As Long = 13

Private Const ADDED_LAW_COL_LAW_REVISION_ID As Long = 5
Private Const ADDED_LAW_COL_LAW_TITLE As Long = 6
Private Const ADDED_LAW_COL_LAW_NUM As Long = 7
Private Const TABLE_DISPLAY_MARKDOWN As String = "markdown"
Private Const TABLE_DISPLAY_ASCII As String = "ascii"
Private Const TABLE_MARKDOWN_MARKER As String = "Markdownテーブル"
Private Const TABLE_ASCII_MARKER As String = "ASCII罫線テーブル"

Public Function TableDisplayModeMarkdown() As String
    TableDisplayModeMarkdown = TABLE_DISPLAY_MARKDOWN
End Function

Public Function TableDisplayModeAscii() As String
    TableDisplayModeAscii = TABLE_DISPLAY_ASCII
End Function

Public Sub ConfigureBodyNavListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 5
    targetListBox.columnWidths = "48 pt;180 pt;0 pt;0 pt;0 pt"
    targetListBox.BoundColumn = 4
End Sub

Public Function PopulateBodyNavListBox(ByVal targetListBox As Object, ByVal addedLawStoreRow As Long) As Long
    On Error GoTo ErrHandler

    ConfigureBodyNavListBox targetListBox
    targetListBox.Clear

    Dim lawRevisionId As String
    lawRevisionId = AddedLawRevisionId(addedLawStoreRow)
    If Len(lawRevisionId) = 0 Then
        PopulateBodyNavListBox = 0
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        PopulateBodyNavListBox = 0
        Exit Function
    End If

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(bodySheet)

    Dim bodyValues As Variant
    LoadBodyCacheValues bodySheet, lastRow, bodyValues

    Dim revisionRows As Collection
    Set revisionRows = RowsForRevision(bodyValues, lawRevisionId)

    Dim rowIndex As Variant
    Dim listIndex As Long
    For Each rowIndex In revisionRows
        If ShouldShowInBodyNav(BodyValueText(bodyValues, CLng(rowIndex), BODY_UNIT_COL_UNIT_KIND)) Then
            targetListBox.AddItem UnitKindDisplay(BodyValueText(bodyValues, CLng(rowIndex), BODY_UNIT_COL_UNIT_KIND))
            listIndex = targetListBox.ListCount - 1
            targetListBox.List(listIndex, 1) = BodyUnitHeading(bodyValues, CLng(rowIndex))
            targetListBox.List(listIndex, 2) = BodyValueText(bodyValues, CLng(rowIndex), BODY_UNIT_COL_SORT_KEY)
            targetListBox.List(listIndex, 3) = CStr(rowIndex)
            targetListBox.List(listIndex, 4) = BodyValueText(bodyValues, CLng(rowIndex), BODY_UNIT_COL_UNIT_KIND)
        End If
    Next rowIndex

    PopulateBodyNavListBox = targetListBox.ListCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawNavigator", "modLawNavigator.PopulateBodyNavListBox", Err.description, "", "", CStr(addedLawStoreRow)
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function BodyPreviewText(ByVal addedLawStoreRow As Long, Optional ByVal maxUnits As Long = 40, Optional ByVal tableDisplayMode As String = TABLE_DISPLAY_MARKDOWN) As String
    On Error GoTo ErrHandler

    Dim lawRevisionId As String
    lawRevisionId = AddedLawRevisionId(addedLawStoreRow)
    If Len(lawRevisionId) = 0 Then
        BodyPreviewText = ""
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        BodyPreviewText = "本文キャッシュがありません。法令追加画面で本文取得に成功した法令を選択してください。"
        Exit Function
    End If

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())

    Dim result As String
    result = AddedLawHeaderText(addedLawStoreRow)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(bodySheet)

    Dim bodyValues As Variant
    LoadBodyCacheValues bodySheet, lastRow, bodyValues

    Dim revisionRows As Collection
    Set revisionRows = RowsForRevision(bodyValues, lawRevisionId)

    Dim articleIndex As Object
    Set articleIndex = BuildArticleUnitIndex(bodyValues, revisionRows)

    BodyPreviewText = BodyPreviewTextFromCache(result, bodyValues, revisionRows, articleIndex, maxUnits, tableDisplayMode)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawNavigator", "modLawNavigator.BodyPreviewText", Err.description, "", "", CStr(addedLawStoreRow)
    BodyPreviewText = "プレビュー生成失敗: " & Err.description
End Function

Private Function BodyPreviewTextFromCache(ByVal headerText As String, ByRef bodyValues As Variant, ByVal revisionRows As Collection, ByVal articleIndex As Object, ByVal maxUnits As Long, Optional ByVal tableDisplayMode As String = TABLE_DISPLAY_MARKDOWN) As String
    Dim result As String
    result = headerText

    Dim rowIndex As Variant
    Dim unitCount As Long
    For Each rowIndex In revisionRows
        If ShouldShowInBodyNav(BodyValueText(bodyValues, CLng(rowIndex), BODY_UNIT_COL_UNIT_KIND)) Then
            unitCount = unitCount + 1
            If unitCount <= maxUnits Then
                result = AppendParagraph(result, FormattedBodyUnitText(bodyValues, CLng(rowIndex), articleIndex, tableDisplayMode))
            End If
        End If
    Next rowIndex

    If unitCount = 0 Then
        BodyPreviewTextFromCache = "本文キャッシュがありません。"
    ElseIf unitCount > maxUnits Then
        BodyPreviewTextFromCache = result & vbCrLf & vbCrLf & "... " & CStr(unitCount - maxUnits) & " 件省略"
    Else
        BodyPreviewTextFromCache = result
    End If
End Function

Public Function BodyUnitPreviewText(ByVal bodyUnitStoreRow As Long, Optional ByVal tableDisplayMode As String = TABLE_DISPLAY_MARKDOWN) As String
    On Error GoTo ErrHandler

    If Not modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        BodyUnitPreviewText = ""
        Exit Function
    End If

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())
    If bodyUnitStoreRow < 2 Or bodyUnitStoreRow > modSheetStore.LastUsedRow(bodySheet) Then
        BodyUnitPreviewText = ""
        Exit Function
    End If

    Dim bodyValues As Variant
    LoadBodyCacheValues bodySheet, modSheetStore.LastUsedRow(bodySheet), bodyValues

    Dim lawRevisionId As String
    lawRevisionId = BodyValueText(bodyValues, bodyUnitStoreRow, BODY_UNIT_COL_LAW_REVISION_ID)

    Dim revisionRows As Collection
    Set revisionRows = RowsForRevision(bodyValues, lawRevisionId)

    Dim articleIndex As Object
    Set articleIndex = BuildArticleUnitIndex(bodyValues, revisionRows)

    BodyUnitPreviewText = FormattedBodyUnitText(bodyValues, bodyUnitStoreRow, articleIndex, tableDisplayMode)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawNavigator", "modLawNavigator.BodyUnitPreviewText", Err.description, "", "", CStr(bodyUnitStoreRow)
    BodyUnitPreviewText = "選択単位プレビュー生成失敗: " & Err.description
End Function

Public Function BodyContextPreviewText(ByVal bodyUnitStoreRow As Long, Optional ByVal tableDisplayMode As String = TABLE_DISPLAY_MARKDOWN) As String
    On Error GoTo ErrHandler

    If Not modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        BodyContextPreviewText = ""
        Exit Function
    End If

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())
    If bodyUnitStoreRow < 2 Or bodyUnitStoreRow > modSheetStore.LastUsedRow(bodySheet) Then
        BodyContextPreviewText = ""
        Exit Function
    End If

    Dim bodyValues As Variant
    LoadBodyCacheValues bodySheet, modSheetStore.LastUsedRow(bodySheet), bodyValues

    Dim lawRevisionId As String
    lawRevisionId = BodyValueText(bodyValues, bodyUnitStoreRow, BODY_UNIT_COL_LAW_REVISION_ID)

    Dim revisionRows As Collection
    Set revisionRows = RowsForRevision(bodyValues, lawRevisionId)

    Dim articleIndex As Object
    Set articleIndex = BuildArticleUnitIndex(bodyValues, revisionRows)

    Dim unitKind As String
    unitKind = BodyValueText(bodyValues, bodyUnitStoreRow, BODY_UNIT_COL_UNIT_KIND)
    If StrComp(unitKind, "Article", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "LawText", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "AppdxTable", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "Appdx", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "SupplementaryProvision", vbBinaryCompare) = 0 Then
        BodyContextPreviewText = FormattedBodyUnitText(bodyValues, bodyUnitStoreRow, articleIndex, tableDisplayMode)
        Exit Function
    End If

    Dim articleKey As String
    articleKey = ArticleKeyFromSortKey(BodyValueText(bodyValues, bodyUnitStoreRow, BODY_UNIT_COL_SORT_KEY))

    Dim articleRows As Collection
    Set articleRows = IndexedRows(articleIndex, "Article|" & articleKey)
    If articleRows.Count > 0 Then
        BodyContextPreviewText = FormattedBodyUnitText(bodyValues, CLng(articleRows(1)), articleIndex, tableDisplayMode)
    Else
        BodyContextPreviewText = FormattedBodyUnitText(bodyValues, bodyUnitStoreRow, articleIndex, tableDisplayMode)
    End If
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawNavigator", "modLawNavigator.BodyContextPreviewText", Err.description, "", "", CStr(bodyUnitStoreRow)
    BodyContextPreviewText = "本文コンテキスト表示失敗: " & Err.description
End Function

Public Function BodyNavigatorPerformanceSmoke(Optional ByVal repetitions As Long = 10) As String
    On Error GoTo ErrHandler

    If repetitions < 1 Then
        repetitions = 1
    End If

    Dim addedRow As Long
    addedRow = PrepareBodyNavigatorSmokeData()

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(bodySheet)

    Dim startedAt As Single
    Dim readMs As Double
    Dim indexMs As Double
    Dim previewMs As Double

    startedAt = Timer
    Dim bodyValues As Variant
    LoadBodyCacheValues bodySheet, lastRow, bodyValues
    readMs = ElapsedMilliseconds(startedAt, Timer)

    startedAt = Timer
    Dim revisionRows As Collection
    Set revisionRows = RowsForRevision(bodyValues, "TESTREV001")
    Dim articleIndex As Object
    Set articleIndex = BuildArticleUnitIndex(bodyValues, revisionRows)
    indexMs = ElapsedMilliseconds(startedAt, Timer)

    startedAt = Timer
    Dim index As Long
    Dim previewText As String
    For index = 1 To repetitions
        previewText = BodyPreviewTextFromCache(AddedLawHeaderText(addedRow), bodyValues, revisionRows, articleIndex, 40)
    Next index
    previewMs = ElapsedMilliseconds(startedAt, Timer)

    If InStr(1, previewText, "第十一条　受給権は差し押えることができない。", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6811, "modLawNavigator.BodyNavigatorPerformanceSmoke", "Preview text mismatch."
    End If

    BodyNavigatorPerformanceSmoke = "ok:rows=" & CStr(revisionRows.Count) _
        & ";repetitions=" & CStr(repetitions) _
        & ";readMs=" & Format$(readMs, "0") _
        & ";indexMs=" & Format$(indexMs, "0") _
        & ";previewMs=" & Format$(previewMs, "0")
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawNavigator", "modLawNavigator.BodyNavigatorPerformanceSmoke", Err.description, "", "", ""
    BodyNavigatorPerformanceSmoke = "error:" & Err.description
End Function

Public Function TableDisplayModeSmoke() As String
    On Error GoTo ErrHandler

    Dim rawText As String
    rawText = "別表第二　身体障害等級及び災害補償表（第七十七条関係）等級災害補償第一級一三四〇日分" & vbCrLf & _
        TABLE_MARKDOWN_MARKER & vbCrLf & _
        "| 等級 | 災害補償 |" & vbCrLf & _
        "| --- | --- |" & vbCrLf & _
        "| 第一級 | 一三四〇日分 |" & vbCrLf & _
        TABLE_ASCII_MARKER & vbCrLf & _
        "+--------+--------------+" & vbCrLf & _
        "| 等級   | 災害補償     |" & vbCrLf & _
        "+--------+--------------+" & vbCrLf & _
        "| 第一級 | 一三四〇日分 |" & vbCrLf & _
        "+--------+--------------+"

    Dim markdownText As String
    markdownText = TableBodyTextForDisplay(rawText, TABLE_DISPLAY_MARKDOWN)
    If InStr(1, markdownText, "| 第一級 | 一三四〇日分 |", vbBinaryCompare) = 0 _
        Or InStr(1, markdownText, TABLE_ASCII_MARKER, vbBinaryCompare) > 0 _
        Or InStr(1, markdownText, "身体障害等級", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 6812, "modLawNavigator.TableDisplayModeSmoke", "Markdown table display extraction mismatch."
    End If

    Dim asciiText As String
    asciiText = TableBodyTextForDisplay(rawText, TABLE_DISPLAY_ASCII)
    If InStr(1, asciiText, "+--------+--------------+", vbBinaryCompare) = 0 _
        Or InStr(1, asciiText, TABLE_MARKDOWN_MARKER, vbBinaryCompare) > 0 _
        Or InStr(1, asciiText, "身体障害等級", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 6813, "modLawNavigator.TableDisplayModeSmoke", "ASCII table display extraction mismatch."
    End If

    TableDisplayModeSmoke = "ok:" & CStr(Len(markdownText)) & ":" & CStr(Len(asciiText))
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawNavigator", "modLawNavigator.TableDisplayModeSmoke", Err.description, "", "", ""
    TableDisplayModeSmoke = "error:" & Err.description
End Function

Public Function BodyNavigatorPerformanceSmokeDefault() As String
    BodyNavigatorPerformanceSmokeDefault = BodyNavigatorPerformanceSmoke(20)
End Function

Public Function BodyNavigatorSmoke() As String
    On Error GoTo ErrHandler

    Dim addedRow As Long
    addedRow = PrepareBodyNavigatorSmokeData()

    If InStr(1, BodyPreviewText(addedRow), "第十一条　受給権は差し押えることができない。", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6801, "modLawNavigator.BodyNavigatorSmoke", "Body preview text mismatch."
    End If
    If InStr(1, BodyPreviewText(addedRow), "別表検索語", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6803, "modLawNavigator.BodyNavigatorSmoke", "Appended table preview text mismatch."
    End If
    If InStr(1, BodyUnitPreviewText(2), "附則側の第十一条", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 6802, "modLawNavigator.BodyNavigatorSmoke", "Unrelated article text was mixed."
    End If
    If InStr(1, BodyUnitPreviewText(4), "第十一条　附則側の第十一条。", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6802, "modLawNavigator.BodyNavigatorSmoke", "Unit preview text mismatch."
    End If

    BodyNavigatorSmoke = "ok:" & CStr(modLawParser.BodyUnitCacheCount("TESTREV001"))
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawNavigator", "modLawNavigator.BodyNavigatorSmoke", Err.description, "", "", ""
    BodyNavigatorSmoke = "error:" & Err.description
End Function

Public Function PrepareBodyNavigatorSmokeData() As Long
    On Error GoTo ErrHandler

    modAddedLawStore.ClearAddedLawStore

    PrepareBodyNavigatorSmokeData = modAddedLawStore.RegisterParsedLawMetadata( _
        "TEST001", _
        "2025-04-01", _
        "TESTREV001", _
        "テスト法", _
        "令和七年法律第一号", _
        "Act", _
        "2025-01-01", _
        modLawRevision.StatusCurrentText(), _
        modLawParser.BodyUnitCacheKey("TESTREV001"), _
        "2026-06-14 00:00:00")

    Dim bodySheet As Worksheet
    Set bodySheet = EnsureBodyUnitSmokeSheet()
    modSheetStore.ClearDataRows bodySheet
    modSheetStore.AppendRow bodySheet, Array("2026-06-14 00:00:00", "TEST001", "TESTREV001", "2025-04-01", "テスト法", "令和七年法律第一号", "Article", "第十一条", "受給権の保護", "", "", "000001", "（受給権の保護）第十一条受給権は差し押えることができない。")
    modSheetStore.AppendRow bodySheet, Array("2026-06-14 00:00:00", "TEST001", "TESTREV001", "2025-04-01", "テスト法", "令和七年法律第一号", "Paragraph", "第十一条", "受給権の保護", "1", "", "000001.0001", "受給権は差し押えることができない。")
    modSheetStore.AppendRow bodySheet, Array("2026-06-14 00:00:00", "TEST001", "TESTREV001", "2025-04-01", "テスト法", "令和七年法律第一号", "Article", "第十一条", "", "", "", "000002", "第十一条附則側の第十一条。２附則側の第二項。")
    modSheetStore.AppendRow bodySheet, Array("2026-06-14 00:00:00", "TEST001", "TESTREV001", "2025-04-01", "テスト法", "令和七年法律第一号", "Paragraph", "第十一条", "", "1", "", "000002.0001", "附則側の第十一条。")
    modSheetStore.AppendRow bodySheet, Array("2026-06-14 00:00:00", "TEST001", "TESTREV001", "2025-04-01", "テスト法", "令和七年法律第一号", "Paragraph", "第十一条", "", "2", "", "000002.0002", "附則側の第二項。")
    modSheetStore.AppendRow bodySheet, Array("2026-06-14 00:00:00", "TEST001", "TESTREV001", "2025-04-01", "テスト法", "令和七年法律第一号", "AppdxTable", "別表第一", "", "", "", "000003", "別表第一　別表検索語を含む簡易テキスト。" & vbCrLf & "Markdownテーブル" & vbCrLf & "| 項目 | 値 |" & vbCrLf & "| --- | --- |" & vbCrLf & "| 別表検索語 | 8 |" & vbCrLf & "ASCII罫線テーブル" & vbCrLf & "+------+-----+" & vbCrLf & "| 項目 | 値  |" & vbCrLf & "+------+-----+" & vbCrLf & "| 別表検索語 | 8 |" & vbCrLf & "+------+-----+")
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawNavigator", "modLawNavigator.PrepareBodyNavigatorSmokeData", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Private Function AddedLawRevisionId(ByVal addedLawStoreRow As Long) As String
    Dim storeSheet As Worksheet
    Set storeSheet = AddedLawSheet()
    If storeSheet Is Nothing Then
        Exit Function
    End If
    If addedLawStoreRow < 2 Or addedLawStoreRow > modSheetStore.LastUsedRow(storeSheet) Then
        Exit Function
    End If
    AddedLawRevisionId = CellText(storeSheet.cells(addedLawStoreRow, ADDED_LAW_COL_LAW_REVISION_ID).Value2)
End Function

Private Function AddedLawHeaderText(ByVal addedLawStoreRow As Long) As String
    Dim storeSheet As Worksheet
    Set storeSheet = AddedLawSheet()
    If storeSheet Is Nothing Then
        Exit Function
    End If
    If addedLawStoreRow < 2 Or addedLawStoreRow > modSheetStore.LastUsedRow(storeSheet) Then
        Exit Function
    End If

    Dim lawTitle As String
    Dim lawNum As String
    lawTitle = CellText(storeSheet.cells(addedLawStoreRow, ADDED_LAW_COL_LAW_TITLE).Value2)
    lawNum = CellText(storeSheet.cells(addedLawStoreRow, ADDED_LAW_COL_LAW_NUM).Value2)

    If Len(lawNum) > 0 Then
        AddedLawHeaderText = lawTitle & vbCrLf & lawNum
    Else
        AddedLawHeaderText = lawTitle
    End If
End Function

Private Function AddedLawSheet() As Worksheet
    If modSheetStore.SheetExists(ThisWorkbook, modAddedLawStore.AddedLawStoreSheetName()) Then
        Set AddedLawSheet = ThisWorkbook.Worksheets(modAddedLawStore.AddedLawStoreSheetName())
    End If
End Function

Private Function EnsureBodyUnitSmokeSheet() As Worksheet
    Dim bodySheet As Worksheet
    If modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())
    Else
        Set bodySheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        bodySheet.Name = modLawParser.BodyUnitSheetName()
    End If

    bodySheet.cells.NumberFormat = "@"
    If Application.CountA(bodySheet.rows(1)) = 0 Then
        WriteBodyUnitHeaders bodySheet
    End If
    bodySheet.Visible = xlSheetVeryHidden
    Set EnsureBodyUnitSmokeSheet = bodySheet
End Function

Private Sub WriteBodyUnitHeaders(ByVal bodySheet As Worksheet)
    Dim headers As Variant
    headers = Array("CachedAt", "LawId", "LawRevisionId", "EnforcementDate", "LawTitle", "LawNum", "UnitKind", "ArticleTitle", "ArticleCaption", "ParagraphNum", "ItemTitle", "SortKey", "Text")

    Dim index As Long
    For index = LBound(headers) To UBound(headers)
        bodySheet.cells(1, index - LBound(headers) + 1).Value2 = headers(index)
    Next index
    bodySheet.rows(1).Font.Bold = True
End Sub

Private Function FormattedBodyUnitText(ByRef bodyValues As Variant, ByVal rowIndex As Long, ByVal articleIndex As Object, Optional ByVal tableDisplayMode As String = TABLE_DISPLAY_MARKDOWN) As String
    Dim unitKind As String
    unitKind = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_UNIT_KIND)

    Select Case unitKind
        Case "Article"
            FormattedBodyUnitText = FormattedArticleText(bodyValues, rowIndex, articleIndex)
        Case "Paragraph"
            FormattedBodyUnitText = FormatParagraphLine(bodyValues, rowIndex, BodyUnitHeading(bodyValues, rowIndex))
        Case "Item"
            FormattedBodyUnitText = BodyUnitHeading(bodyValues, rowIndex) & vbCrLf & BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_TEXT)
        Case "AppdxTable", "Appdx", "SupplementaryProvision"
            FormattedBodyUnitText = FormattedStandaloneBodyText(bodyValues, rowIndex, tableDisplayMode)
        Case Else
            FormattedBodyUnitText = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_TEXT)
    End Select
End Function

Private Function FormattedStandaloneBodyText(ByRef bodyValues As Variant, ByVal rowIndex As Long, Optional ByVal tableDisplayMode As String = TABLE_DISPLAY_MARKDOWN) As String
    Dim heading As String
    heading = BodyUnitHeading(bodyValues, rowIndex)

    Dim bodyText As String
    bodyText = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_TEXT)
    If StrComp(BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_UNIT_KIND), "AppdxTable", vbBinaryCompare) = 0 Then
        bodyText = TableBodyTextForDisplay(bodyText, tableDisplayMode)
    End If

    If Len(heading) = 0 Then
        FormattedStandaloneBodyText = bodyText
    ElseIf Len(bodyText) = 0 Then
        FormattedStandaloneBodyText = heading
    Else
        FormattedStandaloneBodyText = heading & vbCrLf & bodyText
    End If
End Function

Private Function TableBodyTextForDisplay(ByVal rawText As String, ByVal tableDisplayMode As String) As String
    Dim markdownText As String
    markdownText = TextBetweenMarkers(rawText, TABLE_MARKDOWN_MARKER, TABLE_ASCII_MARKER)

    Dim asciiText As String
    asciiText = TextAfterMarker(rawText, TABLE_ASCII_MARKER)

    If Len(markdownText) = 0 And Len(asciiText) = 0 Then
        TableBodyTextForDisplay = rawText
    ElseIf StrComp(tableDisplayMode, TABLE_DISPLAY_ASCII, vbTextCompare) = 0 Then
        If Len(asciiText) > 0 Then
            TableBodyTextForDisplay = asciiText
        Else
            TableBodyTextForDisplay = markdownText
        End If
    ElseIf Len(markdownText) > 0 Then
        TableBodyTextForDisplay = markdownText
    Else
        TableBodyTextForDisplay = asciiText
    End If
End Function

Private Function TextBetweenMarkers(ByVal rawText As String, ByVal startMarker As String, ByVal endMarker As String) As String
    Dim startPos As Long
    startPos = InStr(1, rawText, startMarker, vbBinaryCompare)
    If startPos = 0 Then
        Exit Function
    End If

    startPos = startPos + Len(startMarker)

    Dim endPos As Long
    endPos = InStr(startPos, rawText, endMarker, vbBinaryCompare)
    If endPos = 0 Then
        TextBetweenMarkers = TrimBlock(Mid$(rawText, startPos))
    Else
        TextBetweenMarkers = TrimBlock(Mid$(rawText, startPos, endPos - startPos))
    End If
End Function

Private Function TextAfterMarker(ByVal rawText As String, ByVal marker As String) As String
    Dim startPos As Long
    startPos = InStr(1, rawText, marker, vbBinaryCompare)
    If startPos = 0 Then
        Exit Function
    End If

    TextAfterMarker = TrimBlock(Mid$(rawText, startPos + Len(marker)))
End Function

Private Function TrimBlock(ByVal value As String) As String
    value = Trim$(value)
    Do While Left$(value, 2) = vbCrLf
        value = Mid$(value, 3)
    Loop
    Do While Right$(value, 2) = vbCrLf
        value = Left$(value, Len(value) - 2)
    Loop
    TrimBlock = Trim$(value)
End Function

Private Function FormattedArticleText(ByRef bodyValues As Variant, ByVal articleRow As Long, ByVal articleIndex As Object) As String
    Dim articleTitle As String
    Dim articleCaption As String
    Dim articleSortKey As String
    articleTitle = BodyValueText(bodyValues, articleRow, BODY_UNIT_COL_ARTICLE_TITLE)
    articleCaption = BodyValueText(bodyValues, articleRow, BODY_UNIT_COL_ARTICLE_CAPTION)
    articleSortKey = BodyValueText(bodyValues, articleRow, BODY_UNIT_COL_SORT_KEY)

    Dim result As String
    If Len(articleCaption) > 0 Then
        result = ArticleCaptionDisplay(articleCaption)
    End If

    Dim paragraphRows As Collection
    Set paragraphRows = IndexedRows(articleIndex, "Paragraph|" & articleSortKey)

    Dim rowIndex As Variant
    Dim paragraphOrdinal As Long
    For Each rowIndex In paragraphRows
        paragraphOrdinal = paragraphOrdinal + 1
        result = AppendLine(result, FormatArticleParagraphLine(bodyValues, CLng(rowIndex), articleTitle, paragraphOrdinal))
        result = AppendArticleItemLines(result, bodyValues, articleIndex, BodyValueText(bodyValues, CLng(rowIndex), BODY_UNIT_COL_SORT_KEY))
    Next rowIndex

    If paragraphOrdinal = 0 Then
        result = AppendLine(result, FormatArticleFallbackLine(articleTitle, articleCaption, BodyValueText(bodyValues, articleRow, BODY_UNIT_COL_TEXT)))
    End If

    FormattedArticleText = result
End Function

Private Function AppendArticleItemLines(ByVal baseText As String, ByRef bodyValues As Variant, ByVal articleIndex As Object, ByVal paragraphSortKey As String) As String
    Dim result As String
    result = baseText

    Dim itemRows As Collection
    Set itemRows = IndexedRows(articleIndex, "ItemParent|" & paragraphSortKey)

    Dim rowIndex As Variant
    For Each rowIndex In itemRows
        result = AppendLine(result, FormatItemLine(bodyValues, CLng(rowIndex)))
    Next rowIndex

    AppendArticleItemLines = result
End Function

Private Function FormatArticleParagraphLine(ByRef bodyValues As Variant, ByVal rowIndex As Long, ByVal articleTitle As String, ByVal paragraphOrdinal As Long) As String
    Dim paragraphNum As String
    paragraphNum = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_PARAGRAPH_NUM)

    Dim marker As String
    If paragraphOrdinal = 1 Or Len(paragraphNum) = 0 Or StrComp(paragraphNum, "1", vbTextCompare) = 0 Then
        marker = articleTitle
    Else
        marker = DisplayParagraphNumber(paragraphNum)
    End If

    FormatArticleParagraphLine = marker & "　" & BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_TEXT)
End Function

Private Function FormatParagraphLine(ByRef bodyValues As Variant, ByVal rowIndex As Long, ByVal heading As String) As String
    FormatParagraphLine = heading & vbCrLf & BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_TEXT)
End Function

Private Function FormatItemLine(ByRef bodyValues As Variant, ByVal rowIndex As Long) As String
    Dim itemTitle As String
    itemTitle = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_ITEM_TITLE)
    If Len(itemTitle) = 0 Then
        itemTitle = UnitKindDisplay("Item")
    End If
    FormatItemLine = itemTitle & "　" & BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_TEXT)
End Function

Private Function FormatArticleFallbackLine(ByVal articleTitle As String, ByVal articleCaption As String, ByVal rawText As String) As String
    Dim bodyText As String
    bodyText = Trim$(rawText)
    bodyText = RemoveLeadingText(bodyText, ArticleCaptionDisplay(articleCaption))
    bodyText = RemoveLeadingText(bodyText, articleCaption)
    bodyText = RemoveLeadingText(bodyText, articleTitle)

    If Len(bodyText) = 0 Then
        FormatArticleFallbackLine = BodyUnitFallbackHeading(articleTitle, articleCaption)
    ElseIf Len(articleTitle) > 0 Then
        FormatArticleFallbackLine = articleTitle & "　" & bodyText
    Else
        FormatArticleFallbackLine = bodyText
    End If
End Function

Private Function BodyUnitFallbackHeading(ByVal articleTitle As String, ByVal articleCaption As String) As String
    If Len(articleCaption) > 0 Then
        BodyUnitFallbackHeading = ArticleCaptionDisplay(articleCaption)
    End If
    If Len(articleTitle) > 0 Then
        BodyUnitFallbackHeading = AppendLine(BodyUnitFallbackHeading, articleTitle)
    End If
End Function

Private Function BodyUnitHeading(ByRef bodyValues As Variant, ByVal rowIndex As Long) As String
    Dim unitKind As String
    Dim articleTitle As String
    Dim articleCaption As String
    Dim paragraphNum As String
    Dim itemTitle As String

    unitKind = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_UNIT_KIND)
    articleTitle = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_ARTICLE_TITLE)
    articleCaption = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_ARTICLE_CAPTION)
    paragraphNum = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_PARAGRAPH_NUM)
    itemTitle = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_ITEM_TITLE)

    Dim heading As String
    heading = articleTitle
    If Len(articleCaption) > 0 Then
        heading = AppendHeadingPart(heading, ArticleCaptionDisplay(articleCaption))
    End If
    If Len(paragraphNum) > 0 Then
        heading = AppendHeadingPart(heading, "第" & paragraphNum & "項")
    End If
    If Len(itemTitle) > 0 Then
        heading = AppendHeadingPart(heading, itemTitle)
    End If
    If Len(heading) = 0 Then
        heading = UnitKindDisplay(unitKind)
    End If

    BodyUnitHeading = heading
End Function

Private Function ShouldShowInBodyNav(ByVal unitKind As String) As Boolean
    ShouldShowInBodyNav = (StrComp(unitKind, "Article", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "LawText", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "AppdxTable", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "Appdx", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "SupplementaryProvision", vbBinaryCompare) = 0)
End Function

Private Sub LoadBodyCacheValues(ByVal bodySheet As Worksheet, ByVal lastRow As Long, ByRef bodyValues As Variant)
    If lastRow < 2 Then
        bodyValues = Empty
        Exit Sub
    End If

    bodyValues = bodySheet.Range(bodySheet.cells(1, 1), bodySheet.cells(lastRow, BODY_UNIT_COL_COUNT)).Value2
End Sub

Private Function RowsForRevision(ByRef bodyValues As Variant, ByVal lawRevisionId As String) As Collection
    Dim rows As Collection
    Set rows = New Collection

    If Not IsArray(bodyValues) Then
        Set RowsForRevision = rows
        Exit Function
    End If

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(bodyValues, 1)
        If StrComp(BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_LAW_REVISION_ID), lawRevisionId, vbTextCompare) = 0 Then
            rows.Add rowIndex
        End If
    Next rowIndex

    Set RowsForRevision = rows
End Function

Private Function BuildArticleUnitIndex(ByRef bodyValues As Variant, ByVal revisionRows As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbBinaryCompare

    Dim rowIndex As Variant
    Dim unitKind As String
    Dim sortKey As String
    Dim articleKey As String
    Dim parentKey As String

    For Each rowIndex In revisionRows
        unitKind = BodyValueText(bodyValues, CLng(rowIndex), BODY_UNIT_COL_UNIT_KIND)
        sortKey = BodyValueText(bodyValues, CLng(rowIndex), BODY_UNIT_COL_SORT_KEY)
        articleKey = ArticleKeyFromSortKey(sortKey)
        If Len(unitKind) > 0 And Len(articleKey) > 0 Then
            AddIndexRow result, unitKind & "|" & articleKey, CLng(rowIndex)
        End If
        If StrComp(unitKind, "Item", vbBinaryCompare) = 0 Then
            parentKey = ParentSortKey(sortKey)
            If Len(parentKey) > 0 Then
                AddIndexRow result, "ItemParent|" & parentKey, CLng(rowIndex)
            End If
        End If
    Next rowIndex

    Set BuildArticleUnitIndex = result
End Function

Private Sub AddIndexRow(ByVal targetIndex As Object, ByVal indexKey As String, ByVal rowIndex As Long)
    Dim rows As Collection
    If targetIndex.Exists(indexKey) Then
        Set rows = targetIndex(indexKey)
    Else
        Set rows = New Collection
        targetIndex.Add indexKey, rows
    End If
    rows.Add rowIndex
End Sub

Private Function IndexedRows(ByVal targetIndex As Object, ByVal indexKey As String) As Collection
    If Not targetIndex Is Nothing Then
        If targetIndex.Exists(indexKey) Then
            Set IndexedRows = targetIndex(indexKey)
            Exit Function
        End If
    End If
    Set IndexedRows = New Collection
End Function

Private Function ArticleKeyFromSortKey(ByVal sortKey As String) As String
    Dim separatorPosition As Long
    separatorPosition = InStr(1, sortKey, ".", vbBinaryCompare)
    If separatorPosition > 0 Then
        ArticleKeyFromSortKey = Left$(sortKey, separatorPosition - 1)
    Else
        ArticleKeyFromSortKey = sortKey
    End If
End Function

Private Function ParentSortKey(ByVal sortKey As String) As String
    Dim separatorPosition As Long
    separatorPosition = InStrRev(sortKey, ".")
    If separatorPosition > 0 Then
        ParentSortKey = Left$(sortKey, separatorPosition - 1)
    Else
        ParentSortKey = ""
    End If
End Function

Private Function ArticleCaptionDisplay(ByVal articleCaption As String) As String
    articleCaption = Trim$(articleCaption)
    If Len(articleCaption) = 0 Then
        ArticleCaptionDisplay = ""
    ElseIf Left$(articleCaption, 1) = "（" Or Left$(articleCaption, 1) = "(" Then
        ArticleCaptionDisplay = articleCaption
    Else
        ArticleCaptionDisplay = "（" & articleCaption & "）"
    End If
End Function

Private Function DisplayParagraphNumber(ByVal paragraphNum As String) As String
    DisplayParagraphNumber = ToFullWidthDigits(paragraphNum)
End Function

Private Function ToFullWidthDigits(ByVal value As String) As String
    value = Replace(value, "0", "０")
    value = Replace(value, "1", "１")
    value = Replace(value, "2", "２")
    value = Replace(value, "3", "３")
    value = Replace(value, "4", "４")
    value = Replace(value, "5", "５")
    value = Replace(value, "6", "６")
    value = Replace(value, "7", "７")
    value = Replace(value, "8", "８")
    value = Replace(value, "9", "９")
    ToFullWidthDigits = value
End Function

Private Function UnitKindDisplay(ByVal unitKind As String) As String
    Select Case unitKind
        Case "Article"
            UnitKindDisplay = "条"
        Case "Paragraph"
            UnitKindDisplay = "項"
        Case "Item"
            UnitKindDisplay = "号"
        Case "LawText"
            UnitKindDisplay = "本文"
        Case "AppdxTable"
            UnitKindDisplay = "別表"
        Case "Appdx"
            UnitKindDisplay = "別表等"
        Case "SupplementaryProvision"
            UnitKindDisplay = "附則"
        Case Else
            UnitKindDisplay = unitKind
    End Select
End Function

Private Function AppendHeadingPart(ByVal baseText As String, ByVal partText As String) As String
    If Len(baseText) = 0 Then
        AppendHeadingPart = partText
    ElseIf Len(partText) = 0 Then
        AppendHeadingPart = baseText
    Else
        AppendHeadingPart = baseText & " " & partText
    End If
End Function

Private Function AppendParagraph(ByVal baseText As String, ByVal partText As String) As String
    If Len(baseText) = 0 Then
        AppendParagraph = partText
    ElseIf Len(partText) = 0 Then
        AppendParagraph = baseText
    Else
        AppendParagraph = baseText & vbCrLf & vbCrLf & partText
    End If
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

Private Function RemoveLeadingText(ByVal baseText As String, ByVal leadingText As String) As String
    If Len(leadingText) > 0 And Left$(baseText, Len(leadingText)) = leadingText Then
        RemoveLeadingText = Trim$(Mid$(baseText, Len(leadingText) + 1))
    Else
        RemoveLeadingText = baseText
    End If
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function

Private Function BodyValueText(ByRef bodyValues As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If Not IsArray(bodyValues) Then
        BodyValueText = ""
    ElseIf rowIndex < LBound(bodyValues, 1) Or rowIndex > UBound(bodyValues, 1) _
        Or columnIndex < LBound(bodyValues, 2) Or columnIndex > UBound(bodyValues, 2) Then
        BodyValueText = ""
    Else
        BodyValueText = CellText(bodyValues(rowIndex, columnIndex))
    End If
End Function

Private Function ElapsedMilliseconds(ByVal startedAt As Single, ByVal endedAt As Single) As Double
    If endedAt < startedAt Then
        endedAt = endedAt + 86400!
    End If
    ElapsedMilliseconds = CDbl(endedAt - startedAt) * 1000#
End Function
