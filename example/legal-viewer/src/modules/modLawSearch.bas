Attribute VB_Name = "modLawSearch"
Option Explicit

Private Const SHEET_LAW_LIST As String = "_lv_tmp_law_list"
Private Const SHEET_SEARCH_RESULTS As String = "_lv_tmp_law_search_results"
Private Const DEFAULT_MAX_RESULTS As Long = 50
Private Const LAW_LIST_PAGE_LIMIT As Long = 10000
Private Const LAW_LIST_MAX_PAGES As Long = 20
Private Const LAW_LIST_COLUMN_COUNT As Long = 14
Private Const SEARCH_RESULT_COLUMN_COUNT As Long = 14

Public Function RefreshLawListCache(Optional ByVal maxItems As Long = 0) As Long
    On Error GoTo ErrHandler

    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalCalculation As XlCalculation
    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents
    originalCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim cacheSheet As Worksheet
    Set cacheSheet = EnsureTempSheet(ThisWorkbook, SHEET_LAW_LIST, LawListHeaders())
    modSheetStore.ClearDataRows cacheSheet

    Dim offset As Long
    Dim pageCount As Long
    Dim cachedCount As Long
    Dim cachedAt As String
    cachedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Do
        pageCount = pageCount + 1
        If pageCount > LAW_LIST_MAX_PAGES Then
            Err.Raise vbObjectError + 6202, "modLawSearch.RefreshLawListCache", "Law list paging exceeded safety limit."
        End If

        Dim responseJson As String
        responseJson = modApiClient.ApiGetLaws(PageLimit(maxItems, cachedCount), offset)

        Dim pageRows As Variant
        Dim pageRowCount As Long
        pageRowCount = ExtractLawListRows(responseJson, cachedCount + 1, maxItems, cachedAt, pageRows)
        AppendRows cacheSheet, pageRows, pageRowCount, LAW_LIST_COLUMN_COUNT
        cachedCount = cachedCount + pageRowCount

        If maxItems > 0 And cachedCount >= maxItems Then
            Exit Do
        End If

        Dim nextOffset As Long
        nextOffset = JsonRootLongProperty(responseJson, "next_offset", 0)
        If nextOffset <= 0 Or nextOffset <= offset Then
            Exit Do
        End If

        offset = nextOffset
    Loop

    RefreshLawListCache = cachedCount
    GoTo Cleanup

ErrHandler:
    modLogger.LogErrorSafe "LawSearch", "modLawSearch.RefreshLawListCache", Err.description, "", "", "laws cache refresh"
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    Err.Raise Err.Number, Err.source, Err.description

Cleanup:
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
End Function

Private Function PageLimit(ByVal maxItems As Long, ByVal cachedCount As Long) As Long
    PageLimit = LAW_LIST_PAGE_LIMIT
    If maxItems > 0 And maxItems - cachedCount < PageLimit Then
        PageLimit = maxItems - cachedCount
    End If
    If PageLimit <= 0 Then
        PageLimit = LAW_LIST_PAGE_LIMIT
    End If
End Function

Public Function LawListCacheCount() As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_LAW_LIST) Then
        LawListCacheCount = 0
        Exit Function
    End If

    Dim cacheSheet As Worksheet
    Set cacheSheet = ThisWorkbook.Worksheets(SHEET_LAW_LIST)
    LawListCacheCount = Application.Max(0, modSheetStore.LastUsedRow(cacheSheet) - 1)
End Function

Public Function SearchLaws(ByVal keyword As String, Optional ByVal maxResults As Long = DEFAULT_MAX_RESULTS) As Long
    If LawListCacheCount() > 0 Then
        SearchLaws = SearchCachedLaws(keyword, maxResults)
    Else
        SearchLaws = SearchRemoteLaws(keyword, maxResults)
    End If
End Function

Public Sub ConfigureLawResultListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 7
    targetListBox.columnWidths = "180 pt;95 pt;105 pt;70 pt;0 pt;0 pt;0 pt"
    targetListBox.BoundColumn = 3
End Sub

Public Function SearchLawsForListBox(ByVal keyword As String, ByVal targetListBox As Object, Optional ByVal maxResults As Long = DEFAULT_MAX_RESULTS) As Long
    SearchLaws keyword, maxResults
    SearchLawsForListBox = PopulateSearchResultsListBox(targetListBox, maxResults)
End Function

Public Function SearchRemoteLaws(ByVal keyword As String, Optional ByVal maxResults As Long = DEFAULT_MAX_RESULTS) As Long
    On Error GoTo ErrHandler

    Dim searchKeyword As String
    searchKeyword = ResolveAlias(Trim$(keyword))
    If Len(searchKeyword) = 0 Then
        SearchRemoteLaws = 0
        Exit Function
    End If

    If maxResults <= 0 Then
        maxResults = DEFAULT_MAX_RESULTS
    End If

    Dim resultSheet As Worksheet
    Set resultSheet = EnsureTempSheet(ThisWorkbook, SHEET_SEARCH_RESULTS, SearchResultHeaders())
    modSheetStore.ClearDataRows resultSheet

    Dim resultRows As Variant
    ReDim resultRows(1 To maxResults, 1 To SEARCH_RESULT_COLUMN_COUNT)

    Dim seenKeys As Object
    Set seenKeys = CreateObject("Scripting.Dictionary")

    Dim hitCount As Long
    hitCount = AppendRemoteLawSearch(modApiClient.ApiGetLaws(maxResults, 0, searchKeyword), searchKeyword, resultRows, hitCount, maxResults, seenKeys)

    If hitCount < maxResults And LooksLikeLawNum(searchKeyword) Then
        hitCount = AppendRemoteLawSearch(modApiClient.ApiGetLaws(maxResults, 0, "", searchKeyword), searchKeyword, resultRows, hitCount, maxResults, seenKeys)
    End If

    If hitCount < maxResults And LooksLikeLawId(searchKeyword) Then
        hitCount = AppendRemoteLawSearch(modApiClient.ApiGetLaws(maxResults, 0, "", "", searchKeyword), searchKeyword, resultRows, hitCount, maxResults, seenKeys)
    End If

    AppendRows resultSheet, resultRows, hitCount, SEARCH_RESULT_COLUMN_COUNT
    SearchRemoteLaws = hitCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawSearch", "modLawSearch.SearchRemoteLaws", Err.description, "", "", keyword
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function PopulateSearchResultsListBox(ByVal targetListBox As Object, Optional ByVal maxRows As Long = DEFAULT_MAX_RESULTS) As Long
    On Error GoTo ErrHandler

    ConfigureLawResultListBox targetListBox
    targetListBox.Clear

    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_SEARCH_RESULTS) Then
        PopulateSearchResultsListBox = 0
        Exit Function
    End If

    If maxRows <= 0 Then
        maxRows = DEFAULT_MAX_RESULTS
    End If

    Dim resultSheet As Worksheet
    Set resultSheet = ThisWorkbook.Worksheets(SHEET_SEARCH_RESULTS)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(resultSheet)

    Dim rowIndex As Long
    Dim listIndex As Long
    Dim loadedCount As Long
    For rowIndex = 2 To lastRow
        targetListBox.AddItem CellText(resultSheet.cells(rowIndex, 6).Value2)
        listIndex = targetListBox.ListCount - 1
        targetListBox.List(listIndex, 1) = CellText(resultSheet.cells(rowIndex, 5).Value2)
        targetListBox.List(listIndex, 2) = CellText(resultSheet.cells(rowIndex, 2).Value2)
        targetListBox.List(listIndex, 3) = CellText(resultSheet.cells(rowIndex, 10).Value2)
        targetListBox.List(listIndex, 4) = CellText(resultSheet.cells(rowIndex, 3).Value2)
        targetListBox.List(listIndex, 5) = CellText(resultSheet.cells(rowIndex, 11).Value2)
        targetListBox.List(listIndex, 6) = CellText(resultSheet.cells(rowIndex, 13).Value2)

        loadedCount = loadedCount + 1
        If loadedCount >= maxRows Then
            Exit For
        End If
    Next rowIndex

    PopulateSearchResultsListBox = loadedCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawSearch", "modLawSearch.PopulateSearchResultsListBox", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function SearchCachedLaws(ByVal keyword As String, Optional ByVal maxResults As Long = DEFAULT_MAX_RESULTS) As Long
    On Error GoTo ErrHandler

    Dim searchKeyword As String
    searchKeyword = ResolveAlias(Trim$(keyword))
    If Len(searchKeyword) = 0 Then
        SearchCachedLaws = 0
        Exit Function
    End If

    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_LAW_LIST) Then
        SearchCachedLaws = 0
        Exit Function
    End If

    If maxResults <= 0 Then
        maxResults = DEFAULT_MAX_RESULTS
    End If

    Dim cacheSheet As Worksheet
    Set cacheSheet = ThisWorkbook.Worksheets(SHEET_LAW_LIST)

    Dim resultSheet As Worksheet
    Set resultSheet = EnsureTempSheet(ThisWorkbook, SHEET_SEARCH_RESULTS, SearchResultHeaders())
    modSheetStore.ClearDataRows resultSheet

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(cacheSheet)
    If lastRow < 2 Then
        SearchCachedLaws = 0
        Exit Function
    End If

    Dim cacheValues As Variant
    cacheValues = cacheSheet.Range(cacheSheet.cells(2, 1), cacheSheet.cells(lastRow, LAW_LIST_COLUMN_COUNT)).Value2

    Dim resultRows As Variant
    ReDim resultRows(1 To maxResults, 1 To SEARCH_RESULT_COLUMN_COUNT)

    Dim rowIndex As Long
    Dim hitCount As Long
    For rowIndex = 1 To UBound(cacheValues, 1)
        If LawCacheValuesMatch(cacheValues, rowIndex, searchKeyword) Then
            hitCount = hitCount + 1
            FillSearchResultRow resultRows, hitCount, cacheValues, rowIndex, searchKeyword
            If hitCount >= maxResults Then
                Exit For
            End If
        End If
    Next rowIndex

    AppendRows resultSheet, resultRows, hitCount, SEARCH_RESULT_COLUMN_COUNT
    SearchCachedLaws = hitCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawSearch", "modLawSearch.SearchCachedLaws", Err.description, "", "", keyword
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function LawSearchSmoke(ByVal keyword As String) As String
    On Error GoTo ErrHandler

    Dim cachedCount As Long
    cachedCount = RefreshLawListCache()

    Dim hitCount As Long
    hitCount = SearchCachedLaws(keyword, DEFAULT_MAX_RESULTS)

    LawSearchSmoke = "ok:" & CStr(cachedCount) & ":" & CStr(hitCount)
    Exit Function

ErrHandler:
    LawSearchSmoke = "error:" & Err.description
End Function

Public Function LawSearchPagingSmoke() As String
    On Error GoTo ErrHandler

    Dim cachedCount As Long
    cachedCount = RefreshLawListCache()
    If cachedCount <= 100 Then
        Err.Raise vbObjectError + 6211, "modLawSearch.LawSearchPagingSmoke", "Law list cache did not exceed the old 100 item limit."
    End If

    Dim hitCount As Long
    hitCount = SearchCachedLaws("民法", DEFAULT_MAX_RESULTS)
    If hitCount <= 0 Then
        Err.Raise vbObjectError + 6212, "modLawSearch.LawSearchPagingSmoke", "Expected law search hit was not found."
    End If

    LawSearchPagingSmoke = "ok:" & CStr(cachedCount) & ":" & CStr(hitCount)
    Exit Function

ErrHandler:
    LawSearchPagingSmoke = "error:" & Err.description
End Function

Public Function LawSearchRemoteSmoke() As String
    On Error GoTo ErrHandler

    Dim hitCount As Long
    hitCount = SearchRemoteLaws("民法", DEFAULT_MAX_RESULTS)
    If hitCount <= 0 Then
        Err.Raise vbObjectError + 6213, "modLawSearch.LawSearchRemoteSmoke", "Expected remote law search hit was not found."
    End If

    LawSearchRemoteSmoke = "ok:" & CStr(hitCount) & ":" & CStr(LawListCacheCount())
    Exit Function

ErrHandler:
    LawSearchRemoteSmoke = "error:" & Err.description
End Function

Private Function ResolveAlias(ByVal keyword As String) As String
    ResolveAlias = keyword
    If Len(keyword) = 0 Then
        Exit Function
    End If

    Dim aliasSheet As Worksheet
    Set aliasSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.AliasMasterSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(aliasSheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If StrComp(CStr(aliasSheet.cells(rowIndex, 1).Value2), keyword, vbTextCompare) = 0 Then
            ResolveAlias = CStr(aliasSheet.cells(rowIndex, 2).Value2)
            Exit Function
        End If
    Next rowIndex
End Function

Private Function LawCacheRowMatches(ByVal cacheSheet As Worksheet, ByVal rowIndex As Long, ByVal keyword As String) As Boolean
    LawCacheRowMatches = ContainsText(cacheSheet.cells(rowIndex, 3).Value2, keyword) _
        Or ContainsText(cacheSheet.cells(rowIndex, 4).Value2, keyword) _
        Or ContainsText(cacheSheet.cells(rowIndex, 6).Value2, keyword) _
        Or ContainsText(cacheSheet.cells(rowIndex, 7).Value2, keyword) _
        Or ContainsText(cacheSheet.cells(rowIndex, 8).Value2, keyword) _
        Or ContainsText(cacheSheet.cells(rowIndex, 9).Value2, keyword) _
        Or ContainsText(cacheSheet.cells(rowIndex, 10).Value2, keyword)
End Function

Private Function LawCacheValuesMatch(ByRef cacheValues As Variant, ByVal rowIndex As Long, ByVal keyword As String) As Boolean
    LawCacheValuesMatch = ContainsText(cacheValues(rowIndex, 3), keyword) _
        Or ContainsText(cacheValues(rowIndex, 4), keyword) _
        Or ContainsText(cacheValues(rowIndex, 6), keyword) _
        Or ContainsText(cacheValues(rowIndex, 7), keyword) _
        Or ContainsText(cacheValues(rowIndex, 8), keyword) _
        Or ContainsText(cacheValues(rowIndex, 9), keyword) _
        Or ContainsText(cacheValues(rowIndex, 10), keyword)
End Function

Private Function LooksLikeLawId(ByVal value As String) As Boolean
    If Len(value) < 8 Then
        Exit Function
    End If

    Dim index As Long
    For index = 1 To Len(value)
        Dim codePoint As Long
        codePoint = AscW(Mid$(value, index, 1))
        If Not ((codePoint >= 48 And codePoint <= 57) _
            Or (codePoint >= 65 And codePoint <= 90) _
            Or (codePoint >= 97 And codePoint <= 122)) Then
            Exit Function
        End If
    Next index

    LooksLikeLawId = True
End Function

Private Function LooksLikeLawNum(ByVal value As String) As Boolean
    LooksLikeLawNum = (InStr(1, value, "第", vbTextCompare) > 0 And InStr(1, value, "号", vbTextCompare) > 0) _
        Or InStr(1, value, "法律", vbTextCompare) > 0 _
        Or InStr(1, value, "政令", vbTextCompare) > 0 _
        Or InStr(1, value, "省令", vbTextCompare) > 0 _
        Or InStr(1, value, "府令", vbTextCompare) > 0 _
        Or InStr(1, value, "規則", vbTextCompare) > 0 _
        Or InStr(1, value, "明治", vbTextCompare) > 0 _
        Or InStr(1, value, "大正", vbTextCompare) > 0 _
        Or InStr(1, value, "昭和", vbTextCompare) > 0 _
        Or InStr(1, value, "平成", vbTextCompare) > 0 _
        Or InStr(1, value, "令和", vbTextCompare) > 0
End Function

Private Function ContainsText(ByVal value As Variant, ByVal keyword As String) As Boolean
    If Len(keyword) = 0 Then
        ContainsText = False
    Else
        ContainsText = (InStr(1, CStr(value), keyword, vbTextCompare) > 0)
    End If
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function

Private Sub AppendSearchResultRow(ByVal resultSheet As Worksheet, ByVal cacheSheet As Worksheet, ByVal cacheRowIndex As Long, ByVal hitIndex As Long, ByVal keyword As String)
    modSheetStore.AppendRow resultSheet, Array( _
        hitIndex, _
        cacheSheet.cells(cacheRowIndex, 3).Value2, _
        cacheSheet.cells(cacheRowIndex, 4).Value2, _
        cacheSheet.cells(cacheRowIndex, 5).Value2, _
        cacheSheet.cells(cacheRowIndex, 6).Value2, _
        cacheSheet.cells(cacheRowIndex, 7).Value2, _
        cacheSheet.cells(cacheRowIndex, 8).Value2, _
        cacheSheet.cells(cacheRowIndex, 9).Value2, _
        cacheSheet.cells(cacheRowIndex, 10).Value2, _
        cacheSheet.cells(cacheRowIndex, 11).Value2, _
        cacheSheet.cells(cacheRowIndex, 12).Value2, _
        cacheSheet.cells(cacheRowIndex, 13).Value2, _
        cacheSheet.cells(cacheRowIndex, 14).Value2, _
        keyword)
End Sub

Private Sub FillSearchResultRow(ByRef resultRows As Variant, ByVal hitIndex As Long, ByRef cacheValues As Variant, ByVal cacheRowIndex As Long, ByVal keyword As String)
    resultRows(hitIndex, 1) = hitIndex
    resultRows(hitIndex, 2) = cacheValues(cacheRowIndex, 3)
    resultRows(hitIndex, 3) = cacheValues(cacheRowIndex, 4)
    resultRows(hitIndex, 4) = cacheValues(cacheRowIndex, 5)
    resultRows(hitIndex, 5) = cacheValues(cacheRowIndex, 6)
    resultRows(hitIndex, 6) = cacheValues(cacheRowIndex, 7)
    resultRows(hitIndex, 7) = cacheValues(cacheRowIndex, 8)
    resultRows(hitIndex, 8) = cacheValues(cacheRowIndex, 9)
    resultRows(hitIndex, 9) = cacheValues(cacheRowIndex, 10)
    resultRows(hitIndex, 10) = cacheValues(cacheRowIndex, 11)
    resultRows(hitIndex, 11) = cacheValues(cacheRowIndex, 12)
    resultRows(hitIndex, 12) = cacheValues(cacheRowIndex, 13)
    resultRows(hitIndex, 13) = cacheValues(cacheRowIndex, 14)
    resultRows(hitIndex, 14) = keyword
End Sub

Private Sub FillSearchResultFromLawRow(ByRef resultRows As Variant, ByVal hitIndex As Long, ByRef lawRows As Variant, ByVal lawRowIndex As Long, ByVal keyword As String)
    resultRows(hitIndex, 1) = hitIndex
    resultRows(hitIndex, 2) = lawRows(lawRowIndex, 3)
    resultRows(hitIndex, 3) = lawRows(lawRowIndex, 4)
    resultRows(hitIndex, 4) = lawRows(lawRowIndex, 5)
    resultRows(hitIndex, 5) = lawRows(lawRowIndex, 6)
    resultRows(hitIndex, 6) = lawRows(lawRowIndex, 7)
    resultRows(hitIndex, 7) = lawRows(lawRowIndex, 8)
    resultRows(hitIndex, 8) = lawRows(lawRowIndex, 9)
    resultRows(hitIndex, 9) = lawRows(lawRowIndex, 10)
    resultRows(hitIndex, 10) = lawRows(lawRowIndex, 11)
    resultRows(hitIndex, 11) = lawRows(lawRowIndex, 12)
    resultRows(hitIndex, 12) = lawRows(lawRowIndex, 13)
    resultRows(hitIndex, 13) = lawRows(lawRowIndex, 14)
    resultRows(hitIndex, 14) = keyword
End Sub

Private Function AppendRemoteLawSearch( _
    ByVal responseJson As String, _
    ByVal searchKeyword As String, _
    ByRef resultRows As Variant, _
    ByVal currentHitCount As Long, _
    ByVal maxResults As Long, _
    ByVal seenKeys As Object) As Long

    Dim lawRows As Variant
    Dim rowCount As Long
    rowCount = ExtractLawListRows(responseJson, 1, maxResults, Format$(Now, "yyyy-mm-dd hh:nn:ss"), lawRows)

    Dim rowIndex As Long
    For rowIndex = 1 To rowCount
        Dim uniqueKey As String
        uniqueKey = CStr(lawRows(rowIndex, 3)) & "|" & CStr(lawRows(rowIndex, 4))
        If Len(uniqueKey) > 1 And Not seenKeys.Exists(uniqueKey) Then
            seenKeys.Add uniqueKey, True
            currentHitCount = currentHitCount + 1
            FillSearchResultFromLawRow resultRows, currentHitCount, lawRows, rowIndex, searchKeyword
            If currentHitCount >= maxResults Then
                Exit For
            End If
        End If
    Next rowIndex

    AppendRemoteLawSearch = currentHitCount
End Function

Private Sub AppendRows(ByVal targetSheet As Worksheet, ByRef rowValues As Variant, ByVal rowCount As Long, ByVal columnCount As Long)
    If rowCount <= 0 Then
        Exit Sub
    End If

    Dim writeValues As Variant
    ReDim writeValues(1 To rowCount, 1 To columnCount)

    Dim rowIndex As Long
    Dim columnIndex As Long
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            writeValues(rowIndex, columnIndex) = rowValues(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex

    Dim nextRow As Long
    nextRow = modSheetStore.LastUsedRow(targetSheet) + 1
    If nextRow < 2 Then
        nextRow = 2
    End If

    targetSheet.cells(nextRow, 1).Resize(rowCount, columnCount).Value2 = writeValues
End Sub

Private Function ExtractLawListRows(ByVal responseJson As String, ByVal firstSourceIndex As Long, ByVal maxItems As Long, ByVal cachedAt As String, ByRef pageRows As Variant) As Long
    Dim expectedCount As Long
    expectedCount = JsonRootLongProperty(responseJson, "count", 0)
    If maxItems > 0 And firstSourceIndex + expectedCount - 1 > maxItems Then
        expectedCount = maxItems - firstSourceIndex + 1
    End If
    If expectedCount <= 0 Then
        ReDim pageRows(1 To 1, 1 To LAW_LIST_COLUMN_COUNT)
        ExtractLawListRows = 0
        Exit Function
    End If

    ReDim pageRows(1 To expectedCount, 1 To LAW_LIST_COLUMN_COUNT)

    Dim lawsStart As Long
    lawsStart = FindJsonPropertyValueStart(responseJson, "laws", 1)
    If lawsStart = 0 Then
        Err.Raise vbObjectError + 6201, "modLawSearch.ExtractLawListRows", "API response does not contain laws array."
    End If

    Dim position As Long
    position = NextNonWhitespace(responseJson, lawsStart)
    If Mid$(responseJson, position, 1) <> "[" Then
        Err.Raise vbObjectError + 6201, "modLawSearch.ExtractLawListRows", "API laws property is not an array."
    End If

    position = position + 1

    Dim rowIndex As Long
    Do
        position = NextNonWhitespace(responseJson, position)
        If position = 0 Or position > Len(responseJson) Then
            Exit Do
        End If

        Dim character As String
        character = Mid$(responseJson, position, 1)
        If character = "]" Then
            Exit Do
        End If
        If character <> "{" Then
            Err.Raise vbObjectError + 6203, "modLawSearch.ExtractLawListRows", "Unexpected laws array item."
        End If

        Dim itemEnd As Long
        itemEnd = FindMatchingJsonDelimiter(responseJson, position, "{", "}")
        If itemEnd = 0 Then
            Err.Raise vbObjectError + 6204, "modLawSearch.ExtractLawListRows", "Unclosed laws array item."
        End If

        rowIndex = rowIndex + 1
        If rowIndex > expectedCount Then
            Exit Do
        End If

        FillLawCacheRow pageRows, rowIndex, Mid$(responseJson, position, itemEnd - position + 1), cachedAt, firstSourceIndex + rowIndex - 1

        position = NextNonWhitespace(responseJson, itemEnd + 1)
        If position = 0 Or position > Len(responseJson) Then
            Exit Do
        End If
        character = Mid$(responseJson, position, 1)
        If character = "," Then
            position = position + 1
        ElseIf character = "]" Then
            Exit Do
        End If
    Loop

    ExtractLawListRows = rowIndex
End Function

Private Sub FillLawCacheRow(ByRef pageRows As Variant, ByVal rowIndex As Long, ByVal lawJson As String, ByVal cachedAt As String, ByVal sourceIndex As Long)
    Dim lawInfoJson As String
    lawInfoJson = JsonObjectSlice(lawJson, "law_info")

    Dim revisionInfoJson As String
    revisionInfoJson = JsonObjectSlice(lawJson, "current_revision_info")
    If Len(revisionInfoJson) = 0 Then
        revisionInfoJson = JsonObjectSlice(lawJson, "revision_info")
    End If

    pageRows(rowIndex, 1) = cachedAt
    pageRows(rowIndex, 2) = sourceIndex
    pageRows(rowIndex, 3) = JsonScalarText(lawInfoJson, "law_id")
    pageRows(rowIndex, 4) = JsonScalarText(revisionInfoJson, "law_revision_id")
    pageRows(rowIndex, 5) = JsonScalarText(lawInfoJson, "law_type")
    pageRows(rowIndex, 6) = JsonScalarText(lawInfoJson, "law_num")
    pageRows(rowIndex, 7) = JsonScalarText(revisionInfoJson, "law_title")
    pageRows(rowIndex, 8) = JsonScalarText(revisionInfoJson, "law_title_kana")
    pageRows(rowIndex, 9) = JsonScalarText(revisionInfoJson, "abbrev")
    pageRows(rowIndex, 10) = JsonScalarText(revisionInfoJson, "category")
    pageRows(rowIndex, 11) = JsonScalarText(lawInfoJson, "promulgation_date")
    pageRows(rowIndex, 12) = JsonScalarText(revisionInfoJson, "amendment_enforcement_date")
    pageRows(rowIndex, 13) = JsonScalarText(revisionInfoJson, "repeal_status")
    pageRows(rowIndex, 14) = JsonScalarText(revisionInfoJson, "current_revision_status")
End Sub

Private Function JsonRootLongProperty(ByVal JsonText As String, ByVal propertyName As String, Optional ByVal defaultValue As Long = 0) As Long
    On Error GoTo ErrHandler

    Dim valueText As String
    valueText = JsonScalarText(JsonText, propertyName)
    If Len(valueText) = 0 Then
        JsonRootLongProperty = defaultValue
    Else
        JsonRootLongProperty = CLng(valueText)
    End If
    Exit Function

ErrHandler:
    JsonRootLongProperty = defaultValue
End Function

Private Function JsonObjectSlice(ByVal JsonText As String, ByVal propertyName As String) As String
    Dim valueStart As Long
    valueStart = FindJsonPropertyValueStart(JsonText, propertyName, 1)
    If valueStart = 0 Then
        Exit Function
    End If

    valueStart = NextNonWhitespace(JsonText, valueStart)
    If valueStart > Len(JsonText) Or Mid$(JsonText, valueStart, 1) <> "{" Then
        Exit Function
    End If

    Dim valueEnd As Long
    valueEnd = FindMatchingJsonDelimiter(JsonText, valueStart, "{", "}")
    If valueEnd = 0 Then
        Exit Function
    End If

    JsonObjectSlice = Mid$(JsonText, valueStart, valueEnd - valueStart + 1)
End Function

Private Function JsonScalarText(ByVal JsonText As String, ByVal propertyName As String) As String
    Dim valueStart As Long
    valueStart = FindJsonPropertyValueStart(JsonText, propertyName, 1)
    If valueStart = 0 Then
        Exit Function
    End If

    valueStart = NextNonWhitespace(JsonText, valueStart)
    If valueStart > Len(JsonText) Then
        Exit Function
    End If

    Dim character As String
    character = Mid$(JsonText, valueStart, 1)
    If character = """" Then
        Dim stringEnd As Long
        stringEnd = FindJsonStringEnd(JsonText, valueStart)
        If stringEnd = 0 Then
            Exit Function
        End If
        JsonScalarText = DecodeJsonString(Mid$(JsonText, valueStart + 1, stringEnd - valueStart - 1))
        Exit Function
    End If

    Dim valueEnd As Long
    valueEnd = valueStart
    Do While valueEnd <= Len(JsonText)
        character = Mid$(JsonText, valueEnd, 1)
        If character = "," Or character = "}" Or character = "]" Then
            Exit Do
        End If
        valueEnd = valueEnd + 1
    Loop

    Dim rawValue As String
    rawValue = Trim$(Mid$(JsonText, valueStart, valueEnd - valueStart))
    If StrComp(rawValue, "null", vbTextCompare) = 0 Then
        JsonScalarText = ""
    Else
        JsonScalarText = rawValue
    End If
End Function

Private Function FindJsonPropertyValueStart(ByVal JsonText As String, ByVal propertyName As String, ByVal startAt As Long) As Long
    Dim pattern As String
    pattern = """" & propertyName & """"

    Dim searchAt As Long
    searchAt = startAt
    If searchAt < 1 Then
        searchAt = 1
    End If

    Do
        Dim propertyStart As Long
        propertyStart = InStr(searchAt, JsonText, pattern, vbBinaryCompare)
        If propertyStart = 0 Then
            Exit Function
        End If

        Dim beforeProperty As Long
        beforeProperty = PreviousNonWhitespace(JsonText, propertyStart - 1)

        Dim colonPosition As Long
        colonPosition = NextNonWhitespace(JsonText, propertyStart + Len(pattern))
        If colonPosition <= Len(JsonText) _
            And Mid$(JsonText, colonPosition, 1) = ":" _
            And (beforeProperty = 0 Or Mid$(JsonText, beforeProperty, 1) = "{" Or Mid$(JsonText, beforeProperty, 1) = ",") Then
            FindJsonPropertyValueStart = colonPosition + 1
            Exit Function
        End If

        searchAt = propertyStart + Len(pattern)
    Loop
End Function

Private Function NextNonWhitespace(ByVal text As String, ByVal startAt As Long) As Long
    Dim index As Long
    If startAt < 1 Then
        startAt = 1
    End If

    For index = startAt To Len(text)
        Select Case Mid$(text, index, 1)
            Case " ", vbTab, vbCr, vbLf
            Case Else
                NextNonWhitespace = index
                Exit Function
        End Select
    Next index

    NextNonWhitespace = Len(text) + 1
End Function

Private Function PreviousNonWhitespace(ByVal text As String, ByVal startAt As Long) As Long
    Dim index As Long
    If startAt > Len(text) Then
        startAt = Len(text)
    End If

    For index = startAt To 1 Step -1
        Select Case Mid$(text, index, 1)
            Case " ", vbTab, vbCr, vbLf
            Case Else
                PreviousNonWhitespace = index
                Exit Function
        End Select
    Next index
End Function

Private Function FindMatchingJsonDelimiter(ByVal JsonText As String, ByVal openAt As Long, ByVal openChar As String, ByVal closeChar As String) As Long
    Dim depth As Long
    Dim inString As Boolean
    Dim escaped As Boolean

    Dim index As Long
    For index = openAt To Len(JsonText)
        Dim character As String
        character = Mid$(JsonText, index, 1)

        If inString Then
            If escaped Then
                escaped = False
            ElseIf character = "\" Then
                escaped = True
            ElseIf character = """" Then
                inString = False
            End If
        ElseIf character = """" Then
            inString = True
        ElseIf character = openChar Then
            depth = depth + 1
        ElseIf character = closeChar Then
            depth = depth - 1
            If depth = 0 Then
                FindMatchingJsonDelimiter = index
                Exit Function
            End If
        End If
    Next index
End Function

Private Function FindJsonStringEnd(ByVal JsonText As String, ByVal quoteAt As Long) As Long
    Dim escaped As Boolean
    Dim index As Long
    For index = quoteAt + 1 To Len(JsonText)
        Dim character As String
        character = Mid$(JsonText, index, 1)

        If escaped Then
            escaped = False
        ElseIf character = "\" Then
            escaped = True
        ElseIf character = """" Then
            FindJsonStringEnd = index
            Exit Function
        End If
    Next index
End Function

Private Function DecodeJsonString(ByVal value As String) As String
    If InStr(1, value, "\", vbBinaryCompare) = 0 Then
        DecodeJsonString = value
        Exit Function
    End If

    Dim result As String
    Dim index As Long
    index = 1
    Do While index <= Len(value)
        Dim character As String
        character = Mid$(value, index, 1)

        If character <> "\" Then
            result = result & character
        Else
            index = index + 1
            If index > Len(value) Then
                Exit Do
            End If

            Dim escaped As String
            escaped = Mid$(value, index, 1)
            Select Case escaped
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
                Case "u"
                    If index + 4 <= Len(value) Then
                        result = result & ChrW$(CLng("&H" & Mid$(value, index + 1, 4)))
                        index = index + 4
                    End If
                Case Else
                    result = result & escaped
            End Select
        End If

        index = index + 1
    Loop

    DecodeJsonString = result
End Function

Private Function EnsureTempSheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String, ByVal headers As Variant) As Worksheet
    Dim targetSheet As Worksheet
    If modSheetStore.SheetExists(targetWorkbook, sheetName) Then
        Set targetSheet = targetWorkbook.Worksheets(sheetName)
    Else
        Set targetSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count))
        targetSheet.Name = sheetName
    End If

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

Private Function LawListHeaders() As Variant
    LawListHeaders = Array( _
        "CachedAt", _
        "SourceIndex", _
        "LawId", _
        "LawRevisionId", _
        "LawType", _
        "LawNum", _
        "LawTitle", _
        "LawTitleKana", _
        "Abbrev", _
        "Category", _
        "PromulgationDate", _
        "EnforcementDate", _
        "RepealStatus", _
        "CurrentRevisionStatus")
End Function

Private Function SearchResultHeaders() As Variant
    SearchResultHeaders = Array( _
        "HitIndex", _
        "LawId", _
        "LawRevisionId", _
        "LawType", _
        "LawNum", _
        "LawTitle", _
        "LawTitleKana", _
        "Abbrev", _
        "Category", _
        "PromulgationDate", _
        "EnforcementDate", _
        "RepealStatus", _
        "CurrentRevisionStatus", _
        "MatchedKeyword")
End Function

Private Function JsonObjectProperty(ByVal jsonObject As Object, ByVal propertyName As String) As Object
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

Private Function JsonText(ByVal jsonObject As Object, ByVal propertyName As String) As String
    On Error GoTo ErrHandler

    If jsonObject Is Nothing Then
        Exit Function
    End If

    If Not JsonHasKey(jsonObject, propertyName) Then
        Exit Function
    End If

    Dim value As Variant
    value = jsonObject(propertyName)
    If IsNull(value) Or IsEmpty(value) Then
        JsonText = ""
    Else
        JsonText = CStr(value)
    End If
    Exit Function

ErrHandler:
    JsonText = ""
End Function

Private Function JsonLongProperty(ByVal jsonObject As Object, ByVal propertyName As String, Optional ByVal defaultValue As Long = 0) As Long
    On Error GoTo ErrHandler

    If jsonObject Is Nothing Then
        JsonLongProperty = defaultValue
        Exit Function
    End If

    If Not modJsonUtil.JsonHasKey(jsonObject, propertyName) Then
        JsonLongProperty = defaultValue
        Exit Function
    End If

    Dim value As Variant
    value = jsonObject(propertyName)
    If IsNull(value) Or IsEmpty(value) Then
        JsonLongProperty = defaultValue
    Else
        JsonLongProperty = CLng(value)
    End If
    Exit Function

ErrHandler:
    JsonLongProperty = defaultValue
End Function

Private Function JsonHasKey(ByVal jsonObject As Object, ByVal propertyName As String) As Boolean
    On Error GoTo ErrHandler

    JsonHasKey = CBool(jsonObject.Exists(propertyName))
    Exit Function

ErrHandler:
    JsonHasKey = False
End Function
