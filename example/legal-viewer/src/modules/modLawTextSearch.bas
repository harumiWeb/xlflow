Attribute VB_Name = "modLawTextSearch"
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

Public Sub ConfigureSearchResultsListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 6
    targetListBox.columnWidths = "48 pt;150 pt;310 pt;0 pt;0 pt;0 pt"
    targetListBox.BoundColumn = 4
End Sub

Public Function PopulateBodySearchResultsListBox(ByVal targetListBox As Object, ByVal addedLawStoreRow As Long, ByVal searchText As String, Optional ByVal searchMode As String = "AND") As Long
    On Error GoTo ErrHandler

    ConfigureSearchResultsListBox targetListBox
    targetListBox.Clear

    Dim terms As Collection
    Set terms = ParseSearchTerms(searchText)
    If terms.Count = 0 Then
        PopulateBodySearchResultsListBox = 0
        Exit Function
    End If

    Dim lawRevisionId As String
    lawRevisionId = AddedLawRevisionId(addedLawStoreRow)
    If Len(lawRevisionId) = 0 Then
        PopulateBodySearchResultsListBox = 0
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        PopulateBodySearchResultsListBox = 0
        Exit Function
    End If

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(bodySheet)
    If lastRow < 2 Then
        PopulateBodySearchResultsListBox = 0
        Exit Function
    End If

    Dim bodyValues As Variant
    bodyValues = bodySheet.Range(bodySheet.cells(1, 1), bodySheet.cells(lastRow, BODY_UNIT_COL_COUNT)).Value2

    Dim rowIndex As Long
    Dim candidateText As String
    Dim hitCount As Long
    For rowIndex = 2 To UBound(bodyValues, 1)
        If StrComp(BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_LAW_REVISION_ID), lawRevisionId, vbTextCompare) = 0 _
            And ShouldSearchUnitKind(BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_UNIT_KIND)) Then

            candidateText = SearchableBodyText(bodyValues, rowIndex)
            If MatchesSearchTerms(candidateText, terms, searchMode) Then
                AddSearchResult targetListBox, bodyValues, rowIndex, terms
                hitCount = hitCount + 1
            End If
        End If
    Next rowIndex

    PopulateBodySearchResultsListBox = hitCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawTextSearch", "modLawTextSearch.PopulateBodySearchResultsListBox", Err.description, "", "", searchText
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function BodySearchSmoke() As String
    On Error GoTo ErrHandler

    Dim addedRow As Long
    addedRow = modLawNavigator.PrepareBodyNavigatorSmokeData()

    Load frmMain

    Dim hitCount As Long
    hitCount = PopulateBodySearchResultsListBox(frmMain.lstResults, addedRow, "受給権 差し押える", "AND")
    If hitCount <> 1 Then
        Err.Raise vbObjectError + 7001, "modLawTextSearch.BodySearchSmoke", "AND search hit count mismatch."
    End If

    hitCount = PopulateBodySearchResultsListBox(frmMain.lstResults, addedRow, "受給権 附則側", "OR")
    If hitCount <> 3 Then
        Err.Raise vbObjectError + 7002, "modLawTextSearch.BodySearchSmoke", "OR search hit count mismatch."
    End If

    hitCount = PopulateBodySearchResultsListBox(frmMain.lstResults, addedRow, Chr$(34) & "附則側の第十一条" & Chr$(34), "AND")
    If hitCount <> 1 Then
        Err.Raise vbObjectError + 7003, "modLawTextSearch.BodySearchSmoke", "Quoted phrase search hit count mismatch."
    End If

    hitCount = PopulateBodySearchResultsListBox(frmMain.lstResults, addedRow, "別表検索語", "AND")
    If hitCount <> 1 Then
        Err.Raise vbObjectError + 7004, "modLawTextSearch.BodySearchSmoke", "Appended table search hit count mismatch."
    End If

    BodySearchSmoke = "ok:" & CStr(frmMain.lstResults.ListCount)
    Unload frmMain
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawTextSearch", "modLawTextSearch.BodySearchSmoke", Err.description, "", "", ""
    BodySearchSmoke = "error:" & Err.description
End Function

Private Sub AddSearchResult(ByVal targetListBox As Object, ByRef bodyValues As Variant, ByVal rowIndex As Long, ByVal terms As Collection)
    Dim listIndex As Long
    targetListBox.AddItem UnitKindDisplay(BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_UNIT_KIND))
    listIndex = targetListBox.ListCount - 1
    targetListBox.List(listIndex, 1) = BodyUnitHeading(bodyValues, rowIndex)
    targetListBox.List(listIndex, 2) = MatchSnippet(SearchableBodyText(bodyValues, rowIndex), terms)
    targetListBox.List(listIndex, 3) = CStr(rowIndex)
    targetListBox.List(listIndex, 4) = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_UNIT_KIND)
    targetListBox.List(listIndex, 5) = BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_SORT_KEY)
End Sub

Private Function ParseSearchTerms(ByVal searchText As String) As Collection
    Dim terms As Collection
    Set terms = New Collection

    searchText = Replace(searchText, ChrW$(&H3000), " ")
    searchText = Trim$(searchText)

    Dim buffer As String
    Dim index As Long
    Dim currentChar As String
    Dim inQuote As Boolean

    For index = 1 To Len(searchText)
        currentChar = Mid$(searchText, index, 1)
        If currentChar = Chr$(34) Then
            If inQuote Then
                AddSearchTerm terms, buffer
                buffer = ""
                inQuote = False
            Else
                AddSearchTerm terms, buffer
                buffer = ""
                inQuote = True
            End If
        ElseIf Not inQuote And currentChar = " " Then
            AddSearchTerm terms, buffer
            buffer = ""
        Else
            buffer = buffer & currentChar
        End If
    Next index

    AddSearchTerm terms, buffer
    Set ParseSearchTerms = terms
End Function

Private Sub AddSearchTerm(ByVal terms As Collection, ByVal termText As String)
    termText = Trim$(termText)
    If Len(termText) > 0 Then
        terms.Add termText
    End If
End Sub

Private Function MatchesSearchTerms(ByVal targetText As String, ByVal terms As Collection, ByVal searchMode As String) As Boolean
    Dim term As Variant
    If StrComp(searchMode, "OR", vbTextCompare) = 0 Then
        For Each term In terms
            If InStr(1, targetText, CStr(term), vbTextCompare) > 0 Then
                MatchesSearchTerms = True
                Exit Function
            End If
        Next term
    Else
        For Each term In terms
            If InStr(1, targetText, CStr(term), vbTextCompare) = 0 Then
                MatchesSearchTerms = False
                Exit Function
            End If
        Next term
        MatchesSearchTerms = True
    End If
End Function

Private Function MatchSnippet(ByVal targetText As String, ByVal terms As Collection) As String
    targetText = CompactText(targetText)
    If Len(targetText) = 0 Then
        MatchSnippet = ""
        Exit Function
    End If

    Dim matchPosition As Long
    Dim term As Variant
    For Each term In terms
        matchPosition = InStr(1, targetText, CStr(term), vbTextCompare)
        If matchPosition > 0 Then
            Exit For
        End If
    Next term
    If matchPosition = 0 Then
        matchPosition = 1
    End If

    Dim startPosition As Long
    startPosition = Application.Max(1, matchPosition - 24)

    Dim snippet As String
    snippet = Mid$(targetText, startPosition, 88)
    If startPosition > 1 Then
        snippet = "..." & snippet
    End If
    If startPosition + 88 <= Len(targetText) Then
        snippet = snippet & "..."
    End If

    MatchSnippet = snippet
End Function

Private Function SearchableBodyText(ByRef bodyValues As Variant, ByVal rowIndex As Long) As String
    SearchableBodyText = BodyUnitHeading(bodyValues, rowIndex) & " " & BodyValueText(bodyValues, rowIndex, BODY_UNIT_COL_TEXT)
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

Private Function AddedLawRevisionId(ByVal addedLawStoreRow As Long) As String
    If Not modSheetStore.SheetExists(ThisWorkbook, modAddedLawStore.AddedLawStoreSheetName()) Then
        Exit Function
    End If

    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Worksheets(modAddedLawStore.AddedLawStoreSheetName())
    If addedLawStoreRow < 2 Or addedLawStoreRow > modSheetStore.LastUsedRow(storeSheet) Then
        Exit Function
    End If

    AddedLawRevisionId = CellText(storeSheet.cells(addedLawStoreRow, ADDED_LAW_COL_LAW_REVISION_ID).Value2)
End Function

Private Function ShouldSearchUnitKind(ByVal unitKind As String) As Boolean
    ShouldSearchUnitKind = (StrComp(unitKind, "Paragraph", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "Item", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "LawText", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "AppdxTable", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "Appdx", vbBinaryCompare) = 0 _
        Or StrComp(unitKind, "SupplementaryProvision", vbBinaryCompare) = 0)
End Function

Private Function UnitKindDisplay(ByVal unitKind As String) As String
    Select Case unitKind
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

Private Function AppendHeadingPart(ByVal baseText As String, ByVal partText As String) As String
    If Len(baseText) = 0 Then
        AppendHeadingPart = partText
    ElseIf Len(partText) = 0 Then
        AppendHeadingPart = baseText
    Else
        AppendHeadingPart = baseText & " " & partText
    End If
End Function

Private Function CompactText(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")
    value = Trim$(value)

    Do While InStr(1, value, "  ", vbBinaryCompare) > 0
        value = Replace(value, "  ", " ")
    Loop

    CompactText = value
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
