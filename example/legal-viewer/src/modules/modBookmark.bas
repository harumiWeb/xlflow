Attribute VB_Name = "modBookmark"
Option Explicit

Private Const BOOKMARK_COL_BOOKMARK_ID As Long = 1
Private Const BOOKMARK_COL_LAW_ID As Long = 2
Private Const BOOKMARK_COL_LAW_TITLE As Long = 3
Private Const BOOKMARK_COL_LAW_NUM As Long = 4
Private Const BOOKMARK_COL_ENFORCEMENT_DATE As Long = 5
Private Const BOOKMARK_COL_ARTICLE_NO As Long = 6
Private Const BOOKMARK_COL_PARAGRAPH_NO As Long = 7
Private Const BOOKMARK_COL_ITEM_NO As Long = 8
Private Const BOOKMARK_COL_CAPTION As Long = 9
Private Const BOOKMARK_COL_BOOKMARK_KIND As Long = 10
Private Const BOOKMARK_COL_TAGS As Long = 11
Private Const BOOKMARK_COL_MEMO As Long = 12
Private Const BOOKMARK_COL_CREATED_AT As Long = 13

Private Const BODY_UNIT_COL_LAW_ID As Long = 2
Private Const BODY_UNIT_COL_ENFORCEMENT_DATE As Long = 4
Private Const BODY_UNIT_COL_LAW_TITLE As Long = 5
Private Const BODY_UNIT_COL_LAW_NUM As Long = 6
Private Const BODY_UNIT_COL_UNIT_KIND As Long = 7
Private Const BODY_UNIT_COL_ARTICLE_TITLE As Long = 8
Private Const BODY_UNIT_COL_ARTICLE_CAPTION As Long = 9
Private Const BODY_UNIT_COL_PARAGRAPH_NUM As Long = 10
Private Const BODY_UNIT_COL_ITEM_TITLE As Long = 11
Private Const BODY_UNIT_COL_TEXT As Long = 13

Public Function BookmarkStoreSheetName() As String
    BookmarkStoreSheetName = modSheetStore.BookmarksSheetName()
End Function

Public Function BookmarkCount() As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, BookmarkStoreSheetName()) Then
        BookmarkCount = 0
        Exit Function
    End If

    BookmarkCount = Application.Max(0, modSheetStore.LastUsedRow(ThisWorkbook.Worksheets(BookmarkStoreSheetName())) - 1)
End Function

Public Sub ClearBookmarkStore()
    On Error GoTo ErrHandler

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, BookmarkStoreSheetName())
    modSheetStore.ClearDataRows bookmarkSheet
    modAppStartup.SaveWorkbookState ThisWorkbook
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.ClearBookmarkStore", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Sub ConfigureBookmarkListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 8
    targetListBox.columnWidths = "240 pt;90 pt;120 pt;88 pt;72 pt;96 pt;60 pt;0 pt"
    targetListBox.BoundColumn = 8
End Sub

Public Function PopulateBookmarkListBox(ByVal targetListBox As Object, Optional ByVal filterText As String = "") As Long
    On Error GoTo ErrHandler

    ConfigureBookmarkListBox targetListBox
    targetListBox.Clear

    If Not modSheetStore.SheetExists(ThisWorkbook, BookmarkStoreSheetName()) Then
        PopulateBookmarkListBox = 0
        Exit Function
    End If

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = ThisWorkbook.Worksheets(BookmarkStoreSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(bookmarkSheet)
    If lastRow < 2 Then
        PopulateBookmarkListBox = 0
        Exit Function
    End If

    Dim normalizedFilter As String
    normalizedFilter = NormalizeBookmarkFilterText(filterText)

    Dim rowIndex As Long
    Dim listIndex As Long
    Dim loadedCount As Long
    For rowIndex = lastRow To 2 Step -1
        If Len(normalizedFilter) = 0 Or BookmarkRowContainsText(bookmarkSheet, rowIndex, normalizedFilter) Then
            targetListBox.AddItem BookmarkDisplayText(rowIndex)
            listIndex = targetListBox.ListCount - 1
            targetListBox.List(listIndex, 1) = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_TAGS).Value2)
            targetListBox.List(listIndex, 2) = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_MEMO).Value2)
            targetListBox.List(listIndex, 3) = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_LAW_NUM).Value2)
            targetListBox.List(listIndex, 4) = NormalizeBookmarkDateText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ENFORCEMENT_DATE).Value2)
            targetListBox.List(listIndex, 5) = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_CREATED_AT).Value2)
            targetListBox.List(listIndex, 6) = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_BOOKMARK_KIND).Value2)
            targetListBox.List(listIndex, 7) = CStr(rowIndex)
            loadedCount = loadedCount + 1
        End If
    Next rowIndex

    PopulateBookmarkListBox = loadedCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.PopulateBookmarkListBox", Err.description, "", "", filterText
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function BookmarkRowDisplayText(ByVal rowIndex As Long) As String
    BookmarkRowDisplayText = BookmarkDisplayText(rowIndex)
End Function

Public Function BookmarkRowTags(ByVal rowIndex As Long) As String
    BookmarkRowTags = BookmarkRowText(rowIndex, BOOKMARK_COL_TAGS)
End Function

Public Function BookmarkRowMemo(ByVal rowIndex As Long) As String
    BookmarkRowMemo = BookmarkRowText(rowIndex, BOOKMARK_COL_MEMO)
End Function

Public Function BookmarkRowLawTitle(ByVal rowIndex As Long) As String
    BookmarkRowLawTitle = BookmarkRowText(rowIndex, BOOKMARK_COL_LAW_TITLE)
End Function

Public Function BookmarkRowLawNum(ByVal rowIndex As Long) As String
    BookmarkRowLawNum = BookmarkRowText(rowIndex, BOOKMARK_COL_LAW_NUM)
End Function

Public Function BookmarkRowLawId(ByVal rowIndex As Long) As String
    BookmarkRowLawId = BookmarkRowText(rowIndex, BOOKMARK_COL_LAW_ID)
End Function

Public Function BookmarkRowEnforcementDate(ByVal rowIndex As Long) As String
    BookmarkRowEnforcementDate = BookmarkRowText(rowIndex, BOOKMARK_COL_ENFORCEMENT_DATE)
End Function

Public Function BookmarkRowKind(ByVal rowIndex As Long) As String
    BookmarkRowKind = BookmarkRowText(rowIndex, BOOKMARK_COL_BOOKMARK_KIND)
End Function

Public Function BookmarkRowCaption(ByVal rowIndex As Long) As String
    BookmarkRowCaption = BookmarkRowText(rowIndex, BOOKMARK_COL_CAPTION)
End Function

Public Function BookmarkRowCreatedAt(ByVal rowIndex As Long) As String
    BookmarkRowCreatedAt = BookmarkRowText(rowIndex, BOOKMARK_COL_CREATED_AT)
End Function

Public Function BookmarkRowArticleNo(ByVal rowIndex As Long) As String
    BookmarkRowArticleNo = BookmarkRowText(rowIndex, BOOKMARK_COL_ARTICLE_NO)
End Function

Public Function BookmarkRowParagraphNo(ByVal rowIndex As Long) As String
    BookmarkRowParagraphNo = BookmarkRowText(rowIndex, BOOKMARK_COL_PARAGRAPH_NO)
End Function

Public Function BookmarkRowItemNo(ByVal rowIndex As Long) As String
    BookmarkRowItemNo = BookmarkRowText(rowIndex, BOOKMARK_COL_ITEM_NO)
End Function

Public Function UpdateBookmarkTagsMemo(ByVal rowIndex As Long, ByVal tags As String, ByVal memo As String) As Boolean
    On Error GoTo ErrHandler

    If Not IsValidBookmarkRow(rowIndex) Then
        Exit Function
    End If

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = ThisWorkbook.Worksheets(BookmarkStoreSheetName())
    bookmarkSheet.cells(rowIndex, BOOKMARK_COL_TAGS).Value2 = NormalizeBookmarkText(tags)
    bookmarkSheet.cells(rowIndex, BOOKMARK_COL_MEMO).Value2 = NormalizeBookmarkText(memo)

    modAppStartup.SaveWorkbookState ThisWorkbook
    UpdateBookmarkTagsMemo = True
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.UpdateBookmarkTagsMemo", Err.description, "", "", CStr(rowIndex)
    UpdateBookmarkTagsMemo = False
End Function

Public Function DeleteBookmarkEntry(ByVal rowIndex As Long) As Boolean
    On Error GoTo ErrHandler

    DeleteBookmarkRow rowIndex
    modAppStartup.SaveWorkbookState ThisWorkbook
    DeleteBookmarkEntry = True
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.DeleteBookmarkEntry", Err.description, "", "", CStr(rowIndex)
    DeleteBookmarkEntry = False
End Function

Public Function DeleteBookmarksByFilter(ByVal filterText As String) As Long
    On Error GoTo ErrHandler

    filterText = NormalizeBookmarkFilterText(filterText)
    If Len(filterText) = 0 Then
        DeleteBookmarksByFilter = 0
        Exit Function
    End If

    If Not modSheetStore.SheetExists(ThisWorkbook, BookmarkStoreSheetName()) Then
        DeleteBookmarksByFilter = 0
        Exit Function
    End If

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = ThisWorkbook.Worksheets(BookmarkStoreSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(bookmarkSheet)

    Dim deletedCount As Long
    Dim rowIndex As Long
    For rowIndex = lastRow To 2 Step -1
        If BookmarkRowContainsText(bookmarkSheet, rowIndex, filterText) Then
            bookmarkSheet.rows(rowIndex).Delete
            deletedCount = deletedCount + 1
        End If
    Next rowIndex

    If deletedCount > 0 Then
        modAppStartup.SaveWorkbookState ThisWorkbook
    End If
    DeleteBookmarksByFilter = deletedCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.DeleteBookmarksByFilter", Err.description, "", "", filterText
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function ClearBookmarkEntries() As Long
    On Error GoTo ErrHandler

    ClearBookmarkEntries = BookmarkCount()
    ClearBookmarkStore
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.ClearBookmarkEntries", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function FindBookmarkBodyUnitRow(ByVal rowIndex As Long) As Long
    On Error GoTo ErrHandler

    If Not IsValidBookmarkRow(rowIndex) Then
        Exit Function
    End If

    FindBookmarkBodyUnitRow = FindBookmarkBodyUnitRowByKey( _
        BookmarkRowLawId(rowIndex), _
        BookmarkRowEnforcementDate(rowIndex), _
        BookmarkRowArticleNo(rowIndex), _
        BookmarkRowParagraphNo(rowIndex), _
        BookmarkRowItemNo(rowIndex), _
        BookmarkRowKind(rowIndex))
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.FindBookmarkBodyUnitRow", Err.description, "", "", CStr(rowIndex)
End Function

Public Function FindBookmarkBodyUnitRowByKey(ByVal lawId As String, ByVal enforcementDate As String, ByVal articleNo As String, ByVal paragraphNo As String, ByVal itemNo As String, ByVal bookmarkKind As String) As Long
    On Error GoTo ErrHandler

    lawId = Trim$(lawId)
    enforcementDate = NormalizeBookmarkDateText(enforcementDate)
    articleNo = Trim$(articleNo)
    paragraphNo = Trim$(paragraphNo)
    itemNo = Trim$(itemNo)
    bookmarkKind = Trim$(bookmarkKind)

    If Len(lawId) = 0 Or Len(enforcementDate) = 0 Then
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        Exit Function
    End If

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(bodySheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If StrComp(CellText(bodySheet.cells(rowIndex, 2).Value2), lawId, vbBinaryCompare) = 0 _
            And StrComp(NormalizeBookmarkDateText(bodySheet.cells(rowIndex, 4).Value2), enforcementDate, vbBinaryCompare) = 0 _
            And StrComp(CellText(bodySheet.cells(rowIndex, 8).Value2), articleNo, vbBinaryCompare) = 0 _
            And StrComp(CellText(bodySheet.cells(rowIndex, 10).Value2), paragraphNo, vbBinaryCompare) = 0 _
            And StrComp(CellText(bodySheet.cells(rowIndex, 11).Value2), itemNo, vbBinaryCompare) = 0 _
            And StrComp(BookmarkKindForBodyRowFromSheet(bodySheet, rowIndex), bookmarkKind, vbBinaryCompare) = 0 Then
            FindBookmarkBodyUnitRowByKey = rowIndex
            Exit Function
        End If
    Next rowIndex
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.FindBookmarkBodyUnitRowByKey", Err.description, lawId, enforcementDate, bookmarkKind
End Function

Public Function RegisterBookmarkFromBodyUnitRow(ByVal bodyUnitRow As Long, Optional ByVal tags As String = "", Optional ByVal memo As String = "") As Long
    On Error GoTo ErrHandler

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, BookmarkStoreSheetName())
    bookmarkSheet.cells.NumberFormat = "@"

    Dim bodyValues As Variant
    bodyValues = BodyUnitValues(bodyUnitRow)
    If Not IsArray(bodyValues) Then
        Err.Raise vbObjectError + 7401, "modBookmark.RegisterBookmarkFromBodyUnitRow", "Body unit row is invalid."
    End If

    Dim lawId As String
    Dim lawTitle As String
    Dim lawNum As String
    Dim enforcementDate As String
    Dim articleNo As String
    Dim paragraphNo As String
    Dim itemNo As String
    Dim caption As String
    Dim bookmarkKind As String

    lawId = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_LAW_ID)
    lawTitle = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_LAW_TITLE)
    lawNum = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_LAW_NUM)
    enforcementDate = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_ENFORCEMENT_DATE)
    articleNo = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_ARTICLE_TITLE)
    paragraphNo = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_PARAGRAPH_NUM)
    itemNo = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_ITEM_TITLE)
    caption = BookmarkCaptionFromBodyRow(bodyValues, bodyUnitRow)
    bookmarkKind = BookmarkKindForBodyRow(bodyValues, bodyUnitRow)

    If Len(lawId) = 0 Or Len(lawTitle) = 0 Or Len(enforcementDate) = 0 Then
        Err.Raise vbObjectError + 7402, "modBookmark.RegisterBookmarkFromBodyUnitRow", "Bookmark source fields are missing."
    End If

    Dim existingRow As Long
    existingRow = FindBookmarkRow(lawId, enforcementDate, articleNo, paragraphNo, itemNo, bookmarkKind)
    If existingRow > 0 Then
        RegisterBookmarkFromBodyUnitRow = existingRow
        Exit Function
    End If

    Dim bookmarkId As String
    bookmarkId = BuildBookmarkId()

    RegisterBookmarkFromBodyUnitRow = modSheetStore.AppendRow(bookmarkSheet, Array( _
        bookmarkId, _
        lawId, _
        lawTitle, _
        lawNum, _
        enforcementDate, _
        articleNo, _
        paragraphNo, _
        itemNo, _
        caption, _
        bookmarkKind, _
        NormalizeBookmarkText(tags), _
        NormalizeBookmarkText(memo), _
        Format$(Now, "yyyy-mm-dd hh:nn:ss")))
    modAppStartup.SaveWorkbookState ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.RegisterBookmarkFromBodyUnitRow", Err.description, lawId, enforcementDate, bookmarkKind
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function FindBookmarkRow(ByVal lawId As String, ByVal enforcementDate As String, ByVal articleNo As String, ByVal paragraphNo As String, ByVal itemNo As String, ByVal bookmarkKind As String) As Long
    On Error GoTo ErrHandler

    lawId = Trim$(lawId)
    enforcementDate = Trim$(enforcementDate)
    articleNo = Trim$(articleNo)
    paragraphNo = Trim$(paragraphNo)
    itemNo = Trim$(itemNo)
    bookmarkKind = Trim$(bookmarkKind)

    If Len(lawId) = 0 Or Len(enforcementDate) = 0 Then
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, BookmarkStoreSheetName()) Then
        Exit Function
    End If

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = ThisWorkbook.Worksheets(BookmarkStoreSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(bookmarkSheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If StrComp(CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_LAW_ID).Value2), lawId, vbBinaryCompare) = 0 _
            And StrComp(NormalizeBookmarkDateText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ENFORCEMENT_DATE).Value2), enforcementDate, vbBinaryCompare) = 0 _
            And StrComp(CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ARTICLE_NO).Value2), articleNo, vbBinaryCompare) = 0 _
            And StrComp(CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_PARAGRAPH_NO).Value2), paragraphNo, vbBinaryCompare) = 0 _
            And StrComp(CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ITEM_NO).Value2), itemNo, vbBinaryCompare) = 0 _
            And StrComp(CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_BOOKMARK_KIND).Value2), bookmarkKind, vbBinaryCompare) = 0 Then
            FindBookmarkRow = rowIndex
            Exit Function
        End If
    Next rowIndex
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.FindBookmarkRow", Err.description, lawId, enforcementDate, bookmarkKind
    FindBookmarkRow = 0
End Function

Public Function BookmarkDisplayText(ByVal rowIndex As Long) As String
    If Not modSheetStore.SheetExists(ThisWorkbook, BookmarkStoreSheetName()) Then
        Exit Function
    End If

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = ThisWorkbook.Worksheets(BookmarkStoreSheetName())
    If rowIndex < 2 Or rowIndex > modSheetStore.LastUsedRow(bookmarkSheet) Then
        Exit Function
    End If

    BookmarkDisplayText = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_LAW_TITLE).Value2) _
        & "｜" & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_BOOKMARK_KIND).Value2) _
        & "｜" & BookmarkReferenceText(bookmarkSheet, rowIndex)
End Function

Public Function BookmarkSmoke() As String
    On Error GoTo ErrHandler

    Dim originalCount As Long
    originalCount = BookmarkCount()

    modLawNavigator.PrepareBodyNavigatorSmokeData

    Dim bookmarkRow As Long
    bookmarkRow = RegisterBookmarkFromBodyUnitRow(3, "smoke", "bookmark smoke")
    If BookmarkCount() <> originalCount + 1 Then
        Err.Raise vbObjectError + 7411, "modBookmark.BookmarkSmoke", "Bookmark count mismatch."
    End If
    If Len(BookmarkDisplayText(bookmarkRow)) = 0 Then
        Err.Raise vbObjectError + 7412, "modBookmark.BookmarkSmoke", "Bookmark display text is missing."
    End If

    Dim duplicateRow As Long
    duplicateRow = RegisterBookmarkFromBodyUnitRow(3, "smoke", "bookmark smoke")
    If duplicateRow <> bookmarkRow Then
        Err.Raise vbObjectError + 7413, "modBookmark.BookmarkSmoke", "Bookmark duplicate detection failed."
    End If

    DeleteBookmarkRow bookmarkRow
    If BookmarkCount() <> originalCount Then
        Err.Raise vbObjectError + 7414, "modBookmark.BookmarkSmoke", "Bookmark cleanup failed."
    End If

    BookmarkSmoke = "ok:" & CStr(originalCount)
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.BookmarkSmoke", Err.description, "", "", ""
    BookmarkSmoke = "error:" & Err.description
End Function

Private Function BookmarkRowText(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If Not IsValidBookmarkRow(rowIndex) Then
        Exit Function
    End If

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = ThisWorkbook.Worksheets(BookmarkStoreSheetName())
    BookmarkRowText = CellText(bookmarkSheet.cells(rowIndex, columnIndex).Value2)
End Function

Private Function IsValidBookmarkRow(ByVal rowIndex As Long) As Boolean
    If rowIndex < 2 Then
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, BookmarkStoreSheetName()) Then
        Exit Function
    End If

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = ThisWorkbook.Worksheets(BookmarkStoreSheetName())
    IsValidBookmarkRow = (rowIndex <= modSheetStore.LastUsedRow(bookmarkSheet))
End Function

Private Function NormalizeBookmarkFilterText(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")
    NormalizeBookmarkFilterText = Trim$(value)
End Function

Private Function BookmarkRowContainsText(ByVal bookmarkSheet As Worksheet, ByVal rowIndex As Long, ByVal filterText As String) As Boolean
    Dim haystack As String
    haystack = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_BOOKMARK_ID).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_LAW_ID).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_LAW_TITLE).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_LAW_NUM).Value2) _
        & " " & NormalizeBookmarkDateText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ENFORCEMENT_DATE).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ARTICLE_NO).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_PARAGRAPH_NO).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ITEM_NO).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_CAPTION).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_BOOKMARK_KIND).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_TAGS).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_MEMO).Value2) _
        & " " & CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_CREATED_AT).Value2)

    BookmarkRowContainsText = (InStr(1, haystack, filterText, vbTextCompare) > 0)
End Function

Private Function BookmarkKindForBodyRowFromSheet(ByVal bodySheet As Worksheet, ByVal rowIndex As Long) As String
    Select Case CellText(bodySheet.cells(rowIndex, 7).Value2)
        Case "Article"
            BookmarkKindForBodyRowFromSheet = "条"
        Case "Paragraph"
            BookmarkKindForBodyRowFromSheet = "項"
        Case "Item"
            BookmarkKindForBodyRowFromSheet = "号"
        Case "AppdxTable"
            BookmarkKindForBodyRowFromSheet = "別表"
        Case "Appdx"
            BookmarkKindForBodyRowFromSheet = "別表等"
        Case "SupplementaryProvision"
            BookmarkKindForBodyRowFromSheet = "附則"
        Case Else
            BookmarkKindForBodyRowFromSheet = CellText(bodySheet.cells(rowIndex, 7).Value2)
    End Select
End Function

Private Function BodyUnitValues(ByVal bodyUnitRow As Long) As Variant
    If Not modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        Exit Function
    End If

    Dim bodySheet As Worksheet
    Set bodySheet = ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName())
    If bodyUnitRow < 2 Or bodyUnitRow > modSheetStore.LastUsedRow(bodySheet) Then
        Exit Function
    End If

    BodyUnitValues = bodySheet.Range(bodySheet.cells(1, 1), bodySheet.cells(modSheetStore.LastUsedRow(bodySheet), 13)).Value2
End Function

Private Function BookmarkReferenceText(ByVal bookmarkSheet As Worksheet, ByVal rowIndex As Long) As String
    Dim referenceText As String
    referenceText = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ARTICLE_NO).Value2)

    Dim paragraphNo As String
    paragraphNo = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_PARAGRAPH_NO).Value2)
    If Len(paragraphNo) > 0 Then
        referenceText = AppendPart(referenceText, "第" & paragraphNo & "項")
    End If

    Dim itemNo As String
    itemNo = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_ITEM_NO).Value2)
    If Len(itemNo) > 0 Then
        referenceText = AppendPart(referenceText, itemNo)
    End If

    Dim caption As String
    caption = CellText(bookmarkSheet.cells(rowIndex, BOOKMARK_COL_CAPTION).Value2)
    If Len(caption) > 0 Then
        referenceText = AppendPart(referenceText, caption)
    End If

    BookmarkReferenceText = referenceText
End Function

Private Function BookmarkCaptionFromBodyRow(ByRef bodyValues As Variant, ByVal bodyUnitRow As Long) As String
    Dim itemTitle As String
    itemTitle = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_ITEM_TITLE)
    If Len(itemTitle) > 0 Then
        BookmarkCaptionFromBodyRow = itemTitle
        Exit Function
    End If

    Dim articleCaption As String
    articleCaption = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_ARTICLE_CAPTION)
    If Len(articleCaption) > 0 Then
        BookmarkCaptionFromBodyRow = articleCaption
        Exit Function
    End If

    Dim bodyText As String
    bodyText = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_TEXT)
    bodyText = CompactText(bodyText)
    If Len(bodyText) > 80 Then
        bodyText = Left$(bodyText, 80) & "..."
    End If
    BookmarkCaptionFromBodyRow = bodyText
End Function

Private Function BookmarkKindForBodyRow(ByRef bodyValues As Variant, ByVal bodyUnitRow As Long) As String
    Select Case BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_UNIT_KIND)
        Case "Article"
            BookmarkKindForBodyRow = "条"
        Case "Paragraph"
            BookmarkKindForBodyRow = "項"
        Case "Item"
            BookmarkKindForBodyRow = "号"
        Case "AppdxTable"
            BookmarkKindForBodyRow = "別表"
        Case "Appdx"
            BookmarkKindForBodyRow = "別表等"
        Case "SupplementaryProvision"
            BookmarkKindForBodyRow = "附則"
        Case Else
            BookmarkKindForBodyRow = BodyValueText(bodyValues, bodyUnitRow, BODY_UNIT_COL_UNIT_KIND)
    End Select
End Function

Private Function BuildBookmarkId() As String
    BuildBookmarkId = "BK-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(BookmarkCount() + 1, "0000")
End Function

Private Sub DeleteBookmarkRow(ByVal rowIndex As Long)
    On Error GoTo ErrHandler

    If rowIndex < 2 Then
        Exit Sub
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, BookmarkStoreSheetName()) Then
        Exit Sub
    End If

    Dim bookmarkSheet As Worksheet
    Set bookmarkSheet = ThisWorkbook.Worksheets(BookmarkStoreSheetName())
    If rowIndex > modSheetStore.LastUsedRow(bookmarkSheet) Then
        Exit Sub
    End If

    bookmarkSheet.rows(rowIndex).Delete
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "Bookmark", "modBookmark.DeleteBookmarkRow", Err.description, "", "", CStr(rowIndex)
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Private Function NormalizeBookmarkText(ByVal value As String) As String
    NormalizeBookmarkText = Trim$(Replace(Replace(Replace(value, vbCr, " "), vbLf, " "), vbTab, " "))
End Function

Private Function NormalizeBookmarkDateText(ByVal value As Variant) As String
    If IsDate(value) Then
        NormalizeBookmarkDateText = Format$(CDate(value), "yyyy-mm-dd")
    ElseIf IsNumeric(value) Then
        NormalizeBookmarkDateText = Format$(CDate(CDbl(value)), "yyyy-mm-dd")
    Else
        NormalizeBookmarkDateText = Trim$(CStr(value))
    End If
End Function

Private Function AppendPart(ByVal baseText As String, ByVal partText As String) As String
    If Len(baseText) = 0 Then
        AppendPart = partText
    ElseIf Len(partText) = 0 Then
        AppendPart = baseText
    Else
        AppendPart = baseText & " / " & partText
    End If
End Function

Private Function BodyValueText(ByRef bodyValues As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If Not IsArray(bodyValues) Then
        Exit Function
    End If
    If rowIndex < LBound(bodyValues, 1) Or rowIndex > UBound(bodyValues, 1) _
        Or columnIndex < LBound(bodyValues, 2) Or columnIndex > UBound(bodyValues, 2) Then
        Exit Function
    End If

    If IsError(bodyValues(rowIndex, columnIndex)) Or IsNull(bodyValues(rowIndex, columnIndex)) Or IsEmpty(bodyValues(rowIndex, columnIndex)) Then
        BodyValueText = ""
    Else
        BodyValueText = CStr(bodyValues(rowIndex, columnIndex))
    End If
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
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
