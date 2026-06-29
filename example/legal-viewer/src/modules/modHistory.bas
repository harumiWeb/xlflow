Attribute VB_Name = "modHistory"
Option Explicit

Private Const SHEET_SEARCH_HISTORY As String = "_lv_search_history"
Private Const COL_OCCURRED_AT As Long = 1
Private Const COL_HISTORY_KIND As Long = 2
Private Const COL_SEARCH_TEXT As Long = 3
Private Const COL_LAW_ID As Long = 4
Private Const COL_LAW_TITLE As Long = 5
Private Const COL_ENFORCEMENT_DATE As Long = 6
Private Const COL_SEARCH_MODE As Long = 7
Private Const COL_HIT_COUNT As Long = 8

Public Function SearchHistorySheetName() As String
    SearchHistorySheetName = SHEET_SEARCH_HISTORY
End Function

Public Function SearchHistoryCount() As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_SEARCH_HISTORY) Then
        SearchHistoryCount = 0
        Exit Function
    End If

    SearchHistoryCount = Application.Max(0, modSheetStore.LastUsedRow(ThisWorkbook.Worksheets(SHEET_SEARCH_HISTORY)) - 1)
End Function

Public Function RecordLawSearchHistory(ByVal searchText As String, ByVal hitCount As Long) As Long
    RecordLawSearchHistory = AppendSearchHistory("LawSearch", searchText, "", "", "", "", hitCount)
End Function

Public Function RecordLawSelectionHistory(ByVal searchText As String, ByVal lawId As String, ByVal lawTitle As String, ByVal enforcementDate As String, ByVal hitCount As Long) As Long
    RecordLawSelectionHistory = AppendSearchHistory("LawSelection", searchText, lawId, lawTitle, enforcementDate, "", hitCount)
End Function

Public Function RecordBodySearchHistory(ByVal addedLawStoreRow As Long, ByVal searchText As String, ByVal searchMode As String, ByVal hitCount As Long) As Long
    Dim lawId As String
    Dim lawTitle As String
    Dim enforcementDate As String
    ReadAddedLawSnapshot addedLawStoreRow, lawId, lawTitle, enforcementDate

    RecordBodySearchHistory = AppendSearchHistory("BodySearch", searchText, lawId, lawTitle, enforcementDate, searchMode, hitCount)
End Function

Public Sub ConfigureSearchHistoryListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 8
    targetListBox.columnWidths = "86 pt;72 pt;150 pt;168 pt;72 pt;54 pt;40 pt;0 pt"
    targetListBox.BoundColumn = 8
End Sub

Public Function PopulateSearchHistoryListBox(ByVal targetListBox As Object, Optional ByVal historyKindFilter As String = "") As Long
    On Error GoTo ErrHandler

    ConfigureSearchHistoryListBox targetListBox
    targetListBox.Clear

    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_SEARCH_HISTORY) Then
        PopulateSearchHistoryListBox = 0
        Exit Function
    End If

    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets(SHEET_SEARCH_HISTORY)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(historySheet)
    If lastRow < 2 Then
        PopulateSearchHistoryListBox = 0
        Exit Function
    End If

    Dim normalizedFilter As String
    normalizedFilter = NormalizeHistoryKindFilter(historyKindFilter)

    Dim rowIndex As Long
    Dim listIndex As Long
    Dim loadedCount As Long
    For rowIndex = lastRow To 2 Step -1
        If Len(normalizedFilter) = 0 Or StrComp(CellText(historySheet.cells(rowIndex, COL_HISTORY_KIND).Value2), normalizedFilter, vbTextCompare) = 0 Then
            targetListBox.AddItem CellText(historySheet.cells(rowIndex, COL_OCCURRED_AT).Value2)
            listIndex = targetListBox.ListCount - 1
            targetListBox.List(listIndex, 1) = HistoryKindDisplayText(CellText(historySheet.cells(rowIndex, COL_HISTORY_KIND).Value2))
            targetListBox.List(listIndex, 2) = CellText(historySheet.cells(rowIndex, COL_SEARCH_TEXT).Value2)
            targetListBox.List(listIndex, 3) = CellText(historySheet.cells(rowIndex, COL_LAW_TITLE).Value2)
            targetListBox.List(listIndex, 4) = CellText(historySheet.cells(rowIndex, COL_ENFORCEMENT_DATE).Value2)
            targetListBox.List(listIndex, 5) = CellText(historySheet.cells(rowIndex, COL_SEARCH_MODE).Value2)
            targetListBox.List(listIndex, 6) = CellText(historySheet.cells(rowIndex, COL_HIT_COUNT).Value2)
            targetListBox.List(listIndex, 7) = CStr(rowIndex)
            loadedCount = loadedCount + 1
        End If
    Next rowIndex

    PopulateSearchHistoryListBox = loadedCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.PopulateSearchHistoryListBox", Err.description, "", "", historyKindFilter
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function SearchHistoryDisplayText(ByVal rowIndex As Long) As String
    On Error GoTo ErrHandler

    If Not IsValidSearchHistoryRow(rowIndex) Then
        Exit Function
    End If

    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets(SHEET_SEARCH_HISTORY)

    SearchHistoryDisplayText = CellText(historySheet.cells(rowIndex, COL_OCCURRED_AT).Value2) _
        & "｜" & HistoryKindDisplayText(CellText(historySheet.cells(rowIndex, COL_HISTORY_KIND).Value2)) _
        & "｜" & CellText(historySheet.cells(rowIndex, COL_SEARCH_TEXT).Value2) _
        & "｜" & CellText(historySheet.cells(rowIndex, COL_LAW_TITLE).Value2)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.SearchHistoryDisplayText", Err.description, "", "", CStr(rowIndex)
End Function

Public Function SearchHistoryRowKind(ByVal rowIndex As Long) As String
    SearchHistoryRowKind = SearchHistoryRowText(rowIndex, COL_HISTORY_KIND)
End Function

Public Function SearchHistoryRowSearchText(ByVal rowIndex As Long) As String
    SearchHistoryRowSearchText = SearchHistoryRowText(rowIndex, COL_SEARCH_TEXT)
End Function

Public Function SearchHistoryRowLawId(ByVal rowIndex As Long) As String
    SearchHistoryRowLawId = SearchHistoryRowText(rowIndex, COL_LAW_ID)
End Function

Public Function SearchHistoryRowLawTitle(ByVal rowIndex As Long) As String
    SearchHistoryRowLawTitle = SearchHistoryRowText(rowIndex, COL_LAW_TITLE)
End Function

Public Function SearchHistoryRowEnforcementDate(ByVal rowIndex As Long) As String
    SearchHistoryRowEnforcementDate = SearchHistoryRowText(rowIndex, COL_ENFORCEMENT_DATE)
End Function

Public Function SearchHistoryRowSearchMode(ByVal rowIndex As Long) As String
    SearchHistoryRowSearchMode = SearchHistoryRowText(rowIndex, COL_SEARCH_MODE)
End Function

Public Function SearchHistoryRowHitCount(ByVal rowIndex As Long) As Long
    SearchHistoryRowHitCount = CLng(Val(SearchHistoryRowText(rowIndex, COL_HIT_COUNT)))
End Function

Public Function SearchHistoryRowOccurredAt(ByVal rowIndex As Long) As String
    SearchHistoryRowOccurredAt = SearchHistoryRowText(rowIndex, COL_OCCURRED_AT)
End Function

Public Function DeleteSearchHistoryEntry(ByVal rowIndex As Long) As Boolean
    On Error GoTo ErrHandler

    DeleteSearchHistoryRow rowIndex
    modAppStartup.SaveWorkbookState ThisWorkbook
    DeleteSearchHistoryEntry = True
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.DeleteSearchHistoryEntry", Err.description, "", "", CStr(rowIndex)
    DeleteSearchHistoryEntry = False
End Function

Public Function DeleteSearchHistoryByKind(ByVal historyKind As String) As Long
    On Error GoTo ErrHandler

    historyKind = NormalizeHistoryKindFilter(historyKind)
    If Len(historyKind) = 0 Then
        DeleteSearchHistoryByKind = 0
        Exit Function
    End If

    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_SEARCH_HISTORY) Then
        DeleteSearchHistoryByKind = 0
        Exit Function
    End If

    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets(SHEET_SEARCH_HISTORY)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(historySheet)

    Dim deletedCount As Long
    Dim rowIndex As Long
    For rowIndex = lastRow To 2 Step -1
        If StrComp(CellText(historySheet.cells(rowIndex, COL_HISTORY_KIND).Value2), historyKind, vbTextCompare) = 0 Then
            historySheet.rows(rowIndex).Delete
            deletedCount = deletedCount + 1
        End If
    Next rowIndex

    If deletedCount > 0 Then
        modAppStartup.SaveWorkbookState ThisWorkbook
    End If
    DeleteSearchHistoryByKind = deletedCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.DeleteSearchHistoryByKind", Err.description, "", "", historyKind
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Sub ClearSearchHistory()
    On Error GoTo ErrHandler

    Dim historySheet As Worksheet
    Set historySheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, SHEET_SEARCH_HISTORY)
    modSheetStore.ClearDataRows historySheet
    modAppStartup.SaveWorkbookState ThisWorkbook
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.ClearSearchHistory", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Function SearchHistorySmoke() As String
    On Error GoTo ErrHandler

    Dim originalCount As Long
    originalCount = SearchHistoryCount()

    Dim testRow As Long
    testRow = AppendSearchHistory("Smoke", "履歴確認", "TEST001", "テスト法", "2025-04-01", "AND", 3)

    If SearchHistoryCount() <> originalCount + 1 Then
        Err.Raise vbObjectError + 7301, "modHistory.SearchHistorySmoke", "Search history count mismatch."
    End If
    If Not SearchHistoryRowMatches(testRow, "Smoke", "履歴確認", "TEST001", "テスト法", "2025-04-01", "AND", 3) Then
        Err.Raise vbObjectError + 7302, "modHistory.SearchHistorySmoke", "Search history row mismatch."
    End If

    DeleteSearchHistoryRow testRow
    If SearchHistoryCount() <> originalCount Then
        Err.Raise vbObjectError + 7303, "modHistory.SearchHistorySmoke", "Search history cleanup failed."
    End If

    SearchHistorySmoke = "ok:" & CStr(originalCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.SearchHistorySmoke", Err.description, "", "", ""
    SearchHistorySmoke = "error:" & Err.description
End Function

Private Function AppendSearchHistory( _
    ByVal historyKind As String, _
    ByVal searchText As String, _
    ByVal lawId As String, _
    ByVal lawTitle As String, _
    ByVal enforcementDate As String, _
    ByVal searchMode As String, _
    ByVal hitCount As Long) As Long

    On Error GoTo ErrHandler

    Dim historySheet As Worksheet
    Set historySheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, SHEET_SEARCH_HISTORY)
    historySheet.cells.NumberFormat = "@"

    AppendSearchHistory = modSheetStore.AppendRow(historySheet, Array( _
        Format$(Now, "yyyy-mm-dd hh:nn:ss"), _
        Trim$(historyKind), _
        Trim$(searchText), _
        Trim$(lawId), _
        Trim$(lawTitle), _
        NormalizeHistoryDateText(enforcementDate), _
        Trim$(searchMode), _
        hitCount))
    modAppStartup.SaveWorkbookState ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.AppendSearchHistory", Err.description, historyKind, searchText, lawId
    AppendSearchHistory = 0
End Function

Private Function SearchHistoryRowText(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    On Error GoTo ErrHandler

    If Not IsValidSearchHistoryRow(rowIndex) Then
        Exit Function
    End If

    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets(SHEET_SEARCH_HISTORY)
    SearchHistoryRowText = CellText(historySheet.cells(rowIndex, columnIndex).Value2)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.SearchHistoryRowText", Err.description, "", "", CStr(rowIndex)
End Function

Private Function IsValidSearchHistoryRow(ByVal rowIndex As Long) As Boolean
    If rowIndex < 2 Then
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_SEARCH_HISTORY) Then
        Exit Function
    End If

    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets(SHEET_SEARCH_HISTORY)
    IsValidSearchHistoryRow = (rowIndex <= modSheetStore.LastUsedRow(historySheet))
End Function

Private Function NormalizeHistoryKindFilter(ByVal historyKind As String) As String
    historyKind = Trim$(historyKind)

    Select Case historyKind
        Case "", "すべて", "全部", "All"
            NormalizeHistoryKindFilter = ""
        Case "法令検索", "LawSearch"
            NormalizeHistoryKindFilter = "LawSearch"
        Case "法令選択", "LawSelection"
            NormalizeHistoryKindFilter = "LawSelection"
        Case "条文検索", "BodySearch"
            NormalizeHistoryKindFilter = "BodySearch"
        Case Else
            NormalizeHistoryKindFilter = historyKind
    End Select
End Function

Public Function HistoryKindDisplayText(ByVal historyKind As String) As String
    Select Case Trim$(historyKind)
        Case "LawSearch"
            HistoryKindDisplayText = "法令検索"
        Case "LawSelection"
            HistoryKindDisplayText = "法令選択"
        Case "BodySearch"
            HistoryKindDisplayText = "条文検索"
        Case Else
            HistoryKindDisplayText = historyKind
    End Select
End Function

Private Sub ReadAddedLawSnapshot(ByVal addedLawStoreRow As Long, ByRef lawId As String, ByRef lawTitle As String, ByRef enforcementDate As String)
    On Error GoTo ErrHandler

    If addedLawStoreRow < 2 Then
        Exit Sub
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, modAddedLawStore.AddedLawStoreSheetName()) Then
        Exit Sub
    End If

    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Worksheets(modAddedLawStore.AddedLawStoreSheetName())
    If addedLawStoreRow > modSheetStore.LastUsedRow(storeSheet) Then
        Exit Sub
    End If

    lawId = CellText(storeSheet.cells(addedLawStoreRow, 2).Value2)
    lawTitle = CellText(storeSheet.cells(addedLawStoreRow, 6).Value2)
    enforcementDate = NormalizeHistoryDateText(storeSheet.cells(addedLawStoreRow, 3).Value2)
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.ReadAddedLawSnapshot", Err.description, "", "", CStr(addedLawStoreRow)
End Sub

Private Function NormalizeHistoryDateText(ByVal value As Variant) As String
    If IsDate(value) Then
        NormalizeHistoryDateText = Format$(CDate(value), "yyyy-mm-dd")
    ElseIf IsNumeric(value) Then
        NormalizeHistoryDateText = Format$(CDate(CDbl(value)), "yyyy-mm-dd")
    Else
        NormalizeHistoryDateText = Trim$(CStr(value))
    End If
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function

Private Function SearchHistoryRowMatches( _
    ByVal rowIndex As Long, _
    ByVal historyKind As String, _
    ByVal searchText As String, _
    ByVal lawId As String, _
    ByVal lawTitle As String, _
    ByVal enforcementDate As String, _
    ByVal searchMode As String, _
    ByVal hitCount As Long) As Boolean

    If rowIndex < 2 Then
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_SEARCH_HISTORY) Then
        Exit Function
    End If

    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets(SHEET_SEARCH_HISTORY)
    If rowIndex > modSheetStore.LastUsedRow(historySheet) Then
        Exit Function
    End If

    SearchHistoryRowMatches = StrComp(CellText(historySheet.cells(rowIndex, COL_HISTORY_KIND).Value2), historyKind, vbBinaryCompare) = 0 _
        And StrComp(CellText(historySheet.cells(rowIndex, COL_SEARCH_TEXT).Value2), searchText, vbBinaryCompare) = 0 _
        And StrComp(CellText(historySheet.cells(rowIndex, COL_LAW_ID).Value2), lawId, vbBinaryCompare) = 0 _
        And StrComp(CellText(historySheet.cells(rowIndex, COL_LAW_TITLE).Value2), lawTitle, vbBinaryCompare) = 0 _
        And StrComp(CellText(historySheet.cells(rowIndex, COL_ENFORCEMENT_DATE).Value2), enforcementDate, vbBinaryCompare) = 0 _
        And StrComp(CellText(historySheet.cells(rowIndex, COL_SEARCH_MODE).Value2), searchMode, vbBinaryCompare) = 0 _
        And CLng(Val(CellText(historySheet.cells(rowIndex, COL_HIT_COUNT).Value2))) = hitCount
End Function

Private Sub DeleteSearchHistoryRow(ByVal rowIndex As Long)
    On Error GoTo ErrHandler

    If rowIndex < 2 Then
        Exit Sub
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_SEARCH_HISTORY) Then
        Exit Sub
    End If

    Dim historySheet As Worksheet
    Set historySheet = ThisWorkbook.Worksheets(SHEET_SEARCH_HISTORY)
    If rowIndex > modSheetStore.LastUsedRow(historySheet) Then
        Exit Sub
    End If

    historySheet.rows(rowIndex).Delete
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "History", "modHistory.DeleteSearchHistoryRow", Err.description, "", "", CStr(rowIndex)
    Err.Raise Err.Number, Err.source, Err.description
End Sub
