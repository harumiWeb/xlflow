Attribute VB_Name = "modSheetStore"
Option Explicit

Private Const SHEET_SETTINGS As String = "_lv_settings"
Private Const SHEET_API_LOG As String = "_lv_api_log"
Private Const SHEET_ERROR_LOG As String = "_lv_error_log"
Private Const SHEET_SEARCH_HISTORY As String = "_lv_search_history"
Private Const SHEET_BOOKMARKS As String = "_lv_bookmarks"
Private Const SHEET_BOOKMARK_TAGS As String = "_lv_bookmark_tags"
Private Const SHEET_ALIAS_MASTER As String = "_lv_alias_master"
Private Const TEMP_PREFIX As String = "_lv_tmp_"

Public Function SettingsSheetName() As String
    SettingsSheetName = SHEET_SETTINGS
End Function

Public Function ApiLogSheetName() As String
    ApiLogSheetName = SHEET_API_LOG
End Function

Public Function ErrorLogSheetName() As String
    ErrorLogSheetName = SHEET_ERROR_LOG
End Function

Public Function AliasMasterSheetName() As String
    AliasMasterSheetName = SHEET_ALIAS_MASTER
End Function

Public Function BookmarksSheetName() As String
    BookmarksSheetName = SHEET_BOOKMARKS
End Function

Public Function BookmarkTagsSheetName() As String
    BookmarkTagsSheetName = SHEET_BOOKMARK_TAGS
End Function

Public Function TempSheetPrefix() As String
    TempSheetPrefix = TEMP_PREFIX
End Function

Public Sub EnsurePersistentSheets(ByVal targetWorkbook As Workbook)
    On Error GoTo ErrHandler

    Dim sheetNames As Variant
    sheetNames = KnownPersistentSheetNames()

    Dim index As Long
    For index = LBound(sheetNames) To UBound(sheetNames)
        EnsurePersistentSheet targetWorkbook, CStr(sheetNames(index))
    Next index

    EnsureAliasDefaults targetWorkbook
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modSheetStore.EnsurePersistentSheets", Err.description
End Sub

Public Function EnsurePersistentSheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Worksheet
    On Error GoTo ErrHandler

    Dim targetSheet As Worksheet
    If SheetExists(targetWorkbook, sheetName) Then
        Set targetSheet = targetWorkbook.Worksheets(sheetName)
    Else
        Set targetSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count))
        targetSheet.Name = sheetName
    End If

    EnsureHeaders targetSheet, HeaderRowForSheet(sheetName)
    Set EnsurePersistentSheet = targetSheet
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "modSheetStore.EnsurePersistentSheet", Err.description
End Function

Public Function SheetExists(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Boolean
    Dim targetSheet As Worksheet
    For Each targetSheet In targetWorkbook.Worksheets
        If StrComp(targetSheet.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next targetSheet
End Function

Public Sub ApplyManagedSheetVisibility(ByVal targetWorkbook As Workbook)
    On Error GoTo ErrHandler

    Dim visibilityMode As String
    visibilityMode = modSettings.GetSettingText("ManagedSheetVisibility", "VeryHidden")

    Dim sheetNames As Variant
    sheetNames = KnownPersistentSheetNames()

    Dim index As Long
    For index = LBound(sheetNames) To UBound(sheetNames)
        If SheetExists(targetWorkbook, CStr(sheetNames(index))) Then
            SetWorksheetVisibility targetWorkbook.Worksheets(CStr(sheetNames(index))), visibilityMode
        End If
    Next index
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modSheetStore.ApplyManagedSheetVisibility", Err.description
End Sub

Public Sub DeleteSheetsWithPrefix(ByVal targetWorkbook As Workbook, ByVal prefix As String)
    On Error GoTo ErrHandler

    Dim originalAlerts As Boolean
    originalAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Dim index As Long
    For index = targetWorkbook.Worksheets.Count To 1 Step -1
        If Left$(targetWorkbook.Worksheets(index).Name, Len(prefix)) = prefix Then
            targetWorkbook.Worksheets(index).Visible = xlSheetVisible
            targetWorkbook.Worksheets(index).Delete
        End If
    Next index

Cleanup:
    Application.DisplayAlerts = originalAlerts
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = originalAlerts
    Err.Raise Err.Number, "modSheetStore.DeleteSheetsWithPrefix", Err.description
End Sub

Public Function AppendRow(ByVal targetSheet As Worksheet, ByVal rowValues As Variant) As Long
    On Error GoTo ErrHandler

    Dim nextRow As Long
    nextRow = LastUsedRow(targetSheet) + 1
    If nextRow < 2 Then
        nextRow = 2
    End If

    Dim columnIndex As Long
    For columnIndex = LBound(rowValues) To UBound(rowValues)
        targetSheet.cells(nextRow, columnIndex - LBound(rowValues) + 1).Value2 = rowValues(columnIndex)
    Next columnIndex

    AppendRow = nextRow
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "modSheetStore.AppendRow", Err.description
End Function

Public Function LastUsedRow(ByVal targetSheet As Worksheet) As Long
    If Application.CountA(targetSheet.cells) = 0 Then
        LastUsedRow = 0
    Else
        LastUsedRow = targetSheet.cells(targetSheet.rows.Count, 1).End(xlUp).Row
    End If
End Function

Public Sub ClearDataRows(ByVal targetSheet As Worksheet)
    On Error GoTo ErrHandler

    Dim lastRow As Long
    lastRow = LastUsedRow(targetSheet)
    If lastRow > 1 Then
        targetSheet.rows("2:" & CStr(lastRow)).ClearContents
    End If
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modSheetStore.ClearDataRows", Err.description
End Sub

Private Sub EnsureHeaders(ByVal targetSheet As Worksheet, ByVal headers As Variant)
    On Error GoTo ErrHandler

    If Application.CountA(targetSheet.rows(1)) > 0 Then
        Exit Sub
    End If

    Dim index As Long
    For index = LBound(headers) To UBound(headers)
        targetSheet.cells(1, index - LBound(headers) + 1).Value2 = headers(index)
    Next index
    targetSheet.rows(1).Font.Bold = True
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modSheetStore.EnsureHeaders", Err.description
End Sub

Private Sub SetWorksheetVisibility(ByVal targetSheet As Worksheet, ByVal visibilityMode As String)
    Select Case LCase$(Trim$(visibilityMode))
        Case "visible"
            targetSheet.Visible = xlSheetVisible
        Case "hidden"
            targetSheet.Visible = xlSheetHidden
        Case Else
            targetSheet.Visible = xlSheetVeryHidden
    End Select
End Sub

Private Function KnownPersistentSheetNames() As Variant
    KnownPersistentSheetNames = Array( _
        SHEET_SETTINGS, _
        SHEET_API_LOG, _
        SHEET_ERROR_LOG, _
        SHEET_SEARCH_HISTORY, _
        SHEET_BOOKMARKS, _
        SHEET_BOOKMARK_TAGS, _
        SHEET_ALIAS_MASTER)
End Function

Private Function HeaderRowForSheet(ByVal sheetName As String) As Variant
    Select Case sheetName
        Case SHEET_SETTINGS
            HeaderRowForSheet = Array("Key", "Value", "Description", "UpdatedAt")
        Case SHEET_API_LOG
            HeaderRowForSheet = Array("OccurredAt", "ProcessKind", "Endpoint", "HttpStatus", "Result", "LawId", "EnforcementDate", "ItemCount", "ElapsedMs", "ErrorMessage")
        Case SHEET_ERROR_LOG
            HeaderRowForSheet = Array("OccurredAt", "ProcessKind", "ProcedureName", "ErrorMessage", "LawId", "EnforcementDate", "Detail")
        Case SHEET_SEARCH_HISTORY
            HeaderRowForSheet = Array("OccurredAt", "HistoryKind", "SearchText", "LawId", "LawTitle", "EnforcementDate", "SearchMode", "HitCount")
        Case SHEET_BOOKMARKS
            HeaderRowForSheet = Array("BookmarkId", "LawId", "LawTitle", "LawNum", "EnforcementDate", "ArticleNo", "ParagraphNo", "ItemNo", "Caption", "BookmarkKind", "Tags", "Memo", "CreatedAt")
        Case SHEET_BOOKMARK_TAGS
            HeaderRowForSheet = Array("Tag", "CreatedAt", "UpdatedAt")
        Case SHEET_ALIAS_MASTER
            HeaderRowForSheet = Array("Alias", "LawTitle", "Note", "UpdatedAt")
        Case Else
            HeaderRowForSheet = Array("Key", "Value")
    End Select
End Function

Private Sub EnsureAliasDefaults(ByVal targetWorkbook As Workbook)
    On Error GoTo ErrHandler

    Dim aliasSheet As Worksheet
    Set aliasSheet = EnsurePersistentSheet(targetWorkbook, SHEET_ALIAS_MASTER)
    If LastUsedRow(aliasSheet) > 1 Then
        Exit Sub
    End If

    Dim defaults As Variant
    defaults = Array( _
        Array("労基法", "労働基準法", "初期略称"), _
        Array("安衛法", "労働安全衛生法", "初期略称"), _
        Array("労災法", "労働者災害補償保険法", "初期略称"), _
        Array("雇保法", "雇用保険法", "初期略称"), _
        Array("徴収法", "労働保険の保険料の徴収等に関する法律", "初期略称"), _
        Array("健保法", "健康保険法", "初期略称"), _
        Array("国年法", "国民年金法", "初期略称"), _
        Array("厚年法", "厚生年金保険法", "初期略称"), _
        Array("育介法", "育児休業、介護休業等育児又は家族介護を行う労働者の福祉に関する法律", "初期略称"), _
        Array("均等法", "雇用の分野における男女の均等な機会及び待遇の確保等に関する法律", "初期略称"), _
        Array("パート有期法", "短時間労働者及び有期雇用労働者の雇用管理の改善等に関する法律", "初期略称"), _
        Array("最低賃金法", "最低賃金法", "初期略称"), _
        Array("労契法", "労働契約法", "初期略称"), _
        Array("職安法", "職業安定法", "初期略称"), _
        Array("派遣法", "労働者派遣事業の適正な運営の確保等に関する法律", "初期略称"), _
        Array("高年法", "高年齢者等の雇用の安定等に関する法律", "初期略称"), _
        Array("障害者雇用促進法", "障害者の雇用の促進等に関する法律", "初期略称"))

    Dim index As Long
    For index = LBound(defaults) To UBound(defaults)
        AppendRow aliasSheet, Array(defaults(index)(0), defaults(index)(1), defaults(index)(2), Format$(Now, "yyyy-mm-dd hh:nn:ss"))
    Next index
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modSheetStore.EnsureAliasDefaults", Err.description
End Sub
