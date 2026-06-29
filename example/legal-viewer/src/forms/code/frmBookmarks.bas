Option Explicit

Private Const FORM_WIDTH_POINTS As Double = 820
Private Const FORM_HEIGHT_POINTS As Double = 560

Private Sub UserForm_Initialize()
    ConfigureFormWindow
    ConfigureControls
    modUiState.ApplyDefaultUserFormFont Me
    RefreshBookmarkListForForm
End Sub

Private Sub UserForm_Activate()
    modUiState.ApplyDefaultUserFormFont Me
    modWindowPlacement.CenterUserFormOnExcelMonitor Me
End Sub

Private Sub cmdClearFilter_Click()
    Me.txtFilter.text = ""
    RefreshBookmarkListForForm
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    UpdateSelectedBookmark
End Sub

Private Sub cmdOpen_Click()
    Dim resultText As String
    resultText = OpenSelectedBookmark()
    If Left$(resultText, 3) = "ok:" Then
        Unload Me
    End If
End Sub

Private Sub cmdDelete_Click()
    DeleteSelectedBookmark
End Sub

Private Sub lstBookmarks_Click()
    UpdateSelectedBookmarkFields
End Sub

Private Sub txtFilter_Change()
    RefreshBookmarkListForForm
End Sub

Public Function RefreshBookmarkListForForm() As String
    On Error GoTo ErrHandler

    Dim selectedRow As Long
    selectedRow = SelectedBookmarkStoreRow()

    Dim itemCount As Long
    itemCount = modBookmark.PopulateBookmarkListBox(Me.lstBookmarks, Me.txtFilter.text)

    If itemCount > 0 Then
        SelectBookmarkRowByStoreRow selectedRow
        If Me.lstBookmarks.listIndex < 0 Then
            Me.lstBookmarks.listIndex = 0
        End If
        UpdateSelectedBookmarkFields
    Else
        ClearSelectedBookmarkFields
    End If

    Me.lblStatus.caption = "ブックマーク " & CStr(itemCount) & " 件"
    RefreshBookmarkListForForm = "ok:" & CStr(itemCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.RefreshBookmarkListForForm", Err.description, "", "", ""
    Me.lblStatus.caption = "ブックマークの読み込み失敗: " & Err.description
    RefreshBookmarkListForForm = "error:" & Err.description
End Function

Public Function UpdateSelectedBookmark() As String
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedBookmarkStoreRow()
    If rowIndex <= 0 Then
        Me.lblStatus.caption = "ブックマークを選択してください"
        UpdateSelectedBookmark = "empty"
        Exit Function
    End If

    If Not modBookmark.UpdateBookmarkTagsMemo(rowIndex, Me.txtTags.text, Me.txtMemo.text) Then
        Me.lblStatus.caption = "ブックマークを保存できませんでした"
        UpdateSelectedBookmark = "empty"
        Exit Function
    End If

    RefreshBookmarkListForForm
    SelectBookmarkRowByStoreRow rowIndex
    UpdateSelectedBookmarkFields
    Me.lblStatus.caption = "ブックマークを保存しました"
    UpdateSelectedBookmark = "ok:" & CStr(rowIndex)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.UpdateSelectedBookmark", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "ブックマーク保存失敗: " & Err.description
    UpdateSelectedBookmark = "error:" & Err.description
End Function

Public Function OpenSelectedBookmark() As String
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedBookmarkStoreRow()
    If rowIndex <= 0 Then
        Me.lblStatus.caption = "ブックマークを選択してください"
        OpenSelectedBookmark = "empty"
        Exit Function
    End If

    Dim resolvedBodyRow As Long
    resolvedBodyRow = modBookmark.FindBookmarkBodyUnitRow(rowIndex)

    Dim loadResult As String
    If resolvedBodyRow <= 0 Then
        loadResult = EnsureBookmarkLawLoadedForOpen(rowIndex)
        If Left$(loadResult, 3) <> "ok:" Then
            Me.lblStatus.caption = "ブックマーク対象の法令を再読込できませんでした: " & loadResult
            OpenSelectedBookmark = loadResult
            Exit Function
        End If

        resolvedBodyRow = modBookmark.FindBookmarkBodyUnitRow(rowIndex)
    End If
    If resolvedBodyRow <= 0 Then
        Me.lblStatus.caption = "本文単位を見つけられません"
        OpenSelectedBookmark = "error:body-unit-not-found"
        Exit Function
    End If

    Load frmMain
    frmMain.RefreshAddedLawsForForm
    Dim resolvedAddedRow As Long
    resolvedAddedRow = modAddedLawStore.FindAddedLawRow(modBookmark.BookmarkRowLawId(rowIndex), modBookmark.BookmarkRowEnforcementDate(rowIndex))
    If resolvedAddedRow > 0 Then
        frmMain.SelectAddedLawStoreRowForForm resolvedAddedRow
    End If
    frmMain.SelectBodyUnitStoreRowForForm resolvedBodyRow

    Me.lblStatus.caption = "本文へ移動しました"
    OpenSelectedBookmark = "ok:" & CStr(resolvedBodyRow)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.OpenSelectedBookmark", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "ブックマークを開けません: " & Err.description
    OpenSelectedBookmark = "error:" & Err.description
End Function

Private Function EnsureBookmarkLawLoadedForOpen(ByVal rowIndex As Long) As String
    On Error GoTo ErrHandler

    Dim lawId As String
    lawId = modBookmark.BookmarkRowLawId(rowIndex)
    If Len(lawId) = 0 Then
        EnsureBookmarkLawLoadedForOpen = "empty"
        Exit Function
    End If

    Dim enforcementDate As String
    enforcementDate = modBookmark.BookmarkRowEnforcementDate(rowIndex)
    If Len(enforcementDate) = 0 Then
        EnsureBookmarkLawLoadedForOpen = "empty"
        Exit Function
    End If

    Load frmLawAdd

    Dim searchTerms(1 To 3) As String
    searchTerms(1) = lawId
    searchTerms(2) = modBookmark.BookmarkRowLawNum(rowIndex)
    searchTerms(3) = modBookmark.BookmarkRowLawTitle(rowIndex)

    Dim searchIndex As Long
    For searchIndex = LBound(searchTerms) To UBound(searchTerms)
        If Len(searchTerms(searchIndex)) = 0 Then
            GoTo NextSearchTerm
        End If

        If modLawSearch.SearchLawsForListBox(searchTerms(searchIndex), frmLawAdd.lstResults, 100) <= 0 Then
            GoTo NextSearchTerm
        End If

        If Left$(frmLawAdd.SelectResultForLawKey(lawId, enforcementDate), 3) <> "ok:" Then
            GoTo NextSearchTerm
        End If

        EnsureBookmarkLawLoadedForOpen = frmLawAdd.AddSelectedLawForForm()
        If Left$(EnsureBookmarkLawLoadedForOpen, 3) = "ok:" Then
            Exit Function
        End If

NextSearchTerm:
    Next searchIndex

    EnsureBookmarkLawLoadedForOpen = "empty"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.EnsureBookmarkLawLoadedForOpen", Err.description, "", "", CStr(rowIndex)
    EnsureBookmarkLawLoadedForOpen = "error:" & Err.description
End Function

Public Function DeleteSelectedBookmark() As String
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedBookmarkStoreRow()
    If rowIndex <= 0 Then
        Me.lblStatus.caption = "ブックマークを選択してください"
        DeleteSelectedBookmark = "empty"
        Exit Function
    End If

    If XlflowUI.MsgBox("bookmark-delete-confirm", "選択したブックマークを削除しますか？", vbYesNo + vbQuestion, "法令検索ビューアー", vbNo) <> vbYes Then
        Me.lblStatus.caption = "削除をキャンセルしました"
        DeleteSelectedBookmark = "cancel"
        Exit Function
    End If

    modBookmark.DeleteBookmarkEntry rowIndex
    RefreshBookmarkListForForm
    Me.lblStatus.caption = "ブックマークを削除しました"
    DeleteSelectedBookmark = "ok"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.DeleteSelectedBookmark", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "ブックマーク削除失敗: " & Err.description
    DeleteSelectedBookmark = "error:" & Err.description
End Function

Public Function BookmarkFormSmoke() As String
    On Error GoTo ErrHandler

    Load frmLawAdd
    Dim searchResult As String
    searchResult = frmLawAdd.SearchForKeyword("民法")
    If Left$(searchResult, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7601, "frmBookmarks.BookmarkFormSmoke", "Law search failed."
    End If
    If frmLawAdd.lstResults.ListCount = 0 Then
        Err.Raise vbObjectError + 7602, "frmBookmarks.BookmarkFormSmoke", "Law search returned no results."
    End If

    Dim selectResult As String
    selectResult = frmLawAdd.SelectResultForForm(0)
    If Left$(selectResult, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7603, "frmBookmarks.BookmarkFormSmoke", "Law selection failed."
    End If

    Dim addLawResult As String
    addLawResult = frmLawAdd.AddSelectedLawForForm()
    If Left$(addLawResult, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7604, "frmBookmarks.BookmarkFormSmoke", "Law add failed."
    End If

    Load frmMain
    frmMain.RefreshAddedLawsForForm

    Dim addedRow As Long
    addedRow = modAddedLawStore.FindAddedLawRow(frmLawAdd.SelectedLawId(), frmLawAdd.SelectedEnforcementDate())
    If addedRow <= 0 Then
        Err.Raise vbObjectError + 7605, "frmBookmarks.BookmarkFormSmoke", "Added law row was not found."
    End If

    frmMain.SelectAddedLawStoreRowForForm addedRow
    If frmMain.lstBodyNav.ListCount <= 0 Then
        Err.Raise vbObjectError + 7606, "frmBookmarks.BookmarkFormSmoke", "Body navigation list was not populated."
    End If

    frmMain.SelectBodyUnitStoreRowForForm CLng(Val(frmMain.lstBodyNav.List(0, 3)))
    If frmMain.SelectedCitationBodyUnitStoreRowForForm() <= 0 Then
        Err.Raise vbObjectError + 7607, "frmBookmarks.BookmarkFormSmoke", "Selected body unit row was not resolved."
    End If

    Dim addResult As String
    addResult = frmMain.AddBookmarkForForm()
    If Left$(addResult, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7608, "frmBookmarks.BookmarkFormSmoke", "Bookmark add failed."
    End If

    Dim bookmarkRow As Long
    bookmarkRow = CLng(Val(Mid$(addResult, 4)))
    If bookmarkRow <= 0 Then
        Err.Raise vbObjectError + 7609, "frmBookmarks.BookmarkFormSmoke", "Bookmark row was not returned."
    End If

    Load frmBookmarks
    If Me.cmdUpdate.caption <> "タグ/メモ保存" Then
        Err.Raise vbObjectError + 7614, "frmBookmarks.BookmarkFormSmoke", "Bookmark save caption was not updated."
    End If
    If Me.cmdUpdate.Width < 96 Then
        Err.Raise vbObjectError + 7615, "frmBookmarks.BookmarkFormSmoke", "Bookmark save button width was not expanded."
    End If
    If Me.lblStatus.Top > 512 Or Me.lblStatus.Height < 20 Then
        Err.Raise vbObjectError + 7616, "frmBookmarks.BookmarkFormSmoke", "Bookmark status label was not moved up enough."
    End If
    RefreshBookmarkListForForm
    SelectBookmarkRowByStoreRow bookmarkRow
    If SelectedBookmarkStoreRow() <= 0 Then
        Err.Raise vbObjectError + 7617, "frmBookmarks.BookmarkFormSmoke", "Bookmark row was not selected."
    End If

    Me.txtTags.text = "smoke"
    Me.txtMemo.text = "bookmark smoke"
    Dim updateResult As String
    updateResult = UpdateSelectedBookmark()
    If Left$(updateResult, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7618, "frmBookmarks.BookmarkFormSmoke", "Bookmark update failed."
    End If

    If modSheetStore.SheetExists(ThisWorkbook, modLawParser.BodyUnitSheetName()) Then
        ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName()).Visible = xlSheetVisible
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(modLawParser.BodyUnitSheetName()).Delete
        Application.DisplayAlerts = True
    End If

    Dim openResult As String
    openResult = OpenSelectedBookmark()
    If Left$(openResult, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7619, "frmBookmarks.BookmarkFormSmoke", "Bookmark open failed."
    End If
    If frmMain.SelectedCitationBodyUnitStoreRowForForm() <= 0 Then
        Err.Raise vbObjectError + 7620, "frmBookmarks.BookmarkFormSmoke", "Bookmark open did not select a body unit."
    End If

    modBookmark.DeleteBookmarkEntry bookmarkRow

    BookmarkFormSmoke = "ok:" & CStr(bookmarkRow)
    GoTo Cleanup

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.BookmarkFormSmoke", Err.description, "", "", ""
    BookmarkFormSmoke = "error:" & CStr(Err.Number) & ":" & Err.description

Cleanup:
    DeleteBookmarkSmokeTempCaches
    UnloadFormIfLoaded "frmLawAdd"
    UnloadFormIfLoaded "frmMain"
    UnloadFormIfLoaded "frmBookmarks"
End Function

Private Sub ConfigureFormWindow()
    Me.caption = "ブックマーク"
    Me.Width = FORM_WIDTH_POINTS
    Me.Height = FORM_HEIGHT_POINTS
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub ConfigureControls()
    modBookmark.ConfigureBookmarkListBox Me.lstBookmarks

    Me.txtTags.text = ""
    Me.txtMemo.text = ""
    Me.txtMemo.MultiLine = True
    Me.txtMemo.WordWrap = True
    Me.txtMemo.ScrollBars = 2
    modUiState.ApplyDefaultControlFont Me.txtTags
    modUiState.ApplyDefaultControlFont Me.txtMemo
End Sub

Private Sub UpdateSelectedBookmarkFields()
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedBookmarkStoreRow()
    If rowIndex <= 0 Then
        ClearSelectedBookmarkFields
        Exit Sub
    End If

    Me.txtTags.text = modBookmark.BookmarkRowTags(rowIndex)
    Me.txtMemo.text = modBookmark.BookmarkRowMemo(rowIndex)
    Me.lblStatus.caption = modBookmark.BookmarkRowDisplayText(rowIndex)
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.UpdateSelectedBookmarkFields", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "選択中ブックマークの読み込み失敗: " & Err.description
End Sub

Private Sub ClearSelectedBookmarkFields()
    Me.txtTags.text = ""
    Me.txtMemo.text = ""
End Sub

Private Function SelectedBookmarkStoreRow() As Long
    On Error GoTo ErrHandler

    If Me.lstBookmarks.listIndex < 0 Then
        Exit Function
    End If

    SelectedBookmarkStoreRow = CLng(Val(Me.lstBookmarks.List(Me.lstBookmarks.listIndex, 7)))
    Exit Function

ErrHandler:
    SelectedBookmarkStoreRow = 0
End Function

Private Sub DeleteBookmarkSmokeTempCaches()
    On Error GoTo ErrHandler

    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.DeleteBookmarkSmokeTempCaches", Err.description, "", "", ""
End Sub

Private Sub UnloadFormIfLoaded(ByVal formName As String)
    On Error GoTo ErrHandler

    If Not IsUserFormLoaded(formName) Then Exit Sub

    Select Case formName
        Case "frmLawAdd"
            Unload frmLawAdd
        Case "frmMain"
            Unload frmMain
        Case "frmBookmarks"
            Unload Me
    End Select
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "BookmarkForm", "frmBookmarks.UnloadFormIfLoaded", Err.description, "", "", formName
End Sub

Private Function IsUserFormLoaded(ByVal formName As String) As Boolean
    Dim openForm As Object
    For Each openForm In VBA.UserForms
        If StrComp(TypeName(openForm), formName, vbTextCompare) = 0 Then
            IsUserFormLoaded = True
            Exit Function
        End If
    Next openForm
End Function

Private Sub SelectBookmarkRowByStoreRow(ByVal rowIndex As Long)
    Dim listIndex As Long
    If rowIndex <= 0 Then
        Me.lstBookmarks.listIndex = -1
        Exit Sub
    End If

    For listIndex = 0 To Me.lstBookmarks.ListCount - 1
        If CLng(Val(Me.lstBookmarks.List(listIndex, 7))) = rowIndex Then
            Me.lstBookmarks.listIndex = listIndex
            Exit Sub
        End If
    Next listIndex
End Sub
