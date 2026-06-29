Option Explicit

Private Const FORM_WIDTH_POINTS As Double = 760
Private Const FORM_HEIGHT_POINTS As Double = 500

Private Sub UserForm_Initialize()
    ConfigureFormWindow
    ConfigureControls
    modUiState.ApplyDefaultUserFormFont Me
    RefreshHistoryForForm
End Sub

Private Sub UserForm_Activate()
    modUiState.ApplyDefaultUserFormFont Me
    modWindowPlacement.CenterUserFormOnExcelMonitor Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdReAdd_Click()
    On Error GoTo ErrHandler

    Dim resultText As String
    resultText = ReplaySelectedHistoryEntry(True)
    If Left$(resultText, 3) <> "ok:" Then Exit Sub

    Load frmMain
    frmMain.RefreshAddedLawsForForm
    UnloadFormIfLoaded "frmLawAdd"
    Unload Me
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.cmdReAdd_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "履歴の開く失敗: " & Err.description
End Sub

Private Sub cmdDelete_Click()
    DeleteSelectedHistoryEntry
End Sub

Private Sub cmdDeleteKind_Click()
    DeleteFilteredHistoryEntries
End Sub

Private Sub cmdDeleteAll_Click()
    DeleteAllHistoryEntries
End Sub

Private Sub cmbKindFilter_Change()
    RefreshHistoryForForm
End Sub

Private Sub lstHistory_Click()
    UpdateSelectedHistoryDetail
End Sub

Public Function RefreshHistoryForForm() As String
    On Error GoTo ErrHandler

    Dim itemCount As Long
    itemCount = modHistory.PopulateSearchHistoryListBox(Me.lstHistory, HistoryFilterForForm())

    If itemCount > 0 Then
        Me.lstHistory.listIndex = 0
        UpdateSelectedHistoryDetail
    Else
        ClearDetail
    End If

    Me.lblStatus.caption = "履歴 " & CStr(itemCount) & " 件"
    RefreshHistoryForForm = "ok:" & CStr(itemCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.RefreshHistoryForForm", Err.description, "", "", ""
    Me.lblStatus.caption = "履歴の読み込み失敗: " & Err.description
    RefreshHistoryForForm = "error:" & Err.description
End Function

Public Function ReplaySelectedHistoryEntry(ByVal addAfterSearch As Boolean) As String
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedHistoryStoreRow()
    If rowIndex <= 0 Then
        Me.lblStatus.caption = "履歴を選択してください"
        ReplaySelectedHistoryEntry = "empty"
        Exit Function
    End If

    Dim historyKind As String
    historyKind = modHistory.SearchHistoryRowKind(rowIndex)

    Select Case historyKind
        Case "LawSearch"
            ReplaySelectedHistoryEntry = ReplayLawSearchHistory(rowIndex)
        Case "LawSelection"
            ReplaySelectedHistoryEntry = ReplayLawSelectionHistory(rowIndex, addAfterSearch)
        Case "BodySearch"
            ReplaySelectedHistoryEntry = ReplayBodySearchHistory(rowIndex)
        Case Else
            Me.lblStatus.caption = "未対応の履歴種別です"
            ReplaySelectedHistoryEntry = "empty"
    End Select
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.ReplaySelectedHistoryEntry", Err.description, "", "", CStr(addAfterSearch)
    Me.lblStatus.caption = "履歴の再実行失敗: " & Err.description
    ReplaySelectedHistoryEntry = "error:" & Err.description
End Function

Public Function ReplayLawSearchHistory(ByVal rowIndex As Long) As String
    On Error GoTo ErrHandler

    Load frmLawAdd
    Dim keyword As String
    keyword = modHistory.SearchHistoryRowSearchText(rowIndex)
    If Len(keyword) = 0 Then
        keyword = modHistory.SearchHistoryRowLawTitle(rowIndex)
    End If

    Dim resultText As String
    resultText = frmLawAdd.SearchForKeyword(keyword)
    If Left$(resultText, 3) <> "ok:" Then
        Me.lblStatus.caption = "法令検索の再実行に失敗しました"
        ReplayLawSearchHistory = resultText
        Exit Function
    End If

    XlflowUI.ShowForm frmLawAdd, False
    Me.lblStatus.caption = "法令検索を再実行しました"
    ReplayLawSearchHistory = resultText
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.ReplayLawSearchHistory", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "法令検索の再実行失敗: " & Err.description
    ReplayLawSearchHistory = "error:" & Err.description
End Function

Public Function ReplayLawSelectionHistory(ByVal rowIndex As Long, ByVal addAfterSearch As Boolean) As String
    On Error GoTo ErrHandler

    Load frmLawAdd

    Dim keyword As String
    keyword = modHistory.SearchHistoryRowSearchText(rowIndex)
    If Len(keyword) = 0 Then
        keyword = modHistory.SearchHistoryRowLawTitle(rowIndex)
    End If

    Dim resultText As String
    resultText = frmLawAdd.SearchForKeyword(keyword)
    If Left$(resultText, 3) <> "ok:" Then
        Me.lblStatus.caption = "再検索に失敗しました"
        ReplayLawSelectionHistory = resultText
        Exit Function
    End If

    resultText = frmLawAdd.SelectResultForLawKey(modHistory.SearchHistoryRowLawId(rowIndex), modHistory.SearchHistoryRowEnforcementDate(rowIndex))
    If Left$(resultText, 3) = "error:" Then
        Me.lblStatus.caption = "検索結果の選択に失敗しました"
        ReplayLawSelectionHistory = resultText
        Exit Function
    End If

    If addAfterSearch Then
        resultText = frmLawAdd.AddSelectedLawForForm()
        If Left$(resultText, 3) <> "ok:" Then
            Me.lblStatus.caption = "再追加に失敗しました"
            ReplayLawSelectionHistory = resultText
            Exit Function
        End If
        Me.lblStatus.caption = "法令を再追加しました"
    Else
        Me.lblStatus.caption = "法令検索と候補選択を再実行しました"
    End If

    XlflowUI.ShowForm frmLawAdd, False
    ReplayLawSelectionHistory = resultText
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.ReplayLawSelectionHistory", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "法令選択履歴の再実行失敗: " & Err.description
    ReplayLawSelectionHistory = "error:" & Err.description
End Function

Public Function ReplayBodySearchHistory(ByVal rowIndex As Long) As String
    On Error GoTo ErrHandler

    Dim storeRow As Long
    storeRow = modAddedLawStore.FindAddedLawRow(modHistory.SearchHistoryRowLawId(rowIndex), modHistory.SearchHistoryRowEnforcementDate(rowIndex))
    If storeRow <= 0 Then
        Me.lblStatus.caption = "対象法令を追加済み一覧で見つけられません"
        ReplayBodySearchHistory = "empty"
        Exit Function
    End If

    Dim keyword As String
    keyword = modHistory.SearchHistoryRowSearchText(rowIndex)
    If Len(keyword) = 0 Then
        Me.lblStatus.caption = "検索語が空です"
        ReplayBodySearchHistory = "empty"
        Exit Function
    End If

    Load frmMain
    frmMain.SelectAddedLawStoreRowForForm storeRow
    frmMain.txtBodySearch.text = keyword
    If Len(modHistory.SearchHistoryRowSearchMode(rowIndex)) > 0 Then
        frmMain.cmbSearchMode.value = modHistory.SearchHistoryRowSearchMode(rowIndex)
    End If

    Dim resultText As String
    resultText = frmMain.PerformBodySearchForForm()
    If Left$(resultText, 3) <> "ok:" Then
        Me.lblStatus.caption = "本文検索の再実行に失敗しました"
        ReplayBodySearchHistory = resultText
        Exit Function
    End If

    Me.lblStatus.caption = "本文検索を再実行しました"
    ReplayBodySearchHistory = resultText
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.ReplayBodySearchHistory", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "本文検索履歴の再実行失敗: " & Err.description
    ReplayBodySearchHistory = "error:" & Err.description
End Function

Public Function DeleteSelectedHistoryEntry() As String
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedHistoryStoreRow()
    If rowIndex <= 0 Then
        Me.lblStatus.caption = "履歴を選択してください"
        DeleteSelectedHistoryEntry = "empty"
        Exit Function
    End If

    If XlflowUI.MsgBox("history-delete-confirm", "選択した履歴を削除しますか？", vbYesNo + vbQuestion, "法令検索ビューアー", vbNo) <> vbYes Then
        Me.lblStatus.caption = "削除をキャンセルしました"
        DeleteSelectedHistoryEntry = "cancel"
        Exit Function
    End If

    modHistory.DeleteSearchHistoryEntry rowIndex
    RefreshHistoryForForm
    Me.lblStatus.caption = "履歴を削除しました"
    DeleteSelectedHistoryEntry = "ok"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.DeleteSelectedHistoryEntry", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "履歴削除失敗: " & Err.description
    DeleteSelectedHistoryEntry = "error:" & Err.description
End Function

Public Function DeleteFilteredHistoryEntries() As String
    On Error GoTo ErrHandler

    Dim historyKind As String
    historyKind = HistoryFilterForForm()
    If Len(historyKind) = 0 Then
        Me.lblStatus.caption = "削除対象の種別を選んでください"
        DeleteFilteredHistoryEntries = "empty"
        Exit Function
    End If

    If XlflowUI.MsgBox("history-delete-kind-confirm", "現在の種別の履歴をすべて削除しますか？", vbYesNo + vbQuestion, "法令検索ビューアー", vbNo) <> vbYes Then
        Me.lblStatus.caption = "種別削除をキャンセルしました"
        DeleteFilteredHistoryEntries = "cancel"
        Exit Function
    End If

    Dim deletedCount As Long
    deletedCount = modHistory.DeleteSearchHistoryByKind(historyKind)
    RefreshHistoryForForm
    Me.lblStatus.caption = "種別履歴を " & CStr(deletedCount) & " 件削除しました"
    DeleteFilteredHistoryEntries = "ok:" & CStr(deletedCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.DeleteFilteredHistoryEntries", Err.description, "", "", historyKind
    Me.lblStatus.caption = "種別削除失敗: " & Err.description
    DeleteFilteredHistoryEntries = "error:" & Err.description
End Function

Public Function DeleteAllHistoryEntries() As String
    On Error GoTo ErrHandler

    If XlflowUI.MsgBox("history-delete-all-confirm", "履歴をすべて削除しますか？", vbYesNo + vbQuestion, "法令検索ビューアー", vbNo) <> vbYes Then
        Me.lblStatus.caption = "全削除をキャンセルしました"
        DeleteAllHistoryEntries = "cancel"
        Exit Function
    End If

    modHistory.ClearSearchHistory
    RefreshHistoryForForm
    Me.lblStatus.caption = "履歴をすべて削除しました"
    DeleteAllHistoryEntries = "ok"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.DeleteAllHistoryEntries", Err.description, "", "", ""
    Me.lblStatus.caption = "全削除失敗: " & Err.description
    DeleteAllHistoryEntries = "error:" & Err.description
End Function

Public Function HistoryFormSmoke() As String
    On Error GoTo ErrHandler

    Dim originalCount As Long
    originalCount = modHistory.SearchHistoryCount()

    Dim searchRow As Long
    searchRow = modHistory.RecordLawSearchHistory("法令検索煙突", 12)
    Dim selectionRow As Long
    selectionRow = modHistory.RecordLawSelectionHistory("民法", "TESTLAW01", "テスト法", "2025-04-01", 7)

    modLawNavigator.PrepareBodyNavigatorSmokeData
    Dim smokeAddedRow As Long
    smokeAddedRow = modAddedLawStore.FindAddedLawRow("TEST001", "2025-04-01")
    If smokeAddedRow <= 0 Then
        Err.Raise vbObjectError + 7500, "frmHistory.HistoryFormSmoke", "Smoke law row was not found."
    End If

    Dim bodyHistoryRow As Long
    bodyHistoryRow = modHistory.RecordBodySearchHistory(smokeAddedRow, "受給権", "AND", 1)

    Load frmHistory
    If Me.cmdReAdd.caption <> "開く" Then
        Err.Raise vbObjectError + 7505, "frmHistory.HistoryFormSmoke", "History open caption was not updated."
    End If
    Me.cmbKindFilter.value = "すべて"
    RefreshHistoryForForm
    If Me.lstHistory.ListCount < 3 Then
        Err.Raise vbObjectError + 7501, "frmHistory.HistoryFormSmoke", "History list was not populated."
    End If

    Me.cmbKindFilter.value = "法令検索"
    RefreshHistoryForForm
    If Me.lstHistory.ListCount = 0 Then
        Err.Raise vbObjectError + 7502, "frmHistory.HistoryFormSmoke", "Filtered history list was not populated."
    End If

    Me.cmbKindFilter.value = "すべて"
    RefreshHistoryForForm
    If Left$(ReplayLawSearchHistory(searchRow), 3) = "error:" Then
        Err.Raise vbObjectError + 7503, "frmHistory.HistoryFormSmoke", "Law search replay failed."
    End If

    If Left$(ReplayBodySearchHistory(bodyHistoryRow), 3) = "error:" Then
        Err.Raise vbObjectError + 7504, "frmHistory.HistoryFormSmoke", "Body search replay failed."
    End If

    HistoryFormSmoke = "ok:" & CStr(originalCount)
    Unload frmHistory
    Unload frmLawAdd
    Unload frmMain
    modHistory.DeleteSearchHistoryByKind ("LawSearch")
    modHistory.DeleteSearchHistoryByKind ("LawSelection")
    modHistory.DeleteSearchHistoryByKind ("BodySearch")
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.HistoryFormSmoke", Err.description, "", "", ""
    HistoryFormSmoke = "error:" & CStr(Err.Number) & ":" & Err.description
    Unload frmHistory
    Unload frmLawAdd
    Unload frmMain
End Function

Private Sub ConfigureFormWindow()
    Me.caption = "履歴"
    Me.Width = FORM_WIDTH_POINTS
    Me.Height = FORM_HEIGHT_POINTS
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub ConfigureControls()
    Me.cmbKindFilter.Clear
    Me.cmbKindFilter.AddItem "すべて"
    Me.cmbKindFilter.AddItem "法令検索"
    Me.cmbKindFilter.AddItem "法令選択"
    Me.cmbKindFilter.AddItem "条文検索"
    Me.cmbKindFilter.value = "すべて"

    modHistory.ConfigureSearchHistoryListBox Me.lstHistory
    ConfigureDetailTextBox Me.txtDetail
End Sub

Private Sub ConfigureDetailTextBox(ByVal targetTextBox As Object)
    targetTextBox.MultiLine = True
    targetTextBox.WordWrap = False
    targetTextBox.ScrollBars = 3
    targetTextBox.Locked = True
    modUiState.ApplyDefaultControlFont targetTextBox
End Sub

Private Sub UpdateSelectedHistoryDetail()
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedHistoryStoreRow()
    If rowIndex <= 0 Then
        ClearDetail
        Exit Sub
    End If

    Dim detailText As String
    detailText = "種別: " & modHistory.HistoryKindDisplayText(modHistory.SearchHistoryRowKind(rowIndex)) & vbCrLf & _
        "検索語: " & modHistory.SearchHistoryRowSearchText(rowIndex) & vbCrLf & _
        "法令: " & modHistory.SearchHistoryRowLawTitle(rowIndex) & " / " & modHistory.SearchHistoryRowLawId(rowIndex) & vbCrLf & _
        "施行日: " & modHistory.SearchHistoryRowEnforcementDate(rowIndex) & vbCrLf & _
        "検索モード: " & modHistory.SearchHistoryRowSearchMode(rowIndex) & vbCrLf & _
        "件数: " & CStr(modHistory.SearchHistoryRowHitCount(rowIndex))

    Me.txtDetail.text = detailText
    Me.lblStatus.caption = modHistory.SearchHistoryDisplayText(rowIndex)
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.UpdateSelectedHistoryDetail", Err.description, "", "", ""
    Me.txtDetail.text = "選択内容の読み込み失敗: " & Err.description
End Sub

Private Sub ClearDetail()
    Me.txtDetail.text = ""
End Sub

Private Function SelectedHistoryStoreRow() As Long
    On Error GoTo ErrHandler

    If Me.lstHistory.listIndex < 0 Then
        Exit Function
    End If

    SelectedHistoryStoreRow = CLng(Val(Me.lstHistory.List(Me.lstHistory.listIndex, 7)))
    Exit Function

ErrHandler:
    SelectedHistoryStoreRow = 0
End Function

Private Function HistoryFilterForForm() As String
    Select Case CStr(Me.cmbKindFilter.value)
        Case "法令検索"
            HistoryFilterForForm = "LawSearch"
        Case "法令選択"
            HistoryFilterForForm = "LawSelection"
        Case "条文検索"
            HistoryFilterForForm = "BodySearch"
        Case Else
            HistoryFilterForForm = ""
    End Select
End Function

Private Sub UnloadFormIfLoaded(ByVal formName As String)
    On Error GoTo ErrHandler

    If Not IsUserFormLoaded(formName) Then Exit Sub

    Select Case formName
        Case "frmLawAdd"
            Unload frmLawAdd
        Case "frmMain"
            Unload frmMain
        Case "frmHistory"
            Unload Me
    End Select
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "HistoryForm", "frmHistory.UnloadFormIfLoaded", Err.description, "", "", formName
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





