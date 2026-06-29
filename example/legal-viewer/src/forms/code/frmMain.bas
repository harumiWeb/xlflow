Option Explicit

Private Const FORM_WIDTH_POINTS As Double = 860
Private Const FORM_HEIGHT_POINTS As Double = 610

Private citationTextProgrammaticChange As Boolean
Private citationTextDirty As Boolean

Private Sub UserForm_Initialize()
    ConfigureFormWindow
    modAddedLawStore.ConfigureAddedLawListBox Me.lstAddedLaws
    ConfigurePlaceholderLists
    modUiState.ApplyDefaultUserFormFont Me
    RefreshAddedLawsForForm
End Sub

Private Sub UserForm_Activate()
    modUiState.ApplyDefaultUserFormFont Me
    modWindowPlacement.CenterUserFormOnExcelMonitor Me
End Sub

Private Sub ConfigureFormWindow()
    Me.caption = "法令検索ビューアー"
    Me.Width = FORM_WIDTH_POINTS
    Me.Height = FORM_HEIGHT_POINTS
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub cmdAddLaw_Click()
    On Error GoTo ErrHandler

    Me.lblStatus.caption = "法令追加画面を開いています"
    XlflowUI.ShowForm frmLawAdd, True
    RefreshAddedLawsForForm
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdAddLaw_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "法令追加画面を開けません: " & Err.description
End Sub

Private Sub cmdHistory_Click()
    On Error GoTo ErrHandler

    Me.lblStatus.caption = "履歴画面を開いています"
    XlflowUI.ShowForm frmHistory, True
    RefreshAddedLawsForForm
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdHistory_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "履歴画面を開けません: " & Err.description
End Sub

Private Sub cmdBookmarks_Click()
    On Error GoTo ErrHandler

    Me.lblStatus.caption = "ブックマーク画面を開いています"
    XlflowUI.ShowForm frmBookmarks, True
    RefreshAddedLawsForForm
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdBookmarks_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "ブックマーク画面を開けません: " & Err.description
End Sub

Private Sub cmdBookmarkAdd_Click()
    On Error GoTo ErrHandler

    AddBookmarkForForm
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdBookmarkAdd_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "ブックマークを登録できません: " & Err.description
End Sub

Private Sub cmdAliasMaster_Click()
    On Error GoTo ErrHandler

    Me.lblStatus.caption = "別名マスターを開いています"
    XlflowUI.ShowForm frmAliasMaster, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdAliasMaster_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "別名マスターを開けません: " & Err.description
End Sub

Private Sub cmdCompare_Click()
    On Error GoTo ErrHandler

    Me.lblStatus.caption = "比較画面を開いています"
    XlflowUI.ShowForm frmCompare, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdCompare_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "比較画面を開けません: " & Err.description
End Sub

Private Sub cmdSettings_Click()
    On Error GoTo ErrHandler

    Me.lblStatus.caption = "設定画面を開いています"
    XlflowUI.ShowForm frmSettings, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdSettings_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "設定画面を開けません: " & Err.description
End Sub

Private Sub cmdApiLogs_Click()
    On Error GoTo ErrHandler

    Me.lblStatus.caption = "APIログ画面を開いています"
    Load frmLogs
    frmLogs.SetLogKindForForm "API"
    XlflowUI.ShowForm frmLogs, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdApiLogs_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "APIログ画面を開けません: " & Err.description
End Sub

Private Sub cmdErrorLogs_Click()
    On Error GoTo ErrHandler

    Me.lblStatus.caption = "エラーログ画面を開いています"
    Load frmLogs
    frmLogs.SetLogKindForForm "ERROR"
    XlflowUI.ShowForm frmLogs, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.cmdErrorLogs_Click", Err.description, "", "", ""
    Me.lblStatus.caption = "エラーログ画面を開けません: " & Err.description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstAddedLaws_Click()
    SelectAddedLawForForm Me.lstAddedLaws.listIndex
End Sub

Private Sub lstBodyNav_Click()
    Me.rdoCitationBody.value = True
    UpdateBodyUnitPreview
    ClearCitationIfNotEdited
End Sub

Private Sub cmdSearch_Click()
    PerformBodySearchForForm
End Sub

Private Sub lstResults_Click()
    Me.rdoCitationSearch.value = True
    UpdateSearchResultPreview
    ClearCitationIfNotEdited
End Sub

Private Sub rdoTableMarkdown_Click()
    RefreshTableDisplayMode
End Sub

Private Sub rdoTableAscii_Click()
    RefreshTableDisplayMode
End Sub

Private Sub cmdGenerateCitation_Click()
    GenerateCitationForForm
End Sub

Private Sub cmdCopyCitation_Click()
    CopyCitationForForm
End Sub

Private Sub rdoCitationBody_Click()
    ClearCitationIfNotEdited
End Sub

Private Sub rdoCitationSearch_Click()
    ClearCitationIfNotEdited
End Sub

Private Sub txtCitation_Change()
    If citationTextProgrammaticChange Then Exit Sub
    citationTextDirty = True
End Sub

Public Function RefreshAddedLawsForForm() As String
    On Error GoTo ErrHandler

    SetBusy True, "追加済み法令を読み込み中"

    Dim itemCount As Long
    itemCount = modAddedLawStore.PopulateAddedLawListBox(Me.lstAddedLaws)
    SelectStoredOrFirstAddedLaw
    UpdateSelectedSummary

    Me.lblStatus.caption = "追加済み法令 " & CStr(itemCount) & " 件"
    RefreshAddedLawsForForm = "ok:" & CStr(itemCount)

Cleanup:
    SetBusy False
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.RefreshAddedLawsForForm", Err.description, "", "", ""
    Me.lblStatus.caption = "追加済み法令の読み込み失敗: " & Err.description
    RefreshAddedLawsForForm = "error:" & Err.description
    Resume Cleanup
End Function

Public Function SelectAddedLawForForm(ByVal listIndex As Long) As String
    On Error GoTo ErrHandler

    If listIndex < 0 Or listIndex >= Me.lstAddedLaws.ListCount Then
        Me.lstAddedLaws.listIndex = -1
        UpdateSelectedSummary
        SelectAddedLawForForm = "empty"
        Exit Function
    End If

    Me.lstAddedLaws.listIndex = listIndex

    Dim storeRow As Long
    storeRow = SelectedAddedLawStoreRow()
    If storeRow > 0 Then
        modAddedLawStore.SelectAddedLawRow storeRow
    End If

    UpdateSelectedSummary
    SelectAddedLawForForm = "ok:" & CStr(storeRow)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.SelectAddedLawForForm", Err.description, "", "", CStr(listIndex)
    Me.lblStatus.caption = "法令選択失敗: " & Err.description
    SelectAddedLawForForm = "error:" & Err.description
End Function

Public Function SelectAddedLawStoreRowForForm(ByVal storeRow As Long) As String
    On Error GoTo ErrHandler

    If storeRow <= 0 Then
        Me.lstAddedLaws.listIndex = -1
        UpdateSelectedSummary
        SelectAddedLawStoreRowForForm = "empty"
        Exit Function
    End If

    Dim listIndex As Long
    For listIndex = 0 To Me.lstAddedLaws.ListCount - 1
        If CLng(Val(Me.lstAddedLaws.List(listIndex, 4))) = storeRow Then
            Me.lstAddedLaws.listIndex = listIndex
            modAddedLawStore.SelectAddedLawRow storeRow
            UpdateSelectedSummary
            SelectAddedLawStoreRowForForm = "ok:" & CStr(storeRow)
            Exit Function
        End If
    Next listIndex

    SelectAddedLawStoreRowForForm = "empty"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.SelectAddedLawStoreRowForForm", Err.description, "", "", CStr(storeRow)
    Me.lblStatus.caption = "法令選択失敗: " & Err.description
    SelectAddedLawStoreRowForForm = "error:" & Err.description
End Function

Public Function SelectBodyUnitStoreRowForForm(ByVal bodyUnitRow As Long) As String
    On Error GoTo ErrHandler

    If bodyUnitRow <= 0 Then
        Me.lstBodyNav.listIndex = -1
        UpdateBodyUnitPreview
        SelectBodyUnitStoreRowForForm = "empty"
        Exit Function
    End If

    Dim listIndex As Long
    For listIndex = 0 To Me.lstBodyNav.ListCount - 1
        If CLng(Val(Me.lstBodyNav.List(listIndex, 3))) = bodyUnitRow Then
            Me.lstBodyNav.listIndex = listIndex
            UpdateBodyUnitPreview
            SelectBodyUnitStoreRowForForm = "ok:" & CStr(bodyUnitRow)
            Exit Function
        End If
    Next listIndex

    SelectBodyUnitStoreRowForForm = "empty"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.SelectBodyUnitStoreRowForForm", Err.description, "", "", CStr(bodyUnitRow)
    Me.lblStatus.caption = "本文選択失敗: " & Err.description
    SelectBodyUnitStoreRowForForm = "error:" & Err.description
End Function

Public Function SelectedCitationBodyUnitStoreRowForForm() As Long
    SelectedCitationBodyUnitStoreRowForForm = SelectedCitationBodyUnitStoreRow()
End Function

Public Function AddBookmarkForForm() As String
    On Error GoTo ErrHandler

    Dim bodyUnitRow As Long
    bodyUnitRow = SelectedCitationBodyUnitStoreRowForForm()
    If bodyUnitRow <= 0 Then
        Me.lblStatus.caption = "ブックマーク対象の本文を選択してください"
        AddBookmarkForForm = "empty"
        Exit Function
    End If

    Dim bookmarkRow As Long
    bookmarkRow = modBookmark.RegisterBookmarkFromBodyUnitRow(bodyUnitRow)
    If bookmarkRow <= 0 Then
        Me.lblStatus.caption = "ブックマークを登録できませんでした"
        AddBookmarkForForm = "empty"
        Exit Function
    End If

    Me.lblStatus.caption = "ブックマークを登録しました"
    AddBookmarkForForm = "ok:" & CStr(bookmarkRow)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.AddBookmarkForForm", Err.description, "", "", CStr(bodyUnitRow)
    Me.lblStatus.caption = "ブックマーク登録失敗: " & Err.description
    AddBookmarkForForm = "error:" & Err.description
End Function

Public Function BookmarkAddSmoke() As String
    On Error GoTo ErrHandler

    Dim shouldUnload As Boolean
    Dim addedRow As Long
    Dim bodyUnitRow As Long
    Dim resultText As String
    Dim bookmarkRow As Long

    modLawNavigator.PrepareBodyNavigatorSmokeData

    Load frmMain
    shouldUnload = True
    frmMain.RefreshAddedLawsForForm
    addedRow = modAddedLawStore.FindAddedLawRow("TEST001", "2025-04-01")
    If addedRow <= 0 Then
        Err.Raise vbObjectError + 6930, "frmMain.BookmarkAddSmoke", "Smoke law row was not found."
    End If

    frmMain.SelectAddedLawStoreRowForForm addedRow
    If Me.lstBodyNav.ListCount <= 0 Then
        Err.Raise vbObjectError + 6934, "frmMain.BookmarkAddSmoke", "Body navigation list was not populated."
    End If

    Me.lstBodyNav.listIndex = 0
    UpdateBodyUnitPreview
    bodyUnitRow = SelectedCitationBodyUnitStoreRowForForm()
    If bodyUnitRow <= 0 Then
        Err.Raise vbObjectError + 6935, "frmMain.BookmarkAddSmoke", "Selected body unit row was not resolved."
    End If

    resultText = AddBookmarkForForm()
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 6931, "frmMain.BookmarkAddSmoke", "Bookmark add failed."
    End If

    bookmarkRow = CLng(Val(Mid$(resultText, 4)))
    If bookmarkRow <= 0 Then
        Err.Raise vbObjectError + 6932, "frmMain.BookmarkAddSmoke", "Bookmark row was not returned."
    End If

    If Len(modBookmark.BookmarkRowLawId(bookmarkRow)) = 0 Then
        Err.Raise vbObjectError + 6933, "frmMain.BookmarkAddSmoke", "Bookmark row was not populated."
    End If

    modBookmark.DeleteBookmarkEntry bookmarkRow
    Unload frmMain
    shouldUnload = False
    BookmarkAddSmoke = "ok:" & CStr(bookmarkRow)
    GoTo Cleanup

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.BookmarkAddSmoke", Err.description, CStr(addedRow), CStr(bodyUnitRow), resultText & "|" & CStr(bookmarkRow)
    BookmarkAddSmoke = "error:" & CStr(Err.Number) & ":" & Err.description

Cleanup:
    If shouldUnload Then
        Unload frmMain
    End If
    modTempCache.DeleteAllTempCaches ThisWorkbook
End Function

Private Sub ConfigurePlaceholderLists()
    modLawNavigator.ConfigureBodyNavListBox Me.lstBodyNav
    modLawTextSearch.ConfigureSearchResultsListBox Me.lstResults
    ConfigureSearchControls
    ConfigureTableDisplayControls
    ConfigurePreviewTextBox Me.txtPreview
    ConfigurePreviewTextBox Me.txtUnitPreview
    ConfigureCitationControls
End Sub

Private Sub ConfigureSearchControls()
    Me.cmbSearchMode.Clear
    Me.cmbSearchMode.AddItem "AND"
    Me.cmbSearchMode.AddItem "OR"
    Me.cmbSearchMode.value = modSettings.GetSettingText("DefaultSearchMode", "AND")
    If StrComp(Me.cmbSearchMode.value, "OR", vbTextCompare) <> 0 Then
        Me.cmbSearchMode.value = "AND"
    End If
End Sub

Private Sub ConfigureTableDisplayControls()
    Me.rdoTableMarkdown.GroupName = "table_display"
    Me.rdoTableAscii.GroupName = "table_display"
    Me.rdoTableMarkdown.value = True
    Me.rdoTableAscii.value = False
End Sub

Private Sub ConfigureCitationControls()
    modCitation.ConfigureCitationFormatComboBox Me.cmbCitationFormat, modSettings.GetSettingText("DefaultCitationFormat", modCitation.CitationFormatWithSource())
    ConfigureCitationTextBox Me.txtCitation
    Me.rdoCitationBody.GroupName = "citation_target"
    Me.rdoCitationSearch.GroupName = "citation_target"
    Me.rdoCitationBody.value = True
    Me.rdoCitationSearch.value = False
    citationTextDirty = False
End Sub

Private Sub SelectStoredOrFirstAddedLaw()
    Dim selectedRow As Long
    selectedRow = modAddedLawStore.SelectedAddedLawRow()

    Dim listIndex As Long
    If selectedRow > 0 Then
        For listIndex = 0 To Me.lstAddedLaws.ListCount - 1
            If CLng(Me.lstAddedLaws.List(listIndex, 4)) = selectedRow Then
                Me.lstAddedLaws.listIndex = listIndex
                Exit Sub
            End If
        Next listIndex
    End If

    If Me.lstAddedLaws.ListCount > 0 Then
        Me.lstAddedLaws.listIndex = 0
        modAddedLawStore.SelectAddedLawRow SelectedAddedLawStoreRow()
    Else
        Me.lstAddedLaws.listIndex = -1
    End If
End Sub

Private Sub UpdateSelectedSummary()
    Dim storeRow As Long
    storeRow = SelectedAddedLawStoreRow()

    ClearDependentViews
    If storeRow <= 0 Then
        Me.lblSelected.caption = ""
        Exit Sub
    End If

    Me.lblSelected.caption = modAddedLawStore.AddedLawDetailText(storeRow)
    RefreshBodyViews storeRow
End Sub

Private Sub RefreshBodyViews(ByVal storeRow As Long)
    On Error GoTo ErrHandler

    Dim bodyUnitCount As Long
    Me.lstResults.Clear
    bodyUnitCount = modLawNavigator.PopulateBodyNavListBox(Me.lstBodyNav, storeRow)
    Me.txtPreview.text = modLawNavigator.BodyPreviewText(storeRow, 40, TableDisplayModeForForm())

    If bodyUnitCount > 0 Then
        Me.lstBodyNav.listIndex = 0
        UpdateBodyUnitPreview
        Me.lblStatus.caption = "本文単位 " & CStr(bodyUnitCount) & " 件"
    Else
        Me.txtUnitPreview.text = ""
        Me.lblStatus.caption = "本文キャッシュなし"
    End If
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.RefreshBodyViews", Err.description, "", "", CStr(storeRow)
    Me.lblStatus.caption = "本文表示失敗: " & Err.description
End Sub

Public Function PerformBodySearchForForm() As String
    On Error GoTo ErrHandler

    Dim storeRow As Long
    storeRow = SelectedAddedLawStoreRow()
    If storeRow <= 0 Then
        Me.lstResults.Clear
        Me.lblStatus.caption = "検索対象の法令を選択してください"
        PerformBodySearchForForm = "empty"
        Exit Function
    End If

    If Len(Trim$(Me.txtBodySearch.text)) = 0 Then
        Me.lstResults.Clear
        Me.lblStatus.caption = "検索語を入力してください"
        PerformBodySearchForForm = "empty"
        Exit Function
    End If

    SetBusy True, "検索中"

    Dim hitCount As Long
    hitCount = modLawTextSearch.PopulateBodySearchResultsListBox(Me.lstResults, storeRow, Me.txtBodySearch.text, SearchModeForForm())
    modHistory.RecordBodySearchHistory storeRow, Me.txtBodySearch.text, SearchModeForForm(), hitCount

    If hitCount > 0 Then
        Me.lstResults.listIndex = 0
        UpdateSearchResultPreview
        Me.lblStatus.caption = "検索完了: " & CStr(hitCount) & " 件"
    Else
        Me.txtUnitPreview.text = ""
        Me.lblStatus.caption = "検索結果 0 件"
    End If

    PerformBodySearchForForm = "ok:" & CStr(hitCount)

Cleanup:
    SetBusy False
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.PerformBodySearchForForm", Err.description, "", "", Me.txtBodySearch.text
    Me.lblStatus.caption = "検索失敗: " & Err.description
    PerformBodySearchForForm = "error:" & Err.description
    Resume Cleanup
End Function

Private Sub UpdateBodyUnitPreview()
    On Error GoTo ErrHandler

    Dim bodyUnitRow As Long
    bodyUnitRow = SelectedBodyUnitStoreRow()
    If bodyUnitRow <= 0 Then
        Me.txtUnitPreview.text = ""
    Else
        Me.rdoCitationBody.value = True
        Me.txtUnitPreview.text = modLawNavigator.BodyUnitPreviewText(bodyUnitRow, TableDisplayModeForForm())
        ClearCitationIfNotEdited
    End If
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.UpdateBodyUnitPreview", Err.description, "", "", ""
    Me.txtUnitPreview.text = "選択単位表示失敗: " & Err.description
End Sub

Private Sub UpdateSearchResultPreview()
    On Error GoTo ErrHandler

    Dim bodyUnitRow As Long
    bodyUnitRow = SelectedSearchResultBodyUnitStoreRow()
    If bodyUnitRow <= 0 Then
        Me.txtUnitPreview.text = ""
    Else
        Me.rdoCitationSearch.value = True
        Me.txtPreview.text = modLawNavigator.BodyContextPreviewText(bodyUnitRow, TableDisplayModeForForm())
        Me.txtUnitPreview.text = modLawNavigator.BodyUnitPreviewText(bodyUnitRow, TableDisplayModeForForm())
        ClearCitationIfNotEdited
    End If
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.UpdateSearchResultPreview", Err.description, "", "", ""
    Me.txtUnitPreview.text = "検索結果表示失敗: " & Err.description
End Sub

Private Function SelectedAddedLawStoreRow() As Long
    On Error GoTo ErrHandler

    If Me.lstAddedLaws.listIndex < 0 Then
        SelectedAddedLawStoreRow = 0
    Else
        SelectedAddedLawStoreRow = CLng(Me.lstAddedLaws.List(Me.lstAddedLaws.listIndex, 4))
    End If
    Exit Function

ErrHandler:
    SelectedAddedLawStoreRow = 0
End Function

Private Function SelectedBodyUnitStoreRow() As Long
    On Error GoTo ErrHandler

    If Me.lstBodyNav.listIndex < 0 Then
        SelectedBodyUnitStoreRow = 0
    Else
        SelectedBodyUnitStoreRow = CLng(Me.lstBodyNav.List(Me.lstBodyNav.listIndex, 3))
    End If
    Exit Function

ErrHandler:
    SelectedBodyUnitStoreRow = 0
End Function

Private Function SelectedSearchResultBodyUnitStoreRow() As Long
    On Error GoTo ErrHandler

    If Me.lstResults.listIndex < 0 Then
        SelectedSearchResultBodyUnitStoreRow = 0
    Else
        SelectedSearchResultBodyUnitStoreRow = CLng(Me.lstResults.List(Me.lstResults.listIndex, 3))
    End If
    Exit Function

ErrHandler:
    SelectedSearchResultBodyUnitStoreRow = 0
End Function

Private Function SearchModeForForm() As String
    If StrComp(Me.cmbSearchMode.value, "OR", vbTextCompare) = 0 Then
        SearchModeForForm = "OR"
    Else
        SearchModeForForm = "AND"
    End If
End Function

Private Function TableDisplayModeForForm() As String
    If Me.rdoTableAscii.value Then
        TableDisplayModeForForm = modLawNavigator.TableDisplayModeAscii()
    Else
        TableDisplayModeForForm = modLawNavigator.TableDisplayModeMarkdown()
    End If
End Function

Private Sub RefreshTableDisplayMode()
    On Error GoTo ErrHandler

    Dim storeRow As Long
    storeRow = SelectedAddedLawStoreRow()
    If storeRow > 0 Then
        Me.txtPreview.text = modLawNavigator.BodyPreviewText(storeRow, 40, TableDisplayModeForForm())
    End If

    If Me.lstResults.listIndex >= 0 Then
        UpdateSearchResultPreview
    Else
        UpdateBodyUnitPreview
    End If
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.RefreshTableDisplayMode", Err.description, "", "", ""
    Me.lblStatus.caption = "表表示切替失敗: " & Err.description
End Sub

Public Function GenerateCitationForForm() As String
    On Error GoTo ErrHandler

    If Not ConfirmCitationRegeneration() Then
        Me.lblStatus.caption = "引用文の再生成をキャンセルしました"
        GenerateCitationForForm = "cancel"
        Exit Function
    End If

    Dim bodyUnitRow As Long
    bodyUnitRow = SelectedCitationBodyUnitStoreRow()
    If bodyUnitRow <= 0 Then
        If Me.rdoCitationSearch.value Then
            Me.lblStatus.caption = "引用対象の検索結果を選択してください"
        Else
            Me.lblStatus.caption = "引用対象の本文単位を選択してください"
        End If
        GenerateCitationForForm = "empty"
        Exit Function
    End If

    Dim citationText As String
    citationText = modCitation.CitationTextFromBodyUnitRow(bodyUnitRow, CitationFormatForForm(), TableDisplayModeForForm())
    If Len(Trim$(citationText)) = 0 Then
        Me.lblStatus.caption = "引用文を生成できませんでした"
        GenerateCitationForForm = "empty"
        Exit Function
    End If

    SetCitationText citationText, False
    Me.lblStatus.caption = "引用文を生成しました"
    GenerateCitationForForm = "ok:" & CStr(Len(citationText))
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.GenerateCitationForForm", Err.description, "", "", ""
    Me.lblStatus.caption = "引用文生成失敗: " & Err.description
    GenerateCitationForForm = "error:" & Err.description
End Function

Public Function CopyCitationForForm() As String
    On Error GoTo ErrHandler

    If Len(Trim$(Me.txtCitation.text)) = 0 Then
        Me.lblStatus.caption = "コピーする引用文がありません"
        CopyCitationForForm = "empty"
        Exit Function
    End If

    modClipboard.CopyTextToClipboard Me.txtCitation.text
    Me.lblStatus.caption = "引用文をコピーしました"
    CopyCitationForForm = "ok:" & CStr(Len(Me.txtCitation.text))
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "frmMain.CopyCitationForForm", Err.description, "", "", ""
    XlflowUI.MsgBox "citation-copy-failed", "引用文をコピーできませんでした。" & vbCrLf & Err.description, vbOKOnly + vbExclamation, "法令検索ビューアー", vbOK
    Me.lblStatus.caption = "引用文コピー失敗: " & Err.description
    CopyCitationForForm = "error:" & Err.description
End Function

Private Sub ClearDependentViews()
    Me.lstBodyNav.Clear
    Me.lstResults.Clear
    Me.txtPreview.text = ""
    Me.txtUnitPreview.text = ""
    SetCitationText "", False
End Sub

Private Sub ConfigurePreviewTextBox(ByVal targetTextBox As Object)
    targetTextBox.MultiLine = True
    targetTextBox.WordWrap = False
    targetTextBox.ScrollBars = 3
    targetTextBox.Locked = True
    modUiState.ApplyDefaultControlFont targetTextBox
End Sub

Private Sub ConfigureCitationTextBox(ByVal targetTextBox As Object)
    targetTextBox.MultiLine = True
    targetTextBox.WordWrap = False
    targetTextBox.ScrollBars = 3
    targetTextBox.Locked = False
    modUiState.ApplyDefaultControlFont targetTextBox
End Sub

Private Function CitationFormatForForm() As String
    CitationFormatForForm = CStr(Me.cmbCitationFormat.value)
End Function

Private Function SelectedCitationBodyUnitStoreRow() As Long
    If Me.rdoCitationSearch.value Then
        SelectedCitationBodyUnitStoreRow = SelectedSearchResultBodyUnitStoreRow()
    Else
        SelectedCitationBodyUnitStoreRow = SelectedBodyUnitStoreRow()
    End If
End Function

Private Sub SetCitationText(ByVal citationText As String, ByVal dirtyAfterSet As Boolean)
    citationTextProgrammaticChange = True
    Me.txtCitation.text = citationText
    citationTextProgrammaticChange = False
    citationTextDirty = dirtyAfterSet
End Sub

Private Sub ClearCitationIfNotEdited()
    If Not citationTextDirty Then
        SetCitationText "", False
    End If
End Sub

Private Function ConfirmCitationRegeneration() As Boolean
    If Not citationTextDirty Or Len(Trim$(Me.txtCitation.text)) = 0 Then
        ConfirmCitationRegeneration = True
        Exit Function
    End If

    ConfirmCitationRegeneration = (XlflowUI.MsgBox("citation-regenerate-discard-edits", "編集済みの引用文を再生成しますか？", vbYesNo + vbQuestion, "法令検索ビューアー", vbNo) = vbYes)
End Function

Private Sub SetBusy(ByVal isBusy As Boolean, Optional ByVal statusText As String = "")
    Me.cmdAddLaw.Enabled = Not isBusy
    Me.cmdHistory.Enabled = Not isBusy
    Me.cmdBookmarks.Enabled = Not isBusy
    Me.cmdBookmarkAdd.Enabled = Not isBusy
    Me.cmdCompare.Enabled = Not isBusy
    Me.cmdSettings.Enabled = Not isBusy
    Me.cmdAliasMaster.Enabled = Not isBusy
    Me.cmdApiLogs.Enabled = Not isBusy
    Me.cmdErrorLogs.Enabled = Not isBusy
    Me.cmdClose.Enabled = Not isBusy
    Me.cmdSearch.Enabled = Not isBusy
    Me.cmdGenerateCitation.Enabled = Not isBusy
    Me.cmdCopyCitation.Enabled = Not isBusy
    Me.rdoTableMarkdown.Enabled = Not isBusy
    Me.rdoTableAscii.Enabled = Not isBusy
    Me.rdoCitationBody.Enabled = Not isBusy
    Me.rdoCitationSearch.Enabled = Not isBusy
    Me.txtBodySearch.Enabled = Not isBusy
    Me.cmbSearchMode.Enabled = Not isBusy
    Me.cmbCitationFormat.Enabled = Not isBusy
    Me.lstAddedLaws.Enabled = Not isBusy
    Me.lstBodyNav.Enabled = Not isBusy
    Me.lstResults.Enabled = Not isBusy

    If Len(statusText) > 0 Then
        Me.lblStatus.caption = statusText
    End If
End Sub



