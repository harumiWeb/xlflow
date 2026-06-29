Option Explicit

Private Const MAX_FORM_RESULTS As Long = 100
Private Const FORM_WIDTH_POINTS As Double = 610
Private Const FORM_HEIGHT_POINTS As Double = 410

Private Sub UserForm_Initialize()
    ConfigureFormWindow
    modLawSearch.ConfigureLawResultListBox Me.lstResults
    modLawRevision.ConfigureRevisionCandidateComboBox Me.cboEnforcementDate
    Me.cmdSelect.caption = "追加"
    Me.cboEnforcementDate.Enabled = False
    Me.lblSelected.caption = ""
    Me.lblStatus.caption = "待機中"
    Me.cmdSelect.Enabled = False
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub UserForm_Activate()
    modUiState.ApplyDefaultUserFormFont Me
    modWindowPlacement.CenterUserFormOnExcelMonitor Me
End Sub

Private Sub ConfigureFormWindow()
    Me.caption = "法令追加"
    Me.Width = FORM_WIDTH_POINTS
    Me.Height = FORM_HEIGHT_POINTS
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub cmdSearch_Click()
    SearchForKeyword Trim$(Me.txtKeyword.text)
End Sub

Public Function SearchForKeyword(ByVal keyword As String) As String
    On Error GoTo ErrHandler

    keyword = Trim$(keyword)
    Me.txtKeyword.text = keyword
    If Len(keyword) = 0 Then
        Me.lblStatus.caption = "検索語を入力してください"
        SearchForKeyword = "empty"
        Exit Function
    End If

    SetBusy True, "検索中"

    Dim hitCount As Long
    hitCount = modLawSearch.SearchLawsForListBox(keyword, Me.lstResults, MAX_FORM_RESULTS)
    modHistory.RecordLawSearchHistory keyword, hitCount
    Me.cboEnforcementDate.Clear
    Me.cboEnforcementDate.Enabled = False
    Me.lblSelected.caption = ""
    Me.lblStatus.caption = "検索結果 " & CStr(hitCount) & " 件"
    SearchForKeyword = "ok:" & CStr(hitCount)

Cleanup:
    SetBusy False
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawAddForm", "frmLawAdd.SearchForKeyword", Err.description, "", "", keyword
    Me.lblStatus.caption = "検索失敗: " & Err.description
    SearchForKeyword = "error:" & Err.description
    Resume Cleanup
End Function

Private Sub cmdClear_Click()
    Me.txtKeyword.text = ""
    Me.lstResults.Clear
    Me.cboEnforcementDate.Clear
    Me.cboEnforcementDate.Enabled = False
    Me.lblSelected.caption = ""
    Me.lblStatus.caption = "待機中"
    Me.cmdSelect.Enabled = False
End Sub

Private Sub cmdSelect_Click()
    AddSelectedLawForForm
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstResults_Click()
    LoadEnforcementCandidatesForSelection
End Sub

Private Sub cboEnforcementDate_Change()
    UpdateSelectedSummary
End Sub

Public Function LoadEnforcementCandidatesForSelection() As String
    On Error GoTo ErrHandler

    Dim lawId As String
    lawId = SelectedColumnText(2)
    If Len(lawId) = 0 Then
        Me.cboEnforcementDate.Clear
        Me.cboEnforcementDate.Enabled = False
        Me.cmdSelect.Enabled = False
        Me.lblSelected.caption = ""
        LoadEnforcementCandidatesForSelection = "empty"
        Exit Function
    End If

    SetBusy True, "施行日候補取得中"

    Dim candidateCount As Long
    candidateCount = modLawRevision.PopulateRevisionCandidateComboBox(lawId, Me.cboEnforcementDate)
    Me.cboEnforcementDate.Enabled = (candidateCount > 0)
    UpdateSelectedSummary
    modHistory.RecordLawSelectionHistory Me.txtKeyword.text, lawId, SelectedLawTitle(), SelectedEnforcementDate(), candidateCount
    Me.lblStatus.caption = "施行日候補 " & CStr(candidateCount) & " 件"
    LoadEnforcementCandidatesForSelection = "ok:" & CStr(candidateCount)

Cleanup:
    SetBusy False
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawAddForm", "frmLawAdd.LoadEnforcementCandidatesForSelection", Err.description, lawId, "", ""
    Me.cboEnforcementDate.Clear
    Me.cboEnforcementDate.Enabled = False
    Me.lblStatus.caption = "施行日候補取得失敗: " & Err.description
    LoadEnforcementCandidatesForSelection = "error:" & Err.description
    Resume Cleanup
End Function

Public Function SelectResultForForm(ByVal resultIndex As Long) As String
    If resultIndex < 0 Or resultIndex >= Me.lstResults.ListCount Then
        SelectResultForForm = "empty"
        Exit Function
    End If

    Me.lstResults.listIndex = resultIndex
    SelectResultForForm = LoadEnforcementCandidatesForSelection()
End Function

Public Function SelectResultForLawKey(ByVal lawId As String, Optional ByVal enforcementDate As String = "") As String
    On Error GoTo ErrHandler

    lawId = Trim$(lawId)
    enforcementDate = Trim$(enforcementDate)
    If Len(lawId) = 0 Then
        SelectResultForLawKey = "empty"
        Exit Function
    End If

    Dim resultIndex As Long
    For resultIndex = 0 To Me.lstResults.ListCount - 1
        If StrComp(CStr(Me.lstResults.List(resultIndex, 2)), lawId, vbTextCompare) = 0 Then
            Me.lstResults.listIndex = resultIndex
            SelectResultForLawKey = LoadEnforcementCandidatesForSelection()
            If Len(enforcementDate) > 0 Then
                SelectEnforcementDateForForm enforcementDate
            End If
            Exit Function
        End If
    Next resultIndex

    SelectResultForLawKey = "empty"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawAddForm", "frmLawAdd.SelectResultForLawKey", Err.description, lawId, enforcementDate, ""
    SelectResultForLawKey = "error:" & Err.description
End Function

Public Function SelectEnforcementDateForForm(ByVal enforcementDate As String) As String
    On Error GoTo ErrHandler

    enforcementDate = Trim$(enforcementDate)
    If Len(enforcementDate) = 0 Then
        SelectEnforcementDateForForm = "empty"
        Exit Function
    End If
    If Me.cboEnforcementDate.ListCount = 0 Then
        SelectEnforcementDateForForm = "empty"
        Exit Function
    End If

    Dim candidateIndex As Long
    For candidateIndex = 0 To Me.cboEnforcementDate.ListCount - 1
        If StrComp(CStr(Me.cboEnforcementDate.List(candidateIndex, 2)), enforcementDate, vbTextCompare) = 0 Then
            Me.cboEnforcementDate.listIndex = candidateIndex
            UpdateSelectedSummary
            SelectEnforcementDateForForm = "ok:" & enforcementDate
            Exit Function
        End If
    Next candidateIndex

    SelectEnforcementDateForForm = "empty"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawAddForm", "frmLawAdd.SelectEnforcementDateForForm", Err.description, "", enforcementDate, ""
    SelectEnforcementDateForForm = "error:" & Err.description
End Function

Public Function AddSelectedLawForForm() As String
    On Error GoTo ErrHandler

    Dim lawId As String
    lawId = SelectedLawId()
    If Len(lawId) = 0 Then
        AddSelectedLawForForm = "empty"
        Exit Function
    End If

    If Me.cboEnforcementDate.ListCount = 0 Then
        LoadEnforcementCandidatesForSelection
    End If

    Dim lawRevisionId As String
    lawRevisionId = SelectedLawRevisionId()

    Dim enforcementDate As String
    enforcementDate = SelectedEnforcementDate()

    Dim statusText As String
    statusText = SelectedEnforcementStatus()

    If Len(lawRevisionId) = 0 Or Len(enforcementDate) = 0 Or Len(statusText) = 0 Then
        Me.lblStatus.caption = "施行日を選択してください"
        AddSelectedLawForForm = "missing-enforcement-date"
        Exit Function
    End If

    SetBusy True, "本文取得・解析中"

    Dim unitCount As Long
    unitCount = modLawParser.RefreshLawBodyUnitCache(lawRevisionId, lawId, enforcementDate)
    If unitCount <= 0 Then
        Err.Raise vbObjectError + 6801, "frmLawAdd.AddSelectedLawForForm", "本文を解析できませんでした。"
    End If

    Dim addedRow As Long
    addedRow = modAddedLawStore.RegisterParsedLawMetadata( _
        lawId, _
        enforcementDate, _
        lawRevisionId, _
        SelectedLawTitle(), _
        SelectedLawNum(), _
        "", _
        SelectedPromulgationDate(), _
        statusText, _
        modLawParser.BodyUnitCacheKey(lawRevisionId), _
        Format$(Now, "yyyy-mm-dd hh:nn:ss"))

    Me.lblStatus.caption = "追加完了: 本文単位 " & CStr(unitCount) & " 件"
    AddSelectedLawForForm = "ok:" & CStr(addedRow) & ":" & CStr(unitCount)

Cleanup:
    SetBusy False
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawAddForm", "frmLawAdd.AddSelectedLawForForm", Err.description, lawId, enforcementDate, lawRevisionId
    modUiDialog.NotifyBodyFetchFailed lawId, enforcementDate, Err.description
    Me.lblStatus.caption = "追加失敗: " & Err.description
    AddSelectedLawForForm = "error:" & Err.description
    Resume Cleanup
End Function

Public Function SelectedLawId() As String
    SelectedLawId = SelectedColumnText(2)
End Function

Public Function SelectedLawRevisionId() As String
    If Me.cboEnforcementDate.listIndex >= 0 Then
        SelectedLawRevisionId = CStr(Me.cboEnforcementDate.List(Me.cboEnforcementDate.listIndex, 1))
    Else
        SelectedLawRevisionId = SelectedColumnText(4)
    End If
End Function

Public Function SelectedLawTitle() As String
    SelectedLawTitle = SelectedColumnText(0)
End Function

Public Function SelectedLawNum() As String
    SelectedLawNum = SelectedColumnText(1)
End Function

Public Function SelectedPromulgationDate() As String
    SelectedPromulgationDate = SelectedColumnText(3)
End Function

Public Function SelectedEnforcementDate() As String
    If Me.cboEnforcementDate.listIndex >= 0 Then
        SelectedEnforcementDate = CStr(Me.cboEnforcementDate.List(Me.cboEnforcementDate.listIndex, 2))
    Else
        SelectedEnforcementDate = SelectedColumnText(5)
    End If
End Function

Public Function SelectedEnforcementStatus() As String
    If Me.cboEnforcementDate.listIndex >= 0 Then
        SelectedEnforcementStatus = CStr(Me.cboEnforcementDate.List(Me.cboEnforcementDate.listIndex, 3))
    Else
        SelectedEnforcementStatus = ""
    End If
End Function

Private Sub UpdateSelectedSummary()
    Dim lawTitle As String
    lawTitle = SelectedLawTitle()

    If Len(lawTitle) = 0 Then
        Me.cmdSelect.Enabled = False
        Me.lblSelected.caption = ""
        Exit Sub
    End If

    Me.cmdSelect.Enabled = True
    If Len(SelectedEnforcementDate()) > 0 Then
        Me.lblSelected.caption = lawTitle & " / " & SelectedLawId() & " / " & SelectedEnforcementDate() & " " & SelectedEnforcementStatus()
    Else
        Me.lblSelected.caption = lawTitle & " / " & SelectedLawId()
    End If
    Me.lblStatus.caption = "選択中"
End Sub

Private Function SelectedColumnText(ByVal columnIndex As Long) As String
    If Me.lstResults.listIndex < 0 Then
        SelectedColumnText = ""
    Else
        SelectedColumnText = CStr(Me.lstResults.List(Me.lstResults.listIndex, columnIndex))
    End If
End Function

Private Sub SetBusy(ByVal isBusy As Boolean, Optional ByVal statusText As String = "")
    Me.cmdSearch.Enabled = Not isBusy
    Me.cmdClear.Enabled = Not isBusy
    Me.cmdClose.Enabled = Not isBusy
    Me.cboEnforcementDate.Enabled = (Not isBusy) And (Me.cboEnforcementDate.ListCount > 0)
    Me.cmdSelect.Enabled = (Not isBusy) And (Me.lstResults.listIndex >= 0)

    If Len(statusText) > 0 Then
        Me.lblStatus.caption = statusText
    End If

    Me.Repaint
End Sub







