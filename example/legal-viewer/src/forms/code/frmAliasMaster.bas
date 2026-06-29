Option Explicit

Private Const FORM_WIDTH_POINTS As Double = 680
Private Const FORM_HEIGHT_POINTS As Double = 470

Private Sub UserForm_Initialize()
    ConfigureFormWindow
    ConfigureControls
    modUiState.ApplyDefaultUserFormFont Me
    RefreshAliasesForForm
End Sub

Private Sub UserForm_Activate()
    modUiState.ApplyDefaultUserFormFont Me
    modWindowPlacement.CenterUserFormOnExcelMonitor Me
End Sub

Private Sub cmdNew_Click()
    ClearEditor
    Me.lblStatus.caption = "新しい別名を入力してください"
End Sub

Private Sub cmdSave_Click()
    SaveAliasForForm
End Sub

Private Sub cmdDelete_Click()
    DeleteSelectedAliasForForm
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstAliases_Click()
    UpdateEditorFromSelection
End Sub

Public Function RefreshAliasesForForm() As String
    On Error GoTo ErrHandler

    Dim selectedAlias As String
    selectedAlias = Me.txtAlias.text

    Dim itemCount As Long
    itemCount = modAliasMaster.PopulateAliasListBox(Me.lstAliases)
    If Len(selectedAlias) > 0 Then
        SelectAliasByText selectedAlias
    End If
    If Me.lstAliases.listIndex < 0 And itemCount > 0 Then
        Me.lstAliases.listIndex = 0
    End If

    If Me.lstAliases.listIndex >= 0 Then
        UpdateEditorFromSelection
    Else
        ClearEditor
    End If

    Me.lblStatus.caption = "別名 " & CStr(itemCount) & " 件"
    RefreshAliasesForForm = "ok:" & CStr(itemCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AliasMasterForm", "frmAliasMaster.RefreshAliasesForForm", Err.description, "", "", ""
    Me.lblStatus.caption = "別名一覧の読み込み失敗: " & Err.description
    RefreshAliasesForForm = "error:" & Err.description
End Function

Public Function SaveAliasForForm() As String
    On Error GoTo ErrHandler

    Dim selectedRow As Long
    selectedRow = SelectedAliasStoreRow()

    Dim rowIndex As Long
    rowIndex = modAliasMaster.SaveAlias(Me.txtAlias.text, Me.txtLawTitle.text, Me.txtNote.text, selectedRow)
    Dim savedAlias As String
    savedAlias = Me.txtAlias.text

    RefreshAliasesForForm
    SelectAliasByText savedAlias
    UpdateEditorFromSelection
    If selectedRow > 0 Then
        Me.lblStatus.caption = "選択中の別名を上書きしました"
    Else
        Me.lblStatus.caption = "別名を追加しました"
    End If
    SaveAliasForForm = "ok:" & CStr(rowIndex)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AliasMasterForm", "frmAliasMaster.SaveAliasForForm", Err.description, "", "", Me.txtAlias.text
    Me.lblStatus.caption = "別名保存失敗: " & Err.description
    SaveAliasForForm = "error:" & Err.description
End Function

Public Function DeleteSelectedAliasForForm(Optional ByVal skipConfirm As Boolean = False) As String
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedAliasStoreRow()
    If rowIndex <= 0 Then
        Me.lblStatus.caption = "削除する別名を選択してください"
        DeleteSelectedAliasForForm = "empty"
        Exit Function
    End If

    If Not skipConfirm Then
        If XlflowUI.MsgBox( _
            "alias-master-delete-confirm", _
            "選択中の別名を削除しますか？", _
            vbYesNo + vbQuestion, _
            "別名マスター", _
            vbNo) <> vbYes Then

            Me.lblStatus.caption = "削除をキャンセルしました"
            DeleteSelectedAliasForForm = "cancel"
            Exit Function
        End If
    End If

    modAliasMaster.DeleteAliasRow rowIndex
    RefreshAliasesForForm
    Me.lblStatus.caption = "別名を削除しました"
    DeleteSelectedAliasForForm = "ok"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AliasMasterForm", "frmAliasMaster.DeleteSelectedAliasForForm", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "別名削除失敗: " & Err.description
    DeleteSelectedAliasForForm = "error:" & Err.description
End Function

Public Function AliasMasterFormSmoke() As String
    On Error GoTo ErrHandler

    Load frmAliasMaster
    If Me.cmdNew.caption <> "新規" Or Me.cmdSave.caption <> "保存" Or Me.cmdDelete.caption <> "削除" Then
        Err.Raise vbObjectError + 7971, "frmAliasMaster.AliasMasterFormSmoke", "Alias master action buttons were not initialized."
    End If
    If Me.lstAliases.ListCount <= 0 Then
        Err.Raise vbObjectError + 7972, "frmAliasMaster.AliasMasterFormSmoke", "Alias master list was not populated."
    End If
    If Len(Me.txtAlias.text) = 0 Or Len(Me.txtLawTitle.text) = 0 Then
        Err.Raise vbObjectError + 7973, "frmAliasMaster.AliasMasterFormSmoke", "Alias master selection was not loaded into the editor."
    End If
    If InStr(1, Me.lblUpdatedValue.caption, ".", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7974, "frmAliasMaster.AliasMasterFormSmoke", "Alias timestamp was not formatted."
    End If
    If Not modUiState.UserFormFontMatches(Me) Then
        Err.Raise vbObjectError + 7975, "frmAliasMaster.AliasMasterFormSmoke", "Alias master font was not normalized."
    End If

    AliasMasterFormSmoke = "ok:" & CStr(Me.lstAliases.ListCount)
    Unload frmAliasMaster
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AliasMasterForm", "frmAliasMaster.AliasMasterFormSmoke", Err.description, "", "", ""
    AliasMasterFormSmoke = "error:" & Err.description
    Unload frmAliasMaster
End Function

Private Sub ConfigureFormWindow()
    Me.caption = "別名マスター"
    Me.Width = FORM_WIDTH_POINTS
    Me.Height = FORM_HEIGHT_POINTS
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub ConfigureControls()
    modAliasMaster.ConfigureAliasListBox Me.lstAliases
    Me.txtNote.MultiLine = True
    Me.txtNote.WordWrap = True
    Me.txtNote.ScrollBars = 2
End Sub

Private Sub UpdateEditorFromSelection()
    Dim rowIndex As Long
    rowIndex = SelectedAliasStoreRow()
    If rowIndex <= 0 Then
        ClearEditor
        Exit Sub
    End If

    Me.txtAlias.text = modAliasMaster.AliasRowAlias(rowIndex)
    Me.txtLawTitle.text = modAliasMaster.AliasRowLawTitle(rowIndex)
    Me.txtNote.text = modAliasMaster.AliasRowNote(rowIndex)
    Me.lblUpdatedValue.caption = modAliasMaster.AliasRowUpdatedAt(rowIndex)
    Me.lblStatus.caption = Me.txtAlias.text
End Sub

Private Sub ClearEditor()
    Me.lstAliases.listIndex = -1
    Me.txtAlias.text = ""
    Me.txtLawTitle.text = ""
    Me.txtNote.text = ""
    Me.lblUpdatedValue.caption = ""
End Sub

Private Function SelectedAliasStoreRow() As Long
    On Error GoTo ErrHandler

    If Me.lstAliases.listIndex < 0 Then Exit Function
    SelectedAliasStoreRow = CLng(Val(Me.lstAliases.List(Me.lstAliases.listIndex, 4)))
    Exit Function

ErrHandler:
    SelectedAliasStoreRow = 0
End Function

Private Sub SelectAliasByText(ByVal aliasText As String)
    Dim listIndex As Long
    For listIndex = 0 To Me.lstAliases.ListCount - 1
        If StrComp(CStr(Me.lstAliases.List(listIndex, 0)), Trim$(aliasText), vbTextCompare) = 0 Then
            Me.lstAliases.listIndex = listIndex
            Exit Sub
        End If
    Next listIndex
End Sub



