Option Explicit

Private Const FORM_WIDTH_POINTS As Double = 620
Private Const FORM_HEIGHT_POINTS As Double = 470

Private Sub UserForm_Initialize()
    ConfigureFormWindow
    ConfigureControls
    modUiState.ApplyDefaultUserFormFont Me
    RefreshSettingsForForm
End Sub

Private Sub UserForm_Activate()
    modUiState.ApplyDefaultUserFormFont Me
    modWindowPlacement.CenterUserFormOnExcelMonitor Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    SaveSelectedSettingForForm
End Sub

Private Sub cmdReset_Click()
    ResetSettingsToDefaultsForForm
End Sub

Private Sub lstSettings_Click()
    UpdateSelectedSettingFields
End Sub

Public Function RefreshSettingsForForm() As String
    On Error GoTo ErrHandler

    Dim selectedRow As Long
    selectedRow = SelectedSettingStoreRow()

    Dim itemCount As Long
    itemCount = modSettings.PopulateSettingsListBox(Me.lstSettings)
    If itemCount > 0 Then
        SelectSettingRowByStoreRow selectedRow
        If Me.lstSettings.listIndex < 0 Then
            Me.lstSettings.listIndex = 0
        End If
        UpdateSelectedSettingFields
    Else
        ClearSelectedSettingFields
    End If

    Me.lblStatus.caption = "設定 " & CStr(itemCount) & " 件"
    RefreshSettingsForForm = "ok:" & CStr(itemCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "SettingsForm", "frmSettings.RefreshSettingsForForm", Err.description, "", "", ""
    Me.lblStatus.caption = "設定の読み込み失敗: " & Err.description
    RefreshSettingsForForm = "error:" & Err.description
End Function

Public Function SaveSelectedSettingForForm() As String
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedSettingStoreRow()
    If rowIndex <= 0 Then
        Me.lblStatus.caption = "設定を選択してください"
        SaveSelectedSettingForForm = "empty"
        Exit Function
    End If

    modSettings.SetSettingValue modSettings.SettingRowKey(rowIndex), SelectedSettingValueText(rowIndex), Me.txtDescription.text
    RefreshSettingsForForm
    SelectSettingRowByStoreRow rowIndex
    UpdateSelectedSettingFields
    Me.lblStatus.caption = "設定を保存しました"
    SaveSelectedSettingForForm = "ok:" & CStr(rowIndex)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "SettingsForm", "frmSettings.SaveSelectedSettingForForm", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "設定保存失敗: " & Err.description
    SaveSelectedSettingForForm = "error:" & Err.description
End Function

Public Function ResetSettingsToDefaultsForForm(Optional ByVal skipConfirm As Boolean = False) As String
    On Error GoTo ErrHandler

    If Not skipConfirm Then
        If XlflowUI.MsgBox("settings-reset-confirm", "設定を初期値に戻しますか？", vbYesNo + vbQuestion, "法令検索ビューアー", vbNo) <> vbYes Then
            Me.lblStatus.caption = "初期化をキャンセルしました"
            ResetSettingsToDefaultsForForm = "cancel"
            Exit Function
        End If
    End If

    modSettings.ResetSettingsToDefaults
    RefreshSettingsForForm
    Me.lblStatus.caption = "設定を初期化しました"
    ResetSettingsToDefaultsForForm = "ok"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "SettingsForm", "frmSettings.ResetSettingsToDefaultsForForm", Err.description, "", "", ""
    Me.lblStatus.caption = "設定初期化失敗: " & Err.description
    ResetSettingsToDefaultsForForm = "error:" & Err.description
End Function

Public Function SettingsFormSmoke() As String
    On Error GoTo ErrHandler

    Dim originalValue As String
    Dim originalSearchMode As String
    Dim originalManagedVisibility As String
    Dim originalValueDescription As String
    Dim originalSearchModeDescription As String
    Dim originalManagedVisibilityDescription As String
    Dim resultText As String
    Dim selectedSearchMode As String
    Dim selectedManagedVisibility As String
    originalValue = modSettings.GetSettingText("ApiRetryCount", "1")
    originalSearchMode = modSettings.GetSettingText("DefaultSearchMode", "AND")
    originalManagedVisibility = modSettings.GetSettingText("ManagedSheetVisibility", "VeryHidden")

    Load frmSettings
    If Me.cmdSave.caption <> "設定保存" Then
        Err.Raise vbObjectError + 7801, "frmSettings.SettingsFormSmoke", "Save button caption was not initialized."
    End If
    If Me.cmdReset.caption <> "初期化" Then
        Err.Raise vbObjectError + 7802, "frmSettings.SettingsFormSmoke", "Reset button caption was not initialized."
    End If
    If Me.lblStatus.Top > 438 Or Me.lblStatus.Height < 18 Then
        Err.Raise vbObjectError + 7803, "frmSettings.SettingsFormSmoke", "Status label was not moved up enough."
    End If

    RefreshSettingsForForm
    If Me.lstSettings.ListCount <= 0 Then
        Err.Raise vbObjectError + 7804, "frmSettings.SettingsFormSmoke", "Settings list was not populated."
    End If
    If InStr(1, Me.lstSettings.List(0, 3), ".", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7804, "frmSettings.SettingsFormSmoke", "UpdatedAt column was not formatted."
    End If

    SelectSettingRowByKey "DefaultSearchMode"
    UpdateSelectedSettingFields
    If SelectedSettingStoreRow() <= 0 Then
        Err.Raise vbObjectError + 7805, "frmSettings.SettingsFormSmoke", "Selectable setting row was not selected."
    End If
    If InStr(1, Me.lblUpdatedValue.caption, ".", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7805, "frmSettings.SettingsFormSmoke", "UpdatedAt label was not formatted."
    End If

    If Not Me.cmbValue.Visible Or Me.txtValue.Visible Then
        Err.Raise vbObjectError + 7806, "frmSettings.SettingsFormSmoke", "Selectable setting did not switch to combo box."
    End If
    If Me.cmbValue.ListCount < 2 Then
        Err.Raise vbObjectError + 7807, "frmSettings.SettingsFormSmoke", "Combo box options were not populated."
    End If

    originalSearchModeDescription = modSettings.SettingRowDescription(SelectedSettingStoreRow())
    If StrComp(modSettings.GetSettingText("DefaultSearchMode", "AND"), "AND", vbTextCompare) = 0 Then
        Me.cmbValue.value = "OR"
    Else
        Me.cmbValue.value = "AND"
    End If
    selectedSearchMode = CStr(Me.cmbValue.value)
    resultText = SaveSelectedSettingForForm()
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7808, "frmSettings.SettingsFormSmoke", "Selectable setting save failed."
    End If
    If modSettings.GetSettingText("DefaultSearchMode", "") <> selectedSearchMode Then
        Err.Raise vbObjectError + 7809, "frmSettings.SettingsFormSmoke", "Selectable setting value was not saved."
    End If

    resultText = ResetSettingsToDefaultsForForm(True)
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7810, "frmSettings.SettingsFormSmoke", "Settings reset failed."
    End If
    If modSettings.GetSettingText("DefaultSearchMode", "") <> "AND" Then
        Err.Raise vbObjectError + 7811, "frmSettings.SettingsFormSmoke", "Selectable setting was not reset."
    End If

    SelectSettingRowByKey "ManagedSheetVisibility"
    UpdateSelectedSettingFields
    If SelectedSettingStoreRow() <= 0 Then
        Err.Raise vbObjectError + 7812, "frmSettings.SettingsFormSmoke", "Managed sheet visibility row was not selected."
    End If
    If Not Me.cmbValue.Visible Or Me.txtValue.Visible Then
        Err.Raise vbObjectError + 7813, "frmSettings.SettingsFormSmoke", "Managed sheet visibility did not switch to combo box."
    End If
    If Me.cmbValue.ListCount < 3 Then
        Err.Raise vbObjectError + 7814, "frmSettings.SettingsFormSmoke", "Managed sheet visibility options were not populated."
    End If
    If Not ComboBoxContainsValue(Me.cmbValue, "Visible") Then
        Err.Raise vbObjectError + 7815, "frmSettings.SettingsFormSmoke", "Managed sheet visibility did not include Visible."
    End If
    Me.cmbValue.value = "Visible"
    selectedManagedVisibility = CStr(Me.cmbValue.value)
    resultText = SaveSelectedSettingForForm()
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7816, "frmSettings.SettingsFormSmoke", "Managed sheet visibility save failed."
    End If
    If modSettings.GetSettingText("ManagedSheetVisibility", "") <> selectedManagedVisibility Then
        Err.Raise vbObjectError + 7817, "frmSettings.SettingsFormSmoke", "Managed sheet visibility was not saved."
    End If
    If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetVisible Then
        Err.Raise vbObjectError + 7818, "frmSettings.SettingsFormSmoke", "Managed sheet visibility was not applied."
    End If

    SelectSettingRowByKey "ApiRetryCount"
    UpdateSelectedSettingFields
    If SelectedSettingStoreRow() <= 0 Then
        Err.Raise vbObjectError + 7819, "frmSettings.SettingsFormSmoke", "Textbox setting row was not selected."
    End If
    If Not Me.txtValue.Visible Or Me.cmbValue.Visible Then
        Err.Raise vbObjectError + 7820, "frmSettings.SettingsFormSmoke", "Textbox setting did not switch to text box."
    End If

    originalValueDescription = modSettings.SettingRowDescription(SelectedSettingStoreRow())
    Me.txtValue.text = "2"
    Me.txtDescription.text = "API通信再試行回数"
    resultText = SaveSelectedSettingForForm()
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7821, "frmSettings.SettingsFormSmoke", "Setting save failed."
    End If
    If modSettings.GetSettingText("ApiRetryCount", "") <> "2" Then
        Err.Raise vbObjectError + 7822, "frmSettings.SettingsFormSmoke", "Setting value was not saved."
    End If

    resultText = ResetSettingsToDefaultsForForm(True)
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7823, "frmSettings.SettingsFormSmoke", "Settings reset failed."
    End If
    If modSettings.GetSettingText("ApiRetryCount", "") <> "1" Then
        Err.Raise vbObjectError + 7824, "frmSettings.SettingsFormSmoke", "Settings were not reset."
    End If

    modSettings.SetSettingValue "ApiRetryCount", originalValue, originalValueDescription
    modSettings.SetSettingValue "DefaultSearchMode", originalSearchMode, originalSearchModeDescription
    modSettings.SetSettingValue "ManagedSheetVisibility", originalManagedVisibility, originalManagedVisibilityDescription
    Select Case originalManagedVisibility
        Case "Visible"
            If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetVisible Then
                Err.Raise vbObjectError + 7825, "frmSettings.SettingsFormSmoke", "Managed sheet visibility restore failed."
            End If
        Case "Hidden"
            If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetHidden Then
                Err.Raise vbObjectError + 7825, "frmSettings.SettingsFormSmoke", "Managed sheet visibility restore failed."
            End If
        Case Else
            If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetVeryHidden Then
                Err.Raise vbObjectError + 7825, "frmSettings.SettingsFormSmoke", "Managed sheet visibility restore failed."
            End If
    End Select
    SettingsFormSmoke = "ok:" & CStr(modSettings.SettingsCount())
    Unload frmSettings
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "SettingsForm", "frmSettings.SettingsFormSmoke", Err.description, "", "", ""
    SettingsFormSmoke = "error:" & CStr(Err.Number) & ":" & Err.description
    On Error GoTo CleanupFail
    modSettings.SetSettingValue "ApiRetryCount", originalValue, originalValueDescription
    modSettings.SetSettingValue "DefaultSearchMode", originalSearchMode, originalSearchModeDescription
    modSettings.SetSettingValue "ManagedSheetVisibility", originalManagedVisibility, originalManagedVisibilityDescription
    Unload frmSettings
    GoTo CleanupDone

CleanupFail:
    Resume CleanupDone

CleanupDone:
End Function

Private Sub ConfigureFormWindow()
    Me.caption = "設定"
    Me.Width = FORM_WIDTH_POINTS
    Me.Height = FORM_HEIGHT_POINTS
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub ConfigureControls()
    modSettings.ConfigureSettingsListBox Me.lstSettings
    Me.txtValue.text = ""
    Me.cmbValue.value = ""
    Me.txtValue.Visible = True
    Me.cmbValue.Visible = False
    Me.txtDescription.text = ""
    Me.txtDescription.MultiLine = True
    Me.txtDescription.WordWrap = True
    Me.txtDescription.ScrollBars = 2
    modUiState.ApplyDefaultControlFont Me.txtValue
    modUiState.ApplyDefaultControlFont Me.cmbValue
    modUiState.ApplyDefaultControlFont Me.txtDescription
End Sub

Private Sub UpdateSelectedSettingFields()
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedSettingStoreRow()
    If rowIndex <= 0 Then
        ClearSelectedSettingFields
        Exit Sub
    End If

    Me.lblKeyValue.caption = modSettings.SettingRowKey(rowIndex)
    Me.lblUpdatedValue.caption = modSettings.SettingRowUpdatedAt(rowIndex)
    ConfigureSelectedSettingValueEditor rowIndex
    Me.txtDescription.text = modSettings.SettingRowDescription(rowIndex)
    Me.lblStatus.caption = modSettings.SettingRowKey(rowIndex)
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "SettingsForm", "frmSettings.UpdateSelectedSettingFields", Err.description, "", "", CStr(rowIndex)
    Me.lblStatus.caption = "設定読み込み失敗: " & Err.description
End Sub

Private Sub ClearSelectedSettingFields()
    Me.lblKeyValue.caption = ""
    Me.lblUpdatedValue.caption = ""
    Me.txtValue.text = ""
    Me.cmbValue.Clear
    Me.cmbValue.value = ""
    Me.txtValue.Visible = True
    Me.cmbValue.Visible = False
    Me.txtDescription.text = ""
End Sub

Private Function SelectedSettingStoreRow() As Long
    On Error GoTo ErrHandler

    If Me.lstSettings.listIndex < 0 Then Exit Function
    SelectedSettingStoreRow = CLng(Val(Me.lstSettings.List(Me.lstSettings.listIndex, 4)))
    Exit Function

ErrHandler:
    SelectedSettingStoreRow = 0
End Function

Private Sub SelectSettingRowByStoreRow(ByVal rowIndex As Long)
    Dim listIndex As Long
    If rowIndex <= 0 Then
        Me.lstSettings.listIndex = -1
        Exit Sub
    End If

    For listIndex = 0 To Me.lstSettings.ListCount - 1
        If CLng(Val(Me.lstSettings.List(listIndex, 4))) = rowIndex Then
            Me.lstSettings.listIndex = listIndex
            Exit Sub
        End If
    Next listIndex
End Sub

Private Sub SelectSettingRowByKey(ByVal settingKey As String)
    Dim listIndex As Long
    For listIndex = 0 To Me.lstSettings.ListCount - 1
        If StrComp(CStr(Me.lstSettings.List(listIndex, 0)), settingKey, vbTextCompare) = 0 Then
            Me.lstSettings.listIndex = listIndex
            Exit Sub
        End If
    Next listIndex
End Sub

Private Sub ConfigureSelectedSettingValueEditor(ByVal rowIndex As Long)
    Dim settingKey As String
    settingKey = modSettings.SettingRowKey(rowIndex)

    If modSettings.SettingValueIsSelectable(settingKey) Then
        Me.txtValue.Visible = False
        Me.cmbValue.Visible = True
        modSettings.ConfigureSettingValueComboBox Me.cmbValue, settingKey, modSettings.SettingRowValue(rowIndex)
    Else
        Me.cmbValue.Visible = False
        Me.txtValue.Visible = True
        Me.txtValue.text = modSettings.SettingRowValue(rowIndex)
    End If
End Sub

Private Function SelectedSettingValueText(ByVal rowIndex As Long) As String
    If modSettings.SettingValueIsSelectable(modSettings.SettingRowKey(rowIndex)) Then
        SelectedSettingValueText = CStr(Me.cmbValue.value)
    Else
        SelectedSettingValueText = Me.txtValue.text
    End If
End Function

Private Function ComboBoxContainsValue(ByVal targetComboBox As Object, ByVal expectedValue As String) As Boolean
    Dim index As Long
    For index = 0 To targetComboBox.ListCount - 1
        If StrComp(CStr(targetComboBox.List(index)), expectedValue, vbTextCompare) = 0 Then
            ComboBoxContainsValue = True
            Exit Function
        End If
    Next index
End Function
