Option Explicit

Private Const FORM_WIDTH_POINTS As Double = 840
Private Const FORM_HEIGHT_POINTS As Double = 520

Private currentLogKind As String

Private Sub UserForm_Initialize()
    ConfigureFormWindow
    ConfigureControls
    modUiState.ApplyDefaultUserFormFont Me
    currentLogKind = "API"
    RefreshLogsForForm
End Sub

Private Sub UserForm_Activate()
    modUiState.ApplyDefaultUserFormFont Me
    modWindowPlacement.CenterUserFormOnExcelMonitor Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    ClearCurrentLogForForm
End Sub

Private Sub lstLogs_Click()
    UpdateSelectedLogDetail
End Sub

Public Sub SetLogKindForForm(ByVal logKind As String)
    currentLogKind = NormalizeLogKind(logKind)
    Me.caption = LogKindDisplayText(currentLogKind)
    Me.lblKind.caption = LogKindDisplayText(currentLogKind)
    RefreshLogsForForm
End Sub

Public Function RefreshLogsForForm() As String
    On Error GoTo ErrHandler

    Dim selectedRow As Long
    selectedRow = SelectedLogStoreRow()

    Dim itemCount As Long
    If IsApiLogMode() Then
        itemCount = modLogger.PopulateApiLogListBox(Me.lstLogs)
    Else
        itemCount = modLogger.PopulateErrorLogListBox(Me.lstLogs)
    End If

    If itemCount > 0 Then
        SelectLogRowByStoreRow selectedRow
        If Me.lstLogs.listIndex < 0 Then
            Me.lstLogs.listIndex = 0
        End If
        UpdateSelectedLogDetail
    Else
        Me.txtDetail.text = ""
    End If

    Me.lblStatus.caption = LogKindDisplayText(currentLogKind) & " " & CStr(itemCount) & " 件"
    RefreshLogsForForm = "ok:" & CStr(itemCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LogsForm", "frmLogs.RefreshLogsForForm", Err.description, "", "", currentLogKind
    Me.lblStatus.caption = "ログの読み込み失敗: " & Err.description
    RefreshLogsForForm = "error:" & Err.description
End Function

Public Function ClearCurrentLogForForm(Optional ByVal skipConfirm As Boolean = False) As String
    On Error GoTo ErrHandler

    If Not skipConfirm Then
        If XlflowUI.MsgBox("logs-clear-confirm", LogKindDisplayText(currentLogKind) & "をすべて削除しますか？", vbYesNo + vbQuestion, "法令検索ビューアー", vbNo) <> vbYes Then
            Me.lblStatus.caption = "削除をキャンセルしました"
            ClearCurrentLogForForm = "cancel"
            Exit Function
        End If
    End If

    If IsApiLogMode() Then
        modLogger.ClearApiLog
    Else
        modLogger.ClearErrorLog
    End If

    RefreshLogsForForm
    Me.lblStatus.caption = LogKindDisplayText(currentLogKind) & " をすべて削除しました"
    ClearCurrentLogForForm = "ok"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LogsForm", "frmLogs.ClearCurrentLogForForm", Err.description, "", "", currentLogKind
    Me.lblStatus.caption = "ログ削除失敗: " & Err.description
    ClearCurrentLogForForm = "error:" & Err.description
End Function

Public Function LogsFormSmoke() As String
    On Error GoTo ErrHandler

    Dim apiSheet As Worksheet
    Dim errorSheet As Worksheet
    Set apiSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.ApiLogSheetName())
    Set errorSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.ErrorLogSheetName())

    Dim apiSnapshot As Variant
    Dim errorSnapshot As Variant
    Dim apiTimestampText As String
    Dim errorTimestampText As String
    apiSnapshot = SnapshotSheetValues(apiSheet)
    errorSnapshot = SnapshotSheetValues(errorSheet)

    modLogger.LogApiCall "Smoke", "/logs/api", 200, True, "TEST001", "2025-04-01", 1, 12, ""
    modLogger.LogErrorSafe "Smoke", "frmLogs.LogsFormSmoke", "smoke error", "TEST001", "2025-04-01", "detail"

    Load frmLogs
    SetLogKindForForm "API"
    If Me.cmdClear.caption <> "クリア" Then
        Err.Raise vbObjectError + 7901, "frmLogs.LogsFormSmoke", "Clear button caption was not initialized."
    End If
    If Me.lblStatus.Top > 482 Or Me.lblStatus.Height < 18 Then
        Err.Raise vbObjectError + 7902, "frmLogs.LogsFormSmoke", "Status label was not moved up enough."
    End If
    If Me.lstLogs.ListCount <= 0 Then
        Err.Raise vbObjectError + 7903, "frmLogs.LogsFormSmoke", "API log list was not populated."
    End If
    apiTimestampText = CStr(Me.lstLogs.List(0, 0))
    If IsNumeric(apiTimestampText) _
        Or InStr(1, apiTimestampText, "/", vbBinaryCompare) = 0 _
        Or InStr(1, apiTimestampText, ":", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 79031, "frmLogs.LogsFormSmoke", "API log timestamp was not formatted as text."
    End If

    Dim resultText As String
    resultText = ClearCurrentLogForForm(True)
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7904, "frmLogs.LogsFormSmoke", "API log clear failed."
    End If
    If modLogger.ApiLogCount() <> 0 Then
        Err.Raise vbObjectError + 7905, "frmLogs.LogsFormSmoke", "API log clear did not empty the sheet."
    End If

    SetLogKindForForm "ERROR"
    If Me.lstLogs.ListCount <= 0 Then
        Err.Raise vbObjectError + 7906, "frmLogs.LogsFormSmoke", "Error log list was not populated."
    End If
    errorTimestampText = CStr(Me.lstLogs.List(0, 0))
    If IsNumeric(errorTimestampText) _
        Or InStr(1, errorTimestampText, "/", vbBinaryCompare) = 0 _
        Or InStr(1, errorTimestampText, ":", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 79061, "frmLogs.LogsFormSmoke", "Error log timestamp was not formatted as text."
    End If

    resultText = ClearCurrentLogForForm(True)
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7907, "frmLogs.LogsFormSmoke", "Error log clear failed."
    End If
    If modLogger.ErrorLogCount() <> 0 Then
        Err.Raise vbObjectError + 7908, "frmLogs.LogsFormSmoke", "Error log clear did not empty the sheet."
    End If

    RestoreSheetValues apiSheet, apiSnapshot
    RestoreSheetValues errorSheet, errorSnapshot
    modAppStartup.SaveWorkbookState ThisWorkbook

    LogsFormSmoke = "ok:" & CStr(modLogger.ApiLogCount()) & ":" & CStr(modLogger.ErrorLogCount())
    Unload frmLogs
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LogsForm", "frmLogs.LogsFormSmoke", Err.description, "", "", currentLogKind
    LogsFormSmoke = "error:" & CStr(Err.Number) & ":" & Err.description
    On Error GoTo CleanupFail
    RestoreSheetValues apiSheet, apiSnapshot
    RestoreSheetValues errorSheet, errorSnapshot
    modAppStartup.SaveWorkbookState ThisWorkbook
    Unload frmLogs
    GoTo CleanupDone

CleanupFail:
    Resume CleanupDone

CleanupDone:
End Function

Private Sub ConfigureFormWindow()
    Me.caption = "API通信ログ"
    Me.Width = FORM_WIDTH_POINTS
    Me.Height = FORM_HEIGHT_POINTS
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub ConfigureControls()
    modLogger.ConfigureApiLogListBox Me.lstLogs
    ConfigureDetailTextBox Me.txtDetail
    Me.lblKind.caption = "API通信ログ"
    Me.cmdClear.caption = "クリア"
End Sub

Private Sub ConfigureDetailTextBox(ByVal targetTextBox As Object)
    targetTextBox.MultiLine = True
    targetTextBox.WordWrap = False
    targetTextBox.ScrollBars = 3
    targetTextBox.Locked = True
    modUiState.ApplyDefaultControlFont targetTextBox
End Sub

Private Function NormalizeLogKind(ByVal logKind As String) As String
    If StrComp(Trim$(logKind), "ERROR", vbTextCompare) = 0 Then
        NormalizeLogKind = "ERROR"
    Else
        NormalizeLogKind = "API"
    End If
End Function

Private Function LogKindDisplayText(ByVal logKind As String) As String
    If StrComp(logKind, "ERROR", vbTextCompare) = 0 Then
        LogKindDisplayText = "エラーログ"
    Else
        LogKindDisplayText = "API通信ログ"
    End If
End Function

Private Function IsApiLogMode() As Boolean
    IsApiLogMode = StrComp(currentLogKind, "ERROR", vbTextCompare) <> 0
End Function

Private Function SelectedLogStoreRow() As Long
    On Error GoTo ErrHandler

    If Me.lstLogs.listIndex < 0 Then Exit Function
    SelectedLogStoreRow = CLng(Val(Me.lstLogs.List(Me.lstLogs.listIndex, 5)))
    Exit Function

ErrHandler:
    SelectedLogStoreRow = 0
End Function

Private Sub SelectLogRowByStoreRow(ByVal rowIndex As Long)
    Dim listIndex As Long
    If rowIndex <= 0 Then
        Me.lstLogs.listIndex = -1
        Exit Sub
    End If

    For listIndex = 0 To Me.lstLogs.ListCount - 1
        If CLng(Val(Me.lstLogs.List(listIndex, 5))) = rowIndex Then
            Me.lstLogs.listIndex = listIndex
            Exit Sub
        End If
    Next listIndex
End Sub

Private Sub UpdateSelectedLogDetail()
    On Error GoTo ErrHandler

    Dim rowIndex As Long
    rowIndex = SelectedLogStoreRow()
    If rowIndex <= 0 Then
        Me.txtDetail.text = ""
        Exit Sub
    End If

    If IsApiLogMode() Then
        Me.txtDetail.text = modLogger.ApiLogRowDetailText(rowIndex)
    Else
        Me.txtDetail.text = modLogger.ErrorLogRowDetailText(rowIndex)
    End If
    Me.lblStatus.caption = LogKindDisplayText(currentLogKind) & " " & CStr(rowIndex)
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "LogsForm", "frmLogs.UpdateSelectedLogDetail", Err.description, "", "", currentLogKind
    Me.txtDetail.text = "ログ詳細の読み込み失敗: " & Err.description
End Sub

Private Function SnapshotSheetValues(ByVal targetSheet As Worksheet) As Variant
    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(targetSheet)
    If lastRow <= 0 Then
        SnapshotSheetValues = Empty
        Exit Function
    End If

    Dim lastCol As Long
    lastCol = targetSheet.cells(1, targetSheet.Columns.Count).End(xlToLeft).Column
    SnapshotSheetValues = targetSheet.Range(targetSheet.cells(1, 1), targetSheet.cells(lastRow, lastCol)).Value2
End Function

Private Sub RestoreSheetValues(ByVal targetSheet As Worksheet, ByVal snapshot As Variant)
    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(targetSheet)
    If lastRow > 0 Then
        targetSheet.rows("1:" & CStr(lastRow)).ClearContents
    End If

    If Not IsArray(snapshot) Then
        Exit Sub
    End If

    targetSheet.Range(targetSheet.cells(1, 1), targetSheet.cells(UBound(snapshot, 1), UBound(snapshot, 2))).Value2 = snapshot
End Sub


