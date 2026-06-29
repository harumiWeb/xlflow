Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook ThisWorkbook
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "Main.Run", Err.description, "", "", "xlflow entry point"
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Sub ShowMain()
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook ThisWorkbook
    XlflowUI.ShowForm frmMain, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "Main.ShowMain", Err.description, "", "", "main form entry point"
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Sub ShowCompareForm()
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook ThisWorkbook
    XlflowUI.ShowForm frmCompare, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "CompareForm", "Main.ShowCompareForm", Err.description, "", "", "compare form entry point"
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Sub ShowSettingsForm()
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook ThisWorkbook
    XlflowUI.ShowForm frmSettings, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "SettingsForm", "Main.ShowSettingsForm", Err.description, "", "", "settings form entry point"
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Sub ShowAliasMasterForm()
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook ThisWorkbook
    XlflowUI.ShowForm frmAliasMaster, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AliasMasterForm", "Main.ShowAliasMasterForm", Err.description, "", "", "alias master form entry point"
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Sub ShowApiLogsForm()
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook ThisWorkbook
    Load frmLogs
    frmLogs.SetLogKindForForm "API"
    XlflowUI.ShowForm frmLogs, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "LogsForm", "Main.ShowApiLogsForm", Err.description, "", "", "api log form entry point"
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Sub ShowErrorLogsForm()
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook ThisWorkbook
    Load frmLogs
    frmLogs.SetLogKindForForm "ERROR"
    XlflowUI.ShowForm frmLogs, True
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "LogsForm", "Main.ShowErrorLogsForm", Err.description, "", "", "error log form entry point"
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Function RunStartupSmoke() As String
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook ThisWorkbook
    RunStartupSmoke = "ok:" & CStr(ThisWorkbook.Worksheets.Count)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "Main.RunStartupSmoke", Err.description, "", "", "diagnostic entry point"
    RunStartupSmoke = "error:" & Err.description
End Function

Public Function UiLayoutSmoke() As String
    On Error GoTo ErrHandler

    Load frmMain
    If Not modUiState.UserFormFontMatches(frmMain) Then
        Err.Raise vbObjectError + 6904, "Main.UiLayoutSmoke", "frmMain font was not normalized to " & modUiState.DefaultUserFormFontName() & "."
    End If
    If Not frmMain.rdoTableMarkdown.value Or frmMain.rdoTableAscii.value Then
        Err.Raise vbObjectError + 6905, "Main.UiLayoutSmoke", "Table display option buttons were not initialized."
    End If
    If frmMain.txtPreview.WordWrap Or frmMain.txtUnitPreview.WordWrap Then
        Err.Raise vbObjectError + 6907, "Main.UiLayoutSmoke", "Preview text boxes must not wrap ASCII tables."
    End If
    If frmMain.txtPreview.ScrollBars <> 3 Or frmMain.txtUnitPreview.ScrollBars <> 3 Then
        Err.Raise vbObjectError + 6908, "Main.UiLayoutSmoke", "Preview text boxes must have horizontal and vertical scroll bars."
    End If
    If frmMain.cmbCitationFormat.value <> modCitation.CitationFormatWithSource() Then
        Err.Raise vbObjectError + 6909, "Main.UiLayoutSmoke", "Citation format combo box was not initialized."
    End If
    If Not frmMain.rdoCitationBody.value Or frmMain.rdoCitationSearch.value Then
        Err.Raise vbObjectError + 6911, "Main.UiLayoutSmoke", "Citation target option buttons were not initialized."
    End If
    If frmMain.cmdHistory.caption <> "履歴" Or frmMain.cmdBookmarks.caption <> "ブックマーク" Then
        Err.Raise vbObjectError + 6912, "Main.UiLayoutSmoke", "History or bookmark buttons were not initialized."
    End If
    If frmMain.cmdBookmarks.Width < 70 Then
        Err.Raise vbObjectError + 6914, "Main.UiLayoutSmoke", "Bookmark button width was not expanded."
    End If
    If frmMain.cmdBookmarkAdd.caption <> "ブックマーク追加" Then
        Err.Raise vbObjectError + 6913, "Main.UiLayoutSmoke", "Bookmark add button was not initialized."
    End If
    If frmMain.cmdCompare.caption <> "比較" Or frmMain.cmdSettings.caption <> "設定" Or frmMain.cmdAliasMaster.caption <> "別名マスター" Then
        Err.Raise vbObjectError + 6916, "Main.UiLayoutSmoke", "Compare or settings buttons were not initialized."
    End If
    If frmMain.cmdApiLogs.caption <> "APIログ" Or frmMain.cmdErrorLogs.caption <> "エラーログ" Then
        Err.Raise vbObjectError + 6917, "Main.UiLayoutSmoke", "Log buttons were not initialized."
    End If
    If frmMain.cmdCompare.Width < 52 Or frmMain.cmdSettings.Width < 52 Or frmMain.cmdAliasMaster.Width < 90 Or frmMain.cmdApiLogs.Width < 58 Or frmMain.cmdErrorLogs.Width < 58 Then
        Err.Raise vbObjectError + 6918, "Main.UiLayoutSmoke", "Main action buttons were not widened enough."
    End If
    If frmMain.lblStatus.Top > 558 Or frmMain.lblStatus.Height < 20 Then
        Err.Raise vbObjectError + 6915, "Main.UiLayoutSmoke", "Main status label was not moved up enough."
    End If
    If frmMain.txtCitation.WordWrap Or frmMain.txtCitation.ScrollBars <> 3 Then
        Err.Raise vbObjectError + 6910, "Main.UiLayoutSmoke", "Citation text box must not wrap ASCII table citations."
    End If
    Dim mainText As String
    mainText = "frmMain:" & _
        "width=" & CStr(frmMain.Width) & _
        ";height=" & CStr(frmMain.Height) & _
        ";caption=" & frmMain.caption & _
        ";statusTop=" & CStr(frmMain.lblStatus.Top)
    Unload frmMain

    Load frmLawAdd
    If Not modUiState.UserFormFontMatches(frmLawAdd) Then
        Err.Raise vbObjectError + 6906, "Main.UiLayoutSmoke", "frmLawAdd font was not normalized to " & modUiState.DefaultUserFormFontName() & "."
    End If
    Dim lawAddText As String
    lawAddText = "frmLawAdd:" & _
        "width=" & CStr(frmLawAdd.Width) & _
        ";height=" & CStr(frmLawAdd.Height) & _
        ";caption=" & frmLawAdd.caption & _
        ";statusTop=" & CStr(frmLawAdd.lblStatus.Top)
    Unload frmLawAdd

    modTempCache.DeleteAllTempCaches ThisWorkbook
    UiLayoutSmoke = mainText & " | " & lawAddText
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "UiLayout", "Main.UiLayoutSmoke", Err.description, "", "", "diagnostic entry point"
    UiLayoutSmoke = "error:" & Err.description
End Function

Public Function MainBodyViewSmoke() As String
    On Error GoTo ErrHandler

    modLawNavigator.PrepareBodyNavigatorSmokeData

    Load frmMain
    If frmMain.lstBodyNav.ListCount <> 3 Then
        Err.Raise vbObjectError + 6901, "Main.MainBodyViewSmoke", "Body navigation list was not populated."
    End If
    If InStr(1, frmMain.txtPreview.text, "第十一条　受給権は差し押えることができない。", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6902, "Main.MainBodyViewSmoke", "Body preview text was not populated."
    End If
    If InStr(1, frmMain.txtUnitPreview.text, "附則側の第十一条", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 6903, "Main.MainBodyViewSmoke", "Unrelated article text was mixed."
    End If
    If InStr(1, frmMain.txtUnitPreview.text, "第十一条　受給権は差し押えることができない。", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6903, "Main.MainBodyViewSmoke", "Unit preview text was not populated."
    End If

    MainBodyViewSmoke = "ok:" & CStr(frmMain.lstBodyNav.ListCount)
    Unload frmMain
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "Main.MainBodyViewSmoke", Err.description, "", "", "diagnostic entry point"
    MainBodyViewSmoke = "error:" & Err.description
End Function

Public Function MaintenanceSmoke() As String
    On Error GoTo ErrHandler

    Dim startupResult As String
    Dim layoutResult As String
    Dim navigatorResult As String
    Dim mainBodyResult As String
    Dim parserBulkResult As String
    Dim parserCoverageResult As String
    Dim parserLongTextResult As String
    Dim beforeCloseSaveResult As String
    Dim bodySearchResult As String
    Dim historyResult As String
    Dim bookmarkResult As String
    Dim bookmarkAddResult As String
    Dim historyFormResult As String
    Dim bookmarkFormResult As String
    Dim compareResult As String
    Dim settingsResult As String
    Dim logsResult As String
    Dim aliasMasterResult As String
    Dim aliasMasterFormResult As String
    Dim monitorPlacementResult As String
    Dim mainCitationResult As String
    Dim tableFormatterResult As String
    Dim lawTableParserResult As String
    Dim tableDisplayModeResult As String
    Dim citationResult As String

    startupResult = RunStartupSmoke()
    EnsureSmokeOk "RunStartupSmoke", startupResult

    layoutResult = UiLayoutSmoke()
    EnsureSmokeNotError "UiLayoutSmoke", layoutResult

    navigatorResult = modLawNavigator.BodyNavigatorSmoke()
    EnsureSmokeOk "BodyNavigatorSmoke", navigatorResult

    mainBodyResult = MainBodyViewSmoke()
    EnsureSmokeOk "MainBodyViewSmoke", mainBodyResult

    bodySearchResult = MainBodySearchSmoke()
    EnsureSmokeOk "MainBodySearchSmoke", bodySearchResult

    historyResult = modHistory.SearchHistorySmoke()
    EnsureSmokeOk "SearchHistorySmoke", historyResult

    bookmarkResult = modBookmark.BookmarkSmoke()
    EnsureSmokeOk "BookmarkSmoke", bookmarkResult

    bookmarkAddResult = BookmarkAddSmoke()
    EnsureSmokeOk "BookmarkAddSmoke", bookmarkAddResult

    historyFormResult = HistoryFormSmoke()
    EnsureSmokeOk "HistoryFormSmoke", historyFormResult

    bookmarkFormResult = BookmarkFormSmoke()
    EnsureSmokeOk "BookmarkFormSmoke", bookmarkFormResult

    compareResult = CompareFormSmoke()
    EnsureSmokeOk "CompareFormSmoke", compareResult

    settingsResult = SettingsFormSmoke()
    EnsureSmokeOk "SettingsFormSmoke", settingsResult

    logsResult = LogsFormSmoke()
    EnsureSmokeOk "LogsFormSmoke", logsResult

    aliasMasterResult = modAliasMaster.AliasMasterSmoke()
    EnsureSmokeOk "AliasMasterSmoke", aliasMasterResult

    aliasMasterFormResult = AliasMasterFormSmoke()
    EnsureSmokeOk "AliasMasterFormSmoke", aliasMasterFormResult

    monitorPlacementResult = modWindowPlacement.MonitorPlacementSmoke()
    EnsureSmokeOk "MonitorPlacementSmoke", monitorPlacementResult

    mainCitationResult = MainCitationSmoke()
    EnsureSmokeOk "MainCitationSmoke", mainCitationResult

    tableFormatterResult = modTableFormatter.TableFormatterSmoke()
    EnsureSmokeOk "TableFormatterSmoke", tableFormatterResult

    lawTableParserResult = modLawTableParser.LawTableParserSmoke()
    EnsureSmokeOk "LawTableParserSmoke", lawTableParserResult

    tableDisplayModeResult = modLawNavigator.TableDisplayModeSmoke()
    EnsureSmokeOk "TableDisplayModeSmoke", tableDisplayModeResult

    citationResult = modCitation.CitationSmoke()
    EnsureSmokeOk "CitationSmoke", citationResult

    parserCoverageResult = modLawParser.BodyUnitCoverageSmoke()
    EnsureSmokeOk "BodyUnitCoverageSmoke", parserCoverageResult

    parserBulkResult = modLawParser.BodyUnitBulkWriteSmoke()
    EnsureSmokeOk "BodyUnitBulkWriteSmoke", parserBulkResult

    parserLongTextResult = modLawParser.BodyUnitLongTextSmoke()
    EnsureSmokeOk "BodyUnitLongTextSmoke", parserLongTextResult

    beforeCloseSaveResult = modAppStartup.BeforeCloseSaveSmoke()
    EnsureSmokeOk "BeforeCloseSaveSmoke", beforeCloseSaveResult

    modTempCache.DeleteAllTempCaches ThisWorkbook
    MaintenanceSmoke = "ok:" & startupResult & " | " & layoutResult & " | " & navigatorResult & " | " & mainBodyResult & " | " & bodySearchResult & " | " & historyResult & " | " & bookmarkResult & " | " & historyFormResult & " | " & bookmarkFormResult & " | " & compareResult & " | " & settingsResult & " | " & logsResult & " | " & aliasMasterResult & " | " & aliasMasterFormResult & " | " & monitorPlacementResult & " | " & mainCitationResult & " | " & tableFormatterResult & " | " & lawTableParserResult & " | " & tableDisplayModeResult & " | " & citationResult & " | " & parserCoverageResult & " | " & parserBulkResult & " | " & parserLongTextResult & " | " & beforeCloseSaveResult
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Maintenance", "Main.MaintenanceSmoke", Err.description, "", "", "diagnostic entry point"
    MaintenanceSmoke = "error:" & Err.description
End Function

Public Function MainCitationSmoke() As String
    On Error GoTo ErrHandler

    modLawNavigator.PrepareBodyNavigatorSmokeData

    Load frmMain
    If frmMain.lstBodyNav.ListCount < 3 Then
        Err.Raise vbObjectError + 6970, "Main.MainCitationSmoke", "Body navigation list was not populated for citation."
    End If

    frmMain.lstBodyNav.listIndex = 2
    frmMain.cmbCitationFormat.value = modCitation.CitationFormatTableAscii()

    Dim resultText As String
    resultText = frmMain.GenerateCitationForForm()
    EnsureSmokeOk "frmMain.GenerateCitationForForm", resultText

    If InStr(1, frmMain.txtCitation.text, "+------+-----+", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6971, "Main.MainCitationSmoke", "ASCII table citation was not generated."
    End If
    If InStr(1, frmMain.txtCitation.text, "出典: テスト法（令和七年法律第一号） 別表第一", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6972, "Main.MainCitationSmoke", "Citation source text mismatch."
    End If

    frmMain.cmbCitationFormat.value = modCitation.CitationFormatExcelTsv()
    resultText = frmMain.GenerateCitationForForm()
    EnsureSmokeOk "frmMain.GenerateCitationForForm Excel TSV", resultText
    If InStr(1, frmMain.txtCitation.text, "項目" & vbTab & "値", vbBinaryCompare) = 0 _
        Or InStr(1, frmMain.txtCitation.text, "+------+-----+", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 6973, "Main.MainCitationSmoke", "Excel TSV table citation was not generated."
    End If

    MainCitationSmoke = "ok:" & CStr(Len(frmMain.txtCitation.text))
    Unload frmMain
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "Main.MainCitationSmoke", Err.description, "", "", "diagnostic entry point"
    MainCitationSmoke = "error:" & Err.description
End Function

Public Function MainBodySearchSmoke() As String
    On Error GoTo ErrHandler

    modLawNavigator.PrepareBodyNavigatorSmokeData

    Load frmMain
    frmMain.txtBodySearch.text = "受給権 差し押える"
    frmMain.cmbSearchMode.value = "AND"

    Dim resultText As String
    resultText = frmMain.PerformBodySearchForForm()
    EnsureSmokeOk "frmMain.PerformBodySearchForForm", resultText

    If frmMain.lstResults.ListCount <> 1 Then
        Err.Raise vbObjectError + 6960, "Main.MainBodySearchSmoke", "Search result count mismatch."
    End If
    If InStr(1, frmMain.txtPreview.text, "第十一条　受給権は差し押えることができない。", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6961, "Main.MainBodySearchSmoke", "Search context preview mismatch."
    End If
    If InStr(1, frmMain.txtUnitPreview.text, "受給権は差し押えることができない。", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6962, "Main.MainBodySearchSmoke", "Search unit preview mismatch."
    End If

    MainBodySearchSmoke = "ok:" & CStr(frmMain.lstResults.ListCount)
    Unload frmMain
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "MainForm", "Main.MainBodySearchSmoke", Err.description, "", "", "diagnostic entry point"
    MainBodySearchSmoke = "error:" & Err.description
End Function

Public Function HistoryFormSmoke() As String
    HistoryFormSmoke = frmHistory.HistoryFormSmoke()
End Function

Public Function BookmarkFormSmoke() As String
    BookmarkFormSmoke = frmBookmarks.BookmarkFormSmoke()
End Function

Public Function CompareFormSmoke() As String
    CompareFormSmoke = frmCompare.CompareFormSmoke()
End Function

Public Function BookmarkAddSmoke() As String
    BookmarkAddSmoke = frmMain.BookmarkAddSmoke()
End Function

Public Function SettingsFormSmoke() As String
    SettingsFormSmoke = frmSettings.SettingsFormSmoke()
End Function

Public Function LogsFormSmoke() As String
    LogsFormSmoke = frmLogs.LogsFormSmoke()
End Function

Public Function AliasMasterFormSmoke() As String
    AliasMasterFormSmoke = frmAliasMaster.AliasMasterFormSmoke()
End Function

Private Sub EnsureSmokeOk(ByVal smokeName As String, ByVal resultText As String)
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 6950, "Main.MaintenanceSmoke", smokeName & " failed: " & resultText
    End If
End Sub

Private Sub EnsureSmokeNotError(ByVal smokeName As String, ByVal resultText As String)
    If Left$(resultText, 6) = "error:" Then
        Err.Raise vbObjectError + 6951, "Main.MaintenanceSmoke", smokeName & " failed: " & resultText
    End If
End Sub
