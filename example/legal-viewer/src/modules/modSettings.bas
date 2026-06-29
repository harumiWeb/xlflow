Attribute VB_Name = "modSettings"
Option Explicit

Private Const DEFAULT_API_BASE_URL As String = "https://laws.e-gov.go.jp/api/2"

Public Sub EnsureDefaultSettings(ByVal targetWorkbook As Workbook)
    On Error GoTo ErrHandler

    Dim settingsSheet As Worksheet
    Set settingsSheet = modSheetStore.EnsurePersistentSheet(targetWorkbook, modSheetStore.SettingsSheetName())

    Dim defaults As Variant
    defaults = Array( _
        Array("DefaultCitationFormat", "出典情報つき引用", "引用文の既定形式"), _
        Array("DefaultCitationTarget", "検索結果の最小単位", "引用対象の既定値"), _
        Array("DefaultSearchMode", "AND", "検索モードの既定値"), _
        Array("ApiLogLimit", "500", "API通信ログ保存上限"), _
        Array("ManagedSheetVisibility", "VeryHidden", "管理シート表示状態"), _
        Array("ApiRetryCount", "1", "API通信再試行回数"), _
        Array("TimeoutResolveMs", "10000", "名前解決タイムアウト"), _
        Array("TimeoutConnectMs", "10000", "接続タイムアウト"), _
        Array("TimeoutSendMs", "10000", "送信タイムアウト"), _
        Array("TimeoutReceiveMs", "30000", "受信タイムアウト"), _
        Array("ApiMinIntervalSeconds", "0.5", "API呼び出し最低間隔"), _
        Array("ApiBaseUrl", DEFAULT_API_BASE_URL, "e-Gov法令API Version 2 ベースURL"))

    Dim index As Long
    For index = LBound(defaults) To UBound(defaults)
        If FindSettingRow(settingsSheet, CStr(defaults(index)(0))) = 0 Then
            modSheetStore.AppendRow settingsSheet, Array(defaults(index)(0), defaults(index)(1), defaults(index)(2), Format$(Now, "yyyy-mm-dd hh:nn:ss"))
        End If
    Next index
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modSettings.EnsureDefaultSettings", Err.description
End Sub

Public Function SettingsCount() As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, modSheetStore.SettingsSheetName()) Then
        SettingsCount = 0
        Exit Function
    End If

    SettingsCount = Application.Max(0, modSheetStore.LastUsedRow(ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName())) - 1)
End Function

Public Sub ConfigureSettingsListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 5
    targetListBox.columnWidths = "120 pt;140 pt;250 pt;110 pt;0 pt"
    targetListBox.BoundColumn = 5
End Sub

Public Function SettingValueIsSelectable(ByVal settingKey As String) As Boolean
    Select Case NormalizeSettingKey(settingKey)
        Case "DefaultCitationFormat", "DefaultCitationTarget", "DefaultSearchMode", "ManagedSheetVisibility"
            SettingValueIsSelectable = True
    End Select
End Function

Public Sub ConfigureSettingValueComboBox(ByVal targetComboBox As Object, ByVal settingKey As String, Optional ByVal defaultValue As String = "")
    Dim options As Variant
    options = SettingValueOptions(settingKey)

    targetComboBox.Clear
    targetComboBox.Style = 2
    targetComboBox.columnCount = 1
    targetComboBox.BoundColumn = 1

    Dim index As Long
    If IsArray(options) Then
        For index = LBound(options) To UBound(options)
            targetComboBox.AddItem CStr(options(index))
        Next index
    End If

    If Len(defaultValue) = 0 Then
        defaultValue = GetSettingText(settingKey, "")
    End If

    If Len(defaultValue) > 0 Then
        For index = 0 To targetComboBox.ListCount - 1
            If StrComp(CStr(targetComboBox.List(index)), defaultValue, vbTextCompare) = 0 Then
                targetComboBox.listIndex = index
                Exit Sub
            End If
        Next index
    End If

    If targetComboBox.ListCount > 0 Then
        targetComboBox.listIndex = 0
    End If
End Sub

Public Function SettingValueOptions(ByVal settingKey As String) As Variant
    Select Case NormalizeSettingKey(settingKey)
        Case "DefaultCitationFormat"
            SettingValueOptions = Array( _
                modCitation.CitationFormatSimple(), _
                modCitation.CitationFormatWithSource(), _
                modCitation.CitationFormatMarkdown(), _
                modCitation.CitationFormatTableMarkdown(), _
                modCitation.CitationFormatTableAscii(), _
                modCitation.CitationFormatExcelTsv())
        Case "DefaultCitationTarget"
            SettingValueOptions = Array("本文", "検索結果の最小単位")
        Case "DefaultSearchMode"
            SettingValueOptions = Array("AND", "OR")
        Case "ManagedSheetVisibility"
            SettingValueOptions = Array("Visible", "Hidden", "VeryHidden")
        Case Else
            SettingValueOptions = Empty
    End Select
End Function

Public Function PopulateSettingsListBox(ByVal targetListBox As Object) As Long
    On Error GoTo ErrHandler

    ConfigureSettingsListBox targetListBox
    targetListBox.Clear

    If Not modSheetStore.SheetExists(ThisWorkbook, modSheetStore.SettingsSheetName()) Then
        PopulateSettingsListBox = 0
        Exit Function
    End If

    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName())

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(settingsSheet)
    If lastRow < 2 Then
        PopulateSettingsListBox = 0
        Exit Function
    End If

    Dim rowIndex As Long
    Dim listIndex As Long
    For rowIndex = 2 To lastRow
        targetListBox.AddItem CellText(settingsSheet.cells(rowIndex, 1).Value2)
        listIndex = targetListBox.ListCount - 1
        targetListBox.List(listIndex, 1) = CellText(settingsSheet.cells(rowIndex, 2).Value2)
        targetListBox.List(listIndex, 2) = CellText(settingsSheet.cells(rowIndex, 3).Value2)
        targetListBox.List(listIndex, 3) = SettingRowUpdatedAt(rowIndex)
        targetListBox.List(listIndex, 4) = CStr(rowIndex)
    Next rowIndex

    PopulateSettingsListBox = targetListBox.ListCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Settings", "modSettings.PopulateSettingsListBox", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function SettingRowKey(ByVal rowIndex As Long) As String
    SettingRowKey = SettingRowText(rowIndex, 1)
End Function

Public Function SettingRowValue(ByVal rowIndex As Long) As String
    SettingRowValue = SettingRowText(rowIndex, 2)
End Function

Public Function SettingRowDescription(ByVal rowIndex As Long) As String
    SettingRowDescription = SettingRowText(rowIndex, 3)
End Function

Public Function SettingRowUpdatedAt(ByVal rowIndex As Long) As String
    If rowIndex < 2 Then Exit Function
    If Not modSheetStore.SheetExists(ThisWorkbook, modSheetStore.SettingsSheetName()) Then Exit Function

    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName())
    If rowIndex > modSheetStore.LastUsedRow(settingsSheet) Then Exit Function

    SettingRowUpdatedAt = FormatSettingTimestamp(settingsSheet.cells(rowIndex, 4).Value2)
End Function

Public Function GetSettingText(ByVal settingKey As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrHandler

    Dim settingsSheet As Worksheet
    Set settingsSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.SettingsSheetName())

    Dim rowIndex As Long
    rowIndex = FindSettingRow(settingsSheet, settingKey)
    If rowIndex = 0 Then
        GetSettingText = defaultValue
    Else
        GetSettingText = CStr(settingsSheet.cells(rowIndex, 2).Value2)
    End If
    Exit Function

ErrHandler:
    GetSettingText = defaultValue
End Function

Public Function GetSettingLong(ByVal settingKey As String, ByVal defaultValue As Long) As Long
    On Error GoTo ErrHandler

    Dim valueText As String
    valueText = GetSettingText(settingKey, CStr(defaultValue))
    If Len(Trim$(valueText)) = 0 Then
        GetSettingLong = defaultValue
    Else
        GetSettingLong = CLng(valueText)
    End If
    Exit Function

ErrHandler:
    GetSettingLong = defaultValue
End Function

Public Function GetSettingDouble(ByVal settingKey As String, ByVal defaultValue As Double) As Double
    On Error GoTo ErrHandler

    Dim valueText As String
    valueText = GetSettingText(settingKey, CStr(defaultValue))
    If Len(Trim$(valueText)) = 0 Then
        GetSettingDouble = defaultValue
    Else
        GetSettingDouble = CDbl(valueText)
    End If
    Exit Function

ErrHandler:
    GetSettingDouble = defaultValue
End Function

Public Sub SetSettingValue(ByVal settingKey As String, ByVal settingValue As String, Optional ByVal description As String = "")
    On Error GoTo ErrHandler

    Dim settingsSheet As Worksheet
    Set settingsSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.SettingsSheetName())

    Dim rowIndex As Long
    rowIndex = FindSettingRow(settingsSheet, settingKey)
    If rowIndex = 0 Then
        modSheetStore.AppendRow settingsSheet, Array(settingKey, settingValue, description, Format$(Now, "yyyy-mm-dd hh:nn:ss"))
    Else
        settingsSheet.cells(rowIndex, 2).Value2 = settingValue
        If Len(description) > 0 Then
            settingsSheet.cells(rowIndex, 3).Value2 = description
        End If
        settingsSheet.cells(rowIndex, 4).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    End If
    If StrComp(NormalizeSettingKey(settingKey), "ManagedSheetVisibility", vbTextCompare) = 0 Then
        modSheetStore.ApplyManagedSheetVisibility ThisWorkbook
    End If
    modAppStartup.SaveWorkbookState ThisWorkbook
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modSettings.SetSettingValue", Err.description
End Sub

Public Sub ResetSettingsToDefaults()
    On Error GoTo ErrHandler

    Dim settingsSheet As Worksheet
    Set settingsSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.SettingsSheetName())
    modSheetStore.ClearDataRows settingsSheet
    EnsureDefaultSettings ThisWorkbook
    modSheetStore.ApplyManagedSheetVisibility ThisWorkbook
    modAppStartup.SaveWorkbookState ThisWorkbook
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "Settings", "modSettings.ResetSettingsToDefaults", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Function SettingsSmoke() As String
    On Error GoTo ErrHandler

    Dim settingsSheet As Worksheet
    Set settingsSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.SettingsSheetName())

    Dim snapshot As Variant
    snapshot = SnapshotSheetValues(settingsSheet)

    Dim originalValue As String
    Dim originalManagedVisibility As String
    Dim visibilityOptions As Variant
    Dim optionIndex As Long
    Dim hasVisibleOption As Boolean
    originalValue = GetSettingText("ApiRetryCount", "1")
    originalManagedVisibility = GetSettingText("ManagedSheetVisibility", "VeryHidden")

    If SettingsCount() <= 0 Then
        Err.Raise vbObjectError + 7201, "modSettings.SettingsSmoke", "Settings sheet is empty."
    End If

    visibilityOptions = SettingValueOptions("ManagedSheetVisibility")
    If IsArray(visibilityOptions) Then
        For optionIndex = LBound(visibilityOptions) To UBound(visibilityOptions)
            If StrComp(CStr(visibilityOptions(optionIndex)), "Visible", vbTextCompare) = 0 Then
                hasVisibleOption = True
                Exit For
            End If
        Next optionIndex
    End If
    If Not hasVisibleOption Then
        Err.Raise vbObjectError + 7202, "modSettings.SettingsSmoke", "Managed sheet visibility options were not populated."
    End If

    SetSettingValue "ApiRetryCount", "2", "API通信再試行回数"
    If GetSettingText("ApiRetryCount", "") <> "2" Then
        Err.Raise vbObjectError + 7203, "modSettings.SettingsSmoke", "Setting update failed."
    End If

    SetSettingValue "ManagedSheetVisibility", "Visible", "管理シート表示状態"
    If GetSettingText("ManagedSheetVisibility", "") <> "Visible" Then
        Err.Raise vbObjectError + 7204, "modSettings.SettingsSmoke", "Managed sheet visibility was not saved."
    End If
    If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetVisible Then
        Err.Raise vbObjectError + 7205, "modSettings.SettingsSmoke", "Managed sheet visibility was not applied."
    End If

    ResetSettingsToDefaults
    If GetSettingText("ApiRetryCount", "") <> "1" Then
        Err.Raise vbObjectError + 7206, "modSettings.SettingsSmoke", "Settings reset failed."
    End If
    If GetSettingText("ManagedSheetVisibility", "") <> "VeryHidden" Then
        Err.Raise vbObjectError + 7207, "modSettings.SettingsSmoke", "Managed sheet visibility reset failed."
    End If
    If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetVeryHidden Then
        Err.Raise vbObjectError + 7208, "modSettings.SettingsSmoke", "Managed sheet visibility reset was not applied."
    End If

    RestoreSheetValues settingsSheet, snapshot
    modSettings.SetSettingValue "ManagedSheetVisibility", originalManagedVisibility, "管理シート表示状態"
    modAppStartup.SaveWorkbookState ThisWorkbook
    If GetSettingText("ApiRetryCount", "") <> originalValue Then
        Err.Raise vbObjectError + 7209, "modSettings.SettingsSmoke", "Settings restore failed."
    End If
    Select Case originalManagedVisibility
        Case "Visible"
            If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetVisible Then
                Err.Raise vbObjectError + 7210, "modSettings.SettingsSmoke", "Managed sheet visibility restore failed."
            End If
        Case "Hidden"
            If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetHidden Then
                Err.Raise vbObjectError + 7210, "modSettings.SettingsSmoke", "Managed sheet visibility restore failed."
            End If
        Case Else
            If ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName()).Visible <> xlSheetVeryHidden Then
                Err.Raise vbObjectError + 7210, "modSettings.SettingsSmoke", "Managed sheet visibility restore failed."
            End If
    End Select
    If modSettings.GetSettingText("ManagedSheetVisibility", "") <> originalManagedVisibility Then
        Err.Raise vbObjectError + 7211, "modSettings.SettingsSmoke", "Managed sheet visibility restore value failed."
    End If

    SettingsSmoke = "ok:" & CStr(SettingsCount())
    Exit Function

ErrHandler:
    On Error GoTo CleanupFail
    RestoreSheetValues settingsSheet, snapshot
    modSettings.SetSettingValue "ManagedSheetVisibility", originalManagedVisibility, "管理シート表示状態"
    modAppStartup.SaveWorkbookState ThisWorkbook
    GoTo CleanupDone

CleanupFail:
    Resume CleanupDone

CleanupDone:
    modLogger.LogErrorSafe "Settings", "modSettings.SettingsSmoke", Err.description, "", "", ""
    SettingsSmoke = "error:" & Err.description
End Function

Private Function FindSettingRow(ByVal settingsSheet As Worksheet, ByVal settingKey As String) As Long
    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(settingsSheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If StrComp(CStr(settingsSheet.cells(rowIndex, 1).Value2), settingKey, vbTextCompare) = 0 Then
            FindSettingRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function NormalizeSettingKey(ByVal settingKey As String) As String
    NormalizeSettingKey = Trim$(settingKey)
End Function

Private Function SettingRowText(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If rowIndex < 2 Then
        Exit Function
    End If
    If Not modSheetStore.SheetExists(ThisWorkbook, modSheetStore.SettingsSheetName()) Then
        Exit Function
    End If

    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(modSheetStore.SettingsSheetName())
    If rowIndex > modSheetStore.LastUsedRow(settingsSheet) Then
        Exit Function
    End If

    SettingRowText = CStr(settingsSheet.cells(rowIndex, columnIndex).Value2)
End Function

Private Function FormatSettingTimestamp(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    If IsDate(value) Then
        FormatSettingTimestamp = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss")
    ElseIf IsNumeric(value) Then
        FormatSettingTimestamp = Format$(CDate(CDbl(value)), "yyyy-mm-dd hh:nn:ss")
    Else
        FormatSettingTimestamp = CStr(value)
    End If
End Function

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

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function
