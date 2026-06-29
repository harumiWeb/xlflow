Attribute VB_Name = "modLogger"
Option Explicit

Public Sub LogApiCall( _
    ByVal processKind As String, _
    ByVal endpoint As String, _
    ByVal httpStatus As Long, _
    ByVal succeeded As Boolean, _
    ByVal lawId As String, _
    ByVal enforcementDate As String, _
    ByVal itemCount As Long, _
    ByVal elapsedMs As Long, _
    ByVal ErrorMessage As String)

    On Error GoTo ErrHandler

    Dim logSheet As Worksheet
    Set logSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.ApiLogSheetName())

    Dim resultText As String
    If succeeded Then
        resultText = "成功"
    Else
        resultText = "失敗"
    End If

    modSheetStore.AppendRow logSheet, Array( _
        Format$(Now, "yyyy/mm/dd hh:nn:ss"), _
        processKind, _
        endpoint, _
        httpStatus, _
        resultText, _
        lawId, _
        enforcementDate, _
        itemCount, _
        elapsedMs, _
        ErrorMessage)

    TrimLogRows logSheet, modSettings.GetSettingLong("ApiLogLimit", 500)
    Exit Sub

ErrHandler:
    LogErrorSafe "Logger", "modLogger.LogApiCall", Err.description, lawId, enforcementDate, endpoint
End Sub

Public Function ApiLogCount() As Long
    ApiLogCount = LogSheetCount(modSheetStore.ApiLogSheetName())
End Function

Public Function ErrorLogCount() As Long
    ErrorLogCount = LogSheetCount(modSheetStore.ErrorLogSheetName())
End Function

Public Sub ConfigureApiLogListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 6
    targetListBox.columnWidths = "110 pt;84 pt;170 pt;52 pt;48 pt;0 pt"
    targetListBox.BoundColumn = 6
End Sub

Public Sub ConfigureErrorLogListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 6
    targetListBox.columnWidths = "110 pt;84 pt;170 pt;180 pt;0 pt;0 pt"
    targetListBox.BoundColumn = 6
End Sub

Public Function PopulateApiLogListBox(ByVal targetListBox As Object) As Long
    PopulateApiLogListBox = PopulateLogListBox(targetListBox, modSheetStore.ApiLogSheetName(), True)
End Function

Public Function PopulateErrorLogListBox(ByVal targetListBox As Object) As Long
    PopulateErrorLogListBox = PopulateLogListBox(targetListBox, modSheetStore.ErrorLogSheetName(), False)
End Function

Public Function ApiLogRowDetailText(ByVal rowIndex As Long) As String
    ApiLogRowDetailText = LogRowDetailText(modSheetStore.ApiLogSheetName(), rowIndex, True)
End Function

Public Function ErrorLogRowDetailText(ByVal rowIndex As Long) As String
    ErrorLogRowDetailText = LogRowDetailText(modSheetStore.ErrorLogSheetName(), rowIndex, False)
End Function

Public Sub ClearApiLog()
    ClearLogSheet modSheetStore.ApiLogSheetName()
    modAppStartup.SaveWorkbookState ThisWorkbook
End Sub

Public Sub ClearErrorLog()
    ClearLogSheet modSheetStore.ErrorLogSheetName()
    modAppStartup.SaveWorkbookState ThisWorkbook
End Sub

Public Sub ClearAllLogs()
    ClearApiLog
    ClearErrorLog
End Sub

Public Function LogsSmoke() As String
    On Error GoTo ErrHandler

    Dim apiSheet As Worksheet
    Dim errorSheet As Worksheet
    Set apiSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.ApiLogSheetName())
    Set errorSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.ErrorLogSheetName())

    Dim apiSnapshot As Variant
    Dim errorSnapshot As Variant
    apiSnapshot = SnapshotSheetValues(apiSheet)
    errorSnapshot = SnapshotSheetValues(errorSheet)

    Dim apiCount As Long
    Dim errorCount As Long
    apiCount = ApiLogCount()
    errorCount = ErrorLogCount()

    ClearAllLogs
    If ApiLogCount() <> 0 Or ErrorLogCount() <> 0 Then
        Err.Raise vbObjectError + 7301, "modLogger.LogsSmoke", "Log clear failed."
    End If

    RestoreSheetValues apiSheet, apiSnapshot
    RestoreSheetValues errorSheet, errorSnapshot
    modAppStartup.SaveWorkbookState ThisWorkbook

    If ApiLogCount() <> apiCount Or ErrorLogCount() <> errorCount Then
        Err.Raise vbObjectError + 7302, "modLogger.LogsSmoke", "Log restore failed."
    End If

    LogsSmoke = "ok:" & CStr(apiCount) & ":" & CStr(errorCount)
    Exit Function

ErrHandler:
    On Error GoTo CleanupFail
    RestoreSheetValues apiSheet, apiSnapshot
    RestoreSheetValues errorSheet, errorSnapshot
    modAppStartup.SaveWorkbookState ThisWorkbook
    GoTo CleanupDone

CleanupFail:
    Resume CleanupDone

CleanupDone:
    modLogger.LogErrorSafe "Logger", "modLogger.LogsSmoke", Err.description, "", "", ""
    LogsSmoke = "error:" & Err.description
End Function

Public Sub LogError( _
    ByVal processKind As String, _
    ByVal procedureName As String, _
    ByVal ErrorMessage As String, _
    ByVal lawId As String, _
    ByVal enforcementDate As String, _
    ByVal detail As String)

    On Error GoTo ErrHandler

    Dim logSheet As Worksheet
    Set logSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.ErrorLogSheetName())

    modSheetStore.AppendRow logSheet, Array( _
        Format$(Now, "yyyy/mm/dd hh:nn:ss"), _
        processKind, _
        procedureName, _
        ErrorMessage, _
        lawId, _
        enforcementDate, _
        detail)
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modLogger.LogError", Err.description
End Sub

Public Sub LogErrorSafe( _
    ByVal processKind As String, _
    ByVal procedureName As String, _
    ByVal ErrorMessage As String, _
    ByVal lawId As String, _
    ByVal enforcementDate As String, _
    ByVal detail As String)

    On Error GoTo FailSafe
    LogError processKind, procedureName, ErrorMessage, lawId, enforcementDate, detail
    Exit Sub

FailSafe:
End Sub

Private Sub TrimLogRows(ByVal logSheet As Worksheet, ByVal maxRows As Long)
    On Error GoTo ErrHandler

    If maxRows <= 0 Then
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(logSheet)

    Dim maxStoredLastRow As Long
    maxStoredLastRow = maxRows + 1
    If lastRow <= maxStoredLastRow Then
        Exit Sub
    End If

    Dim deleteUntilRow As Long
    deleteUntilRow = lastRow - maxRows
    If deleteUntilRow >= 2 Then
        logSheet.rows("2:" & CStr(deleteUntilRow)).Delete
    End If
    Exit Sub

ErrHandler:
    LogErrorSafe "Logger", "modLogger.TrimLogRows", Err.description, "", "", logSheet.Name
End Sub

Private Function PopulateLogListBox(ByVal targetListBox As Object, ByVal sheetName As String, ByVal apiMode As Boolean) As Long
    On Error GoTo ErrHandler

    If apiMode Then
        ConfigureApiLogListBox targetListBox
    Else
        ConfigureErrorLogListBox targetListBox
    End If
    targetListBox.Clear

    If Not modSheetStore.SheetExists(ThisWorkbook, sheetName) Then
        PopulateLogListBox = 0
        Exit Function
    End If

    Dim logSheet As Worksheet
    Set logSheet = ThisWorkbook.Worksheets(sheetName)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(logSheet)
    If lastRow < 2 Then
        PopulateLogListBox = 0
        Exit Function
    End If

    Dim rowIndex As Long
    Dim listIndex As Long
    For rowIndex = lastRow To 2 Step -1
        If apiMode Then
            targetListBox.AddItem LogTimestampText(logSheet.cells(rowIndex, 1).Value2)
            listIndex = targetListBox.ListCount - 1
            targetListBox.List(listIndex, 1) = LogCellText(logSheet.cells(rowIndex, 2).Value2)
            targetListBox.List(listIndex, 2) = LogCellText(logSheet.cells(rowIndex, 3).Value2)
            targetListBox.List(listIndex, 3) = LogCellText(logSheet.cells(rowIndex, 4).Value2) & " / " & LogCellText(logSheet.cells(rowIndex, 5).Value2)
            targetListBox.List(listIndex, 4) = LogCellText(logSheet.cells(rowIndex, 10).Value2)
            targetListBox.List(listIndex, 5) = CStr(rowIndex)
        Else
            targetListBox.AddItem LogTimestampText(logSheet.cells(rowIndex, 1).Value2)
            listIndex = targetListBox.ListCount - 1
            targetListBox.List(listIndex, 1) = LogCellText(logSheet.cells(rowIndex, 2).Value2)
            targetListBox.List(listIndex, 2) = LogCellText(logSheet.cells(rowIndex, 3).Value2)
            targetListBox.List(listIndex, 3) = LogCellText(logSheet.cells(rowIndex, 4).Value2)
            targetListBox.List(listIndex, 4) = LogCellText(logSheet.cells(rowIndex, 7).Value2)
            targetListBox.List(listIndex, 5) = CStr(rowIndex)
        End If
    Next rowIndex

    PopulateLogListBox = targetListBox.ListCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "Logger", "modLogger.PopulateLogListBox", Err.description, "", "", sheetName
    Err.Raise Err.Number, Err.source, Err.description
End Function

Private Function LogRowDetailText(ByVal sheetName As String, ByVal rowIndex As Long, ByVal apiMode As Boolean) As String
    If rowIndex < 2 Then Exit Function
    If Not modSheetStore.SheetExists(ThisWorkbook, sheetName) Then Exit Function

    Dim logSheet As Worksheet
    Set logSheet = ThisWorkbook.Worksheets(sheetName)
    If rowIndex > modSheetStore.LastUsedRow(logSheet) Then Exit Function

    If apiMode Then
        LogRowDetailText = "OccurredAt: " & LogTimestampText(logSheet.cells(rowIndex, 1).Value2) & vbCrLf & _
            "ProcessKind: " & LogCellText(logSheet.cells(rowIndex, 2).Value2) & vbCrLf & _
            "Endpoint: " & LogCellText(logSheet.cells(rowIndex, 3).Value2) & vbCrLf & _
            "HttpStatus: " & LogCellText(logSheet.cells(rowIndex, 4).Value2) & vbCrLf & _
            "Result: " & LogCellText(logSheet.cells(rowIndex, 5).Value2) & vbCrLf & _
            "LawId: " & LogCellText(logSheet.cells(rowIndex, 6).Value2) & vbCrLf & _
            "EnforcementDate: " & LogCellText(logSheet.cells(rowIndex, 7).Value2) & vbCrLf & _
            "ItemCount: " & LogCellText(logSheet.cells(rowIndex, 8).Value2) & vbCrLf & _
            "ElapsedMs: " & LogCellText(logSheet.cells(rowIndex, 9).Value2) & vbCrLf & _
            "ErrorMessage: " & LogCellText(logSheet.cells(rowIndex, 10).Value2)
    Else
        LogRowDetailText = "OccurredAt: " & LogTimestampText(logSheet.cells(rowIndex, 1).Value2) & vbCrLf & _
            "ProcessKind: " & LogCellText(logSheet.cells(rowIndex, 2).Value2) & vbCrLf & _
            "ProcedureName: " & LogCellText(logSheet.cells(rowIndex, 3).Value2) & vbCrLf & _
            "ErrorMessage: " & LogCellText(logSheet.cells(rowIndex, 4).Value2) & vbCrLf & _
            "LawId: " & LogCellText(logSheet.cells(rowIndex, 5).Value2) & vbCrLf & _
            "EnforcementDate: " & LogCellText(logSheet.cells(rowIndex, 6).Value2) & vbCrLf & _
            "Detail: " & LogCellText(logSheet.cells(rowIndex, 7).Value2)
    End If
End Function

Private Sub ClearLogSheet(ByVal sheetName As String)
    Dim logSheet As Worksheet
    Set logSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, sheetName)
    modSheetStore.ClearDataRows logSheet
End Sub

Private Function LogSheetCount(ByVal sheetName As String) As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, sheetName) Then Exit Function
    LogSheetCount = Application.Max(0, modSheetStore.LastUsedRow(ThisWorkbook.Worksheets(sheetName)) - 1)
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

Private Function LogCellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        LogCellText = ""
    Else
        LogCellText = CStr(value)
    End If
End Function

Private Function LogTimestampText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        LogTimestampText = ""
    ElseIf IsDate(value) Then
        LogTimestampText = Format$(CDate(value), "yyyy/mm/dd hh:nn:ss")
    ElseIf IsNumeric(value) Then
        LogTimestampText = Format$(CDate(CDbl(value)), "yyyy/mm/dd hh:nn:ss")
    Else
        LogTimestampText = CStr(value)
    End If
End Function
