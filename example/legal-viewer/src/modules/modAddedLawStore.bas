Attribute VB_Name = "modAddedLawStore"
Option Explicit

Private Const SHEET_ADDED_LAWS As String = "_lv_tmp_added_laws"

Public Function AddedLawStoreSheetName() As String
    AddedLawStoreSheetName = SHEET_ADDED_LAWS
End Function

Public Function EnsureAddedLawStore() As Worksheet
    Set EnsureAddedLawStore = EnsureTempSheet(ThisWorkbook, SHEET_ADDED_LAWS, AddedLawHeaders())
End Function

Public Sub ClearAddedLawStore()
    On Error GoTo ErrHandler

    Dim storeSheet As Worksheet
    Set storeSheet = EnsureAddedLawStore()
    modSheetStore.ClearDataRows storeSheet
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AddedLawStore", "modAddedLawStore.ClearAddedLawStore", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Function AddedLawCount() As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_ADDED_LAWS) Then
        AddedLawCount = 0
        Exit Function
    End If

    AddedLawCount = Application.Max(0, modSheetStore.LastUsedRow(ThisWorkbook.Worksheets(SHEET_ADDED_LAWS)) - 1)
End Function

Public Function FindAddedLawRow(ByVal lawId As String, ByVal enforcementDate As String) As Long
    On Error GoTo ErrHandler

    lawId = Trim$(lawId)
    enforcementDate = modLawRevision.NormalizeEnforcementDate(enforcementDate)
    If Len(lawId) = 0 Or Len(enforcementDate) = 0 Then
        FindAddedLawRow = 0
        Exit Function
    End If

    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_ADDED_LAWS) Then
        FindAddedLawRow = 0
        Exit Function
    End If

    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Worksheets(SHEET_ADDED_LAWS)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(storeSheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If StrComp(CellText(storeSheet.cells(rowIndex, 2).Value2), lawId, vbTextCompare) = 0 _
            And StrComp(NormalizedCellDateText(storeSheet.cells(rowIndex, 3).Value2), enforcementDate, vbTextCompare) = 0 Then
            FindAddedLawRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    FindAddedLawRow = 0
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AddedLawStore", "modAddedLawStore.FindAddedLawRow", Err.description, lawId, enforcementDate, ""
    FindAddedLawRow = 0
End Function

Public Function RegisterParsedLawMetadata( _
    ByVal lawId As String, _
    ByVal enforcementDate As String, _
    ByVal lawRevisionId As String, _
    ByVal lawTitle As String, _
    ByVal lawNum As String, _
    ByVal lawType As String, _
    ByVal promulgationDate As String, _
    ByVal statusText As String, _
    ByVal bodyCacheKey As String, _
    ByVal fetchedAt As String) As Long

    On Error GoTo ErrHandler

    lawId = Trim$(lawId)
    enforcementDate = modLawRevision.NormalizeEnforcementDate(enforcementDate)
    statusText = Trim$(statusText)

    If Len(lawId) = 0 Then
        Err.Raise vbObjectError + 6501, "modAddedLawStore.RegisterParsedLawMetadata", "LawId is required."
    End If
    If Len(enforcementDate) = 0 Then
        Err.Raise vbObjectError + 6502, "modAddedLawStore.RegisterParsedLawMetadata", "Valid enforcement date is required."
    End If
    If Not modLawRevision.IsKnownEnforcementStatus(statusText) Then
        Err.Raise vbObjectError + 6503, "modAddedLawStore.RegisterParsedLawMetadata", "Known enforcement status is required."
    End If
    If Len(Trim$(bodyCacheKey)) = 0 Then
        Err.Raise vbObjectError + 6504, "modAddedLawStore.RegisterParsedLawMetadata", "Body cache key is required."
    End If

    Dim storeSheet As Worksheet
    Set storeSheet = EnsureAddedLawStore()

    Dim existingRow As Long
    existingRow = FindAddedLawRow(lawId, enforcementDate)
    If existingRow > 0 Then
        SelectAddedLawRow existingRow
        RegisterParsedLawMetadata = existingRow
        Exit Function
    End If

    Dim addedAt As String
    addedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If Len(Trim$(fetchedAt)) = 0 Then
        fetchedAt = addedAt
    End If

    RegisterParsedLawMetadata = modSheetStore.AppendRow(storeSheet, Array( _
        addedAt, _
        lawId, _
        enforcementDate, _
        statusText, _
        lawRevisionId, _
        lawTitle, _
        lawNum, _
        lawType, _
        promulgationDate, _
        bodyCacheKey, _
        fetchedAt, _
        ""))
    SelectAddedLawRow RegisterParsedLawMetadata
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AddedLawStore", "modAddedLawStore.RegisterParsedLawMetadata", Err.description, lawId, enforcementDate, lawTitle
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Sub SelectAddedLawRow(ByVal targetRow As Long)
    On Error GoTo ErrHandler

    Dim storeSheet As Worksheet
    Set storeSheet = EnsureAddedLawStore()

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(storeSheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If rowIndex = targetRow Then
            storeSheet.cells(rowIndex, 12).Value2 = "1"
        Else
            storeSheet.cells(rowIndex, 12).Value2 = ""
        End If
    Next rowIndex
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AddedLawStore", "modAddedLawStore.SelectAddedLawRow", Err.description, "", "", CStr(targetRow)
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Function SelectedAddedLawRow() As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_ADDED_LAWS) Then
        SelectedAddedLawRow = 0
        Exit Function
    End If

    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Worksheets(SHEET_ADDED_LAWS)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(storeSheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If CStr(storeSheet.cells(rowIndex, 12).Value2) = "1" Then
            SelectedAddedLawRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Public Function AddedLawDisplayText(ByVal rowIndex As Long) As String
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_ADDED_LAWS) Then
        AddedLawDisplayText = ""
        Exit Function
    End If

    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Worksheets(SHEET_ADDED_LAWS)
    If rowIndex < 2 Or rowIndex > modSheetStore.LastUsedRow(storeSheet) Then
        AddedLawDisplayText = ""
        Exit Function
    End If

    AddedLawDisplayText = CellText(storeSheet.cells(rowIndex, 6).Value2) _
        & "｜" & CellText(storeSheet.cells(rowIndex, 4).Value2) _
        & "｜" & NormalizedCellDateText(storeSheet.cells(rowIndex, 3).Value2)
End Function

Public Function AddedLawDetailText(ByVal rowIndex As Long) As String
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_ADDED_LAWS) Then
        AddedLawDetailText = ""
        Exit Function
    End If

    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Worksheets(SHEET_ADDED_LAWS)
    If rowIndex < 2 Or rowIndex > modSheetStore.LastUsedRow(storeSheet) Then
        AddedLawDetailText = ""
        Exit Function
    End If

    Dim lawTitle As String
    Dim lawNum As String
    Dim enforcementDate As String
    Dim statusText As String
    Dim lawId As String
    lawTitle = CellText(storeSheet.cells(rowIndex, 6).Value2)
    lawNum = CellText(storeSheet.cells(rowIndex, 7).Value2)
    enforcementDate = NormalizedCellDateText(storeSheet.cells(rowIndex, 3).Value2)
    statusText = CellText(storeSheet.cells(rowIndex, 4).Value2)
    lawId = CellText(storeSheet.cells(rowIndex, 2).Value2)

    If Len(lawTitle) = 0 Then
        lawTitle = lawId
    End If

    Dim detailLine As String
    detailLine = lawNum
    If Len(enforcementDate) > 0 Then
        detailLine = AppendDetailPart(detailLine, enforcementDate)
    End If
    If Len(statusText) > 0 Then
        detailLine = AppendDetailPart(detailLine, statusText)
    End If
    If Len(detailLine) = 0 Then
        detailLine = lawId
    End If

    If Len(detailLine) = 0 Then
        AddedLawDetailText = lawTitle
    Else
        AddedLawDetailText = lawTitle & vbCrLf & detailLine
    End If
End Function

Public Sub ConfigureAddedLawListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 5
    targetListBox.columnWidths = "220 pt;55 pt;80 pt;0 pt;0 pt"
    targetListBox.BoundColumn = 4
End Sub

Public Function PopulateAddedLawListBox(ByVal targetListBox As Object) As Long
    On Error GoTo ErrHandler

    ConfigureAddedLawListBox targetListBox
    targetListBox.Clear

    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_ADDED_LAWS) Then
        PopulateAddedLawListBox = 0
        Exit Function
    End If

    Dim storeSheet As Worksheet
    Set storeSheet = ThisWorkbook.Worksheets(SHEET_ADDED_LAWS)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(storeSheet)

    Dim rowIndex As Long
    Dim listIndex As Long
    For rowIndex = 2 To lastRow
        targetListBox.AddItem CellText(storeSheet.cells(rowIndex, 6).Value2)
        listIndex = targetListBox.ListCount - 1
        targetListBox.List(listIndex, 1) = CellText(storeSheet.cells(rowIndex, 4).Value2)
        targetListBox.List(listIndex, 2) = NormalizedCellDateText(storeSheet.cells(rowIndex, 3).Value2)
        targetListBox.List(listIndex, 3) = CellText(storeSheet.cells(rowIndex, 2).Value2)
        targetListBox.List(listIndex, 4) = CStr(rowIndex)
    Next rowIndex

    PopulateAddedLawListBox = Application.Max(0, lastRow - 1)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AddedLawStore", "modAddedLawStore.PopulateAddedLawListBox", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function AddedLawStoreSmoke() As String
    On Error GoTo ErrHandler

    ClearAddedLawStore

    Dim firstRow As Long
    firstRow = RegisterParsedLawMetadata( _
        "TEST001", _
        "2025-04-01", _
        "TESTREV001", _
        "テスト法", _
        "令和七年法律第一号", _
        "Act", _
        "2025-01-01", _
        modLawRevision.StatusCurrentText(), _
        "_lv_tmp_body_TEST001_20250401", _
        "2026-06-13 00:00:00")

    Dim duplicateRow As Long
    duplicateRow = RegisterParsedLawMetadata( _
        "TEST001", _
        "2025-04-01", _
        "TESTREV001", _
        "テスト法", _
        "令和七年法律第一号", _
        "Act", _
        "2025-01-01", _
        modLawRevision.StatusCurrentText(), _
        "_lv_tmp_body_TEST001_20250401", _
        "2026-06-13 00:00:00")

    If firstRow <> 2 Or duplicateRow <> 2 Then
        Err.Raise vbObjectError + 6511, "modAddedLawStore.AddedLawStoreSmoke", "Duplicate detection failed."
    End If
    If AddedLawCount() <> 1 Then
        Err.Raise vbObjectError + 6512, "modAddedLawStore.AddedLawStoreSmoke", "Added law count mismatch."
    End If
    If SelectedAddedLawRow() <> 2 Then
        Err.Raise vbObjectError + 6513, "modAddedLawStore.AddedLawStoreSmoke", "Selection state mismatch."
    End If

    AddedLawStoreSmoke = "ok:1"
    ClearAddedLawStore
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AddedLawStore", "modAddedLawStore.AddedLawStoreSmoke", Err.description, "", "", ""
    AddedLawStoreSmoke = "error:" & Err.description
End Function

Public Function AddedLawDetailSmoke() As String
    On Error GoTo ErrHandler

    ClearAddedLawStore

    Dim firstRow As Long
    firstRow = RegisterParsedLawMetadata( _
        "TEST001", _
        "2025-04-01", _
        "TESTREV001", _
        "テスト法", _
        "令和七年法律第一号", _
        "Act", _
        "2025-01-01", _
        modLawRevision.StatusCurrentText(), _
        "_lv_tmp_body_TEST001_20250401", _
        "2026-06-13 00:00:00")

    Dim detailText As String
    detailText = AddedLawDetailText(firstRow)
    If InStr(1, detailText, "テスト法", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6521, "modAddedLawStore.AddedLawDetailSmoke", "Law title is missing."
    End If
    If InStr(1, detailText, "令和七年法律第一号", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6522, "modAddedLawStore.AddedLawDetailSmoke", "Law number is missing."
    End If
    If InStr(1, detailText, "2025-04-01", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6523, "modAddedLawStore.AddedLawDetailSmoke", "Enforcement date is missing."
    End If
    If InStr(1, detailText, modLawRevision.StatusCurrentText(), vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 6524, "modAddedLawStore.AddedLawDetailSmoke", "Status is missing."
    End If

    AddedLawDetailSmoke = "ok:" & CStr(Len(detailText))
    ClearAddedLawStore
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AddedLawStore", "modAddedLawStore.AddedLawDetailSmoke", Err.description, "", "", ""
    AddedLawDetailSmoke = "error:" & Err.description
End Function

Private Function EnsureTempSheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String, ByVal headers As Variant) As Worksheet
    Dim targetSheet As Worksheet
    If modSheetStore.SheetExists(targetWorkbook, sheetName) Then
        Set targetSheet = targetWorkbook.Worksheets(sheetName)
    Else
        Set targetSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count))
        targetSheet.Name = sheetName
    End If

    targetSheet.cells.NumberFormat = "@"
    If Application.CountA(targetSheet.rows(1)) = 0 Then
        WriteHeaders targetSheet, headers
    End If
    targetSheet.Visible = xlSheetVeryHidden
    Set EnsureTempSheet = targetSheet
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function

Private Function AppendDetailPart(ByVal baseText As String, ByVal partText As String) As String
    If Len(baseText) = 0 Then
        AppendDetailPart = partText
    ElseIf Len(partText) = 0 Then
        AppendDetailPart = baseText
    Else
        AppendDetailPart = baseText & " / " & partText
    End If
End Function

Private Function NormalizedCellDateText(ByVal value As Variant) As String
    If IsDate(value) Then
        NormalizedCellDateText = Format$(CDate(value), "yyyy-mm-dd")
    ElseIf IsNumeric(value) Then
        NormalizedCellDateText = Format$(CDate(CDbl(value)), "yyyy-mm-dd")
    Else
        NormalizedCellDateText = modLawRevision.NormalizeEnforcementDate(CellText(value))
    End If
End Function

Private Sub WriteHeaders(ByVal targetSheet As Worksheet, ByVal headers As Variant)
    Dim index As Long
    For index = LBound(headers) To UBound(headers)
        targetSheet.cells(1, index - LBound(headers) + 1).Value2 = headers(index)
    Next index
    targetSheet.rows(1).Font.Bold = True
End Sub

Private Function AddedLawHeaders() As Variant
    AddedLawHeaders = Array( _
        "AddedAt", _
        "LawId", _
        "EnforcementDate", _
        "Status", _
        "LawRevisionId", _
        "LawTitle", _
        "LawNum", _
        "LawType", _
        "PromulgationDate", _
        "BodyCacheKey", _
        "FetchedAt", _
        "Selected")
End Function
