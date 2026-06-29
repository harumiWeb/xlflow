Attribute VB_Name = "modAliasMaster"
Option Explicit

Public Function AliasCount() As Long
    Dim aliasSheet As Worksheet
    Set aliasSheet = EnsureAliasSheet()
    AliasCount = Application.Max(0, modSheetStore.LastUsedRow(aliasSheet) - 1)
End Function

Public Sub ConfigureAliasListBox(ByVal targetListBox As Object)
    targetListBox.columnCount = 5
    targetListBox.columnWidths = "62 pt;140 pt;88 pt;90 pt;0 pt"
End Sub

Public Function PopulateAliasListBox(ByVal targetListBox As Object) As Long
    On Error GoTo ErrHandler

    ConfigureAliasListBox targetListBox
    targetListBox.Clear

    Dim aliasSheet As Worksheet
    Set aliasSheet = EnsureAliasSheet()

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(aliasSheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        targetListBox.AddItem CellText(aliasSheet.cells(rowIndex, 1).Value2)
        Dim listIndex As Long
        listIndex = targetListBox.ListCount - 1
        targetListBox.List(listIndex, 1) = CellText(aliasSheet.cells(rowIndex, 2).Value2)
        targetListBox.List(listIndex, 2) = CellText(aliasSheet.cells(rowIndex, 3).Value2)
        targetListBox.List(listIndex, 3) = FormatAliasTimestamp(aliasSheet.cells(rowIndex, 4).Value2)
        targetListBox.List(listIndex, 4) = CStr(rowIndex)
    Next rowIndex

    PopulateAliasListBox = targetListBox.ListCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AliasMaster", "modAliasMaster.PopulateAliasListBox", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function AliasRowAlias(ByVal rowIndex As Long) As String
    AliasRowAlias = AliasRowText(rowIndex, 1)
End Function

Public Function AliasRowLawTitle(ByVal rowIndex As Long) As String
    AliasRowLawTitle = AliasRowText(rowIndex, 2)
End Function

Public Function AliasRowNote(ByVal rowIndex As Long) As String
    AliasRowNote = AliasRowText(rowIndex, 3)
End Function

Public Function AliasRowUpdatedAt(ByVal rowIndex As Long) As String
    Dim aliasSheet As Worksheet
    Set aliasSheet = EnsureAliasSheet()
    If rowIndex < 2 Or rowIndex > modSheetStore.LastUsedRow(aliasSheet) Then Exit Function

    AliasRowUpdatedAt = FormatAliasTimestamp(aliasSheet.cells(rowIndex, 4).Value2)
End Function

Public Function SaveAlias( _
    ByVal aliasText As String, _
    ByVal lawTitle As String, _
    Optional ByVal note As String = "", _
    Optional ByVal targetRow As Long = 0) As Long

    On Error GoTo ErrHandler

    aliasText = Trim$(aliasText)
    lawTitle = Trim$(lawTitle)
    note = Trim$(note)

    If Len(aliasText) = 0 Then
        Err.Raise vbObjectError + 7951, "modAliasMaster.SaveAlias", "別名を入力してください。"
    End If
    If Len(lawTitle) = 0 Then
        Err.Raise vbObjectError + 7952, "modAliasMaster.SaveAlias", "法令名を入力してください。"
    End If

    Dim aliasSheet As Worksheet
    Set aliasSheet = EnsureAliasSheet()

    Dim rowIndex As Long
    If targetRow > 0 Then
        If targetRow < 2 Or targetRow > modSheetStore.LastUsedRow(aliasSheet) Then
            Err.Raise vbObjectError + 7957, "modAliasMaster.SaveAlias", "更新対象の別名が見つかりません。"
        End If

        Dim duplicateRow As Long
        duplicateRow = FindAliasRow(aliasSheet, aliasText)
        If duplicateRow > 0 And duplicateRow <> targetRow Then
            Err.Raise vbObjectError + 7958, "modAliasMaster.SaveAlias", "同じ別名が既に登録されています。"
        End If
        rowIndex = targetRow
    Else
        rowIndex = FindAliasRow(aliasSheet, aliasText)
    End If

    If rowIndex = 0 Then
        rowIndex = modSheetStore.AppendRow(aliasSheet, Array(aliasText, lawTitle, note, Format$(Now, "yyyy-mm-dd hh:nn:ss")))
    Else
        aliasSheet.cells(rowIndex, 1).Value2 = aliasText
        aliasSheet.cells(rowIndex, 2).Value2 = lawTitle
        aliasSheet.cells(rowIndex, 3).Value2 = note
        aliasSheet.cells(rowIndex, 4).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    End If

    modAppStartup.SaveWorkbookState ThisWorkbook
    SaveAlias = rowIndex
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "modAliasMaster.SaveAlias", Err.description
End Function

Public Sub DeleteAliasRow(ByVal rowIndex As Long)
    On Error GoTo ErrHandler

    Dim aliasSheet As Worksheet
    Set aliasSheet = EnsureAliasSheet()
    If rowIndex < 2 Or rowIndex > modSheetStore.LastUsedRow(aliasSheet) Then
        Err.Raise vbObjectError + 7953, "modAliasMaster.DeleteAliasRow", "削除対象の別名が見つかりません。"
    End If

    aliasSheet.rows(rowIndex).Delete
    modAppStartup.SaveWorkbookState ThisWorkbook
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modAliasMaster.DeleteAliasRow", Err.description
End Sub

Public Function AliasMasterSmoke() As String
    On Error GoTo ErrHandler

    Dim aliasSheet As Worksheet
    Set aliasSheet = EnsureAliasSheet()

    Dim snapshot As Variant
    snapshot = SnapshotSheetValues(aliasSheet)

    Dim smokeAlias As String
    smokeAlias = "__xlflow_alias_smoke__"

    Dim rowIndex As Long
    rowIndex = SaveAlias(smokeAlias, "試験法令", "追加")
    If AliasRowLawTitle(rowIndex) <> "試験法令" Then
        Err.Raise vbObjectError + 7954, "modAliasMaster.AliasMasterSmoke", "Alias insert failed."
    End If

    Dim countAfterInsert As Long
    countAfterInsert = AliasCount()

    Dim renamedAlias As String
    renamedAlias = smokeAlias & "_renamed"
    Dim updatedRow As Long
    updatedRow = SaveAlias(renamedAlias, "試験法令改", "更新", rowIndex)
    If updatedRow <> rowIndex Then
        Err.Raise vbObjectError + 7955, "modAliasMaster.AliasMasterSmoke", "Selected alias row changed during update."
    End If
    If AliasCount() <> countAfterInsert Then
        Err.Raise vbObjectError + 7955, "modAliasMaster.AliasMasterSmoke", "Selected alias update inserted a new row."
    End If
    If AliasRowAlias(rowIndex) <> renamedAlias Or AliasRowLawTitle(rowIndex) <> "試験法令改" Or AliasRowNote(rowIndex) <> "更新" Then
        Err.Raise vbObjectError + 7955, "modAliasMaster.AliasMasterSmoke", "Selected alias row overwrite failed."
    End If

    DeleteAliasRow rowIndex
    If FindAliasRow(aliasSheet, renamedAlias) <> 0 Then
        Err.Raise vbObjectError + 7956, "modAliasMaster.AliasMasterSmoke", "Alias delete failed."
    End If

    RestoreSheetValues aliasSheet, snapshot
    modAppStartup.SaveWorkbookState ThisWorkbook
    AliasMasterSmoke = "ok:" & CStr(AliasCount())
    Exit Function

ErrHandler:
    On Error GoTo CleanupFail
    RestoreSheetValues aliasSheet, snapshot
    modAppStartup.SaveWorkbookState ThisWorkbook
    GoTo CleanupDone

CleanupFail:
    Resume CleanupDone

CleanupDone:
    modLogger.LogErrorSafe "AliasMaster", "modAliasMaster.AliasMasterSmoke", Err.description, "", "", ""
    AliasMasterSmoke = "error:" & Err.description
End Function

Private Function EnsureAliasSheet() As Worksheet
    Set EnsureAliasSheet = modSheetStore.EnsurePersistentSheet(ThisWorkbook, modSheetStore.AliasMasterSheetName())
End Function

Private Function FindAliasRow(ByVal aliasSheet As Worksheet, ByVal aliasText As String) As Long
    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(aliasSheet)

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CellText(aliasSheet.cells(rowIndex, 1).Value2)), Trim$(aliasText), vbTextCompare) = 0 Then
            FindAliasRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function AliasRowText(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    Dim aliasSheet As Worksheet
    Set aliasSheet = EnsureAliasSheet()
    If rowIndex < 2 Or rowIndex > modSheetStore.LastUsedRow(aliasSheet) Then Exit Function

    AliasRowText = CellText(aliasSheet.cells(rowIndex, columnIndex).Value2)
End Function

Private Function FormatAliasTimestamp(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    If IsDate(value) Then
        FormatAliasTimestamp = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss")
    ElseIf IsNumeric(value) Then
        FormatAliasTimestamp = Format$(CDate(CDbl(value)), "yyyy-mm-dd hh:nn:ss")
    Else
        FormatAliasTimestamp = CStr(value)
    End If
End Function

Private Function SnapshotSheetValues(ByVal targetSheet As Worksheet) As Variant
    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(targetSheet)
    If lastRow <= 0 Then
        SnapshotSheetValues = Empty
        Exit Function
    End If

    Dim lastColumn As Long
    lastColumn = targetSheet.cells(1, targetSheet.Columns.Count).End(xlToLeft).Column
    SnapshotSheetValues = targetSheet.Range(targetSheet.cells(1, 1), targetSheet.cells(lastRow, lastColumn)).Value2
End Function

Private Sub RestoreSheetValues(ByVal targetSheet As Worksheet, ByVal snapshot As Variant)
    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(targetSheet)
    If lastRow > 0 Then
        targetSheet.rows("1:" & CStr(lastRow)).ClearContents
    End If

    If Not IsArray(snapshot) Then Exit Sub

    targetSheet.Range( _
        targetSheet.cells(1, 1), _
        targetSheet.cells(UBound(snapshot, 1), UBound(snapshot, 2))).Value2 = snapshot
End Sub

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function
