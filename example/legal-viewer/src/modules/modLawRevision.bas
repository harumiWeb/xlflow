Attribute VB_Name = "modLawRevision"
Option Explicit

Private Const STATUS_OLD As String = "旧版"
Private Const STATUS_CURRENT As String = "現行"
Private Const STATUS_FUTURE As String = "未施行"
Private Const SHEET_REVISION_CANDIDATES As String = "_lv_tmp_law_revisions"

Public Function StatusOldText() As String
    StatusOldText = STATUS_OLD
End Function

Public Function StatusCurrentText() As String
    StatusCurrentText = STATUS_CURRENT
End Function

Public Function StatusFutureText() As String
    StatusFutureText = STATUS_FUTURE
End Function

Public Function NormalizeEnforcementDate(ByVal enforcementDateText As String) As String
    Dim parsedDate As Date
    If TryParseIsoDate(enforcementDateText, parsedDate) Then
        NormalizeEnforcementDate = Format$(parsedDate, "yyyy-mm-dd")
    Else
        NormalizeEnforcementDate = ""
    End If
End Function

Public Function CurrentEffectiveDateFromCsv(ByVal enforcementDateCsv As String, Optional ByVal baseDate As Date = 0) As String
    On Error GoTo ErrHandler

    Dim targetDate As Date
    targetDate = EffectiveBaseDate(baseDate)

    Dim parts As Variant
    parts = Split(enforcementDateCsv, ",")

    Dim bestDate As Date
    Dim hasBestDate As Boolean
    Dim index As Long
    For index = LBound(parts) To UBound(parts)
        Dim parsedDate As Date
        If TryParseIsoDate(Trim$(CStr(parts(index))), parsedDate) Then
            If parsedDate <= targetDate Then
                If Not hasBestDate Or parsedDate > bestDate Then
                    bestDate = parsedDate
                    hasBestDate = True
                End If
            End If
        End If
    Next index

    If hasBestDate Then
        CurrentEffectiveDateFromCsv = Format$(bestDate, "yyyy-mm-dd")
    Else
        CurrentEffectiveDateFromCsv = ""
    End If
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawRevision", "modLawRevision.CurrentEffectiveDateFromCsv", Err.description, "", "", enforcementDateCsv
    CurrentEffectiveDateFromCsv = ""
End Function

Public Function EnforcementStatus(ByVal enforcementDateText As String, ByVal currentEffectiveDateText As String, Optional ByVal baseDate As Date = 0) As String
    On Error GoTo ErrHandler

    Dim enforcementDate As Date
    If Not TryParseIsoDate(enforcementDateText, enforcementDate) Then
        EnforcementStatus = ""
        Exit Function
    End If

    Dim targetDate As Date
    targetDate = EffectiveBaseDate(baseDate)

    If enforcementDate > targetDate Then
        EnforcementStatus = STATUS_FUTURE
        Exit Function
    End If

    Dim currentDate As Date
    If TryParseIsoDate(currentEffectiveDateText, currentDate) Then
        If enforcementDate = currentDate Then
            EnforcementStatus = STATUS_CURRENT
        Else
            EnforcementStatus = STATUS_OLD
        End If
    Else
        EnforcementStatus = STATUS_OLD
    End If
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawRevision", "modLawRevision.EnforcementStatus", Err.description, "", enforcementDateText, currentEffectiveDateText
    EnforcementStatus = ""
End Function

Public Function EnforcementStatusFromCsv(ByVal enforcementDateText As String, ByVal enforcementDateCsv As String, Optional ByVal baseDate As Date = 0) As String
    Dim currentDateText As String
    currentDateText = CurrentEffectiveDateFromCsv(enforcementDateCsv, baseDate)
    EnforcementStatusFromCsv = EnforcementStatus(enforcementDateText, currentDateText, baseDate)
End Function

Public Function FormatEnforcementCandidate(ByVal enforcementDateText As String, ByVal enforcementDateCsv As String, Optional ByVal baseDate As Date = 0) As String
    Dim normalizedDate As String
    normalizedDate = NormalizeEnforcementDate(enforcementDateText)
    If Len(normalizedDate) = 0 Then
        FormatEnforcementCandidate = ""
        Exit Function
    End If

    FormatEnforcementCandidate = normalizedDate & "｜" & EnforcementStatusFromCsv(normalizedDate, enforcementDateCsv, baseDate)
End Function

Public Function IsKnownEnforcementStatus(ByVal statusText As String) As Boolean
    IsKnownEnforcementStatus = (StrComp(statusText, STATUS_OLD, vbTextCompare) = 0) _
        Or (StrComp(statusText, STATUS_CURRENT, vbTextCompare) = 0) _
        Or (StrComp(statusText, STATUS_FUTURE, vbTextCompare) = 0)
End Function

Public Function RefreshRevisionCandidates(ByVal lawId As String) As Long
    On Error GoTo ErrHandler

    lawId = Trim$(lawId)
    If Len(lawId) = 0 Then
        Err.Raise vbObjectError + 6411, "modLawRevision.RefreshRevisionCandidates", "LawId is required."
    End If

    Dim responseJson As String
    responseJson = modApiClient.ApiGetLawRevisions(lawId)

    Dim root As Object
    Set root = modJsonUtil.ParseJsonObject(responseJson)

    Dim lawInfo As Object
    Set lawInfo = modJsonUtil.JsonObjectProperty(root, "law_info")

    Dim revisions As Object
    Set revisions = modJsonUtil.JsonObjectProperty(root, "revisions")
    If revisions Is Nothing Then
        Err.Raise vbObjectError + 6412, "modLawRevision.RefreshRevisionCandidates", "API response does not contain revisions array."
    End If

    Dim targetSheet As Worksheet
    Set targetSheet = EnsureTempSheet(ThisWorkbook, SHEET_REVISION_CANDIDATES, RevisionCandidateHeaders())
    modSheetStore.ClearDataRows targetSheet

    Dim dateCsv As String
    dateCsv = RevisionDateCsv(revisions)

    Dim index As Long
    For index = 1 To revisions.Count
        AppendRevisionCandidateRow targetSheet, lawInfo, revisions(index), index, dateCsv
    Next index

    RefreshRevisionCandidates = revisions.Count
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawRevision", "modLawRevision.RefreshRevisionCandidates", Err.description, lawId, "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function RevisionCandidateCount() As Long
    If Not modSheetStore.SheetExists(ThisWorkbook, SHEET_REVISION_CANDIDATES) Then
        RevisionCandidateCount = 0
        Exit Function
    End If

    RevisionCandidateCount = Application.Max(0, modSheetStore.LastUsedRow(ThisWorkbook.Worksheets(SHEET_REVISION_CANDIDATES)) - 1)
End Function

Public Sub ConfigureRevisionCandidateComboBox(ByVal targetComboBox As Object)
    targetComboBox.columnCount = 4
    targetComboBox.columnWidths = "140 pt;0 pt;0 pt;0 pt"
    targetComboBox.BoundColumn = 2
End Sub

Public Function PopulateRevisionCandidateComboBox(ByVal lawId As String, ByVal targetComboBox As Object) As Long
    On Error GoTo ErrHandler

    ConfigureRevisionCandidateComboBox targetComboBox
    targetComboBox.Clear

    Dim candidateCount As Long
    candidateCount = RefreshRevisionCandidates(lawId)
    If candidateCount = 0 Then
        PopulateRevisionCandidateComboBox = 0
        Exit Function
    End If

    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Worksheets(SHEET_REVISION_CANDIDATES)

    Dim lastRow As Long
    lastRow = modSheetStore.LastUsedRow(targetSheet)

    Dim currentIndex As Long
    currentIndex = -1

    Dim rowIndex As Long
    Dim listIndex As Long
    For rowIndex = 2 To lastRow
        targetComboBox.AddItem CellText(targetSheet.cells(rowIndex, 12).Value2)
        listIndex = targetComboBox.ListCount - 1
        targetComboBox.List(listIndex, 1) = CellText(targetSheet.cells(rowIndex, 3).Value2)
        targetComboBox.List(listIndex, 2) = CellText(targetSheet.cells(rowIndex, 6).Value2)
        targetComboBox.List(listIndex, 3) = CellText(targetSheet.cells(rowIndex, 11).Value2)

        If currentIndex < 0 And StrComp(CellText(targetSheet.cells(rowIndex, 11).Value2), STATUS_CURRENT, vbTextCompare) = 0 Then
            currentIndex = listIndex
        End If
    Next rowIndex

    If targetComboBox.ListCount > 0 Then
        If currentIndex < 0 Then
            currentIndex = 0
        End If
        targetComboBox.listIndex = currentIndex
    End If

    PopulateRevisionCandidateComboBox = targetComboBox.ListCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawRevision", "modLawRevision.PopulateRevisionCandidateComboBox", Err.description, lawId, "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function RevisionCandidatesSmoke(ByVal lawId As String) As String
    On Error GoTo ErrHandler

    Dim candidateCount As Long
    candidateCount = RefreshRevisionCandidates(lawId)
    RevisionCandidatesSmoke = "ok:" & CStr(candidateCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawRevision", "modLawRevision.RevisionCandidatesSmoke", Err.description, lawId, "", ""
    RevisionCandidatesSmoke = "error:" & Err.description
End Function

Public Function LawRevisionSmoke() As String
    On Error GoTo ErrHandler

    Dim datesCsv As String
    datesCsv = "2024-04-01,2025-04-01,2026-04-01"

    Dim baseDate As Date
    baseDate = DateSerial(2025, 6, 13)

    If CurrentEffectiveDateFromCsv(datesCsv, baseDate) <> "2025-04-01" Then
        Err.Raise vbObjectError + 6401, "modLawRevision.LawRevisionSmoke", "Current effective date mismatch."
    End If

    If EnforcementStatusFromCsv("2024-04-01", datesCsv, baseDate) <> STATUS_OLD Then
        Err.Raise vbObjectError + 6402, "modLawRevision.LawRevisionSmoke", "Old status mismatch."
    End If

    If EnforcementStatusFromCsv("2025-04-01", datesCsv, baseDate) <> STATUS_CURRENT Then
        Err.Raise vbObjectError + 6403, "modLawRevision.LawRevisionSmoke", "Current status mismatch."
    End If

    If EnforcementStatusFromCsv("2026-04-01", datesCsv, baseDate) <> STATUS_FUTURE Then
        Err.Raise vbObjectError + 6404, "modLawRevision.LawRevisionSmoke", "Future status mismatch."
    End If

    LawRevisionSmoke = "ok"
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawRevision", "modLawRevision.LawRevisionSmoke", Err.description, "", "", ""
    LawRevisionSmoke = "error:" & Err.description
End Function

Private Function EffectiveBaseDate(ByVal baseDate As Date) As Date
    If baseDate = 0 Then
        EffectiveBaseDate = Date
    Else
        EffectiveBaseDate = baseDate
    End If
End Function

Private Sub AppendRevisionCandidateRow(ByVal targetSheet As Worksheet, ByVal lawInfo As Object, ByVal revisionInfo As Object, ByVal sourceIndex As Long, ByVal dateCsv As String)
    Dim enforcementDate As String
    enforcementDate = CandidateDate(revisionInfo)

    Dim statusText As String
    statusText = EnforcementStatusFromCsv(enforcementDate, dateCsv)

    Dim rowValues As Variant
    rowValues = Array( _
        Format$(Now, "yyyy-mm-dd hh:nn:ss"), _
        sourceIndex, _
        modJsonUtil.JsonTextProperty(revisionInfo, "law_revision_id"), _
        modJsonUtil.JsonTextProperty(lawInfo, "law_id"), _
        modJsonUtil.JsonTextProperty(lawInfo, "law_num"), _
        enforcementDate, _
        modJsonUtil.JsonTextProperty(revisionInfo, "law_title"), _
        modJsonUtil.JsonTextProperty(revisionInfo, "law_title_kana"), _
        modJsonUtil.JsonTextProperty(revisionInfo, "abbrev"), _
        modJsonUtil.JsonTextProperty(revisionInfo, "updated"), _
        statusText, _
        enforcementDate & "｜" & statusText, _
        modJsonUtil.JsonTextProperty(revisionInfo, "current_revision_status"), _
        modJsonUtil.JsonTextProperty(revisionInfo, "repeal_status"), _
        modJsonUtil.JsonTextProperty(revisionInfo, "amendment_promulgate_date"), _
        modJsonUtil.JsonTextProperty(revisionInfo, "amendment_scheduled_enforcement_date"))

    modSheetStore.AppendRow targetSheet, rowValues
End Sub

Private Function CandidateDate(ByVal revisionInfo As Object) As String
    CandidateDate = NormalizeEnforcementDate(modJsonUtil.JsonTextProperty(revisionInfo, "amendment_enforcement_date"))
    If Len(CandidateDate) = 0 Then
        CandidateDate = NormalizeEnforcementDate(modJsonUtil.JsonTextProperty(revisionInfo, "amendment_scheduled_enforcement_date"))
    End If
End Function

Private Function RevisionDateCsv(ByVal revisions As Object) As String
    Dim result As String
    Dim index As Long
    For index = 1 To revisions.Count
        Dim candidate As String
        candidate = CandidateDate(revisions(index))
        If Len(candidate) > 0 Then
            If Len(result) > 0 Then
                result = result & ","
            End If
            result = result & candidate
        End If
    Next index
    RevisionDateCsv = result
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

Private Sub WriteHeaders(ByVal targetSheet As Worksheet, ByVal headers As Variant)
    Dim index As Long
    For index = LBound(headers) To UBound(headers)
        targetSheet.cells(1, index - LBound(headers) + 1).Value2 = headers(index)
    Next index
    targetSheet.rows(1).Font.Bold = True
End Sub

Private Function RevisionCandidateHeaders() As Variant
    RevisionCandidateHeaders = Array( _
        "CachedAt", _
        "SourceIndex", _
        "LawRevisionId", _
        "LawId", _
        "LawNum", _
        "EnforcementDate", _
        "LawTitle", _
        "LawTitleKana", _
        "Abbrev", _
        "Updated", _
        "Status", _
        "DisplayText", _
        "CurrentRevisionStatus", _
        "RepealStatus", _
        "AmendmentPromulgateDate", _
        "ScheduledEnforcementDate")
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function

Private Function TryParseIsoDate(ByVal dateText As String, ByRef parsedDate As Date) As Boolean
    On Error GoTo ErrHandler

    dateText = Trim$(dateText)
    If Len(dateText) < 10 Then
        TryParseIsoDate = False
        Exit Function
    End If

    dateText = Left$(dateText, 10)
    If Mid$(dateText, 5, 1) <> "-" Or Mid$(dateText, 8, 1) <> "-" Then
        TryParseIsoDate = False
        Exit Function
    End If

    Dim yearValue As Long
    Dim monthValue As Long
    Dim dayValue As Long
    yearValue = CLng(Left$(dateText, 4))
    monthValue = CLng(Mid$(dateText, 6, 2))
    dayValue = CLng(Right$(dateText, 2))

    parsedDate = DateSerial(yearValue, monthValue, dayValue)
    TryParseIsoDate = (Year(parsedDate) = yearValue And Month(parsedDate) = monthValue And Day(parsedDate) = dayValue)
    Exit Function

ErrHandler:
    TryParseIsoDate = False
End Function
