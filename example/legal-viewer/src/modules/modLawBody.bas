Attribute VB_Name = "modLawBody"
Option Explicit

Public Function GetLawDataJson(ByVal lawRevisionId As String, Optional ByVal lawId As String = "", Optional ByVal enforcementDate As String = "") As String
    On Error GoTo ErrHandler

    lawRevisionId = Trim$(lawRevisionId)
    If Len(lawRevisionId) = 0 Then
        Err.Raise vbObjectError + 6601, "modLawBody.GetLawDataJson", "LawRevisionId is required."
    End If

    GetLawDataJson = modApiClient.ApiGetLawData(lawRevisionId, lawId, enforcementDate)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawBody", "modLawBody.GetLawDataJson", Err.description, lawId, enforcementDate, lawRevisionId
    Err.Raise Err.Number, Err.source, Err.description
End Function

Public Function LawDataSmoke(ByVal lawRevisionId As String) As String
    On Error GoTo ErrHandler

    Dim responseJson As String
    responseJson = GetLawDataJson(lawRevisionId)

    Dim root As Object
    Set root = modJsonUtil.ParseJsonObject(responseJson)

    Dim fullText As Object
    Set fullText = modJsonUtil.JsonObjectProperty(root, "law_full_text")
    If fullText Is Nothing Then
        Err.Raise vbObjectError + 6602, "modLawBody.LawDataSmoke", "law_full_text is missing."
    End If

    Dim revisionInfo As Object
    Set revisionInfo = modJsonUtil.JsonObjectProperty(root, "revision_info")

    LawDataSmoke = "ok:" & modJsonUtil.JsonTextProperty(revisionInfo, "law_revision_id") & ":" & modJsonUtil.JsonTextProperty(fullText, "tag")
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "LawBody", "modLawBody.LawDataSmoke", Err.description, "", "", lawRevisionId
    LawDataSmoke = "error:" & Err.description
End Function
