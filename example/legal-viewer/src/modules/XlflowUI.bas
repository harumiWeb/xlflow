Attribute VB_Name = "XlflowUI"
Option Explicit

Public Function MsgBox( _
    ByVal DialogId As String, _
    ByVal prompt As String, _
    Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional ByVal Title As String = "", _
    Optional ByVal DefaultResponse As VbMsgBoxResult = vbOK) As VbMsgBoxResult

    On Error GoTo ErrHandler

    Dim scriptedResponse As String
    scriptedResponse = ReadScriptedResponse("msgbox", DialogId)
    If Len(scriptedResponse) > 0 Then
        MsgBox = MsgBoxResultFromText(scriptedResponse, DefaultResponse)
        Exit Function
    End If

    If ShouldUseDefaultResponse() Then
        MsgBox = DefaultResponse
    Else
        MsgBox = VBA.Interaction.MsgBox(prompt, Buttons, Title)
    End If
    Exit Function

ErrHandler:
    MsgBox = DefaultResponse
End Function

Public Function ShowForm(ByVal targetForm As Object, Optional ByVal modal As Boolean = True) As Boolean
    On Error GoTo ErrHandler

    If ShouldUseDefaultResponse() Then
        ShowForm = False
        Exit Function
    End If

    If modal Then
        CallByName targetForm, "Show", VbMethod, vbModal
    Else
        CallByName targetForm, "Show", VbMethod, vbModeless
    End If
    ShowForm = True
    Exit Function

ErrHandler:
    ShowForm = False
End Function

Private Function ShouldUseDefaultResponse() As Boolean
    On Error GoTo ErrHandler

    ShouldUseDefaultResponse = (Not Application.Visible)
    Exit Function

ErrHandler:
    ShouldUseDefaultResponse = True
End Function

Private Function ReadScriptedResponse(ByVal kind As String, ByVal DialogId As String) As String
    Dim normalizedId As String
    normalizedId = NormalizeDialogId(DialogId)

    Dim candidateNames As Variant
    candidateNames = Array( _
        "__xlflow_ui_" & kind & "_" & normalizedId, _
        "_xlflow_ui_" & kind & "_" & normalizedId, _
        "xlflow_ui_" & kind & "_" & normalizedId)

    Dim workbookName As Name
    For Each workbookName In ThisWorkbook.Names
        Dim localName As String
        localName = LocalWorkbookName(workbookName.Name)

        Dim candidate As Variant
        For Each candidate In candidateNames
            If StrComp(localName, CStr(candidate), vbTextCompare) = 0 Then
                ReadScriptedResponse = Trim$(WorkbookNameText(workbookName))
                Exit Function
            End If
        Next candidate
    Next workbookName
End Function

Private Function LocalWorkbookName(ByVal fullName As String) As String
    Dim separatorIndex As Long
    separatorIndex = InStrRev(fullName, "!")
    If separatorIndex > 0 Then
        LocalWorkbookName = Mid$(fullName, separatorIndex + 1)
    Else
        LocalWorkbookName = fullName
    End If
End Function

Private Function WorkbookNameText(ByVal workbookName As Name) As String
    On Error GoTo FormulaValue

    WorkbookNameText = CStr(workbookName.RefersToRange.Value2)
    Exit Function

FormulaValue:
    On Error GoTo FailSafe

    Dim refersToText As String
    refersToText = workbookName.RefersTo
    If Left$(refersToText, 1) = "=" Then
        refersToText = Mid$(refersToText, 2)
    End If
    WorkbookNameText = UnquoteExcelString(refersToText)
    Exit Function

FailSafe:
    WorkbookNameText = ""
End Function

Private Function UnquoteExcelString(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) >= 2 And Left$(value, 1) = Chr$(34) And Right$(value, 1) = Chr$(34) Then
        value = Mid$(value, 2, Len(value) - 2)
        value = Replace(value, Chr$(34) & Chr$(34), Chr$(34))
    End If
    UnquoteExcelString = value
End Function

Private Function NormalizeDialogId(ByVal DialogId As String) As String
    Dim result As String
    Dim previousWasSeparator As Boolean

    Dim index As Long
    For index = 1 To Len(DialogId)
        Dim charCode As Long
        charCode = AscW(Mid$(DialogId, index, 1))

        If (charCode >= 48 And charCode <= 57) Or (charCode >= 97 And charCode <= 122) Then
            result = result & Chr$(charCode)
            previousWasSeparator = False
        ElseIf charCode >= 65 And charCode <= 90 Then
            result = result & Chr$(charCode + 32)
            previousWasSeparator = False
        ElseIf Len(result) > 0 And Not previousWasSeparator Then
            result = result & "_"
            previousWasSeparator = True
        End If
    Next index

    Do While Right$(result, 1) = "_"
        result = Left$(result, Len(result) - 1)
    Loop

    If Len(result) = 0 Then
        Err.Raise vbObjectError + 6901, "XlflowUI.NormalizeDialogId", "DialogId must contain an ASCII letter or digit."
    End If

    NormalizeDialogId = result
End Function

Private Function MsgBoxResultFromText(ByVal value As String, ByVal DefaultResponse As VbMsgBoxResult) As VbMsgBoxResult
    Select Case LCase$(Trim$(value))
        Case "abort"
            MsgBoxResultFromText = vbAbort
        Case "cancel"
            MsgBoxResultFromText = vbCancel
        Case "ignore"
            MsgBoxResultFromText = vbIgnore
        Case "no"
            MsgBoxResultFromText = vbNo
        Case "ok"
            MsgBoxResultFromText = vbOK
        Case "retry"
            MsgBoxResultFromText = vbRetry
        Case "yes"
            MsgBoxResultFromText = vbYes
        Case Else
            If IsNumeric(value) Then
                MsgBoxResultFromText = CLng(value)
            Else
                MsgBoxResultFromText = DefaultResponse
            End If
    End Select
End Function
