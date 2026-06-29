Attribute VB_Name = "modUiDialog"
Option Explicit

Private Const DIALOG_BODY_FETCH_FAILED As String = "law-add-body-fetch-failed"

Public Sub NotifyBodyFetchFailed(ByVal lawId As String, ByVal enforcementDate As String, ByVal detail As String)
    On Error GoTo ErrHandler

    Dim prompt As String
    prompt = "法令本文の取得または解析に失敗しました。" & vbCrLf _
        & "法令ID: " & DisplayTextOrDash(lawId) & vbCrLf _
        & "施行日: " & DisplayTextOrDash(enforcementDate) & vbCrLf _
        & "詳細: " & DisplayTextOrDash(detail)

    XlflowUI.MsgBox DIALOG_BODY_FETCH_FAILED, prompt, vbOKOnly + vbExclamation, "法令追加", vbOK
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "UiDialog", "modUiDialog.NotifyBodyFetchFailed", Err.description, lawId, enforcementDate, detail
End Sub

Private Function DisplayTextOrDash(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) = 0 Then
        DisplayTextOrDash = "-"
    Else
        DisplayTextOrDash = value
    End If
End Function
