Attribute VB_Name = "modClipboard"
Option Explicit

Public Sub CopyTextToClipboard(ByVal value As String)
    On Error GoTo ErrHandler

    Dim dataObject As Object
    Set dataObject = CreateObject("Forms.DataObject.1")
    dataObject.SetText value
    dataObject.PutInClipboard
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "Clipboard", "modClipboard.CopyTextToClipboard", Err.description, "", "", ""
    Err.Raise Err.Number, Err.source, Err.description
End Sub
