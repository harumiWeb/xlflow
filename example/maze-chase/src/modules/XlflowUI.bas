Attribute VB_Name = "XlflowUI"
Option Explicit

' XlflowUI keeps one workbook-side dialog API that works in both interactive Excel and xlflow headless runs.
' Use stable dialog ids because xlflow maps scripted responses to those ids during unattended execution.
Private Const xlflowResponseErrorBase As Long = vbObjectError + 520
Private Const xlflowErrInvalidDialogId As Long = xlflowResponseErrorBase + 1
Private Const xlflowErrInvalidMsgBoxResponse As Long = xlflowResponseErrorBase + 2
Private Const xlflowErrMissingInputResponse As Long = xlflowResponseErrorBase + 3
Private Const xlflowErrMissingFileDialogResponse As Long = xlflowResponseErrorBase + 4
Private Const xlflowErrInvalidFileDialogResponse As Long = xlflowResponseErrorBase + 5
Private Const xlflowFileDialogCancelToken As String = "@cancel"
Private Const xlflowStreamHelperName As String = "__XLFLOW_UI_STREAM_HELPER__"
Private Const xlflowStreamRedactInputName As String = "__XLFLOW_UI_STREAM_REDACT_INPUT__"

' Wrapper for VBA.MsgBox. In headless-like modes xlflow resolves the response from --msgbox.
Public Function MsgBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal DefaultResponse As String = "") As VbMsgBoxResult
    Dim ResponseSource As String
    Dim ResponseToken As String
    ValidateDialogId Id, "XlflowUI.MsgBox"
    If XlflowRuntime.IsHeadless() Then
        On Error GoTo HandleHeadlessFailure
        MsgBox = ResolveMsgBoxResponse(Id, DefaultResponse, ResponseSource, ResponseToken)
        EmitHeadlessUIEvent "msgbox", Id, Prompt, Title, ResponseSource, ResponseToken, "", False, ""
        Exit Function
    End If

    MsgBox = VBA.Interaction.MsgBox(Prompt, Buttons, Title)
    Exit Function

    HandleHeadlessFailure:
    EmitHeadlessUIEvent "msgbox", Id, Prompt, Title, "error", "", "", False, Err.Description
    Err.Raise Err.Number, Err.source, Err.Description
End Function

' Wrapper for VBA.InputBox. Default is interactive-only UI text, while DefaultValue is the headless fallback.
Public Function InputBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "", Optional ByVal DefaultValue As String = "") As String
    Dim ResponseSource As String
    Dim responseValue As String
    Dim DisplayValue As String
    Dim Redacted As Boolean
    ValidateDialogId Id, "XlflowUI.InputBox"
    If XlflowRuntime.IsHeadless() Then
        On Error GoTo HandleHeadlessFailure
        responseValue = ResolveInputResponse(Id, DefaultValue, ResponseSource)
        Redacted = ShouldRedactInputStream()
        DisplayValue = responseValue
        If Redacted Then
            DisplayValue = "[redacted]"
        End If
        EmitHeadlessUIEvent "inputbox", Id, Prompt, Title, ResponseSource, "", DisplayValue, Redacted, ""
        InputBox = responseValue
        Exit Function
    End If

    InputBox = VBA.Interaction.InputBox(Prompt, Title, Default)
    Exit Function

    HandleHeadlessFailure:
    Redacted = ShouldRedactInputStream()
    DisplayValue = ""
    If Redacted Then
        DisplayValue = "[redacted]"
    End If
    EmitHeadlessUIEvent "inputbox", Id, Prompt, Title, "error", "", DisplayValue, Redacted, Err.Description
    Err.Raise Err.Number, Err.source, Err.Description
End Function

' Wrapper for Application.GetOpenFilename.
' In headless mode, repeat --filedialog get-open:<id>=<path> to simulate multi-select results.
Public Function GetOpenFilename(ByVal Id As String, Optional ByVal FileFilter As String = "", Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal DefaultValue As Variant) As Variant
    Dim ResponseSource As String
    Dim DisplayValue As String
    ValidateDialogId Id, "XlflowUI.GetOpenFilename"
    If XlflowRuntime.IsHeadless() Then
        On Error GoTo HandleHeadlessFailure
        GetOpenFilename = ResolveFileDialogResponse("get-open", Id, MultiSelect, DefaultValue, ResponseSource, DisplayValue)
        EmitHeadlessUIEvent "get-open", Id, "", Title, ResponseSource, "", DisplayValue, False, ""
        Exit Function
    End If

    GetOpenFilename = Application.GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)
    Exit Function

    HandleHeadlessFailure:
    EmitHeadlessUIEvent "get-open", Id, "", Title, "error", "", "", False, Err.Description
    Err.Raise Err.Number, Err.source, Err.Description
End Function

' Wrapper for Application.FileDialog(msoFileDialogOpen).
' In headless mode, this uses the same --filedialog file-open:<id>=<path> contract as other xlflow runs.
Public Function FileDialogOpen(ByVal Id As String, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal DefaultValue As Variant) As Variant
    Dim ResponseSource As String
    Dim DisplayValue As String
    Dim dialog As FileDialog
    ValidateDialogId Id, "XlflowUI.FileDialogOpen"
    If XlflowRuntime.IsHeadless() Then
        On Error GoTo HandleHeadlessFailure
        FileDialogOpen = ResolveFileDialogResponse("file-open", Id, MultiSelect, DefaultValue, ResponseSource, DisplayValue)
        EmitHeadlessUIEvent "file-open", Id, "", Title, ResponseSource, "", DisplayValue, False, ""
        Exit Function
    End If

    Set dialog = Application.FileDialog(msoFileDialogOpen)
    With dialog
        If Len(Title) > 0 Then .Title = Title
        If Len(ButtonText) > 0 Then .ButtonName = ButtonText
        .AllowMultiSelect = MultiSelect
        If .Show <> -1 Then
            FileDialogOpen = False
        ElseIf MultiSelect Then
            FileDialogOpen = SelectedItemsToVariantArray(.SelectedItems)
        Else
            FileDialogOpen = CStr(.SelectedItems(1))
        End If
    End With
    Exit Function

    HandleHeadlessFailure:
    EmitHeadlessUIEvent "file-open", Id, "", Title, "error", "", "", False, Err.Description
    Err.Raise Err.Number, Err.source, Err.Description
End Function

' Wrapper for Application.GetSaveAsFilename.
' Use --filedialog save-as:<id>=<path> in unattended runs, or @cancel to simulate Cancel.
Public Function GetSaveAsFilename(ByVal Id As String, Optional ByVal InitialFileName As String = "", Optional ByVal FileFilter As String = "", Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal DefaultValue As Variant) As Variant
    Dim ResponseSource As String
    Dim DisplayValue As String
    ValidateDialogId Id, "XlflowUI.GetSaveAsFilename"
    If XlflowRuntime.IsHeadless() Then
        On Error GoTo HandleHeadlessFailure
        GetSaveAsFilename = ResolveFileDialogResponse("save-as", Id, False, DefaultValue, ResponseSource, DisplayValue)
        EmitHeadlessUIEvent "save-as", Id, "", Title, ResponseSource, "", DisplayValue, False, ""
        Exit Function
    End If

    GetSaveAsFilename = Application.GetSaveAsFilename(InitialFileName, FileFilter, FilterIndex, Title, ButtonText)
    Exit Function

    HandleHeadlessFailure:
    EmitHeadlessUIEvent "save-as", Id, "", Title, "error", "", "", False, Err.Description
    Err.Raise Err.Number, Err.source, Err.Description
End Function

' Wrapper for Application.FileDialog(msoFileDialogFolderPicker).
' Use --filedialog folder:<id>=<path> in unattended runs, or @cancel to simulate Cancel.
Public Function FolderPicker(ByVal Id As String, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal InitialPath As String = "", Optional ByVal DefaultValue As Variant) As Variant
    Dim ResponseSource As String
    Dim DisplayValue As String
    Dim dialog As FileDialog
    ValidateDialogId Id, "XlflowUI.FolderPicker"
    If XlflowRuntime.IsHeadless() Then
        On Error GoTo HandleHeadlessFailure
        FolderPicker = ResolveFileDialogResponse("folder", Id, False, DefaultValue, ResponseSource, DisplayValue)
        EmitHeadlessUIEvent "folder", Id, "", Title, ResponseSource, "", DisplayValue, False, ""
        Exit Function
    End If

    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    With dialog
        If Len(Title) > 0 Then .Title = Title
        If Len(ButtonText) > 0 Then .ButtonName = ButtonText
        If Len(InitialPath) > 0 Then .InitialFileName = InitialPath
        If .Show <> -1 Then
            FolderPicker = False
        Else
            FolderPicker = CStr(.SelectedItems(1))
        End If
    End With
    Exit Function

    HandleHeadlessFailure:
    EmitHeadlessUIEvent "folder", Id, "", Title, "error", "", "", False, Err.Description
    Err.Raise Err.Number, Err.source, Err.Description
End Function

' Resolves a headless MsgBox response from xlflow markers or workbook defaults.
Private Function ResolveMsgBoxResponse(ByVal Id As String, Optional ByVal DefaultResponse As String = "", Optional ByRef ResponseSource As String = "", Optional ByRef ResponseToken As String = "") As VbMsgBoxResult
    Dim response As String
    On Error GoTo UseDefault
    response = LCase$(Trim$(ReadResponseValue("msgbox", Id)))
    ResponseSource = "scripted"
    GoTo Resolve

    UseDefault:
    If Len(Trim$(DefaultResponse)) = 0 Then
        ResponseSource = "error"
        Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Missing scripted MsgBox response for dialog id '" & Id & "'."
    End If
    response = LCase$(Trim$(DefaultResponse))
    ResponseSource = "default"

    Resolve:
    ResponseToken = response

    Select Case response
    Case "abort"
        ResolveMsgBoxResponse = vbAbort
    Case "cancel"
        ResolveMsgBoxResponse = vbCancel
    Case "ignore"
        ResolveMsgBoxResponse = vbIgnore
    Case "no"
        ResolveMsgBoxResponse = vbNo
    Case "ok"
        ResolveMsgBoxResponse = vbOK
    Case "retry"
        ResolveMsgBoxResponse = vbRetry
    Case "yes"
        ResolveMsgBoxResponse = vbYes
    Case Else
        If ResponseSource = "default" Then
            Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Invalid default MsgBox response for dialog id '" & Id & "'."
        End If
        ResponseSource = "error"
        Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Invalid scripted MsgBox response for dialog id '" & Id & "'."
    End Select
End Function

Private Function ResolveInputResponse(ByVal Id As String, Optional ByVal DefaultValue As String = "", Optional ByRef ResponseSource As String = "") As String
    On Error GoTo UseDefault
    ResponseSource = "scripted"
    ResolveInputResponse = ReadResponseValue("input", Id)
    Exit Function

    UseDefault:
    If Len(DefaultValue) = 0 Then
        ResponseSource = "error"
        Err.Raise xlflowErrMissingInputResponse, "XlflowUI.InputBox", "Missing scripted InputBox response for dialog id '" & Id & "'."
    End If
    ResponseSource = "default"
    ResolveInputResponse = DefaultValue
End Function

' Resolves file dialog responses from injected markers.
' Cancel is represented as False, and multi-select responses return a Variant string array.
Private Function ResolveFileDialogResponse(ByVal Kind As String, ByVal Id As String, ByVal MultiSelect As Boolean, Optional ByVal DefaultValue As Variant, Optional ByRef ResponseSource As String = "", Optional ByRef DisplayValue As String = "") As Variant
    Dim MarkerValue As String

    On Error GoTo UseDefault
    MarkerValue = ReadFileDialogResponseValue(Kind, Id)
    On Error GoTo InvalidScripted
    ResponseSource = "scripted"
    ResolveFileDialogResponse = NormalizeFileDialogResponse(Kind, Id, MultiSelect, MarkerValue, DisplayValue)
    Exit Function

    UseDefault:
    If IsEmpty(DefaultValue) Then
        ResponseSource = "error"
        Err.Raise xlflowErrMissingFileDialogResponse, FileDialogSourceName(Kind), "Missing scripted file dialog response for dialog id '" & Id & "'."
    End If
    ResponseSource = "default"
    ResolveFileDialogResponse = NormalizeFileDialogDefault(Kind, Id, MultiSelect, DefaultValue, DisplayValue)
    Exit Function

    InvalidScripted:
    ResponseSource = "error"
    Err.Raise Err.Number, Err.source, Err.Description
End Function

' Converts injected newline-delimited marker values into the same Variant shape Excel would normally return.
Private Function NormalizeFileDialogResponse(ByVal Kind As String, ByVal Id As String, ByVal MultiSelect As Boolean, ByVal MarkerValue As String, Optional ByRef DisplayValue As String = "") As Variant
    Dim values As Variant
    Dim count As Long

    If LCase$(Trim$(MarkerValue)) = xlflowFileDialogCancelToken Then
        DisplayValue = xlflowFileDialogCancelToken
        NormalizeFileDialogResponse = False
        Exit Function
    End If

    values = SplitFileDialogValues(MarkerValue)
    count = FileDialogValueCount(values)
    If MultiSelect Then
        DisplayValue = JoinStringArray(values, " | ")
        NormalizeFileDialogResponse = values
        Exit Function
    End If
    If count <> 1 Then
        Err.Raise xlflowErrInvalidFileDialogResponse, FileDialogSourceName(Kind), "File dialog '" & Id & "' expected one scripted path but received " & CStr(count) & "."
    End If
    DisplayValue = CStr(values(LBound(values)))
    NormalizeFileDialogResponse = CStr(values(LBound(values)))
End Function

Private Function NormalizeFileDialogDefault(ByVal Kind As String, ByVal Id As String, ByVal MultiSelect As Boolean, ByVal DefaultValue As Variant, Optional ByRef DisplayValue As String = "") As Variant
    Dim values As Variant
    Dim count As Long

    If VarType(DefaultValue) = vbBoolean Then
        If CBool(DefaultValue) = False Then
            DisplayValue = xlflowFileDialogCancelToken
            NormalizeFileDialogDefault = False
            Exit Function
        End If
    End If

    If IsArray(DefaultValue) Then
        values = VariantArrayToStringArray(DefaultValue)
        count = FileDialogValueCount(values)
        If MultiSelect Then
            DisplayValue = JoinStringArray(values, " | ")
            NormalizeFileDialogDefault = values
            Exit Function
        End If
        If count <> 1 Then
            Err.Raise xlflowErrInvalidFileDialogResponse, FileDialogSourceName(Kind), "File dialog '" & Id & "' default value must contain exactly one path."
        End If
        DisplayValue = CStr(values(LBound(values)))
        NormalizeFileDialogDefault = CStr(values(LBound(values)))
        Exit Function
    End If

    If MultiSelect Then
        values = Array(CStr(DefaultValue))
        DisplayValue = JoinStringArray(values, " | ")
        NormalizeFileDialogDefault = values
        Exit Function
    End If

    DisplayValue = CStr(DefaultValue)
    NormalizeFileDialogDefault = DisplayValue
End Function

' Reads the workbook-defined name that xlflow injected for one file dialog wrapper call.
Private Function ReadFileDialogResponseValue(ByVal Kind As String, ByVal Id As String) As String
    On Error GoTo Missing
    ReadFileDialogResponseValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(BuildFileDialogResponseName(Kind, Id)).refersTo)
    Exit Function

    Missing:
    Err.Raise xlflowErrMissingFileDialogResponse, FileDialogSourceName(Kind), "Missing scripted file dialog response for dialog id '" & Id & "'."
End Function

Private Function BuildFileDialogResponseName(ByVal Kind As String, ByVal Id As String) As String
    Dim normalizedKind As String

    normalizedKind = LCase$(Trim$(Kind))
    Select Case normalizedKind
    Case "get-open"
        BuildFileDialogResponseName = "__XLFLOW_UI_FILEDIALOG_GET_OPEN_" & NormalizeResponseId(Id) & "__"
    Case "file-open"
        BuildFileDialogResponseName = "__XLFLOW_UI_FILEDIALOG_FILE_OPEN_" & NormalizeResponseId(Id) & "__"
    Case "save-as"
        BuildFileDialogResponseName = "__XLFLOW_UI_FILEDIALOG_SAVE_AS_" & NormalizeResponseId(Id) & "__"
    Case "folder"
        BuildFileDialogResponseName = "__XLFLOW_UI_FILEDIALOG_FOLDER_" & NormalizeResponseId(Id) & "__"
    Case Else
        Err.Raise xlflowErrInvalidFileDialogResponse, FileDialogSourceName(Kind), "Unsupported file dialog kind '" & Kind & "'."
    End Select
End Function

Private Function FileDialogSourceName(ByVal Kind As String) As String
    Select Case LCase$(Trim$(Kind))
    Case "get-open"
        FileDialogSourceName = "XlflowUI.GetOpenFilename"
    Case "file-open"
        FileDialogSourceName = "XlflowUI.FileDialogOpen"
    Case "save-as"
        FileDialogSourceName = "XlflowUI.GetSaveAsFilename"
    Case "folder"
        FileDialogSourceName = "XlflowUI.FolderPicker"
    Case Else
        FileDialogSourceName = "XlflowUI.FileDialog"
    End Select
End Function

' Headless multi-select responses are stored as newline-delimited paths.
Private Function SplitFileDialogValues(ByVal MarkerValue As String) As Variant
    Dim normalizedValue As String
    Dim values As Variant
    Dim i As Long

    normalizedValue = Replace$(MarkerValue, vbCrLf, vbLf)
    normalizedValue = Replace$(normalizedValue, vbCr, vbLf)
    If Len(normalizedValue) = 0 Then
        Err.Raise xlflowErrInvalidFileDialogResponse, "XlflowUI.FileDialog", "Scripted file dialog response must contain at least one path."
    End If
    values = Split(normalizedValue, vbLf)
    For i = LBound(values) To UBound(values)
        If Len(CStr(values(i))) = 0 Then
            Err.Raise xlflowErrInvalidFileDialogResponse, "XlflowUI.FileDialog", "Scripted file dialog response contains an empty path entry."
        End If
    Next i
    SplitFileDialogValues = values
End Function

Private Function VariantArrayToStringArray(ByVal values As Variant) As Variant
    Dim result() As String
    Dim i As Long

    ReDim result(LBound(values) To UBound(values))
    For i = LBound(values) To UBound(values)
        result(i) = CStr(values(i))
        If Len(result(i)) = 0 Then
            Err.Raise xlflowErrInvalidFileDialogResponse, "XlflowUI.FileDialog", "File dialog default values cannot contain empty paths."
        End If
    Next i
    VariantArrayToStringArray = result
End Function

Private Function FileDialogValueCount(ByVal values As Variant) As Long
    FileDialogValueCount = UBound(values) - LBound(values) + 1
End Function

Private Function JoinStringArray(ByVal values As Variant, ByVal Separator As String) As String
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If i > LBound(values) Then
            JoinStringArray = JoinStringArray & Separator
        End If
        JoinStringArray = JoinStringArray & CStr(values(i))
    Next i
End Function

Private Function SelectedItemsToVariantArray(ByVal SelectedItems As Office.FileDialogSelectedItems) As Variant
    Dim values() As String
    Dim i As Long

    ReDim values(0 To SelectedItems.count - 1)
    For i = 1 To SelectedItems.count
        values(i - 1) = CStr(SelectedItems(i))
    Next i
    SelectedItemsToVariantArray = values
End Function

Private Sub EmitHeadlessUIEvent(ByVal Kind As String, ByVal Id As String, ByVal Prompt As String, ByVal Title As String, ByVal ResponseSource As String, Optional ByVal ResolvedResult As String = "", Optional ByVal ResolvedValue As String = "", Optional ByVal Redacted As Boolean = False, Optional ByVal FailureMessage As String = "")
    Dim helperMacro As String
    Dim payload As String

    helperMacro = ResolveStreamHelperMacro()
    If Len(helperMacro) = 0 Then
        Exit Sub
    End If

    payload = "{" & _
    JsonProperty("event", "ui_dialog") & "," & _
    JsonProperty("kind", Kind) & "," & _
    JsonProperty("dialog_id", Id) & "," & _
    JsonProperty("prompt", Prompt) & "," & _
    JsonProperty("title", Title) & "," & _
    JsonProperty("response_source", ResponseSource) & "," & _
    JsonProperty("resolved_result", ResolvedResult) & "," & _
    JsonProperty("resolved_value", ResolvedValue) & "," & _
    JsonBooleanProperty("redacted", Redacted) & "," & _
    JsonProperty("runtime_mode", XlflowRuntime.ModeName()) & "," & _
    JsonProperty("error", FailureMessage) & "}"

    InvokeStreamHelper helperMacro, payload
End Sub

Private Sub InvokeStreamHelper(ByVal helperMacro As String, ByVal payload As String)
    On Error GoTo Swallow
    Application.Run "'" & ThisWorkbook.name & "'!" & helperMacro, payload
    Exit Sub

    Swallow:
End Sub

Private Function ResolveStreamHelperMacro() As String
    ResolveStreamHelperMacro = ReadOptionalDefinedNameValue(xlflowStreamHelperName)
End Function

Private Function ShouldRedactInputStream() As Boolean
    Dim configured As String

    configured = LCase$(Trim$(ReadOptionalDefinedNameValue(xlflowStreamRedactInputName)))
    ShouldRedactInputStream = (configured <> "false")
End Function

Private Function ReadOptionalDefinedNameValue(ByVal name As String) As String
    On Error GoTo Missing
    ReadOptionalDefinedNameValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(name).refersTo)
    Exit Function

    Missing:
    ReadOptionalDefinedNameValue = ""
End Function

Private Function JsonEscape(ByVal value As String) As String
    value = Replace$(value, "\", "\\")
    value = Replace$(value, Chr$(34), Chr$(92) & Chr$(34))
    value = Replace$(value, vbCrLf, "\n")
    value = Replace$(value, vbCr, "\n")
    value = Replace$(value, vbLf, "\n")
    value = Replace$(value, vbTab, "\t")
    JsonEscape = value
End Function

Private Function JsonProperty(ByVal name As String, ByVal value As String) As String
    JsonProperty = Chr$(34) & JsonEscape(name) & Chr$(34) & ":" & Chr$(34) & JsonEscape(value) & Chr$(34)
End Function

Private Function JsonBooleanProperty(ByVal name As String, ByVal value As Boolean) As String
    JsonBooleanProperty = Chr$(34) & JsonEscape(name) & Chr$(34) & ":" & LCase$(CStr(value))
End Function

Private Function ReadResponseValue(ByVal Kind As String, ByVal Id As String) As String
    On Error GoTo Missing
    ReadResponseValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(BuildResponseName(Kind, Id)).refersTo)
    Exit Function

    Missing:
    If LCase$(Trim$(Kind)) = "msgbox" Then
        Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Missing scripted MsgBox response for dialog id '" & Id & "'."
    End If
    Err.Raise xlflowErrMissingInputResponse, "XlflowUI.InputBox", "Missing scripted InputBox response for dialog id '" & Id & "'."
End Function

Private Function BuildResponseName(ByVal Kind As String, ByVal Id As String) As String
    BuildResponseName = "__XLFLOW_UI_" & UCase$(Trim$(Kind)) & "_" & NormalizeResponseId(Id) & "__"
End Function

Private Function NormalizeResponseId(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim normalized As String
    Dim lastSeparator As Boolean

    value = LCase$(Trim$(value))
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch Like "[a-z0-9]" Then
            normalized = normalized & ch
            lastSeparator = False
        ElseIf Len(normalized) > 0 And Not lastSeparator Then
            normalized = normalized & "_"
            lastSeparator = True
        End If
    Next i

    Do While Len(normalized) > 0 And Right$(normalized, 1) = "_"
        normalized = Left$(normalized, Len(normalized) - 1)
    Loop

    NormalizeResponseId = normalized
End Function

Private Sub ValidateDialogId(ByVal Id As String, ByVal SourceName As String)
    If Len(NormalizeResponseId(Id)) = 0 Then
        Err.Raise xlflowErrInvalidDialogId, SourceName, "Dialog id must contain at least one ASCII letter or digit."
    End If
End Sub

Private Function DecodeWorkbookDefinedName(ByVal refersTo As String) As String
    If Len(refersTo) = 0 Then
        DecodeWorkbookDefinedName = ""
        Exit Function
    End If
    If Left$(refersTo, 1) = "=" Then
        refersTo = Mid$(refersTo, 2)
    End If
    If Len(refersTo) >= 2 Then
        If Left$(refersTo, 1) = Chr$(34) And Right$(refersTo, 1) = Chr$(34) Then
            refersTo = Mid$(refersTo, 2, Len(refersTo) - 2)
        End If
    End If
    DecodeWorkbookDefinedName = Replace$(refersTo, Chr$(34) & Chr$(34), Chr$(34))
End Function
