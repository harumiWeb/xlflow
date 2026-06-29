Attribute VB_Name = "XlflowDebug"
Option Explicit

' XlflowDebug mirrors workbook-side debug output to the terminal during xlflow runs.
' Use XlflowDebug.Log instead of raw Debug.Print when terminal-visible logs are desired.
Private Const xlflowDebugPipeName As String = "__XLFLOW_DEBUG_PIPE__"

#If VBA7 Then
        Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
        Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
        Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
        Private Const xlflowInvalidHandleValue As LongPtr = -1
#Else
        Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
        Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
        Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
        Private Const xlflowInvalidHandleValue As Long = -1
#End If

Private Const xlflowGenericWrite As Long = &H40000000
Private Const xlflowOpenExisting As Long = 3

Public Sub Log(ParamArray parts() As Variant)
        Dim index As Long
        Dim lowerBound As Long
        Dim upperBound As Long
        Dim Message As String
        Dim errorNumber As Long
        Dim errorSource As String
        Dim errorDescription As String

        On Error GoTo EmptyParts
        lowerBound = LBound(parts)
        upperBound = UBound(parts)
        On Error GoTo 0

        For index = lowerBound To upperBound
                If index > lowerBound Then
                        Message = Message & " "
                End If
                Message = Message & StringifyValue(parts(index))
        Next index

GoTo PrintMessage

EmptyParts:
        errorNumber = Err.Number
        errorSource = Err.source
        errorDescription = Err.description
        Err.Clear
        On Error GoTo 0
        If errorNumber <> 9 Then
                Err.Raise errorNumber, errorSource, errorDescription
        End If

PrintMessage:
        If Len(Message) = 0 Then
                Debug.Print
        Else
                Debug.Print Message
        End If

        EmitDebugEvent Message
End Sub

Private Sub EmitDebugEvent(ByVal Message As String)
        Dim PipeName As String
        Dim Payload As String

        PipeName = ResolveDebugPipeName()
        If Len(PipeName) = 0 Then
                Exit Sub
        End If

        Payload = "{" & _
                JsonProperty("event", "debug_log") & "," & _
                JsonProperty("message", Message) & "," & _
                JsonProperty("runtime_mode", XlflowRuntime.ModeName()) & "," & _
                JsonProperty("source", "XlflowDebug.Log") & "}"

        SendPipeText PipeName, Payload & vbLf
End Sub

Private Function StringifyValue(ByVal value As Variant) As String
        If IsObject(value) Then
                On Error GoTo ObjectFallback
                StringifyValue = "[Object " & TypeName(value) & "]"
                On Error GoTo 0
                Exit Function
ObjectFallback:
                Err.Clear
                On Error GoTo 0
                StringifyValue = "[Object]"
                Exit Function
        End If

        If IsArray(value) Then
                StringifyValue = "[Array]"
                Exit Function
        End If

        If IsEmpty(value) Then
                StringifyValue = "[Empty]"
                Exit Function
        End If

        If IsNull(value) Then
                StringifyValue = "[Null]"
                Exit Function
        End If

        Select Case VarType(value)
                Case vbBoolean
                        If CBool(value) Then
                                StringifyValue = "True"
                        Else
                                StringifyValue = "False"
                        End If
                Case vbError
                        StringifyValue = "[Error " & CStr(CLng(value)) & "]"
                Case Else
                        On Error GoTo UnsupportedValue
                        StringifyValue = CStr(value)
                        On Error GoTo 0
                        Exit Function
UnsupportedValue:
                        Err.Clear
                        On Error GoTo 0
                        StringifyValue = "[Unsupported Variant]"
        End Select
End Function

Private Function ResolveDebugPipeName() As String
        ResolveDebugPipeName = ReadOptionalDefinedNameValue(xlflowDebugPipeName)
End Function

Private Function ReadOptionalDefinedNameValue(ByVal Name As String) As String
        On Error GoTo Missing
        ReadOptionalDefinedNameValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(Name).RefersTo)
        Exit Function

Missing:
        ReadOptionalDefinedNameValue = ""
End Function

Private Function DecodeWorkbookDefinedName(ByVal RefersTo As String) As String
        If Len(RefersTo) = 0 Then
                DecodeWorkbookDefinedName = ""
                Exit Function
        End If
        If Left$(RefersTo, 1) = "=" Then
                RefersTo = Mid$(RefersTo, 2)
        End If
        If Len(RefersTo) >= 2 Then
                If Left$(RefersTo, 1) = Chr$(34) And Right$(RefersTo, 1) = Chr$(34) Then
                        RefersTo = Mid$(RefersTo, 2, Len(RefersTo) - 2)
                End If
        End If
        DecodeWorkbookDefinedName = Replace$(RefersTo, Chr$(34) & Chr$(34), Chr$(34))
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

Private Function JsonProperty(ByVal Name As String, ByVal value As String) As String
        JsonProperty = Chr$(34) & JsonEscape(Name) & Chr$(34) & ":" & Chr$(34) & JsonEscape(value) & Chr$(34)
End Function

Private Sub SendPipeText(ByVal PipeName As String, ByVal Payload As String)
        Dim bytesWritten As Long
#If VBA7 Then
        Dim pipeHandle As LongPtr
#Else
        Dim pipeHandle As Long
#End If

        pipeHandle = CreateFileW(StrPtr(PipeName), xlflowGenericWrite, 0, 0, xlflowOpenExisting, 0, 0)
        If pipeHandle = xlflowInvalidHandleValue Then
                Exit Sub
        End If

        On Error GoTo Cleanup
        Call WriteFile(pipeHandle, StrPtr(Payload), Len(Payload) * 2, bytesWritten, 0)

Cleanup:
        If pipeHandle <> xlflowInvalidHandleValue Then
                Call CloseHandle(pipeHandle)
        End If
End Sub
