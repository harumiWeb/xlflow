Attribute VB_Name = "XlflowRuntime"
Option Explicit

' XlflowRuntime exposes the execution mode that xlflow injected before user VBA started.
' Use these helpers when workbook code must branch between interactive and unattended flows.
Private Const xlflowInteractive As Long = 0
Private Const xlflowHeadless As Long = 1
Private Const xlflowCI As Long = 2
Private Const xlflowAgent As Long = 3
Private Const xlflowTest As Long = 4

' Returns a stable numeric mode value for lightweight branching.
Public Function Mode() As Long
        Select Case ModeName()
                Case "headless"
                        Mode = xlflowHeadless
                Case "ci"
                        Mode = xlflowCI
                Case "agent"
                        Mode = xlflowAgent
                Case "test"
                        Mode = xlflowTest
                Case Else
                        Mode = xlflowInteractive
        End Select
End Function

' Returns the normalized mode name injected by xlflow.
Public Function ModeName() As String
        Dim raw As String
        raw = ReadWorkbookModeName()
        If Len(raw) = 0 Then
                raw = Environ$("XLFLOW_MODE")
        End If
        raw = LCase$(Trim$(raw))

        Select Case raw
                Case "headless", "ci", "agent", "test"
                        ModeName = raw
                Case Else
                        ModeName = "interactive"
        End Select
End Function

' True only for normal human-driven Excel usage.
Public Function IsInteractive() As Boolean
        IsInteractive = (Mode() = xlflowInteractive)
End Function

' True for all unattended-style modes such as headless, CI, agent, and test.
Public Function IsHeadless() As Boolean
        Select Case Mode()
                Case xlflowHeadless, xlflowCI, xlflowAgent, xlflowTest
                        IsHeadless = True
                Case Else
                        IsHeadless = False
        End Select
End Function

Public Function IsCI() As Boolean
        IsCI = (Mode() = xlflowCI)
End Function

Public Function IsAgent() As Boolean
        IsAgent = (Mode() = xlflowAgent)
End Function

Public Function IsTest() As Boolean
        IsTest = (Mode() = xlflowTest)
End Function

Private Function ReadWorkbookModeName() As String
        On Error GoTo Missing
        ReadWorkbookModeName = DecodeWorkbookDefinedName(ThisWorkbook.Names("__XLFLOW_MODE__").RefersTo)
        Exit Function

Missing:
        ReadWorkbookModeName = ""
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
