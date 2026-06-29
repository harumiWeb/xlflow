Attribute VB_Name = "modWindowPlacement"
Option Explicit

Private Const MONITOR_DEFAULTTONEAREST As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type monitorInfo
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Declare PtrSafe Function MonitorFromWindow Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function GetMonitorInfoW Lib "user32" ( _
    ByVal hMonitor As LongPtr, _
    ByRef monitorInfo As monitorInfo) As Long

Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByRef windowRect As RECT) As Long

Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hwndInsertAfter As LongPtr, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal flags As Long) As Long

Private Declare PtrSafe Function FindWindowExW Lib "user32" ( _
    ByVal hwndParent As LongPtr, _
    ByVal hwndChildAfter As LongPtr, _
    ByVal classNamePointer As LongPtr, _
    ByVal windowNamePointer As LongPtr) As LongPtr

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByRef processId As Long) As Long

Public Function CenterUserFormOnExcelMonitor(ByVal targetForm As Object) As Boolean
    On Error GoTo ErrHandler

    Dim excelHwnd As LongPtr
    excelHwnd = ExcelWindowHandle()
    If excelHwnd = 0 Then Exit Function

    Dim formHwnd As LongPtr
    formHwnd = FindUserFormWindow(CStr(targetForm.caption), excelHwnd)
    If formHwnd = 0 Then Exit Function

    Dim monitorHandle As LongPtr
    monitorHandle = MonitorFromWindow(excelHwnd, MONITOR_DEFAULTTONEAREST)
    If monitorHandle = 0 Then Exit Function

    Dim monitorInfo As monitorInfo
    monitorInfo.cbSize = LenB(monitorInfo)
    If GetMonitorInfoW(monitorHandle, monitorInfo) = 0 Then Exit Function

    Dim formRect As RECT
    If GetWindowRect(formHwnd, formRect) = 0 Then Exit Function

    Dim formWidth As Long
    Dim formHeight As Long
    formWidth = formRect.Right - formRect.Left
    formHeight = formRect.Bottom - formRect.Top

    Dim targetX As Long
    Dim targetY As Long
    targetX = monitorInfo.rcWork.Left + ((monitorInfo.rcWork.Right - monitorInfo.rcWork.Left - formWidth) \ 2)
    targetY = monitorInfo.rcWork.Top + ((monitorInfo.rcWork.Bottom - monitorInfo.rcWork.Top - formHeight) \ 2)

    If targetX < monitorInfo.rcWork.Left Then targetX = monitorInfo.rcWork.Left
    If targetY < monitorInfo.rcWork.Top Then targetY = monitorInfo.rcWork.Top

    CenterUserFormOnExcelMonitor = (SetWindowPos( _
        formHwnd, _
        0, _
        targetX, _
        targetY, _
        0, _
        0, _
        SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE) <> 0)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "WindowPlacement", "modWindowPlacement.CenterUserFormOnExcelMonitor", Err.description, "", "", TypeName(targetForm)
    CenterUserFormOnExcelMonitor = False
End Function

Public Function MonitorPlacementSmoke() As String
    On Error GoTo ErrHandler

    Dim excelHwnd As LongPtr
    excelHwnd = ExcelWindowHandle()
    If excelHwnd = 0 Then
        Err.Raise vbObjectError + 7961, "modWindowPlacement.MonitorPlacementSmoke", "Excel window handle was not found."
    End If

    Dim monitorHandle As LongPtr
    monitorHandle = MonitorFromWindow(excelHwnd, MONITOR_DEFAULTTONEAREST)
    If monitorHandle = 0 Then
        Err.Raise vbObjectError + 7962, "modWindowPlacement.MonitorPlacementSmoke", "Excel monitor was not found."
    End If

    Dim monitorInfo As monitorInfo
    monitorInfo.cbSize = LenB(monitorInfo)
    If GetMonitorInfoW(monitorHandle, monitorInfo) = 0 Then
        Err.Raise vbObjectError + 7963, "modWindowPlacement.MonitorPlacementSmoke", "Excel monitor work area was not available."
    End If
    If monitorInfo.rcWork.Right <= monitorInfo.rcWork.Left Or monitorInfo.rcWork.Bottom <= monitorInfo.rcWork.Top Then
        Err.Raise vbObjectError + 7964, "modWindowPlacement.MonitorPlacementSmoke", "Excel monitor work area was invalid."
    End If

    MonitorPlacementSmoke = "ok:" & _
        CStr(monitorInfo.rcWork.Left) & "," & _
        CStr(monitorInfo.rcWork.Top) & "," & _
        CStr(monitorInfo.rcWork.Right) & "," & _
        CStr(monitorInfo.rcWork.Bottom)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "WindowPlacement", "modWindowPlacement.MonitorPlacementSmoke", Err.description, "", "", ""
    MonitorPlacementSmoke = "error:" & Err.description
End Function

Private Function ExcelWindowHandle() As LongPtr
    If Not Application.ActiveWindow Is Nothing Then
        If Not Application.ActiveWorkbook Is Nothing Then
            If Application.ActiveWorkbook Is ThisWorkbook Then
                ExcelWindowHandle = Application.ActiveWindow.hwnd
                Exit Function
            End If
        End If
    End If

    If ThisWorkbook.Windows.Count > 0 Then
        ExcelWindowHandle = ThisWorkbook.Windows(1).hwnd
    Else
        ExcelWindowHandle = Application.hwnd
    End If
End Function

Private Function FindUserFormWindow(ByVal formCaption As String, ByVal excelHwnd As LongPtr) As LongPtr
    Dim excelProcessId As Long
    GetWindowThreadProcessId excelHwnd, excelProcessId

    Dim candidateHwnd As LongPtr
    candidateHwnd = FindWindowExW(0, 0, 0, StrPtr(formCaption))

    Do While candidateHwnd <> 0
        Dim candidateProcessId As Long
        GetWindowThreadProcessId candidateHwnd, candidateProcessId
        If candidateProcessId = excelProcessId Then
            FindUserFormWindow = candidateHwnd
            Exit Function
        End If
        candidateHwnd = FindWindowExW(0, candidateHwnd, 0, StrPtr(formCaption))
    Loop
End Function
