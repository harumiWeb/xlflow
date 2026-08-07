VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_010_Date_Calendar_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private mc As clsMonthCal

Private Sub Form_Activate()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Create an instance of our Class ** YOU MUST SET THE hWndForm Property!
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2005     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    Set mc = New clsMonthCal
    mc.hWndForm = Me.hwnd

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("Form_Activate", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Create an instance of our Class ** YOU MUST SET THE hWndForm Property!
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2005     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    Set mc = New clsMonthCal
    mc.hWndForm = Me.hwnd

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("Form_Load", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              This is required in case user Closes Form with the Calendar still open.
'                       It also handles when the user closes the application with the Calendar still open.
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2005     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    If Not mc Is Nothing Then
        If mc.IsCalendar Then
            Cancel = 1
            Exit Sub
        End If
        Set mc = Nothing
    End If

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("Form_Unload", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Sub SetControlDate()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Retrieve the currently selected date(s).
'                       Initialize the Calendar to select the date contained in this control.
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2005     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim ctrl As Control
Dim blnRet As Boolean
Dim dteStart As Date
Dim dteEnd As Date

    Set ctrl = Screen.ActiveControl
    dteStart = DateValue(Nz(ctrl.Value, Now()))
    gdteTimeValue = timeValue(Nz(ctrl.Value, Now()))
    dteEnd = 0
    gdteFillCtrl = ctrl.Name
    blnRet = ShowMonthCalendar(mc, dteStart, dteEnd)
    If blnRet = True Then
        ctrl.Value = CDate(dteStart)
        'ctrl.Value = CDate(dteStart & "  " & gdteTimeValue)
    End If

ExitProcedure:
    On Error Resume Next
    Set ctrl = Nothing
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("SetControlDate", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Function UpdateDateControl(intKeyAscii As Integer, Optional strD As String = "D") As Integer
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Increase or decrease the date with the plus and minus keys
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2005     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    Select Case intKeyAscii
        Case 43            ' Plus key
            'Screen.ActiveControl = Screen.ActiveControl + 1
            Screen.ActiveControl = DateAdd(strD, 1, Screen.ActiveControl)
            UpdateDateControl = 0
        Case 61            ' Plus key with [SHIFT]
            'Screen.ActiveControl = Screen.ActiveControl + 1
            Screen.ActiveControl = DateAdd(strD, 1, Screen.ActiveControl)
            UpdateDateControl = 0
        Case 45            ' Minus key
            Screen.ActiveControl = DateAdd(strD, -1, Screen.ActiveControl)
            UpdateDateControl = 0
        Case 95            ' Minus key with [SHIFT]
            Screen.ActiveControl = DateAdd(strD, -1, Screen.ActiveControl)
            UpdateDateControl = 0
        Case Else
            UpdateDateControl = intKeyAscii
    End Select

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("UpdateDateControl", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Private Sub txtEnterDate_DblClick(Cancel As Integer)
On Error Resume Next

    Call SetControlDate

End Sub

Private Sub txtEnterDate_KeyPress(KeyAscii As Integer)
On Error Resume Next

    KeyAscii = UpdateDateControl(KeyAscii)
    KeyAscii = 0
        
End Sub
