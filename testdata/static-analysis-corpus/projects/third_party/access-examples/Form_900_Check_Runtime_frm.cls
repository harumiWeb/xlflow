VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_900_Check_Runtime_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim TotalElapsedMilliSec As Long
Dim StartTickCount As Long
Dim lngCummulatTime As Long
Dim strLastObjName As String
Dim intObjectCount As Integer
Dim strPage As String

Public Function SetListboxSource(pintType As Integer) As String
On Error Resume Next
Dim strSql As String
Dim strFilter As String

                strFilter = "WHERE "
    strFilter = strFilter & "Name Not Like '~*' "
    strFilter = strFilter & "AND "
    strFilter = strFilter & "Type = " & pintType & " "

             strSql = "SELECT "
    strSql = strSql & "Name "
    strSql = strSql & "FROM "
    strSql = strSql & "MSysObjects "
    strSql = strSql & strFilter 'filter string
    strSql = strSql & "ORDER BY 1 "
    SetListboxSource = strSql

End Function

Public Function OpenObject()
On Error GoTo ErrTrap
Dim tStart As Double
Dim lngStart As Long
Dim lngElapsed As Long
Dim strName  As String
Dim strMsg As String
Dim strType As String

    strName = Me.lstObjectName
   
    Select Case Me.tabObject
        Case Is = Me.pgeQueries.PageIndex
            DoCmd.Close acQuery, strName
            strType = Me.pgeQueries.Caption
        Case Is = Me.pgeReports.PageIndex
            DoCmd.Close acReport, strName
            strType = Me.pgeReports.Caption
    End Select
    
    Me.txtMessage = "Started Opening " & strType & vbCrLf & strName & vbCrLf & "at " & Format(time(), "medium time")
    Me.Repaint
    lngStart = GetTickCount()
    
    Select Case Me.tabObject
        Case Is = Me.pgeQueries.PageIndex
            DoCmd.OpenQuery strName, acNormal
        Case Is = Me.pgeReports.PageIndex
            DoCmd.OpenReport strName, acViewPreview
    End Select
      
    Beep
    lngElapsed = GetTickCount() - lngStart
    strMsg = "Time to Open " & strType & vbCrLf & "   " & strName & vbCrLf & "      =  " & FormatElapsedTime(lngElapsed)
    If strLastObjName = strName Then
        intObjectCount = intObjectCount + 1
        lngCummulatTime = lngCummulatTime + lngElapsed
        strMsg = strMsg & vbCrLf & vbCrLf & "  Has been run  " & CStr(intObjectCount) & "  Times"
        strMsg = strMsg & vbCrLf & "  Average Run Time = " & FormatElapsedTime(lngCummulatTime / intObjectCount)
    
    Else
        strLastObjName = strName
        intObjectCount = 1
        lngCummulatTime = lngElapsed
    End If
   
    Me.txtMessage = strMsg
    Me.Repaint
    Beep

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("OpenObject", Me.Name, Err.Number, Err.Description)
            GoTo ExitProcedure
        Case Else
            GoTo ExitProcedure
    End Select
    
End Function

Private Sub cmdClose_Click()

    DoCmd.Close

End Sub

Private Sub cmdRefresh_Click()

    Me.lstObjectName.Requery

End Sub

Private Sub cmdRun_Click()

    Call OpenObject

End Sub

Private Sub Form_Load()
On Error Resume Next

    strLastObjName = "No Object Has been run yet"
    Me.lstObjectName.RowSource = SetListboxSource(Me.pgeQueries.Tag)
   
End Sub

Private Function FormatElapsedTime(lngTime As Long) As String
On Error GoTo ErrTrap
Dim Hours As String
Dim Minutes As String
Dim Seconds As String
Dim MilliSec As String
   
        Hours = Format((lngTime \ 3600000), "00")
        Minutes = Format((lngTime \ 60000) Mod 60, "00")
        Seconds = Format((lngTime \ 1000) Mod 60, "00")
        MilliSec = Format((lngTime Mod 1000) \ 10, "00")
        
        FormatElapsedTime = Hours & ":" & Minutes & ":" & Seconds & "." & MilliSec

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("FormatElapsedTime", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Sub lstObjectName_DblClick(Cancel As Integer)

    Call OpenObject

End Sub

Private Sub lstObjectName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTrap

    If KeyCode = 13 Then          'Enter Key has been pressed
        Call OpenObject
    End If

ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("lstObjectName_KeyDown", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub tabObject_Change()
On Error GoTo ErrTrap
Dim intType As Integer

    Select Case Me.tabObject
        Case Is = Me.pgeQueries.PageIndex
            intType = Me.pgeQueries.Tag
        Case Is = Me.pgeReports.PageIndex
            intType = Me.pgeReports.Tag
    End Select
    Me.lstObjectName.RowSource = SetListboxSource(intType)

ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("tabObject_Change", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub
