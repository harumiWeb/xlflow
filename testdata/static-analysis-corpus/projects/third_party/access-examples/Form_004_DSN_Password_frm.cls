VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_004_DSN_Password_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Create a form to login to an Oracle database as ReadOnly or
'                       allow user to login in with the same form.  Create message to
'                       let the user know they entered the incorrect UID and/or PWD
'
' Ver.  Date            Author              Details
' 1.00  13 SEP 2001     Anthony Duguid      Initial version, the towers fell
'--------------------------------------------------------------------------------------------------------------------
Option Explicit
Public bLoginChk As Boolean

Public Function GetOsUser() As String
' Returns the Network Login Name
    GetOsUser = Environ("USERNAME")
    
End Function

Private Sub cmdCancel_Click()
On Error Resume Next

    If bLoginChk = True Then
        DoCmd.Close
    End If
    DoCmd.Quit 'Exit the application

End Sub

Private Sub cmdOK_Click()
On Error Resume Next

    Call sLogin
    If bLoginChk = True Then
        DoCmd.Close
    End If

End Sub

Public Function sLogin() As Boolean
On Error GoTo ErrTrap
Dim lngTemp  As Long
Dim x As Integer
Dim strUid As String
Dim strPwd As String
       
100     DoCmd.Hourglass True    'turn on the hourglass; wouldn't want people clicking on everthing...

    'Did they enter the "USER NAME"?
    If IsNull(Me.txtUid) = True Then
        DoCmd.Hourglass False   'turn off the hourglass; done with this
        MsgBox "Please enter a User Name.", vbInformation
        Me.txtUid.SetFocus
        Exit Function
    End If
    
    'Did they enter the "PASSWORD"?
    If IsNull(Me.txtPwd) = True Then
        DoCmd.Hourglass False   'turn off the hourglass; done with this
        MsgBox "Please enter a Password.", vbInformation
        Me.txtPwd.SetFocus
        Exit Function
    End If
    
    'If they entered both, proceed.
    If IsNull(Me.txtUid) = False And IsNull(Me.txtPwd) = False Then
        Call OpenDbs(Me.txtUid.Value, Me.txtPwd.Value)
        DoCmd.Hourglass False   'turn off the hourglass; done with this
    End If

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            MsgBox Err.Description
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Sub cmdReadOnly_Click()
On Error GoTo ErrTrap

    DoCmd.Hourglass True    'turn on the hourglass
    Call OpenDbs(Me.lblUid, Me.lblPwd) 'open the database as read only

ExitProcedure:
    On Error Resume Next
    DoCmd.Hourglass False   'turn off the hourglass
    If bLoginChk = True Then
        DoCmd.Close
    End If
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
        
End Sub

Public Function OpenDbs(strUid As String, strPwd As String) As Boolean
On Error GoTo ErrTrap
Dim dbsCurr As Database
Dim qdfPt As QueryDef
Dim strOdbc As String
Dim strDbs As String
Dim strDsn As String
Dim strTbl As String
Dim x As Integer

    strDbs = Me.lblDbs      'Database Name
    strDsn = Me.lblDsn      'DSN
    strTbl = Me.lblTbl      'Table name for use in the pass-through query
        
    'Connect String
    strOdbc = "ODBC;DSN=" & strDsn & ";UID=" & strUid & ";PWD=" & strPwd & ";DBQ=" & strDbs & ";DBA=W;APA=T;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;FRL=F;MTS=T;CSR=F;PFC=10;TLO=0;"
    Set dbsCurr = CurrentDb
   'Create a pass-through query
    Set qdfPt = dbsCurr.CreateQueryDef("")
    qdfPt.Connect = strOdbc 'connect string
    qdfPt.SQL = "SELECT COUNT (*) FROM " & strTbl 'pass-through SQL string
    qdfPt.ReturnsRecords = True
    qdfPt.OpenRecordset
    bLoginChk = True
    
    '----------------------------------
    'set globals for changing password
    '----------------------------------
    'gstrUID = strUID
    'gstrPWD = strPWD
    'gstrDBS = strDBS
    '----------------------------------

ExitProcedure:
    On Error Resume Next
    Set qdfPt = Nothing
    Set dbsCurr = Nothing
    'DoCmd.Close  'close current form
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is = 3151
            MsgBox "Your User Name and/or Password are incorrect." & vbCrLf & _
            "Please re-enter them.", vbCritical, "Login Error"
            Me.txtUid.SetFocus
            Resume ExitProcedure
        Case Is = 3146
            Resume ExitProcedure 'stupid read only access
        Case Is = 2501
            MsgBox "You have been logged in as Read Only", vbInformation, "Login Information"
            Resume ExitProcedure 'this one's for Raymond who likes to do the {ALT} {F4} combination (closes active window)
        Case Else
            MsgBox "Contact your system administrator." & vbCrLf & Err.Description, vbCritical, "Login Failure:  # " & Err.Number
            Resume ExitProcedure
    End Select

End Function

Private Sub Form_Close()
On Error Resume Next
    
    'If bLoginChk = False Then
    '    Call OpenDBS(Me.lblUid.Caption, Me.lblPwd.Caption)
    'End If

End Sub

Private Sub Form_Load()
On Error Resume Next
      
    Me.txtUid.SetFocus

End Sub

Private Sub lblExit_Click()
On Error Resume Next

    DoCmd.Close

End Sub

Private Sub lblPageVisible_DblClick(Cancel As Integer)
On Error Resume Next

    Me.pgeDefault.Visible = True

End Sub

Private Sub txtPwd_AfterUpdate()
On Error GoTo ErrTrap

    If IsNull(Me.txtUid) = False And IsNull(Me.txtPwd) = False Then
        Me.SetFocus
        Me.cmdOk.SetFocus
        Call cmdOK_Click
    End If

ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is = 2501
            Resume ExitProcedure
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbCritical, "Form_PWD_frm" & ": " & "txtPwd_AfterUpdate"
            Resume ExitProcedure
    End Select
    
End Sub
