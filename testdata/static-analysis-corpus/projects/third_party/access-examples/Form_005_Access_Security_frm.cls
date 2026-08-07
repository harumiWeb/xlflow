VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_005_Access_Security_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Security Information about the database
' Notes:                Accounts and Permissions are only enabled for the Admins Group
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

' Declare a record type to break down the user info
Private Type UserRec
   bMach(1 To 32) As Byte  ' 1st 32 bytes hold machine name
   bUser(1 To 32) As Byte  ' 2nd 32 bytes hold user name
End Type

Private Sub cmdClearPassword_Click()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Clear the current user's password
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next

    If IsNull(Me.lstUserList) Then
        MsgBox "Select a User.", vbInformation, "No action taken"
        Exit Sub
    End If
    If ClearPassword(Me.lstUserList.Column(4)) = True Then
        MsgBox "Password cleared for " & Me.lstUserList.Column(1), vbInformation, "Done."
    End If

End Sub

Private Sub cmdClose_Click()
On Error Resume Next

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim varCurrPwd As Variant
Dim varNewPwd As Variant
Dim varCnfPwd As Variant

    Select Case Me.tabSecurity.Value
        Case Is = Me.pgePassword.PageIndex
            varCurrPwd = Me.txtCurrentPassword
            varNewPwd = Me.txtNewPassword
            varCnfPwd = Me.txtConfirmPassword
        
            If IsNull(varCurrPwd) And IsNull(varNewPwd) And IsNull(varCnfPwd) Then
                Call CheckPassword
            Else
                Call ChangePassword
                DoCmd.Close acForm, Me.Name        'commented out the user may want the form to stay open
            End If
        Case Is = Me.pgeUsers.PageIndex
            Me.lstUserList.RowSource = GetUserList()
    End Select

ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            MsgBox Err.Description
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Function CheckPassword() As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Check to see if the user enter all fields
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strMsg As String
Dim strInput As String
Dim Msg, style, TITLE, Help, Ctxt, Response, MyString

    Msg = "Your Current Password has not been entered.  Would you like to exit this form?"   ' Define message
    style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons
    TITLE = CurrentUser()
    Response = MsgBox(Msg, style, TITLE)

    If Response = vbYes Then    ' User chose Yes.
        MyString = "Yes"
        DoCmd.Close acForm, Me.Name
    Else    ' User chose No.
        MyString = "No"
        Me.txtCurrentPassword.SetFocus
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

Private Sub ChangePassword()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Check to see if the user's password is correct
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim varCurrPwd As Variant
Dim varNewPwd As Variant
Dim varCnfPwd As Variant

    varCurrPwd = Me.txtCurrentPassword
    varNewPwd = Me.txtNewPassword
    varCnfPwd = Me.txtConfirmPassword

SET_PWD:
    If IsNull(varCurrPwd) Then varCurrPwd = ""
    If IsNull(varNewPwd) And Not IsNull(varCnfPwd) Then GoTo ENTRY_ERROR
    If Not IsNull(varNewPwd) And IsNull(varCnfPwd) Then GoTo ENTRY_ERROR
    If IsNull(varNewPwd) And IsNull(varCnfPwd) Then GoTo PWD_VARIFY
    If StrComp(varNewPwd, varCnfPwd, 0) <> 0 Then GoTo ENTRY_ERROR


PWD_VARIFY:
    If IsNull(varCurrPwd) Then varCurrPwd = ""
    If IsNull(varNewPwd) Then GoTo ENTRY_ERROR
    DBEngine.Workspaces(0).Users(CurrentUser()).NewPassword varCurrPwd, varNewPwd 'error here.
    MsgBox "Your Password has been changed. ", vbInformation, CurrentUser()
    DoCmd.Close acForm, Me.Name
    Exit Sub

ENTRY_ERROR:
    DoCmd.Beep
    MsgBox "Your entries for the new password are not identical. Please enter both passwords again.", vbInformation, CurrentUser()
    Me.txtNewPassword = Null
    Me.txtConfirmPassword = Null
    Me.txtNewPassword.SetFocus
    Exit Sub

ErrTrap:
    DoCmd.Beep
    MsgBox "There has been an error. Please verify that you have the correct OLD password and a valid NEW password for this user ID.", vbInformation, CurrentUser()
    Me.txtCurrentPassword.SetFocus
    Exit Sub

End Sub

Public Function IsAdminsGrp() As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Check to see if the user is in the Admins Group
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strGroup As String
Dim strUser As String
Dim ws As Workspace
Dim grp As GROUP
Dim strUserName As String
    strGroup = "Admins"
    strUser = CurrentUser()

100     Set ws = DBEngine.Workspaces(0)
200     Set grp = ws.Groups(strGroup) 'errors here.
300     strUserName = ws.Groups(strGroup).Users(strUser).Name
        IsAdminsGrp = True
        
ExitProcedure:
    On Error Resume Next
    Set ws = Nothing
    Set grp = Nothing
    Exit Function

ErrTrap:
    If Erl = 100 Or Erl = 200 Or Erl = 300 Then 'user group doesn't equal Admins
        IsAdminsGrp = False
        Resume ExitProcedure
    End If
    
    Select Case Err.Number
        Case Is <> 0
            MsgBox Err.Description
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Sub cmdAccounts_Click()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Open the MS Access Form for User and Group Accounts
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    DoCmd.RunCommand acCmdUserAndGroupAccounts
    
ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            MsgBox Err.Description
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdPermissions_Click()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Open the MS Access Form for User and Group Permissions
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    DoCmd.RunCommand acCmdUserAndGroupPermissions
    
ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            MsgBox Err.Description
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub Form_Open(Cancel As Integer)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Check to see if the user is in the Admins Group on the opening of the form.
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    If IsAdminsGrp Then
        Me.cmdAccounts.Enabled = True
        Me.cmdPermissions.Enabled = True
        Me.pgeUsers.Visible = True
        Me.lstUserList.RowSource = GetUserList()
    Else
        Me.cmdAccounts.Enabled = False
        Me.cmdPermissions.Enabled = False
        Me.pgeUsers.Visible = False
    End If

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            MsgBox Err.Description
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Function GetLoggedUsers(strUser As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Will read *.LDB file and read who's currently logged on and their machine name.
' Notes:                The LDB file has a 64 byte record.  The station name starts at byte 1 and is null
'                       terminated.  Log-in names start at the 33rd byte and are also null terminated.
'                       I had to change the way the file was accessed because the Input() function did not
'                       return nulls, so there was no way to see where the names ended.
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim intLdbFile As Integer
Dim intStart As Integer
Dim intLOF As Integer
Dim i As Integer
Dim strPath As String
Dim x As String
Dim strLogString As String
Dim strLoginUsers As String
Dim strMachine As String
Dim strUserNme As String
Dim recUserNme As UserRec    ' Defined in General
Dim dbCurrent As Database

   Set dbCurrent = DBEngine.Workspaces(0).Databases(0)
   strPath = dbCurrent.Name ' path of where we will find the .ldb file
   dbCurrent.Close
   strPath = Left(strPath, InStr(1, strPath, ".")) + "LDB" ' Iterate thru dbCurrent.LDB file for login names.

   x = Dir(strPath) ' Test for valid file, else Error
   intStart = 1
   intLdbFile = FreeFile

   Open strPath For Binary Access Read Shared As intLdbFile
   intLOF = LOF(intLdbFile)
   Do While Not EOF(intLdbFile)
      Get intLdbFile, , recUserNme
      With recUserNme
         i = 1
         strMachine = ""
         While .bMach(i) <> 0
            strMachine = strMachine & Chr(.bMach(i))
            i = i + 1
         Wend
         i = 1
         strUserNme = ""
         While .bUser(i) <> 0
            strUserNme = strUserNme & Chr(.bUser(i))
            i = i + 1
         Wend
      End With
      If strUser = strUserNme Then
        GetLoggedUsers = strMachine & "; " & strUserNme & "; " & Date & "; " & time()
        GoTo ExitProcedure
      Else
        intStart = intStart + 64 'increment to next record offset
        GetLoggedUsers = ""
      End If
   Loop
    GetLoggedUsers = ""
    
ExitProcedure:
    On Error Resume Next
    Close intLdbFile
    Set dbCurrent = Nothing
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case 68
            MsgBox "Couldn't populate the list", 48, "No LDB File"
            DoCmd.Close
            Resume ExitProcedure
        Case Is <> 0
            MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
            Close intLdbFile
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Private Sub lblUser_Click()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              show user's information
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next

    MsgBox "Machine: " & vbTab & Me.lstUserList & vbCrLf & _
    "User: " & vbTab & Me.lstUserList.Column(1) & vbCrLf & _
    "Id: " & vbTab & Me.lstUserList.Column(4) & vbCrLf & _
    "Date: " & vbTab & Me.lstUserList.Column(2) & vbCrLf & _
    "Time: " & vbTab & Me.lstUserList.Column(3) & vbCrLf, vbOKOnly, "Selected User Informtion"

End Sub

Private Sub tabSecurity_Change()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Change settings to buttons and pages, also update the user list
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next

    Select Case Me.tabSecurity
        Case Is = Me.pgePassword.PageIndex
            Me.cmdOk.Caption = "&Save"
            Me.cmdClose.Caption = "&Cancel"
            Me.cmdOk.ControlTipText = "Save Password"
            Me.cmdClearPassword.Visible = False
        Case Is = Me.pgeUsers.PageIndex
            Me.cmdOk.Caption = "Refresh &List"
            Me.cmdClose.Caption = "&Close"
            Me.cmdOk.ControlTipText = "Update Current User List"
            Me.lstUserList.RowSource = GetUserList()
            Me.cmdClearPassword.Visible = True
    End Select
    
End Sub

Public Function GetUserList() As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Get a list of all the users in the security file (.mdw)
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim MyWorkspace         As Workspace
Dim strUserList         As String
Dim strTempRec          As String
Dim strTempUser         As String
Dim TempUser            As User
Dim intNumUser          As Integer
Dim intSource           As Integer
Dim x                   As Integer

    DoCmd.Hourglass True
    Set MyWorkspace = DBEngine.Workspaces(0)
    intNumUser = MyWorkspace.Users.Count
    For x = 0 To intNumUser - 1
        Set TempUser = MyWorkspace.Users(x)
        strTempUser = TempUser.Name
        strTempRec = GetLoggedUsers(TempUser.Name)
        If Len(strTempRec) = 0 Then
            If TempUser.Name = "Creator" Or TempUser.Name = "Engine" Then
                'DO NOTHING, MAY WANT TO SHOW THIS AT SOME POINT
            Else
                strUserList = strUserList & " " & "; " & TempUser.Name & "; " & " " & "; " & " " & "; " & x & "; "
            End If
        Else
            strUserList = strUserList & strTempRec & "; " & x & "; "
        End If
    Next x
    GetUserList = Left(strUserList, Len(strUserList) - 1)
    Me.pgeUsers.Caption = "User &List [" & x - 2 & "]"

ExitProcedure:
    On Error Resume Next
    DoCmd.Hourglass False
    Set MyWorkspace = Nothing
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            'MsgBox "An error occured while populating the list of users: " & Error$
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Private Function ClearPassword(intSerNum As Integer) As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Call this to clear the user's password
'
' Ver.  Date            Author              Details
' 1.00  03 FEB 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    DBEngine.Workspaces(0).Users(intSerNum).NewPassword "", ""
    ClearPassword = True 'returns true if successful

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            ClearPassword = False
            'MsgBox "There has been an error in clearing the password. Please verify that you are logged in as a member of the Admins group, which is required to clear passwords. Error message: " & Error$
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function
