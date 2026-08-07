VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_001_About_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error Resume Next

    DoCmd.Close

End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              load form defaults
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
  
    Me.Caption = "About " & GetProperties("Title")
    Me.txtVersion = GetProperties("Title")
    Me.txtDescription = GetProperties("Subject")
    Me.txtPath = Application.CurrentProject.FullName
    Me.txtComputerName = Environ("COMPUTERNAME")
    Me.txtUserName = Environ("USERNAME")
    Me.txtAuthor = GetProperties("Author")
    Me.txtManager = GetProperties("Manager")
    Me.txtEmail.Caption = GetProperties("Hyperlink base")
    Me.txtEmail.HyperlinkAddress = "mailto:" & Me.txtEmail.Caption & "?subject=" & Me.txtVersion & " - " & Me.txtDescription '& "&body=%0A this is a test"
    Me.txtCopyright = GetProperties("Company") & "  " & Format(DateValue(GetProperties("LastUpdated")), "dd-mmm-yyyy")
    Me.txtWarning = GetProperties("Keywords")
    Me.lblHidden.Caption = ""
    Me.imgLogo.Picture = GetProperties("Category")

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2220 'reference to logo not valid
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("Form_Load", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Function GetProperties(pstrName As String) As Variant
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Returns information from the database properties.
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    GetProperties = Trim(CodeDb.Containers!Databases.Documents!SummaryInfo.Properties(pstrName))

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is = 3270 'property not assigned
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("GetProperties", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
End Function

Private Sub lblHidden_DblClick(Cancel As Integer)
On Error Resume Next

    If IsTestDbs = True Then
        DoCmd.RunCommand acCmdDatabaseProperties
    End If
    
End Sub

Public Function IsTestDbs() As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Is the database test or production?
'
' Ver.  Date            Author              Details
' 1.00  03 MAY 2008     Anthony  Duguid     Initial version.
' 1.01  01 MAY 2009     Anthony  Duguid     added "accdb" to 2007 version case
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strApp As String

    IsTestDbs = False
    strApp = Application.CurrentProject.Name
    Select Case Application.Version
        Case Is = "12.0" 'Access 2007
            If Right(strApp, 5) = "accdb" Or Right(strApp, 3) = "mdb" Then
                IsTestDbs = True
            End If
        
        Case Else 'Access 95, 97, 2000, 2002, 2003
            If Right(strApp, 3) = "mdb" Then
                IsTestDbs = True
            End If

    End Select

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("IsTestDbs", "modUtilities", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

