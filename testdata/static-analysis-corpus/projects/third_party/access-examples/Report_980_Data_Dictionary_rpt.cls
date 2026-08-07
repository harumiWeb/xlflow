VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_980_Data_Dictionary_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Data Dictionary of current database
'
' Ver.  Date            Author              Details
' 1.00  22 JUL 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Function FormatName(strFrmNme As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Format the object name for the title
'
' Ver.  Date            Author              Details
' 1.00  01 MAY 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strTemp As String
Dim intL As Integer
    intL = Len(strFrmNme)
    
    If intL < 8 Then
        FormatName = "Not a valid object name."
        GoTo ExitProcedure
    End If
    'remove the first 4 characters from the naming convention i.e. '123_'
    'remove the last 4 characters from the naming convention i.e. '_rpt'
    strTemp = Mid(strFrmNme, 5, intL - 8) ' 4 for "frm" or "rpt" to appear
    'remove all the '_' from the string
    strTemp = ReplaceStr(strTemp, "_", " ", 1)
    FormatName = strTemp

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("FormatName", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
            
End Function

Public Function ReplaceStr(strText As String, _
strSearch As String, _
strReplacement As String, _
intCompMode As Integer) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Replaces the pvarSearchStr string with pvarReplacement string in the pvarTextIn string.
'                       Uses pintCompMode to determine comparison mode as an option
'
' Ver.  Date            Author              Details
' 1.00  06 MAR 2008     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Dim strTemp As String
Dim Pointer As Integer

    strTemp = strText
    Pointer = InStr(1, strTemp, strSearch, intCompMode)
    Do While Pointer > 0
      strTemp = Left(strTemp, Pointer - 1) & strReplacement & Mid(strTemp, Pointer + Len(strSearch))
      Pointer = InStr(Pointer + Len(strReplacement), strTemp, strSearch, intCompMode)
    Loop
    ReplaceStr = strTemp

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("ReplaceStr", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Function GetObjectDescription(strObjtNme As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              get the description of the object
'
' Ver.  Date            Author              Details
' 1.00  01 MAY 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim dbCurrent       As Database         ' current database
Dim conTest         As Container        ' Container object
Dim docTest         As Document         ' Document object
Dim objTest         As Control        ' Control object

    ' open the database
    Set dbCurrent = DBEngine.Workspaces(0).Databases(0)
    
    ' loop through the containers displaying each ones name
    For Each conTest In dbCurrent.Containers
        ' and the documents for this container
        For Each docTest In conTest.Documents
            If docTest.Name = strObjtNme Then
                GetObjectDescription = docTest.Properties("Description")
            End If
        Next
    Next

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case 3270
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("GetObjectDescription", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Sub ErrorMsg(strProcedure As String, _
strModule As String, _
dblErrNbr As Double, _
strErrDes As String, _
Optional varErrLine As Variant = 0, _
Optional strTitle As String = "Unexpected Error")
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Global error message for all procedures
'
' Ver.  Date            Author              Details
' 1.00  20 JUN 2002     Anthony Duguid      Initial version
' 1.01  22 FEB 2008     Anthony Duguid      added error line
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim strMsg As String
        
    strMsg = "Contact your system administrator." & vbCrLf & _
    "Module: " & strModule & vbCrLf & _
    "Procedure: " & strProcedure & vbCrLf & _
    IIf(varErrLine = 0, "", "Error Line: " & varErrLine & vbCrLf) & _
    "Error #: " & dblErrNbr & vbCrLf & _
    "Error Description: " & strErrDes

    MsgBox strMsg, vbCritical, strTitle

End Sub
