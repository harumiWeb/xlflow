Attribute VB_Name = "modMain"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Main procedures for database
'
' Ver.  Date            Author              Details
' 1.00  11-MAY-2001     Anthony  Duguid     Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Function OSUserID() As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Returns the network login name
'
' Ver.  Date            Author              Details
' 1.00  11-MAY-2001     Anthony  Duguid     Initial version.
' 1.01  22-AUG-2002     Anthony  Duguid     Referencing new function
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    OSUserID = Environ("USERNAME")

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("OSUserID", "modMain", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Sub ErrorMsg( _
ByVal pstrProcedure As String, _
ByVal pstrModule As String, _
dblErrNbr As Double, _
strErrDes As String, _
Optional ByVal pvarErrLine As Variant = 0, _
Optional ByVal pstrTitle As String = "Unexpected Error")
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Global error message for all procedures
'
' Ver.  Date            Author              Details
' 1.00  20-JUN-2002     Anthony  Duguid     Initial version.
' 1.01  22-FEB-2008     Anthony  Duguid     added line number
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim strMsg As String

    strMsg = "Contact your system administrator."
    strMsg = strMsg & vbCrLf & "Module: " & pstrModule
    strMsg = strMsg & vbCrLf & "Procedure: " & pstrProcedure
    strMsg = strMsg & IIf(pvarErrLine = 0, "", vbCrLf & "Error Line: " & pvarErrLine)
    strMsg = strMsg & "Error #: " & dblErrNbr & vbCrLf
    strMsg = strMsg & "Error Description: " & strErrDes
    MsgBox strMsg, vbCritical, pstrTitle

End Sub

Public Function GetObjectDescription(ByVal pstrObjectName As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Get the description of the object
'
' Ver.  Date            Author              Details
' 1.00  01-MAY-2004     Anthony  Duguid     Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim dbs       As Database         ' current database
Dim con       As Container        ' Container object
Dim doc       As Document         ' Document object

    ' open the database
    Set dbs = DBEngine.Workspaces(0).Databases(0)
    
    ' loop through the containers displaying each ones name
    For Each con In dbs.Containers
        'Debug.Print "Container: "; conTest.NAME 'i.e. report, form, module.
        ' and the documents for this container
        For Each doc In con.Documents
            If doc.Name = pstrObjectName Then
                GetObjectDescription = doc.Properties("Description")
            End If
        Next
    Next

ExitProcedure:
    On Error Resume Next
    Set dbs = Nothing
    Set con = Nothing
    Set doc = Nothing
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case 3270 '
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("GetObjectDescription", "modMain", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function GetObjectTypeName(ByVal pdblTypeNbr As Double) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Format the object type code for the description
'
' Ver.  Date            Author              Details
' 1.00  01-MAY-2004     Anthony  Duguid     Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    Select Case pdblTypeNbr
        Case Is = -32758
            GetObjectTypeName = "Admin"
        Case Is = -32768
            GetObjectTypeName = "Form"
        Case Is = -32764
            GetObjectTypeName = "Report"
        Case Is = -32766
            GetObjectTypeName = "Macro"
        Case Is = -32761
            GetObjectTypeName = "Module"
        Case Is = 1
            GetObjectTypeName = "Access Table"
        Case Is = 4
            GetObjectTypeName = "Linked ODBC Table"
        Case Is = 5
            GetObjectTypeName = "Query"
        Case Is = 6
            GetObjectTypeName = "Linked Access Table"
        Case Is = 8
            GetObjectTypeName = "Relationship"
        Case Else
            GetObjectTypeName = CStr(pdblTypeNbr)
    End Select

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("GetObjectTypeName", "modMain", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function
