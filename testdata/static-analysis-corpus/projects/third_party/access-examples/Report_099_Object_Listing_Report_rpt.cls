VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_099_Object_Listing_Report_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Report List of all objects within database
' References:           Microsoft Visual Basic for Applications Extensibility 5.3
'                       Microsoft DAO 3.6 Object Library - dao360.dll
'
' Ver.  Date            Author              Details
' 1.00  17 JUL 2002     Anthony Duguid      Initial version
' 1.01  13 MAR 2008     Anthony Duguid      added the subroutines in modules
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Private Sub Report_Open(Cancel As Integer)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Load the source of the report
'
' Ver.  Date            Author              Details
' 1.00  01 MAY 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim dbs          As Object
Dim objAccess    As AccessObject
Dim strModule    As String

    Me.RecordSource = GetReportSource()
    'must open all the modules to get the procedure list of each
    Set dbs = Application.CurrentProject
    For Each objAccess In dbs.AllModules
        strModule = objAccess.Name
        DoCmd.OpenModule strModule
    Next objAccess

ExitProcedure:
    On Error Resume Next
    Set dbs = Nothing
    Set objAccess = Nothing
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("Report_Open", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub secDetail_Format(Cancel As Integer, FormatCount As Integer)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Alternate colors between the lines
'
' Ver.  Date            Author              Details
' 1.00  20 JUN 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next

   If Me.txtObjectCount Mod 2 = 1 Then
      Me.SECTION(0).BackColor = 16777215 'white
   Else
      Me.SECTION(0).BackColor = 15269864 'light green
   End If
   
End Sub

Public Function FormatName(ByVal pstrObjectName As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Format the object name for the title
'
' Ver.  Date            Author              Details
' 1.00  01 MAY 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strTemp As String
Dim intL As Integer

    intL = Len(pstrObjectName)
    'remove the first 4 characters from the naming convention i.e. '123_'
    'remove the last 4 characters from the naming convention i.e. '_rpt'
    strTemp = Mid(pstrObjectName, 5, intL - 8) ' 4 for "frm" or "rpt" to appear
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

Public Function GetObjectDescription(ByVal pstrObjectName As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              get the description of the object
'
' Ver.  Date            Author              Details
' 1.00  01 MAY 2004     Anthony Duguid      Initial version
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
            Call ErrorMsg("GetObjectDescription", Me.Name, Err.Number, Err.Description)
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
' 1.00  01 MAY 2004     Anthony Duguid      Initial version
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
            Call ErrorMsg("GetObjectTypeName", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Function GetProcedureList(Optional ByVal pstrModule As String = "") As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              get procedure names in module(s)
' Note:                 must open all modules before running this code
'
' Ver.  Date            Author              Details
' 1.00  13 MAR 2008     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim dbs          As Object
Dim mdl          As Module
Dim objAccess    As AccessObject
Dim lngLineNum   As Long
Dim strProcName  As String
Dim strProcList  As String
Dim strModule    As String
Dim intProcType  As Long
    
    intProcType = 3 '0=vbext_pk_Proc, 1=vbext_pk_Let, 2=vbext_pk_Set, 3=vbext_pk_Get
    If Len(pstrModule) = 0 Then 'get complete list of modules and there procedures
        Set dbs = Application.CurrentProject
        ' Search for open AccessObject objects in AllModules collection.
        For Each objAccess In dbs.AllModules
            strModule = objAccess.Name
            'DoCmd.OpenModule strModule ' *** DOES NOT WORK FROM QUERY, FORM OR REPORT ***
            Set mdl = Modules(strModule)
            With mdl
                lngLineNum = .CountOfDeclarationLines + 1
                strProcName = .ProcOfLine(lngLineNum, intProcType)
                Do Until lngLineNum >= .CountOfLines
                    strProcList = strProcList & strProcName & vbCrLf 'build list of procedures
                    lngLineNum = lngLineNum + .ProcCountLines(strProcName, intProcType) + 1
                    strProcName = .ProcOfLine(lngLineNum, intProcType)
                Loop
            End With
        Next objAccess
    Else    'only get the procedures for the passed in module
        'open module here
        'DoCmd.OpenModule pstrModule  ' *** DOES NOT WORK FROM QUERY, FORM OR REPORT ***
        Set mdl = Modules(pstrModule)
        With mdl
            lngLineNum = .CountOfDeclarationLines + 1
            strProcName = .ProcOfLine(lngLineNum, intProcType)
            Do Until lngLineNum >= .CountOfLines
                strProcList = strProcList & strProcName & vbCrLf 'build list of procedures
                lngLineNum = lngLineNum + .ProcCountLines(strProcName, intProcType) + 1
                strProcName = .ProcOfLine(lngLineNum, intProcType)
            Loop
        End With
    End If
    GetProcedureList = strProcList

ExitProcedure:
    On Error Resume Next
    Set dbs = Nothing
    Set mdl = Nothing
    Set objAccess = Nothing
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is = 35 'class module
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("GetProcedureList", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Function GetReportSource() As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Create the report source
'
' Ver.  Date            Author              Details
' 1.00  01 MAY 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strSql As String

                      strSql = "SELECT "
    strSql = strSql & vbCrLf & "MSysRelationships.*, "
    strSql = strSql & vbCrLf & "[Type], "
    strSql = strSql & vbCrLf & "[Name], "
    strSql = strSql & vbCrLf & "[DateUpdate], "
    strSql = strSql & vbCrLf & "IIf(IsNull([Database]),Null,'Attachment = ' & [Database]) AS Attachment, "
    strSql = strSql & vbCrLf & "IIf(IsNull([Connect]),Null,'Connection = ' & [Connect]) AS [Connection], "
    strSql = strSql & vbCrLf & "IIf(IsNull([ForeignName]),Null,'Foreign Name = ' & [ForeignName]) AS [Foreign Name] "
    strSql = strSql & vbCrLf & "FROM "
    strSql = strSql & vbCrLf & "MSysObjects "
    strSql = strSql & vbCrLf & "LEFT JOIN "
    strSql = strSql & vbCrLf & "MSysRelationships "
    strSql = strSql & vbCrLf & "ON "
    strSql = strSql & vbCrLf & "MSysObjects.Name = MSysRelationships.szRelationship "
    strSql = strSql & vbCrLf & "WHERE "
    strSql = strSql & vbCrLf & "[Name] Not Like 'Msys*' " 'don't show system objects
    strSql = strSql & vbCrLf & "AND "
    strSql = strSql & vbCrLf & "[Name] Not Like 'AccessLayout' "
    strSql = strSql & vbCrLf & "AND "
    strSql = strSql & vbCrLf & "[Name] Not Like '~sq_*' " 'don't show temporary queries
    strSql = strSql & vbCrLf & "AND "
    strSql = strSql & vbCrLf & "[Name] Not Like '*~*' " 'don't show temporary objects
    strSql = strSql & vbCrLf & "AND "
    strSql = strSql & vbCrLf & "[Name] Not Like 'zzz*' " 'don't show objects I may want to delete
    strSql = strSql & vbCrLf & "AND "
    strSql = strSql & vbCrLf & "[Type] Not In (-32757,-32758, 3) "
    strSql = strSql & vbCrLf & "ORDER BY "
    strSql = strSql & vbCrLf & "1, "
    strSql = strSql & vbCrLf & "2"
    'Debug.Print strSql
    GetReportSource = strSql

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("GetReportSource", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function ReplaceStr( _
ByVal pvarTextIn As Variant, _
ByVal pvarSearchStr As String, _
ByVal pvarReplacement As String, _
Optional ByVal pintCompMode As Integer = vbBinaryCompare)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Replaces the pvarSearchStr string with pvarReplacement string in the pvarTextIn string.
'                       Uses pintCompMode to determine comparison mode as an option
'
' Ver.  Date            Author              Details
' 1.00  06 MAR 2008     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strWorkText As String
Dim intPointer As Integer

  If IsNull(pvarTextIn) Then
    ReplaceStr = Null
  Else
    strWorkText = pvarTextIn
    intPointer = InStr(1, strWorkText, pvarSearchStr, pintCompMode)
    Do While intPointer > 0
      strWorkText = Left(strWorkText, intPointer - 1) & pvarReplacement & Mid(strWorkText, intPointer + Len(pvarSearchStr))
      intPointer = InStr(intPointer + Len(pvarReplacement), strWorkText, pvarSearchStr, pintCompMode)
    Loop
    ReplaceStr = strWorkText
  End If

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("ReplaceStr", "modCompare", Err.Number, Err.Description)
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
' 1.00  20 JUN 2002     Anthony Duguid      Initial version
' 1.01  22 FEB 2008     Anthony Duguid      added error line
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

