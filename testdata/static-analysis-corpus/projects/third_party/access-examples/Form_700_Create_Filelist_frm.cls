VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_700_Create_Filelist_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error Resume Next

    Me.Refresh
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:
'
' Ver.  Date            Author              Details
' 1.00  18-NOV-2004     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strDescription As String
Dim strPathName As String
Dim strFileType As String

    strDescription = Me.txtDescription
    strPathName = Me.txtPathName
    strFileType = Me.cboFileType
    Call GetFiles(strDescription, strPathName, strFileType)
    
ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("cmdOk_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Sub GetFiles( _
ByVal pstrDescription As String, _
ByVal pstrPath As String, _
Optional ByVal pstrFileType As String = "*.dwg")
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Get all files of a user defined type
'
' Ver.  Date            Author              Details
' 1.00  03-JUN-2008     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim colFiles    As New Collection
Dim varFile     As Variant 'Path and File name variable in loop
Dim strFile     As String  'File Name
Dim strPath     As String  'File Path
Dim intCount    As Integer 'File count

    Call RecursiveDir(colFiles, pstrPath, pstrFileType, True)
    DoCmd.Hourglass True
    Application.SetOption "Confirm Action Queries", False

    For Each varFile In colFiles
        strPath = varFile
        strPath = ReplaceStr(strPath, "'", "!")
        strFile = CutLastWord(varFile, "", "\")
        Call InsertFiles(pstrDescription, strPath, strFile)
        intCount = intCount + 1
        Me.lblFileCount.Caption = "File Count = " & intCount
        Me.Repaint
    Next varFile

ExitProcedure:
    On Error Resume Next
    Application.SetOption "Confirm Action Queries", True
    Set colFiles = Nothing
    DoCmd.Hourglass False
    MsgBox "Done", vbInformation, "Completed"
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("GetFiles", "modCompare", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Function InsertFiles( _
ByVal pstrDesc As String, _
ByVal pstrPath As String, _
ByVal pstrFile As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Insert records for the comparison
'
' Ver.  Date            Author              Details
' 1.00  03-JUN-2008     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strSql As String

                      strSql = "INSERT INTO "
    strSql = strSql & vbCrLf & "tblFileList ("
    strSql = strSql & vbCrLf & "FILE_DESC, "
    strSql = strSql & vbCrLf & "PATH_NAME, "
    strSql = strSql & vbCrLf & "FILE_NAME "
    strSql = strSql & vbCrLf & ") "
    strSql = strSql & vbCrLf & "VALUES ("
    strSql = strSql & vbCrLf & "'" & pstrDesc & "', "
    strSql = strSql & vbCrLf & "'" & pstrPath & "', "
    strSql = strSql & vbCrLf & "'" & pstrFile & "' "
    strSql = strSql & vbCrLf & ")"
    'Debug.Print strSql
    DoCmd.RunSQL strSql

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("InsertFiles", "modCompare", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function RecursiveDir( _
ByVal pcolFiles As Collection, _
ByVal pstrFolder As String, _
ByVal pstrFileSpec As String, _
ByVal pbIncludeSubfolders As Boolean) As Variant
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              loop through all directories and subdirectories
'
' Ver.  Date            Author              Details
' 1.00  03-JUN-2008     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim colFolders  As New Collection
Dim varTemp     As Variant
Dim varFolder   As Variant

    'Add files in strFolder matching strFileSpec to colFiles
    pstrFolder = TrailingSlash(pstrFolder)
    varTemp = Dir(pstrFolder & pstrFileSpec)
    Do While varTemp <> vbNullString
        pcolFiles.Add pstrFolder & varTemp
        varTemp = Dir
    Loop

    If pbIncludeSubfolders Then
        'Fill colFolders with list of subdirectories of strFolder
        varTemp = Dir(pstrFolder, vbDirectory)
        Do While varTemp <> vbNullString
            If (varTemp <> ".") And (varTemp <> "..") Then
                If (GetAttr(pstrFolder & varTemp) And vbDirectory) <> 0 Then
                    colFolders.Add varTemp
                End If
            End If
            varTemp = Dir
        Loop

        'Call RecursiveDir for each subfolder in colFolders
        For Each varFolder In colFolders
            Call RecursiveDir(pcolFiles, pstrFolder & varFolder, pstrFileSpec, True)
        Next varFolder
    End If

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is = 75 'Access Denied to folder
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("RecursiveDir", "modCompare", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function TrailingSlash(ByVal pstrFolder As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              handles the backslash on directory path
'
' Ver.  Date            Author              Details
' 1.00  03-JUN-2008     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    If Len(pstrFolder) > 0 Then
        If Right(pstrFolder, 1) = "\" Then
            TrailingSlash = pstrFolder
        Else
            TrailingSlash = pstrFolder & "\"
        End If
    End If
    
ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("TrailingSlash", "modCompare", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function CutFirstWord( _
ByVal pvarWord As Variant, _
ByRef pvarRemainder As Variant, _
Optional ByVal pstrSearch As String = " ") As Variant
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              CutFirstWord:returns the first word in pvarWord.  varRemainder:returns the rest
'
' Ver.  Date            Author              Details
' 1.00  03-JUN-2008     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim varTemp     As Variant
Dim intP        As Integer

    varTemp = Trim(pvarWord)
    intP = InStr(varTemp, pstrSearch)
    If intP = 0 Then
        CutFirstWord = varTemp
        pvarRemainder = Null
    Else
        CutFirstWord = Left(varTemp, intP - 1)
        pvarRemainder = Trim(Mid(varTemp, intP + 1))
    End If

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("CutFirstWord", "modCompare", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function CutLastWord( _
ByVal pvarWord As Variant, _
ByRef pvarRemainder As Variant, _
Optional ByVal pstrSearch As String = " ") As Variant
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              CutLastWord:returns the first word in pvarWord.  varRemainder:returns the rest
'
' Ver.  Date            Author              Details
' 1.00  03-JUN-2008     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim varTemp     As Variant
Dim intP        As Integer
Dim intX        As Integer

    varTemp = Trim(pvarWord)
    intP = 1
    For intX = Len(varTemp) To 1 Step -1
        If Mid(varTemp, intX, 1) = pstrSearch Then
            intP = intX + 1
            Exit For
        End If
    Next intX
    If intP = 1 Then
        CutLastWord = varTemp
        pvarRemainder = Null
    Else
        CutLastWord = Mid(varTemp, intP)
        pvarRemainder = Trim(Left(varTemp, intP - 1))
    End If

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("CutLastWord", "modCompare", Err.Number, Err.Description)
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
' 1.00  03-JUN-2008     Anthony Duguid      Initial version.
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
' 1.00  20-JUN-2002     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim strMsg As String

    strMsg = "Contact your system administrator."
    strMsg = strMsg & vbCrLf & "Module: " & pstrModule
    strMsg = strMsg & vbCrLf & "Procedure: " & pstrProcedure
    strMsg = strMsg & IIf(pvarErrLine = 0, "", vbCrLf & "Error Line: " & pvarErrLine & vbCrLf)
    strMsg = strMsg & "Error #: " & dblErrNbr & vbCrLf
    strMsg = strMsg & "Error Description: " & strErrDes
    MsgBox strMsg, vbCritical, pstrTitle

End Sub
