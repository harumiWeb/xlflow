Attribute VB_Name = "modFileFunctions"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Creates a filter string from the passed in arguments.
'
' Ver.  Date            Author              Details
' 1.00  01-JAN-2001     Microsoft           Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

'API Functions
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
    "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias _
    "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
Private Declare Function SHBrowseForFolder Lib "shell32" _
                                        (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                        (ByVal pidList As Long, _
                                        ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                        (ByVal lpString1 As String, ByVal _
                                        lpString2 As String) As Long

'constants
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Const ALLFILES = "All Files"
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHOWHELP = &H10

'types
Private Type BrowseInfo
         hWndOwner      As Long
         pIDLRoot       As Long
         pszDisplayName As Long
         lpszTitle      As Long
         ulFlags        As Long
         lpfnCallback   As Long
         lparam         As Long
         iImage         As Long
End Type

Private Type MSA_OPENFILENAME
    ' Filter string used for the Open dialog filters.
    ' Use MSA_CreateFilterString() to create this.
    ' Default = All Files, *.*
    strFilter As String
    ' Initial Filter to display.
    ' Default = 1.
    lngFilterIndex As Long
    ' Initial directory for the dialog to open in.
    ' Default = Current working directory.
    strInitialDir As String
    ' Initial file name to populate the dialog with.
    ' Default = "".
    strInitialFile As String
    strDialogTitle As String
    ' Default extension to append to file if user didn't specify one.
    ' Default = System Values (Open File, Save File).
    strDefaultExtension As String
    ' Flags (see constant list) to be used.
    ' Default = no flags.
    lngFlags As Long
    ' Full path of file picked.  When the File Open dialog box is
    ' presented, if the user picks a nonexistent file,
    ' only the text in the "File Name" box is returned.
    strFullPathReturned As String
    ' File name of file picked.
    strFileNameReturned As String
    ' Offset in full path (strFullPathReturned) where the file name
    ' (strFileNameReturned) begins.
    intFileOffset As Integer
    ' Offset in full path (strFullPathReturned) where the file extension begins.
    intFileExtension As Integer
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As Long
    nMaxCustrFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustrData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Public Function RefreshLinks(pstrFileName As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              refresh linked tables in a database
' Variables:            strFileName - FileName of Database
' Return:               True if successful.
'
' Ver.  Date            Author              Details
' 1.00  04 JUN 2008     Anthony  Duguid     Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim dbs As Database
Dim tdf As TableDef

    ' Loop through all tables in the database.
    Set dbs = CurrentDb
    For Each tdf In dbs.TableDefs
        ' If the table has a connect string, it's a linked table.
        If Len(tdf.Connect) > 0 Then
            tdf.Connect = ";DATABASE=" & pstrFileName
            Err = 0
            On Error Resume Next
            tdf.RefreshLink         ' Relink the table.
            If Err <> 0 Then
                RefreshLinks = False
                Exit Function
            End If
        End If
    Next tdf
    RefreshLinks = True        ' Relinking complete.

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("RefreshLinks", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Function GetDirectory(pstrTitle As String, pfrmSource As Form) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Opens a Treeview control that displays the directories in a computer, API wrapper for SHBrowseForFolder Function
' Variables:            szTitle - Text Prompt in Dialog Box
'                       frmSource - Form that is to act as owner of dialog (usually Me)
' Return:               Selected Directory Path
'
' Ver.  Date            Author              Details
' 1.00  04-JUN-2008     Anthony  Duguid     Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim lpIDList As Long
Dim sBuffer As String
Dim tBrowseInfo As BrowseInfo

    With tBrowseInfo
        .hWndOwner = pfrmSource.hwnd
        .lpszTitle = lstrcat(pstrTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        GetDirectory = sBuffer
    End If

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("GetDirectory", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Function FindFile( _
Optional pstrPath As String = "C:\", _
Optional pstrTitle As String = "Please Select a File", _
Optional pstrFilterTitle As String = "Excel Files", _
Optional pstrFilterValue As String = "*.xls") As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Displays the Open dialog box for the user to locate a File.
' Variables:            pstrPath - Initial Path to set dialog to
'                       pstrTitle - Title of the dialog box
'                       pstrFilterTitle - frendly name for type of files to be located (E.G. "Excel Files")
'                       pstrFilterValue - Wildcard Patern for Files (E.G. *.XLS)
' Return:               Returns the full path to File.
' Example:              Me.txtFilePath = FindFile() 'defaults for Excel
'                       Me.txtFilePath = FindFile("C:\", "Please Select a File", "Excel Files", "*.xls")
'
' Ver.  Date            Author              Details
' 1.00  02-NOV-2008     Anthony  Duguid     Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim msaof As MSA_OPENFILENAME

    ' Set options for the dialog box.
    msaof.strDialogTitle = pstrTitle
    msaof.strInitialDir = pstrPath
    msaof.strFilter = MSA_CreateFilterString(pstrFilterTitle, pstrFilterValue)
    
    ' Call the Open dialog routine.
    MSA_GetOpenFileName msaof
    
    ' Return the path and file name.
    FindFile = Trim(msaof.strFullPathReturned)

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("FindFile", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Sub ReturnAllFiles( _
ByVal pstrPath As String, _
Optional ByVal pstrFileType As String = "*.*")
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Return all files of a user defined type
'
' Ver.  Date            Author              Details
' 1.00  04 JUN 2008     Anthony  Duguid     Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim colFiles    As New Collection
Dim varFile     As Variant 'Path and File name variable in loop
Dim strFile     As String  'File Name
Dim strPath     As String  'File Path

    DoCmd.Hourglass True
    Call RecursiveDir(colFiles, pstrPath, pstrFileType, True)
    For Each varFile In colFiles
        strPath = varFile
        'strPath = ReplaceStr(strPath, "'", "!")
        strFile = CutLastWord(varFile, "", "\")
        Debug.Print strPath
        Debug.Print strFile
    Next varFile

ExitProcedure:
    On Error Resume Next
    Set colFiles = Nothing
    DoCmd.Hourglass False
    MsgBox "Done", vbInformation, "Completed"
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("ReturnAllFiles", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Function RecursiveDir( _
ByVal pcolFiles As Collection, _
ByVal pstrFolder As String, _
ByVal pstrFileSpec As String, _
ByVal pbIncludeSubfolders As Boolean) As Variant
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              loop through all directories and subdirectories
'
' Ver.  Date            Author              Details
' 1.00  04 JUN 2008     Anthony  Duguid     Initial version.
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
        Case Is <> 0
            Call ErrorMsg("RecursiveDir", "modFileFunctions", Err.Number, Err.Description)
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
' 1.00  04 JUN 2008     Anthony  Duguid     Initial version.
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
            Call ErrorMsg("TrailingSlash", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Function MSA_CreateFilterString(ParamArray varFilt() As Variant) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Creates a filter string from the passed in arguments.
' Notes:                Expects an even number of argumentss (filter name, extension), but
'                       if an odd number is passed in, it appends "*.*".
' Returns:              "" if no argumentss are passed in.
'
' Ver.  Date            Author              Details
' 1.00  01 JAN 2001     Microsoft           Initial version.
' 1.01  04 JUN 2008     Anthony  Duguid     added error trap
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strFilter As String
Dim intRet As Integer
Dim intNum As Integer

    intNum = UBound(varFilt)
    If (intNum <> -1) Then
        For intRet = 0 To intNum
            strFilter = strFilter & varFilt(intRet) & vbNullChar
        Next
        If intNum Mod 2 = 0 Then
            strFilter = strFilter & "*.*" & vbNullChar
        End If
        
        strFilter = strFilter & vbNullChar
    Else
        strFilter = ""
    End If
    MSA_CreateFilterString = strFilter
    
ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("MSA_CreateFilterString", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Function MSA_ConvertFilterString(pstrFilterIn As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Creates a filter string from a bar ("|") separated string.
' Notes:                The string should pairs of filter|extension strings, i.e. "Access Databases|*.mdb|All Files|*.*"
'                       If no extensions exists for the last filter pair, *.* is added.
'                       This code will ignore any empty strings, i.e. "||" pairs.
' Returns:              "" if the strings passed in is empty.
'
' Ver.  Date            Author              Details
' 1.00  01-JAN-2001     Microsoft           Initial version.
' 1.01  04-JUN-2008     Anthony  Duguid     added error trap
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strFilter As String
Dim intNum As Integer
Dim intPos As Integer
Dim intLastPos As Integer

    strFilter = ""
    intNum = 0
    intPos = 1
    intLastPos = 1

    ' Add strings as long as we find bars.
    ' Ignore any empty strings (not allowed).
    Do
        intPos = InStr(intLastPos, pstrFilterIn, "|")
        If (intPos > intLastPos) Then
            strFilter = strFilter & Mid(pstrFilterIn, intLastPos, intPos - intLastPos) & vbNullChar
            intNum = intNum + 1
            intLastPos = intPos + 1
        ElseIf (intPos = intLastPos) Then
            intLastPos = intPos + 1
        End If
    Loop Until (intPos = 0)
        
    ' Get last string if it exists (assuming strFilterIn was not bar terminated).
    intPos = Len(pstrFilterIn)
    If (intPos >= intLastPos) Then
        strFilter = strFilter & Mid(pstrFilterIn, intLastPos, intPos - intLastPos + 1) & vbNullChar
        intNum = intNum + 1
    End If
    
    ' Add *.* if there's no extension for the last string.
    If intNum Mod 2 = 1 Then
        strFilter = strFilter & "*.*" & vbNullChar
    End If
    
    ' Add terminating NULL if we have any filter.
    If strFilter <> "" Then
        strFilter = strFilter & vbNullChar
    End If
    MSA_ConvertFilterString = strFilter

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("MSA_ConvertFilterString", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Function MSA_GetSaveFileName(msaof As MSA_OPENFILENAME) As Integer
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Opens the file save dialog.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim of As OPENFILENAME
Dim intRet As Integer

    MSAOF_to_OF msaof, of
    of.Flags = of.Flags Or OFN_HIDEREADONLY
    intRet = GetSaveFileName(of)
    If intRet Then
        OF_to_MSAOF of, msaof
    End If
    MSA_GetSaveFileName = intRet

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("MSA_GetSaveFileName", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Function MSA_SimpleGetSaveFileName() As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Opens the file save dialog with default values.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim msaof As MSA_OPENFILENAME
Dim intRet As Integer
Dim strRet As String
    
    intRet = MSA_GetSaveFileName(msaof)
    If intRet Then
        strRet = msaof.strFullPathReturned
    End If
    MSA_SimpleGetSaveFileName = strRet

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("MSA_SimpleGetSaveFileName", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Function MSA_GetOpenFileName(msaof As MSA_OPENFILENAME) As Integer
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Opens the Open dialog.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim of As OPENFILENAME
Dim intRet As Integer

    MSAOF_to_OF msaof, of
    intRet = GetOpenFileName(of)
    If intRet Then
        OF_to_MSAOF of, msaof
    End If
    MSA_GetOpenFileName = intRet

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("MSA_GetOpenFileName", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Function MSA_SimpleGetOpenFileName() As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Opens the Open dialog with default values.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim msaof As MSA_OPENFILENAME
Dim intRet As Integer
Dim strRet As String
    
    intRet = MSA_GetOpenFileName(msaof)
    If intRet Then
        strRet = msaof.strFullPathReturned
    End If
    MSA_SimpleGetOpenFileName = strRet

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("MSA_SimpleGetOpenFileName", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Sub OF_to_MSAOF(of As OPENFILENAME, msaof As MSA_OPENFILENAME)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              This sub converts from the Win32 structure to the Microsoft Access structure.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
    msaof.strFullPathReturned = Left(of.lpstrFile, InStr(of.lpstrFile, vbNullChar) - 1)
    msaof.strFileNameReturned = of.lpstrFileTitle
    msaof.intFileOffset = of.nFileOffset
    msaof.intFileExtension = of.nFileExtension

ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("OF_to_MSAOF", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub MSAOF_to_OF(msaof As MSA_OPENFILENAME, of As OPENFILENAME)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              This sub converts from the Microsoft Access structure to the Win32 structure.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strFile As String * 512

    ' Initialize some parts of the structure.
    of.hWndOwner = Application.hWndAccessApp
    of.hInstance = 0
    of.lpstrCustomFilter = 0
    of.nMaxCustrFilter = 0
    of.lpfnHook = 0
    of.lpTemplateName = 0
    of.lCustrData = 0
    If msaof.strFilter = "" Then
        of.lpstrFilter = MSA_CreateFilterString(ALLFILES)
    Else
        of.lpstrFilter = msaof.strFilter
    End If
    of.nFilterIndex = msaof.lngFilterIndex
    of.lpstrFile = msaof.strInitialFile & String(512 - Len(msaof.strInitialFile), 0)
    of.nMaxFile = 511
    of.lpstrFileTitle = String(512, 0)
    of.nMaxFileTitle = 511
    of.lpstrTitle = msaof.strDialogTitle
    of.lpstrInitialDir = msaof.strInitialDir
    of.lpstrDefExt = msaof.strDefaultExtension
    of.Flags = msaof.lngFlags
    of.lStructSize = Len(of)

ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("MSAOF_to_OF", "modFileFunctions", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub
