VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_800_Link_Manager_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Change or refresh ODBC links
'
' Ver.  Date            Author              Details
' 1.00  05-JAN-2001     Roger K Saunders    Initial version.
' 1.01  01-JUL-2010     Anthony L Duguid    error trapped and contained in one form
'--------------------------------------------------------------------------------------------------------------------
Option Explicit
Dim dbCurentDatabase        As Database
Dim gblnCancel              As Boolean 'used to cancel procedures
Dim gblnFirstConnect        As Boolean 'used to create the headers for the listbox only once
Dim strFailedConnect()      As String
Dim strSelectedConnects()   As String
Dim ArrayOfLinks()          As String
Dim varConnectStrings()     As Variant
Dim strFullConnectString    As String
Dim intTablesReConnected    As Integer
Dim intFailedConnect        As Integer
Dim intQueryConnect         As Integer
Dim intItemsMatching        As Integer
'SortArray
Dim varA                    As Variant
Dim varB                    As Variant
Dim varC                    As Variant
'Listbox objects
Dim lstOld                  As ListBox
Dim lstLnk                  As ListBox

Public Sub LoadArrayOfLinks(intSortCol As Integer)
On Error GoTo ErrTrap
Dim dbCurr  As Database
Dim tdf     As TableDef
Dim qdf     As QueryDef
Dim intLink As Integer
Dim varConn As Variant

    'varConn = Me.lstConnections
    'If IsNull(varConn) Then varConn = "ODBC;*"
    varConn = "ODBC;*"
    Set dbCurr = CurrentDb
    ReDim ArrayOfLinks(2, 0)            'Empty array for reloading
    intLink = -1                        'Will equal 0 in loop

'****************************** Load Names of Linked Tables into Array *****************************
    dbCurr.TableDefs.Refresh            'Refresh TableDefs collection.
    For Each tdf In dbCurr.TableDefs    'Loop through List of tables
        If tdf.Connect Like varConn And Not tdf.Name Like "*~*" Then  'If it is an ODBC Connection then load into array
            intLink = intLink + 1
            ReDim Preserve ArrayOfLinks(2, intLink)
            ArrayOfLinks(0, intLink) = "Table"
            ArrayOfLinks(1, intLink) = tdf.Name
            ArrayOfLinks(2, intLink) = tdf.Connect
       End If
    Next tdf

'************************** Load Names of Pass Through Queries into Array **************************
    dbCurr.QueryDefs.Refresh            ' Refresh QueryDefs collection.
    For Each qdf In dbCurr.QueryDefs    'Loop through List of Queries
        If qdf.Connect Like varConn And Not qdf.Name Like "*~*" Then   'If it is an ODBC Connection then load into array
            intLink = intLink + 1
            ReDim Preserve ArrayOfLinks(2, intLink)
            ArrayOfLinks(0, intLink) = "Query"
            ArrayOfLinks(1, intLink) = qdf.Name
            ArrayOfLinks(2, intLink) = qdf.Connect
       End If
    Next qdf
          
'******************** If It is set then force remove from memory when finished ********************
    'sort here
    'array.sort ArrayOfLinks (1)
    Call SortArray(ArrayOfLinks(), intSortCol)

ExitProcedure:
    On Error Resume Next
    Set tdf = Nothing
    Set qdf = Nothing
    Set dbCurr = Nothing
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 13 'Type mismatch
            gblnFirstConnect = False
            Resume ExitProcedure
        Case Is <> 0
            gblnFirstConnect = False
            Call ErrorMsg("LoadArrayOfLinks", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Public Function SortArray(TempArray As Variant, intCol As Integer)
On Error GoTo ErrTrap
Dim Temp        As Variant
Dim i           As Integer
Dim NoExchanges As Integer

    ' Loop until no more "exchanges" are made.
    Do
        NoExchanges = True
        ' Loop through each element in the array.
        For i = LBound(TempArray, 2) To UBound(TempArray, 2) - 1
            ' If the element is greater than the element
            ' following it, exchange the two elements.
            If TempArray(intCol, i) > TempArray(intCol, i + 1) Then
                NoExchanges = False
                varA = TempArray(0, i)
                varB = TempArray(1, i)
                varC = TempArray(2, i)
                TempArray(0, i) = TempArray(0, i + 1)
                TempArray(1, i) = TempArray(1, i + 1)
                TempArray(2, i) = TempArray(2, i + 1)
                TempArray(0, i + 1) = varA
                TempArray(1, i + 1) = varB
                TempArray(2, i + 1) = varC
            End If
        Next i
    Loop While Not (NoExchanges)

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("SortArray", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Sub LinkedTables(rsLinkedTables As Recordset)
On Error GoTo ErrTrap
Dim s As String
Dim dbCurr As Database

    Set dbCurr = CurrentDb()
                      s = "SELECT "
    s = s & vbCrLf & " Left([Connect],InStr([Connect],';')-1) AS DSN "
    s = s & vbCrLf & ", Left([ForeignName],InStr([ForeignName],'.')-1) AS Owner "
    s = s & vbCrLf & ", Mid([ForeignName],InStr([ForeignName],'.')+1) AS ForeignTable "
    s = s & vbCrLf & ", Name as Link_Name "
    s = s & vbCrLf & ", ForeignName "
    s = s & vbCrLf & ", Connect "
    s = s & vbCrLf & "FROM "
    s = s & vbCrLf & "MsysObjects "
    s = s & vbCrLf & "WHERE "
    's = s & vbCrLf & "[Connect] Not Like '*Microsoft*' And "
    s = s & vbCrLf & "[Type] = 4 "
    s = s & vbCrLf & "ORDER BY "
    s = s & vbCrLf & "1, 2, 4 "
    Set rsLinkedTables = dbCurr.OpenRecordset(s)

ExitProcedure:
    On Error Resume Next
    Set dbCurr = Nothing
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Debug.Print s
            Call ErrorMsg("LinkedTables", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub LoadSelectedConnects()
On Error GoTo ErrTrap
Dim intConnects As Integer
Dim bytSelected As Byte
Dim varItm As Variant

    ReDim strSelectedConnects(0)
    bytSelected = 0
    For Each varItm In lstOld.ItemsSelected
       bytSelected = bytSelected + 1
       ReDim Preserve strSelectedConnects(bytSelected)
       strSelectedConnects(bytSelected) = lstOld.Column(0, varItm) & "*"
    Next varItm

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("LoadSelectedConnects", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub chkIncludePassword_AfterUpdate()
On Error GoTo ErrTrap
Dim intStart As Integer
Dim intStop As Integer

    If Me.chkIncludePassword = True Then
        Me.txtNewConnect = strFullConnectString
     Else
        intStart = InStr(strFullConnectString, ";UID=")
        intStop = InStr(InStr(strFullConnectString, ";PWD=") + 4, strFullConnectString, ";") + 1
        Me.txtNewConnect = Left(strFullConnectString, intStart)
        If intStop > 1 Then
            Me.txtNewConnect = Me.txtNewConnect & Mid(strFullConnectString, intStop)
        End If
    End If

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("chkIncludePassword_AfterUpdate", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub cboListSort_AfterUpdate()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Allow sorting of List box
'
' Ver.  Date            Author              Details
' 1.00  05-JAN-2001     Roger K Saunders    Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Const intTwip As Integer = 1440 'use to convert from inches to twips - an Access unit of measure

    If Me.cboListSort.Value = 0 Then
        Me.lstObjects.ColumnWidths = CStr(0.5 * intTwip) & ";" & CStr(2.2 * intTwip) & ";" & CStr(3 * intTwip)
     Else
        Me.lstObjects.ColumnWidths = CStr(0 * intTwip) & ";" & CStr(2.2 * intTwip) & ";" & CStr(3 * intTwip)
    End If
    Me.lstObjects.Requery

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("cboListSort_AfterUpdate", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub cmdRefresh_Click()
On Error GoTo ErrTrap
Dim blnUpdate As Boolean
    
    blnUpdate = IIf(Len(Me.txtNewConnect) > 0, True, False)
    Me.cmdCancel.Caption = "&Cancel"
    Select Case Me.tabConnections
        Case Me.pgeConnections.PageIndex 'ODBC Connect Strings
            If Me.lstConnections.ItemsSelected.Count = 0 Then GoTo ExitProcedure
            Call UpdateConnections(blnUpdate) 'Update Links by Connect Sting
        Case Me.pgeObjects.PageIndex 'Linked Objects
            If Me.lstObjects.ItemsSelected.Count = 0 Then GoTo ExitProcedure
            Call UpdateObjects(blnUpdate) 'Update Link by Selected Link
    End Select
    Call LoadArrayOfLinks(Me.cboListSort) 'Reload Array For List box
    Me.lstObjects.Requery

ExitProcedure:
    On Error Resume Next
    Me.cmdCancel.Caption = "&Close"
    Me.cmdCancel.SetFocus
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            'Call ErrorMsg("cmdRefresh_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Public Sub UpdateConnections(pblnChange As Boolean)
On Error GoTo ErrTrap  'Flags Message for the help desk if there is a problem
Dim intTotalConnect As Integer
Dim intConnects As Integer
Dim strExistingConnect As String
Dim strNewConnect As String
Dim strConnect As String
Dim fConnect As Boolean
Dim qdf As QueryDef
Dim tdf As TableDef
Dim intLastLength As Integer
Dim intCurLength As Integer

    DoCmd.Hourglass True
    intTablesReConnected = 0
    intFailedConnect = 0
    intQueryConnect = 0
    intItemsMatching = 0
    gblnCancel = False
    fConnect = True
    Me.txtReconnected = ""
    Me.txtConnectFailed = ""
    If pblnChange = True Then        'If Passed In Parameter for new connection says to change connection
        strNewConnect = Me.txtNewConnect & ""
    End If

    Call LoadSelectedConnects              'Loads Selected Connect Strings into arrays

'************************************* Update Table Attachments ************************************
   dbCurentDatabase.TableDefs.Refresh
   For Each tdf In dbCurentDatabase.TableDefs
      DoEvents
      If gblnCancel = True Then
        DoCmd.Hourglass False
        Call FillConnectionList
      End If
      strExistingConnect = tdf.Connect
      If strExistingConnect Like "ODBC*" Then
         For intConnects = 1 To UBound(strSelectedConnects, 1)
            If strExistingConnect Like strSelectedConnects(intConnects) Then
               intItemsMatching = intItemsMatching + 1
               Call UpdateLinks("Table", tdf.Name, strNewConnect, intItemsMatching)
            End If
         Next intConnects
      End If
   Next tdf

'****************************** Update Pass Through Query Attachments ******************************
   dbCurentDatabase.QueryDefs.Refresh
   For Each qdf In dbCurentDatabase.QueryDefs
      DoEvents
      If gblnCancel = True Then
        DoCmd.Hourglass False
        Call FillConnectionList
      End If
      strExistingConnect = qdf.Connect
      If strExistingConnect Like "ODBC*" Then
         For intConnects = 1 To UBound(strSelectedConnects, 1)
            If strExistingConnect Like strSelectedConnects(intConnects) Then
                intItemsMatching = intItemsMatching + 1
                Call UpdateLinks("Query", qdf.Name, strNewConnect, intItemsMatching)
               intLastLength = Len(Me.txtReconnected)
               End If
         Next intConnects
      End If
   Next qdf

ExitProcedure:
    On Error Resume Next
    Call FillConnectionList
    DoCmd.Hourglass False
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            'Call ErrorMsg("UpdateConnections", Me.Name, Err.Number, Err.Description)
            Resume Next
        Case Else
            Resume Next
    End Select

End Sub

Private Sub Form_Close()
On Error Resume Next

    DoCmd.Hourglass False
    
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrTrap
   
   If Me.cmdCancel.Caption = "&Close" Then
      DoCmd.Close
    Else
      gblnCancel = True
      Me.cmdCancel.Caption = "&Close"
   End If

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("cmdCancel_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo ErrTrap
    
    Me.lstConnections.RowSourceType = "RowSourceTypeConnections"
    Me.lstObjects.RowSourceType = "RowSourceTypeObjects"
    
    Set lstOld = Me.lstConnections
    Set lstLnk = Me.lstObjects
    Set dbCurentDatabase = CurrentDb()           'Current Database

    Call FillConnectionList        'Load List of ODBC Attachments
    Call cboListSort_AfterUpdate

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("Form_Open", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Function RowSourceTypeConnections(fld As Control, Id As Variant, Row As Variant, Col As Variant, Code As Variant) As Variant
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              The following function uses a Select Case statement to fill a
'                       multiple column list box. The function fills the first
'                       column of the list box with the dates of the next four Mondays. The
'                       second column is filled with the dates of the next four Tuesdays.
' Reference:            ACC2000: How to Use Code to Fill a Multiple-Column List Box
'
' Ver.  Date            Author              Details
' 1.00  05-JAN-2001     Microsoft           Initial version.
' 1.01  05-JAN-2001     Roger K Saunders    edited for specific listbox
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Static bytArraySize As Integer
Dim ReturnVal As Variant
    
    ReturnVal = Null
    Select Case Code
        Case acLBInitialize           ' Initialize.
            bytArraySize = Nz(UBound(varConnectStrings, 2), 0)
            ReturnVal = bytArraySize
        Case acLBOpen                 ' Open.
            ReturnVal = Timer          ' Generate unique ID for control.
        Case acLBGetRowCount          ' Get number of rows.
            ReturnVal = bytArraySize + 1
        Case acLBGetColumnCount    ' Get number of columns.
            ReturnVal = UBound(varConnectStrings, 1) + 1
        Case acLBGetColumnWidth    ' Column width.
        
        Case acLBGetValue             ' Get data.
            ReturnVal = varConnectStrings(Col, Row)
        Case acLBEnd                  ' End.
        
    End Select
    RowSourceTypeConnections = ReturnVal

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is = 6
            bytArraySize = 0
            Resume ExitProcedure
        Case Is = 9 'Subscript out of range (error on acLBInitialize)
            RowSourceTypeConnections = Null
            Resume ExitProcedure
        Case Is <> 0
            RowSourceTypeConnections = Null
            Call ErrorMsg("RowSourceTypeConnections", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Public Function RowSourceTypeObjects(fld As Control, Id As Variant, Row As Variant, Col As Variant, Code As Variant) As Variant
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              The following function uses a Select Case statement to fill a
'                       multiple column list box. The function fills the first
'                       column of the list box with the dates of the next four Mondays. The
'                       second column is filled with the dates of the next four Tuesdays.
' Reference:            ACC2000: How to Use Code to Fill a Multiple-Column List Box
'
' Ver.  Date            Author              Details
' 1.00  05-JAN-2001     Microsoft           Initial version.
' 1.01  05-JAN-2001     Roger K Saunders    edited for specific listbox
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    Static intArraySize As Integer
    Dim ReturnVal As Variant
    ReturnVal = Null
    
    Select Case Code
     Case acLBInitialize                         ' Initialize.
        Call LoadArrayOfLinks(Me.cboListSort)    'Call Function to fill array Link Names)
        intArraySize = UBound(ArrayOfLinks, 2)
        ReturnVal = intArraySize
     Case acLBOpen                               ' Open.
        ReturnVal = Timer                        ' Generate unique ID for control.
     Case acLBGetRowCount                        ' Get number of rows.
        ReturnVal = intArraySize + 1
     Case acLBGetColumnCount                     ' Get number of columns.
        ReturnVal = UBound(ArrayOfLinks, 1) + 1
     Case acLBGetColumnWidth                     ' Column width.
     Case acLBGetValue                           ' Get data.
         ReturnVal = ArrayOfLinks(Col, Row)
     Case acLBEnd                                ' End.
    End Select
    RowSourceTypeObjects = ReturnVal

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("RowSourceTypeObjects", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Private Sub FillConnectionList()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              ReLoad List of ODBC connections
'
' Ver.  Date            Author              Details
' 1.00  05-JAN-2001     Roger K Saunders    Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next

   ReDim varConnectStrings(3, 0)
   Call ArrayOfODBCConnects
   Me.lstConnections.RowSourceType = "RowSourceTypeConnections"

End Sub

Public Sub ArrayOfODBCConnects()
On Error GoTo ErrTrap
Dim qdf As QueryDef
Dim tdf As TableDef
Dim strObject As String

    gblnFirstConnect = True
    dbCurentDatabase.TableDefs.Refresh           ' Refresh TableDefs collection.
    For Each tdf In dbCurentDatabase.TableDefs
       If tdf.Connect Like "ODBC;*" And Not tdf.Name Like "*~*" Then
          Call LoadConnectArray(tdf.Connect, 1, 0)
       End If
    Next tdf
    
    dbCurentDatabase.QueryDefs.Refresh           ' Refresh QueryDefs collection.
    For Each qdf In dbCurentDatabase.QueryDefs
       If qdf.Connect Like "ODBC;*" And Not qdf.Name Like "*~*" Then
          Call LoadConnectArray(qdf.Connect, 0, 1)
       End If
    Next qdf

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 3258
            Call ErrorMsg("ArrayOfODBCConnects", Me.Name, Err.Number, Err.Description, Erl, "Query definition error")
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("ArrayOfODBCConnects", Me.Name, Err.Number, Err.Description, Erl)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Public Sub LoadConnectArray(ByVal strConnect As String, bytTable As Byte, bytQuery As Byte)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:
' Example:              varConnectStrings(0, # ) = ODBC Connection
'                       varConnectStrings(1, # ) = Qty Tables with connection
'                       varConnectStrings(2, # ) = Qty Queries with connection
'                       varConnectStrings(3, # ) = ODBC Connection For Display
'
' Ver.  Date            Author              Details
' 1.00  05-JAN-2001     Roger K Saunders    Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim bytArraySize As Integer
Dim bytConnectString As Integer

If strConnect Like "*DATABASE=" Then      'Strip the DATABASE= from the string
   strConnect = Left(strConnect, Len(strConnect) - 9)
End If

If gblnFirstConnect = True Then
   gblnFirstConnect = False
   bytArraySize = 1
   ReDim varConnectStrings(3, 1)
   varConnectStrings(0, 0) = "Connection"
   varConnectStrings(1, 0) = "tbl"
   varConnectStrings(2, 0) = "qry"
   varConnectStrings(3, 0) = "Connection"
   varConnectStrings(1, 1) = bytTable
   varConnectStrings(2, 1) = bytQuery
   varConnectStrings(0, 1) = strConnect
   varConnectStrings(3, 1) = strConnect
   

 Else
   For bytArraySize = 1 To UBound(varConnectStrings, 2)
      If varConnectStrings(0, bytArraySize) = strConnect Then
         'DBQ is already loaded into array so, increment count but do not load again
         varConnectStrings(1, bytArraySize) = varConnectStrings(1, bytArraySize) + bytTable
         varConnectStrings(2, bytArraySize) = varConnectStrings(2, bytArraySize) + bytQuery
         GoTo Exit_LoadConnectArray
      End If
   Next bytArraySize
   bytArraySize = UBound(varConnectStrings, 2) + 1
   ReDim Preserve varConnectStrings(3, bytArraySize)
   For bytConnectString = bytArraySize To 1 Step -1
      If varConnectStrings(0, bytConnectString - 1) > strConnect Then
         varConnectStrings(0, bytConnectString) = varConnectStrings(0, bytConnectString - 1)
         varConnectStrings(1, bytConnectString) = varConnectStrings(1, bytConnectString - 1)
         varConnectStrings(2, bytConnectString) = varConnectStrings(2, bytConnectString - 1)
         varConnectStrings(3, bytConnectString) = varConnectStrings(0, bytConnectString - 1)
       Else
         varConnectStrings(0, bytConnectString) = strConnect
         varConnectStrings(1, bytConnectString) = bytTable
         varConnectStrings(2, bytConnectString) = bytQuery
         varConnectStrings(3, bytConnectString) = strConnect
         GoTo Exit_LoadConnectArray
      End If
   Next bytConnectString
End If

Exit_LoadConnectArray:

End Sub

Private Sub lstObjects_AfterUpdate()
On Error GoTo ErrTrap
Dim varItm As Variant
        
    Me.txtReconnected = Null
    Me.txtConnectFailed = Null

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("lstObjects_AfterUpdate", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub lstConnections_AfterUpdate()
On Error GoTo ErrTrap
Dim varItm As Variant
Dim intTableTotal As Integer
Dim intQueryTotal As Integer

    intTableTotal = 0
    intQueryTotal = 0
    For Each varItm In lstOld.ItemsSelected
       intTableTotal = intTableTotal + lstOld.Column(1, varItm)
       intQueryTotal = intQueryTotal + lstOld.Column(2, varItm)
    Next varItm
    Me.txtTablesSelected = intTableTotal
    Me.txtQueriesSelected = intQueryTotal
    Me.txtReconnected = Null
    Me.txtConnectFailed = Null
    
ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("lstConnections_AfterUpdate", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub tabConnections_Change()
On Error GoTo ErrTrap

    Select Case Me.tabConnections
     Case Me.pgeConnections.PageIndex 'ODBC Connect Strings
        Call lstConnections_AfterUpdate
     Case Me.pgeObjects.PageIndex 'Linked Objects
        Call lstObjects_AfterUpdate
        'Call LoadArrayOfLinks(Me.cboListSort) 'Reload Array For List box
        'Me.lstObjects.Requery
    End Select
    
ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("tabConnections_Change", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub txtConnectFailed_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTrap
Dim intNotValidConnect As Integer
Dim strMsg As String
Dim bytResponse As Byte

    If KeyCode = 46 Then ' user pressed the Delete button
        DoCmd.Hourglass True        'Turn on Hourglass
        strMsg = "Please Confirm You have chosen to"
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & vbCrLf & "DELETE Links to "
        strMsg = strMsg & vbCrLf & CStr(UBound(strFailedConnect, 1))
        strMsg = strMsg & "  ODBC Tables"
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & vbCrLf & "        -Yes- to Delete Links"
          
        bytResponse = MsgBox(strMsg, vbYesNo + vbDefaultButton2, "Confirm Link Delete")
        
        If bytResponse = 6 Then            'If User selects Yes
           For intNotValidConnect = 1 To UBound(strFailedConnect, 1)
              DoCmd.DeleteObject acTable, strFailedConnect(intNotValidConnect)
           Next intNotValidConnect
           Me.txtConnectFailed = ""
           Call FillConnectionList        'Load List of ODBC Attachments
           Me.cmdCancel.SetFocus
         Else
        End If
        
    End If

ExitProcedure:
    On Error Resume Next
    DoCmd.Hourglass False
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("txtConnectFailed_KeyDown", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub txtNewConnect_DblClick(Cancel As Integer)
On Error GoTo ErrTrap
Dim MyWS As Workspace
Dim MyDB As Database
Dim conPubs As Connection

    Set MyWS = DBEngine.Workspaces(0)
    
ReTryConnection:
    Set MyDB = MyWS.OpenDatabase("", False, False, "ODBC;")
    strFullConnectString = MyDB.Connect
    If Len(strFullConnectString) > 0 Then
       Call chkIncludePassword_AfterUpdate
       Me.txtNewConnect.Tag = ""
    End If

ExitProcedure:
    On Error Resume Next
    MyDB.Close ' Close database but keep connection.
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 91
            If MsgBox("Connect String was not Successfully built, Do you wish to retry?", vbRetryCancel) = vbRetry Then
                Resume ReTryConnection
            Else
                Resume ExitProcedure
            End If
        Case Is <> 0
            'Call ErrorMsg("txtNewConnect_DblClick", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub txtNewConnect_Enter()
On Error GoTo ErrTrap

   Me.txtNewConnect.Tag = "Typed Connection"
  
ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("txtNewConnect_Enter", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Private Sub UpdateObjects(pblnChange As Boolean)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Change or refresh connection
' Example:              Call UpdateObjects(False)
'
' Ver.  Date            Author              Details
' 1.00  05-JAN-2001     Roger K Saunders    Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo 0
Dim varItm As Variant
Dim tdf As TableDef
Dim strNewConnect As String

    intItemsMatching = 0
    DoCmd.Hourglass True
    gblnCancel = False
    If pblnChange = True Then 'Changes Connection otherwise just refreshes it
        strNewConnect = Me.txtNewConnect & ""
    End If
    For Each varItm In lstLnk.ItemsSelected
      DoEvents
      If gblnCancel = True Then
        GoTo ExitProcedure
      End If
        intItemsMatching = intItemsMatching + 1
        Call UpdateLinks(lstLnk.Column(0, varItm), lstLnk.Column(1, varItm), strNewConnect, intItemsMatching)
        lstLnk.Selected(varItm) = False
    Next varItm

ExitProcedure:
    On Error Resume Next
    Set tdf = Nothing
    Call FillConnectionList
    DoCmd.Hourglass False
    Exit Sub
    
ErrTrap:


End Sub

Private Function UpdateLinks(strType As String, strName As String, strNewConnect As String, intLink As Integer) As Integer
On Error GoTo ErrTrap   'Flags Message for the help desk if there is a problem
Dim tdf As TableDef
Dim qdf As QueryDef
Dim fConnect As Boolean
Dim intLastLength As Integer
Dim intCurLength As Integer
'Initilize variables
    fConnect = True
    If intLink = 1 Then     'this is the first link of the set so reset messages
        intTablesReConnected = 0
        intFailedConnect = 0
        intQueryConnect = 0
        Me.txtReconnected = ""
        Me.txtConnectFailed = ""
    End If

    Me.Repaint

    Select Case strType
'************************************* Update Table Attachments ************************************
     Case "Table"
        Set tdf = dbCurentDatabase.TableDefs(strName)
        If strNewConnect > "" Then        'If New Connection is passed in then use it
            tdf.Connect = strNewConnect
        End If
52      tdf.RefreshLink          'Refresh Attachments
        If fConnect = True And (gblnCancel = False Or tdf.Connect = strNewConnect) Then
        '************************* Build Message for Successful Links ******************************
            intLastLength = Len(Me.txtReconnected)
            Me.txtReconnected = Me.txtReconnected & "table: " & tdf.Name & vbCrLf
            intCurLength = Len(Me.txtReconnected) - intLastLength - 7
            Me.txtReconnected.SetFocus
            Me.txtReconnected.SelStart = intLastLength + 7
            Me.txtReconnected.SelLength = intCurLength
            intTablesReConnected = intTablesReConnected + 1
         Else
        '************************** Build Message for Failed Links *********************************
            intLastLength = Len(Me.txtConnectFailed)
            Me.txtConnectFailed = Me.txtConnectFailed & "table: " & tdf.Name & vbCrLf
            intCurLength = Len(Me.txtConnectFailed) - intLastLength - 7
            Me.txtConnectFailed.SetFocus
            Me.txtConnectFailed.SelStart = intLastLength + 7
            Me.txtConnectFailed.SelLength = intCurLength
            intFailedConnect = intFailedConnect + 1
            ReDim Preserve strFailedConnect(intFailedConnect)
            strFailedConnect(intFailedConnect) = tdf.Name
            fConnect = True      'reset flag
        End If
'****************************** Update Pass Through Query Attachments ******************************
     Case "Query"
        If Len(strNewConnect) > 0 Then  'If New Connection is passed in then use it
            Set qdf = dbCurentDatabase.QueryDefs(strName)
            qdf.Connect = strNewConnect
            Me.txtReconnected = Me.txtReconnected & "query: " & qdf.Name & vbCrLf
            intCurLength = Len(Me.txtReconnected) - intLastLength - 7
            Me.txtReconnected.SetFocus
            Me.txtReconnected.SelStart = intLastLength + 7
            Me.txtReconnected.SelLength = intCurLength
            intQueryConnect = intQueryConnect + 1
        End If
     End Select

ExitProcedure:
    UpdateLinks = Abs(fConnect)
    Exit Function

ErrTrap:
    If Erl = 52 Then
       fConnect = False
       Resume Next
    End If
    Resume Next
    
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
' 1.01  22-FEB-2008     Anthony Duguid      added error line.
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim strMsg As String

    strMsg = "Contact your system administrator."
    strMsg = strMsg & vbCrLf & "Module: " & pstrModule
    strMsg = strMsg & vbCrLf & "Procedure: " & pstrProcedure
    strMsg = strMsg & IIf(pvarErrLine = 0, "", vbCrLf & "Error Line: " & pvarErrLine)
    strMsg = strMsg & vbCrLf & "Error #: " & dblErrNbr
    strMsg = strMsg & vbCrLf & "Error Description: " & strErrDes
    MsgBox strMsg, vbCritical, pstrTitle

End Sub
