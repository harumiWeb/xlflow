Attribute VB_Name = "modOutlook"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Outlook subroutines
'
' Ver.  Date            Author              Details
' 1.00  09 SEP 2009     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Sub ImportContactsFromOutlook( _
Optional pstrTableName As String = "tblEvents", _
Optional pstrMailbox As String = "YourMailBoxName")
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              load exchange contacts to table
'
' Ver.  Date            Author              Details
' 1.00  09 SEP 2009     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim oapp        As New Outlook.Application
Dim onsp        As Outlook.NameSpace
Dim ofdr        As Outlook.MAPIFolder
Dim orcp        As Outlook.Recipient
Dim oItem       As Outlook.AppointmentItem
Dim oItems      As Outlook.Items
Dim rst         As DAO.Recordset
Dim dblCnt      As Double
Dim i           As Integer

    Set onsp = oapp.GetNamespace("MAPI")
    Set ofdr = onsp.GetDefaultFolder(olFolderContacts)
    Set oItems = ofdr.Items
    dblCnt = oItems.Count
    If dblCnt <> 0 Then
        For i = 1 To dblCnt
            If TypeName(oItems(i)) = "ContactItem" Then
                Set oItem = oItems(i)
                rst.AddNew
                rst!FirstName = oItem.FirstName
                rst!LastName = oItem.LastName
                rst!Address = oItem.BusinessAddressStreet
                rst!City = oItem.BusinessAddressCity
                rst!State = oItem.BusinessAddressState
                rst!Zip_Code = oItem.BusinessAddressPostalCode
                ' Custom Outlook properties would look like this:
                ' rst!AccessFieldName = oItem.UserProperties("OutlookPropertyName")
                rst.Update
            End If
        Next i
        rst.Close
        MsgBox "Finished.", vbInformation, "Contact Import"
    Else
        MsgBox "No contacts to export.", vbInformation, "Contact Import"
    End If

ExitProcedure:
    On Error Resume Next
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("ImportContactsFromOutlook", "modOutlook", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Public Sub ImportCalendarFromOutlook( _
Optional pstrTableName As String = "tblEvents", _
Optional pstrMailbox As String = "YourMailBoxName", _
Optional pdblNbrDays As Double = 730) 'not more than 2 years
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Load exchange calendar to table
' Example:              Subject = "Item 1; Item 2; Item 3; Item 4"
'                       pstrMailbox  = "YourMailBoxName"
' Notes:                Requires "Microsoft Outlook ##.# Object Library"
'
' Ver.  Date            Author              Details
' 1.00  09 SEP 2009     Anthony Duguid      Initial version.
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim oapp        As New Outlook.Application
Dim onsp        As Outlook.NameSpace
Dim ofdr        As Outlook.MAPIFolder
Dim orcp        As Outlook.Recipient
Dim oItem       As Outlook.AppointmentItem
Dim oItems      As Outlook.Items
Dim rst         As DAO.Recordset
Dim arrList()   As String 'break up the subject into individual columns
Dim i           As Double

    Set rst = CurrentDb.OpenRecordset(pstrTableName, dbOpenDynaset, [dbSeeChanges])
    CurrentDb.Execute "DELETE * FROM " & pstrTableName, [dbSeeChanges]
    Set onsp = oapp.GetNamespace("MAPI")
    Set orcp = onsp.CreateRecipient(pstrMailbox)
    Set ofdr = onsp.GetSharedDefaultFolder(orcp, olFolderCalendar)
    Set oItems = ofdr.Items
    oItems.Sort "[Start]" 'must use this for reoccurring events
    oItems.IncludeRecurrences = True 'must use this for reoccurring events
    If oItems.Count <> 0 Then 'check to see if there are any events
        For i = 1 To oItems.Count 'loop through all events
            If TypeName(oItems(i)) = "AppointmentItem" Then   'get only events from calendar
                Set oItem = oItems(i)
                If oItem.GetRecurrencePattern.NoEndDate = False Then 'there is an end date
                    pdblNbrDays = CInt(DateDiff("d", rst!PatternStartDate, rst!PatternEndDate))
                End If
                If oItem.Start > DateAdd("d", pdblNbrDays, oItem.GetRecurrencePattern.PatternStartDate) Then Exit For
                rst.AddNew 'create record
                rst!Subject = oItem.Subject
                arrList = Split(oItem.Subject, ";")
                If UBound(arrList) >= 3 Then
                    rst!EventCategory = Trim(arrList(0))
                    rst!TimeGroupDesc = Trim(arrList(1))
                    rst!PayrollCode = Trim(arrList(2))
                    rst!CycleType = Trim(arrList(3))
                End If
                rst!STARTDATE = oItem.Start
                rst!EndDate = oItem.End
            End If
            rst.Update 'reload the table
        Next i
        rst.Close
        MsgBox "Finished.", vbInformation, "Calendar Import"
    Else
        MsgBox "No events to export.", vbInformation, "Calendar Import"
    End If

ExitProcedure:
    On Error Resume Next
    Set rst = Nothing
    Set oapp = Nothing
    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("ImportCalendarFromOutlook", "modMail", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Public Function PushAppointments()
On Error Resume Next
    Dim Outlook As Outlook.Application
    Dim onsp As Outlook.NameSpace
    Dim olRecip As Outlook.Recipient
    Dim olSharedFolder As Outlook.MAPIFolder
    Dim rs As ADODB.Recordset
    Dim addRS As ADODB.Recordset
    Dim olItems
    Dim itm
    
    Set rs = New ADODB.Recordset
    Set addRS = New ADODB.Recordset
    
    'Establish connection to mySQL
    'ConnectDB
    
    'Open users recordset
   ' rs.Open "SELECT * FROM users;", oConn, adOpenDynamic, adLockOptimistic
        
    'Set Outlook variables
    Set Outlook = New Outlook.Application
    Set onsp = Outlook.GetNamespace("MAPI")
    
    Do Until rs.EOF
    Set olRecip = onsp.CreateRecipient(rs.Fields("name"))
    Set olSharedFolder = onsp.GetSharedDefaultFolder(olRecip, olFolderCalendar)
    
olItems.Sort "[Start]"
olItems.IncludeRecurrences = True

    Set olItems = olSharedFolder.Items
    
    For Each itm In olItems
        If itm.Start > Date And itm.End < Date + 1 Then
               ' addRS.Open "INSERT INTO " & _
               '             "calendar ( uid, subject, location, date, start, end ) " & _
               '             "VALUES " & _
               '             "( '" & rs.Fields("id") & "', " & _
               '             "'" & itm.Subject & "', " & _
               '             "'" & itm.Location & "', " & _
               '             "'" & Format(itm.Start, "yyyy-mm-dd") & "', " & _
               '             "'" & Format(itm.Start, "hh:mm am/pm") & "', " & _
               '             "'" & Format(itm.End, "hh:mm am/pm") & "' )", oConn, adOpenDynamic, adLockOptimistic
        End If
    Next
    
    Debug.Print rs.Fields("name") & " DONE."
    rs.MoveNext
    
    Loop
 
End Function
