VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_805_Timeline_Report_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Time Line Report "Canned"
'
' Ver.  Date            Author              Details
' 1.00  11 JUN 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Private mdatEarliest     As Date
Private mdatLatest       As Date
Private mintDayDiff      As Double
Private bNoData          As Boolean

Private Sub Report_Activate()
On Error GoTo ErrTrap
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Get the minimum, maximum, and difference of two dates in the recordset
' Notes:                Must use a saved query to run this report
'
' Ver.  Date            Author              Details
' 1.00  11 JUN 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Dim dteMinStart         As Date
Dim dteMaxStart         As Date
Dim dteMinEnd           As Date
Dim dteMaxEnd           As Date

    If bNoData Then GoTo ExitProcedure ' no data for this report.
    dteMinStart = DateValue(DMin("[" & Me.txtStartDate.ControlSource & "]", Me.RecordSource, "IsNull([" & Me.txtStartDate.ControlSource & "]) = False"))
    dteMaxStart = DateValue(DMax("[" & Me.txtStartDate.ControlSource & "]", Me.RecordSource, "IsNull([" & Me.txtStartDate.ControlSource & "]) = False"))
    dteMinEnd = DateValue(DMin("[" & Me.txtEndDate.ControlSource & "]", Me.RecordSource, "IsNull([" & Me.txtEndDate.ControlSource & "]) = False"))
    dteMaxEnd = DateValue(DMax("[" & Me.txtEndDate.ControlSource & "]", Me.RecordSource, "IsNull([" & Me.txtEndDate.ControlSource & "]) = False"))
    If dteMinEnd < dteMinStart Then ' get the lowest of the two date ranges
        dteMinStart = dteMinEnd
    End If
    If dteMaxEnd < dteMaxStart Then ' get the highest of the two date ranges
        dteMaxEnd = dteMaxStart
    End If
    mdatEarliest = dteMinStart
    mdatLatest = dteMaxEnd
    mintDayDiff = DateDiff("d", mdatEarliest, mdatLatest)
    If mintDayDiff > 3650 Then ' if you want to see more than 10 years get a copy of MS Project
        MsgBox "Please select a smaller date range." & vbCrLf & "Your date range exceeds 10 years." & vbCrLf & "Reset the " & Me.txtStartDate.ControlSource & " and " & Me.txtEndDate.ControlSource, vbInformation, "Exceeded Maximum Range."
        DoCmd.Close
        End
    End If
    Me.txtMinStartDate.Caption = Format(mdatEarliest, "dd/mm/yyyy")
    Me.txtMaxEndDate.Caption = Format(mdatLatest, "dd/mm/yyyy")

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("Report_Activate", Me.Name, Err.Number, Err.Description, Erl)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub Report_NoData(Cancel As Integer)
On Error Resume Next
'must make sure that there is data to graph

    bNoData = True

End Sub

Private Sub secDetail_Format(Cancel As Integer, FormatCount As Integer)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Compares two dates and gives a time line, negative days are welcome.
'
' Ver.  Date            Author              Details
' 1.00  16 JUN 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim intStartDayDiff     As Double
Dim intDayDiff          As Double
Dim sngFactor           As Double
Dim intDay              As Double
Dim dteStart            As Date
Dim dteEnd              As Date
Dim dblLabelLeft        As Double

    'can't go over the page.  need to do this every time.
    Me.shpGrowForDateNeg.Width = 0.0001
    Me.shpGrowForDateNeg.Left = Me.shpMaxDays.Left
    Me.shpGrowForDate.Width = 0.0001
    Me.shpGrowForDate.Left = Me.shpMaxDays.Left
    Me.shpGrowForDateNeg.Visible = False
    Me.shpGrowForDate.Visible = False
    Me.txtTotalDays.Visible = False

    If bNoData Or IsNull(Me.txtStartDate) And IsNull(Me.txtEndDate) Then
        GoTo ExitProcedure 'there is no data to graph for this record.
    End If
    Me.txtTotalDays.Visible = True
    If mintDayDiff = 0 Then mintDayDiff = 1
    sngFactor = Me.shpMaxDays.Width / mintDayDiff
    Me.ScaleMode = 1 'Twips
    If Not IsNull(Me.txtStartDate) And Not IsNull(Me.txtEndDate) Then
        dteStart = DateValue(Me.txtStartDate)
        dteEnd = DateValue(Me.txtEndDate)
        'If dteStart = #10/18/2003# And dteEnd = #10/20/2003# Then Stop       'for testing only
        intDayDiff = DateDiff("d", dteStart, dteEnd)
        Me.txtTotalDays = intDayDiff & " Day(s)"
        If dteEnd < dteStart Then 'negative values handled here.
            intStartDayDiff = (DateDiff("d", mdatEarliest, dteEnd))
            Me.shpGrowForDate.Visible = False
            Me.shpGrowForDateNeg.Visible = True
            Me.shpGrowForDateNeg.Width = Abs(intDayDiff) * sngFactor
            Me.shpGrowForDateNeg.Left = Me.shpMaxDays.Left + (intStartDayDiff * sngFactor)
            dblLabelLeft = Me.shpGrowForDateNeg.Left
        Else
            intStartDayDiff = (DateDiff("d", mdatEarliest, dteStart))
            Me.shpGrowForDateNeg.Visible = False
            Me.shpGrowForDate.Visible = True
            Me.shpGrowForDate.Width = Abs(intDayDiff) * sngFactor
            Me.shpGrowForDate.Left = Me.shpMaxDays.Left + (intStartDayDiff * sngFactor)
            dblLabelLeft = Me.shpGrowForDate.Left
        End If
        'If dblLabelLeft = 13360 Then Stop  'for testing only
        'Debug.Print dblLabelLeft   'for testing only
    End If
    If Not IsNull(Me.txtStartDate) And IsNull(Me.txtEndDate) Then
        intDayDiff = DateDiff("d", Me.txtStartDate, Me.txtStartDate)
        Me.txtTotalDays = " No Day(s)"
        intStartDayDiff = (DateDiff("d", mdatEarliest, Me.txtStartDate))
        Me.shpGrowForDate.Visible = False
        Me.shpGrowForDateNeg.Visible = True
        Me.shpGrowForDateNeg.Width = Abs(intDayDiff) * sngFactor
        Me.shpGrowForDateNeg.Left = Me.shpMaxDays.Left + (intStartDayDiff * sngFactor)
        dblLabelLeft = Me.shpGrowForDateNeg.Left
    End If
    If IsNull(Me.txtStartDate) And Not IsNull(Me.txtEndDate) Then
        intDayDiff = DateDiff("d", Me.txtEndDate, Me.txtEndDate)
        Me.txtTotalDays = " No Day(s)"
        intStartDayDiff = (DateDiff("d", mdatEarliest, Me.txtEndDate))
        Me.shpGrowForDate.Visible = False
        Me.shpGrowForDateNeg.Visible = True
        Me.shpGrowForDateNeg.Width = Abs(intDayDiff) * sngFactor
        Me.shpGrowForDateNeg.Left = Me.shpMaxDays.Left + (intStartDayDiff * sngFactor)
        dblLabelLeft = Me.shpGrowForDateNeg.Left
    End If
    'If dblLabelLeft > 13360 Then 'to far right for the control, so I'm going to reset the Left of the control to the max.
    '    dblLabelLeft = 13360
    'End If
    Me.txtTotalDays.Left = dblLabelLeft
    
ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 6 'overflow
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("secDetail_Format", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Sub

Public Function FormatName(strObjectName As String) As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Format the name of objects in the database to be more readable.
'
' Ver.  Date            Author              Details
' 1.00  20 MAY 2004     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strTemp As String
Dim intL As Integer
    intL = Len(strObjectName)
    
    strTemp = Mid(strObjectName, 5, intL - 8) ' 4 for "frm" or "rpt" to appear
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

