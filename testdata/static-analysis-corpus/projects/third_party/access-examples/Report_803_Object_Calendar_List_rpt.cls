VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_803_Object_Calendar_List_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              create a calendar from a table in access
'
' Ver.  Date            Author              Details
' 1.00  17 JUL 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Const TwentyEight = 28

Private Type udtPrtMips_str
strRGB As String * TwentyEight
End Type

' ' Structure from ADH97
Private Type udtPrtMips
    lngLeft As Long
    lngTop As Long
    lngRight As Long
    lngBottom As Long
    lngDataOnly As Long
    lngWidth As Long
    lngHeight As Long
    lngDefaultSize As Long
    lngItemsAcross As Long
    lngColumnSpacing As Long
    lngRowSpacing As Long
    lngItemLayout As Long
    lngFastPrinting As Long
    lngDataSheet As Long
End Type

Public bNoData As Boolean

Private Sub Report_Activate()
    bNoData = False
    DoCmd.Maximize
End Sub

Private Sub Report_NoData(Cancel As Integer)
On Error Resume Next
    bNoData = True
    Me.secDetail.Visible = False
    'Me.secPageFooter.Visible = False
    Me.secPageHeader.Visible = False
    Me.secReportHeader.Visible = True

End Sub

Private Sub Report_Page()
If bNoData = True Then Exit Sub
Const TWIPSPERINCH = 1440
' Offset from right edge of Control where our
' Vertical Line will start - adjust as you desire.
' Expressed in Twips
Dim intLineMargin As Integer
' Need Control object to loop through Control Collection
Dim CtlDetail As Control
' Y position where we will start Drawing our Grid from
Dim startY As Integer
' Var will be set to the Margin you
' set for the Bottom of the Page
' and Top of the Page via
' the Page Setup Menu
' Expressed in TWIPS.
Dim varBMargin As Variant
Dim varTMargin As Variant
' PrtMips structs
Dim MIPSstr As udtPrtMips_str
Dim MIPS As udtPrtMips

    ' See Access help for PrtMip'
    MIPSstr.strRGB = Me.PrtMip
    LSet MIPS = MIPSstr
    ' Copy value into our page Margin vars
    varBMargin = MIPS.lngBottom
    varTMargin = MIPS.lngTop
    ' intLineMargin is expressed in TWIPS
    intLineMargin = 60
    Me.DrawWidth = 6
    ' We need to draw a vertical line at the left
    ' edge of our report. All other vertical lines are to the
    ' right edge of the existing controls
    ' You should place the right most control at the extreme
    ' right edge of your report - intMargin
    With Me
        Me.Line ((0), .SECTION(acPageHeader).Height) _
        -(0, .ScaleHeight - _
        (.SECTION(acPageFooter).Height))
    End With
    ' OK lets draw a vertical line to seperate each field
    ' for each control in details control
    For Each CtlDetail In Me.SECTION(acDetail).Controls
            ' Don't draw a Vertical line at end of Horizontal Line control - not needed
        If Not TypeOf CtlDetail Is Access.Line Then
            With CtlDetail
            'setup your
             Me.Line ((.Left + .Width), Me.SECTION(acPageHeader).Height)-(.Left + .Width, Me.ScaleHeight - (Me.SECTION(acPageFooter).Height))
            '--------------------------------------------------------------------------------
            'ALD. 11/01/2006  use this one if you don't have exact measured controls.
             'Me.Line ((.Left + .Width + intLineMargin), Me.Section(acPageHeader).Height) _
             '-(.Left + .Width + intLineMargin, Me.ScaleHeight - _
             '(Me.Section(acPageFooter).Height))
            '--------------------------------------------------------------------------------
            End With
        End If
    Next
    'Draw our Horizontal Lines
    'With Me
    '    startY = .Section(acPageHeader).Height
    '    Do While startY <= (.ScaleHeight + varBMargin) - (.Section(acPageFooter).Height)
    '        Me.Line (0, startY)-(.Width, startY)
    '        startY = startY + .Section(acDetail).Height
    '    Loop
    'End With
    'CleanUp time
    Set CtlDetail = Nothing

End Sub

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
' Purpose:              Replace character(s) in a string value
' Notes:                Replaces the strSearch string with strReplacement string in the strText string.
'                       Uses intCompMode to determine comparison mode
'
' Ver.  Date            Author              Details
' 1.00  01 MAY 2004     Anthony Duguid      Initial version
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

Private Sub Report_Close()
    DoCmd.Restore
End Sub

Private Sub Report_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub

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

