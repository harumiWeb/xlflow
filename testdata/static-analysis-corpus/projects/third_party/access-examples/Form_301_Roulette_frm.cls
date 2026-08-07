VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_301_Roulette_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Would you like to play a game?
'
' Ver.  Date            Author              Details
' 1.00  05 NOV 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Option Explicit
Dim marrRoulette()  As Variant
Const mdblGreen     As Double = 65280
Const mdblBlack     As Double = 0
Const mdblRed       As Double = 255
Const mdblGray      As Double = 9868950
Const mdblWhite     As Double = 16777215

Public Function RotateGraph()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Rotate the graph each loop
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim objGraph        As Object
Dim intNewRotation  As Integer
Dim intOldRotation  As Integer

    Set objGraph = Me.gphRouletteWheel.Object.Application.chart

    ' Determine the initial rotation setting.
    intOldRotation = objGraph.Rotation

    ' Rotate graph from initial setting to 360 degrees.
    For intNewRotation = intOldRotation To 360 Step 5
        objGraph.Rotation = intNewRotation
        If objGraph.Rotation = 360 Then objGraph.Rotation = 10
        DoEvents
        DoCmd.RepaintObject
    Next

ExitProcedure:
    On Error Resume Next
    Set objGraph = Nothing
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("RotateGraph", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Private Sub cmdRun_Click()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Select random number from array
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim x               As Integer 'array loop
Dim strRanNbr       As String 'random number string
Dim dblForeColor    As Double 'font color
Dim dblBackColor    As Double 'backround color

    Me.lblWinNbr.Caption = ""
    Me.lblWinNbr.BackColor = mdblGray
    Call RotateGraph
    strRanNbr = GetRndNbr
        For x = 0 To UBound(marrRoulette)
            If strRanNbr = marrRoulette(x, 0) Then
                dblBackColor = marrRoulette(x, 1) 'get the backround color of the number
            End If
        Next x
    Select Case dblBackColor 'set the font color
        Case Is = mdblRed
            dblForeColor = mdblBlack
        Case Is = mdblGreen
            dblForeColor = mdblBlack
        Case Is = mdblBlack
            dblForeColor = mdblWhite
    End Select
    Me.lblWinNbr.BackColor = dblBackColor
    Me.lblWinNbr.ForeColor = dblForeColor
    Me.lblWinNbr.Caption = strRanNbr
    Me.Requery
    
End Sub

Private Sub Form_Current()
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Set the color of the label
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim intWinNbr As Integer
    intWinNbr = Me.lblWinNbr.Caption
    
    If intWinNbr > 0 Then
        Me.TimerInterval = 300
    Else
        Me.TimerInterval = 0
        Me.lblWinNbr.ForeColor = vbBlack
    End If
    
End Sub

Public Function Randomizer() As Integer
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Initialize the random number generator
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Static AlreadyDone As Integer
        If AlreadyDone = False Then Randomize: AlreadyDone = True
    Randomizer = 0
    
End Function

Public Function GetRndNbr() As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Select a random number from a select group of values
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Dim MyValue As Variant
Randomize   ' Initialize random-number generator.

    MyValue = Int((38 * Rnd) + 1)    ' Generate random value 37 = 00
    If MyValue = "37" Then
        GetRndNbr = "0"
    ElseIf MyValue = "38" Then
        GetRndNbr = "00"
    Else
    GetRndNbr = MyValue
    End If
        
End Function

Public Function LoadArray() As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              used to return the number and color of the backround
'
' Ver.  Date            Author              Details
' 1.00  22 FEB 2008     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim intChart    As Integer
Dim intColumn   As Integer

    intChart = 5
    intColumn = 37
    ReDim marrRoulette(intColumn, 2)
    marrRoulette(0, 0) = "2":   marrRoulette(0, 1) = mdblBlack:   marrRoulette(0, 2) = intChart
    marrRoulette(1, 0) = "0":   marrRoulette(1, 1) = mdblGreen:   marrRoulette(1, 2) = intChart
    marrRoulette(2, 0) = "28":  marrRoulette(2, 1) = mdblBlack:   marrRoulette(2, 2) = intChart
    marrRoulette(3, 0) = "9":   marrRoulette(3, 1) = mdblRed:     marrRoulette(3, 2) = intChart
    marrRoulette(4, 0) = "26":  marrRoulette(4, 1) = mdblBlack:   marrRoulette(4, 2) = intChart
    marrRoulette(5, 0) = "30":  marrRoulette(5, 1) = mdblRed:     marrRoulette(5, 2) = intChart
    marrRoulette(6, 0) = "11":  marrRoulette(6, 1) = mdblBlack:   marrRoulette(6, 2) = intChart
    marrRoulette(7, 0) = "7":   marrRoulette(7, 1) = mdblRed:     marrRoulette(7, 2) = intChart
    marrRoulette(8, 0) = "20":  marrRoulette(8, 1) = mdblBlack:   marrRoulette(8, 2) = intChart
    marrRoulette(9, 0) = "32":  marrRoulette(9, 1) = mdblRed:     marrRoulette(9, 2) = intChart
    marrRoulette(10, 0) = "17": marrRoulette(10, 1) = mdblBlack:  marrRoulette(10, 2) = intChart
    marrRoulette(11, 0) = "5":  marrRoulette(11, 1) = mdblRed:    marrRoulette(11, 2) = intChart
    marrRoulette(12, 0) = "22": marrRoulette(12, 1) = mdblBlack:  marrRoulette(12, 2) = intChart
    marrRoulette(13, 0) = "34": marrRoulette(13, 1) = mdblRed:    marrRoulette(13, 2) = intChart
    marrRoulette(14, 0) = "15": marrRoulette(14, 1) = mdblBlack:  marrRoulette(14, 2) = intChart
    marrRoulette(15, 0) = "3":  marrRoulette(15, 1) = mdblRed:    marrRoulette(15, 2) = intChart
    marrRoulette(16, 0) = "24": marrRoulette(16, 1) = mdblBlack:  marrRoulette(16, 2) = intChart
    marrRoulette(17, 0) = "36": marrRoulette(17, 1) = mdblRed:    marrRoulette(17, 2) = intChart
    marrRoulette(18, 0) = "13": marrRoulette(18, 1) = mdblBlack:  marrRoulette(18, 2) = intChart
    marrRoulette(19, 0) = "1":  marrRoulette(19, 1) = mdblRed:    marrRoulette(19, 2) = intChart
    marrRoulette(20, 0) = "00": marrRoulette(20, 1) = mdblGreen:  marrRoulette(20, 2) = intChart
    marrRoulette(21, 0) = "27": marrRoulette(21, 1) = mdblRed:    marrRoulette(21, 2) = intChart
    marrRoulette(22, 0) = "10": marrRoulette(22, 1) = mdblBlack:  marrRoulette(22, 2) = intChart
    marrRoulette(23, 0) = "25": marrRoulette(23, 1) = mdblRed:    marrRoulette(23, 2) = intChart
    marrRoulette(24, 0) = "29": marrRoulette(24, 1) = mdblBlack:  marrRoulette(24, 2) = intChart
    marrRoulette(25, 0) = "12": marrRoulette(25, 1) = mdblRed:    marrRoulette(25, 2) = intChart
    marrRoulette(26, 0) = "8":  marrRoulette(26, 1) = mdblBlack:  marrRoulette(26, 2) = intChart
    marrRoulette(27, 0) = "19": marrRoulette(27, 1) = mdblRed:    marrRoulette(27, 2) = intChart
    marrRoulette(28, 0) = "31": marrRoulette(28, 1) = mdblBlack:  marrRoulette(28, 2) = intChart
    marrRoulette(29, 0) = "18": marrRoulette(29, 1) = mdblRed:    marrRoulette(29, 2) = intChart
    marrRoulette(30, 0) = "6":  marrRoulette(30, 1) = mdblBlack:  marrRoulette(30, 2) = intChart
    marrRoulette(31, 0) = "21": marrRoulette(31, 1) = mdblRed:    marrRoulette(31, 2) = intChart
    marrRoulette(32, 0) = "33": marrRoulette(32, 1) = mdblBlack:  marrRoulette(32, 2) = intChart
    marrRoulette(33, 0) = "16": marrRoulette(33, 1) = mdblRed:    marrRoulette(33, 2) = intChart
    marrRoulette(34, 0) = "4":  marrRoulette(34, 1) = mdblBlack:  marrRoulette(34, 2) = intChart
    marrRoulette(35, 0) = "23": marrRoulette(35, 1) = mdblRed:    marrRoulette(35, 2) = intChart
    marrRoulette(36, 0) = "35": marrRoulette(36, 1) = mdblBlack:  marrRoulette(36, 2) = intChart
    marrRoulette(37, 0) = "14": marrRoulette(37, 1) = mdblRed:    marrRoulette(37, 2) = intChart
    LoadArray = True
    
ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("LoadArray", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select

End Function

Private Sub Form_Open(Cancel As Integer)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              load array, set as source for graph
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
' 1.01  19 FEB 2008     Anthony Duguid      added array for source
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim x As Integer
Dim strTemp As String
    
    Call LoadArray 'set table values
    'strTemp = "'Column1'; 'Column2'; "
    'For x = 0 To UBound(marrRoulette)
    '    strTemp = strTemp & "'" & marrRoulette(x, 0) & "' " 'wheel number
    '    strTemp = strTemp & "," & marrRoulette(x, 2) & "; " 'wheel size
    'Next x
    'Debug.Print strTemp
    'Me.gphRouletteWheel.Visible = True
    'Me.gphRouletteWheel.RowSourceType = "Value List"
    'Me.gphRouletteWheel.RowSource = strTemp
    'Me.gphRouletteWheel.Requery
    'DoEvents
    'DoCmd.Save
    'Me.Repaint

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
