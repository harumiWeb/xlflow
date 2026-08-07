VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_002_Splash_Screen_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
On Error Resume Next

    Me.imgLogo.Picture = GetProperties("Category")
    Me.TimerInterval = Me.txtSeconds * 1000 'convert to milliseconds\
    
End Sub

Private Sub Form_Timer()
On Error Resume Next

    DoCmd.Close
    
End Sub

Public Function GetProperties(pstrName As String) As Variant
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Returns information from the database properties.
'
' Ver.  Date            Author              Details
' 1.00  22 AUG 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    GetProperties = Trim(CodeDb.Containers!Databases.Documents!SummaryInfo.Properties(pstrName))

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    Select Case Err.Number
        Case Is = 3270 'property not assigned
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("GetProperties", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function
