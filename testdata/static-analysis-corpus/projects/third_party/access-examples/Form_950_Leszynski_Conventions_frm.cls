VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_950_Leszynski_Conventions_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function OpenDetailReport() As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Opens the detail report
'
' Ver.  Date            Author              Details
' 1.00  08 JUL 2007     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error GoTo ErrTrap

    If Me.FilterOn = False Then
        DoCmd.OpenReport "950_Leszynski_Conventions_rpt", acViewPreview
    Else
        DoCmd.OpenReport "950_Leszynski_Conventions_rpt", acViewPreview, , Me.Filter
    End If

ExitProcedure:
    On Error Resume Next
    Exit Function
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("OpenDetailReport", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function
