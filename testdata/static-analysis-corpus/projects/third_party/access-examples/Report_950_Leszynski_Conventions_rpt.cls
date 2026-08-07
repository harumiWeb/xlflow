VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_950_Leszynski_Conventions_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              naming convention report
'
' Ver.  Date            Author              Details
' 1.00  05 NOV 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Private Sub Report_Close()
On Error Resume Next

    Me.Filter = ""

End Sub

Private Sub secDetail_Format(Cancel As Integer, FormatCount As Integer)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Alternate colors between the lines
'
' Ver.  Date            Author              Details
' 1.00  20 JUN 2002     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next

   If Me.txtRecordCount Mod 2 = 1 Then
      Me.SECTION(0).BackColor = 16777215 'white
   Else
      Me.SECTION(0).BackColor = 15269864 'light green
   End If
   
End Sub

Public Function HeaderString() As String
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              To show the filters from the main form
'
' Ver.  Date            Author              Details
' 1.00  09 JUL 2007     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next

    HeaderString = IIf(Len(Me.Filter) = 0, "", "Filter by: " & Me.Filter)
    
End Function
