VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_804_Weekly_Report_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              To show numbers for 52 week data
'
' Ver.  Date            Author              Details
' 1.00  05 NOV 2001     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

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

Private Sub secPageHeader_Format(Cancel As Integer, FormatCount As Integer)
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Set the column labels for the report as a date
'
' Ver.  Date            Author              Details
' 1.00  22 JAN 2007     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim intYear As String
    'intYear = CInt(DatePart("yyyy", Forms![100_Main_Menu_frm]![txtWED_END_DTE]))
    intYear = CInt(DatePart("yyyy", Now())) 'CHANGE THIS TO YOUR FILTER
Dim i As Integer

    For i = 1 To 52
        Me.Controls("txtWeek" & i).Value = WeekStart(i, intYear)
    Next i

End Sub

Public Function WeekStart(WhichWeek As Integer, ByVal WhichYear As Integer) As Date
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Find the week date based on the year
'
' Ver.  Date            Author              Details
' 1.00  22 JAN 2007     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next

    WeekStart = YearStart(WhichYear) + ((WhichWeek - 1) * 7)

End Function

Public Function YearStart(WhichYear As Integer) As Date
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Used in the WeekStart function to define which year to set the labels.
'
' Ver.  Date            Author              Details
' 1.00  22 JAN 2007     Anthony Duguid      Initial version
'--------------------------------------------------------------------------------------------------------------------
On Error Resume Next
Dim WeekDay As Integer
Dim NewYear As Date
Dim dteTemp As Date

    NewYear = DateSerial(WhichYear, 1, 1)
    WeekDay = (NewYear - 2) Mod 7 'Generate weekday index where Monday = 0
    
    If WeekDay < 4 Then
        dteTemp = NewYear - WeekDay
    Else
        dteTemp = NewYear - WeekDay + 7
    End If
    
    YearStart = Format(DateAdd("d", 5, dteTemp), "MMM DD") ' change to Saturday

End Function
