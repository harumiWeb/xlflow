VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarDayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Button As MSForms.CommandButton
Attribute Button.VB_VarHelpID = -1

Private mHost As Object
Private mDateValue As Date

Public Sub Bind(ByVal host As Object, ByVal sourceButton As MSForms.CommandButton, ByVal dateValue As Date)
  Set mHost = host
  Set Button = sourceButton
  mDateValue = dateValue
End Sub

Private Sub Button_Click()
  CallByName mHost, "HandleDayButtonClick", VbMethod, mDateValue
End Sub
