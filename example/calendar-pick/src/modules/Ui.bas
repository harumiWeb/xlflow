Attribute VB_Name = "Ui"
Option Explicit

Public Sub RunFromButton()
  Main.Run
End Sub

Public Sub InstallCalendarPickerButton()
  App.InstallLaunchButton ThisWorkbook
End Sub
