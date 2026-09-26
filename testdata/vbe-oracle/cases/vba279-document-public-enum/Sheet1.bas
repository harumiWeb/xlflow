Attribute VB_Name = "Sheet1"
Option Explicit

Public Enum SheetState
    SheetIdle
    SheetBusy
End Enum

Public Sub Touch()
    Dim state As SheetState
    state = SheetIdle
End Sub
