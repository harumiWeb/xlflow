Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Predeclared.Value = 1
    Debug.Print Predeclared.Value
    Set Predeclared = Nothing
End Sub
