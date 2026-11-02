Attribute VB_Name = "Main"
Option Explicit
#Const devMode = True

Public Sub Work()
    GoSub UsedLabel: Debug.Print 1
#If devMode Then
    Debug.Print 2
#End If
    Exit Sub
UsedLabel:
    Return
UnusedLabel:
    Return
End Sub
