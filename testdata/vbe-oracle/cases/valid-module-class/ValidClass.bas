Attribute VB_Name = "ValidClass"
Option Explicit
Public Event Changed()
Friend Sub Run()
    Debug.Print TypeName(Me)
End Sub
