Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeMissingSetObjectAssignment()
    Dim value As Object
    value = CreateObject("Scripting.Dictionary")
End Sub
