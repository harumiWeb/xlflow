Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim path As String
    path = "C:\temp\oracle.txt"
    Open path For Input As #1
    Close #1
End Sub
