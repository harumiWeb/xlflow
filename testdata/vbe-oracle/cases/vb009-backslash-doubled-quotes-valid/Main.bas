Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Const dateFormat As String = "\""yyyy-mm-dd hh:nn:ss\"""
    Dim value As String
    value = "prefix \""quoted\"""
    Debug.Print dateFormat, value
End Sub
