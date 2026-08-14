Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
    Dim value As Object
    If TypeOf value Is Collection Then
        Debug.Print "collection"
    End If
End Sub
