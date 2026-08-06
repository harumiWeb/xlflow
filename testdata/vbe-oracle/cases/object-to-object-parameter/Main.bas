Attribute VB_Name = "Main"
Option Explicit

Public Sub ProbeObjectToObjectParameter()
    Dim value As Object
    Set value = CreateObject("Scripting.Dictionary")
    AcceptObject value
End Sub

Private Sub AcceptObject(ByVal value As Object)
End Sub
