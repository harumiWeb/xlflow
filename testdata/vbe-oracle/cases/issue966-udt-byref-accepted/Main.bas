Attribute VB_Name = "Main"
Option Explicit

Private Type Payload
    member As Long
End Type

Private Sub PrintValue(ByRef value As Payload)
    Debug.Print value.member
End Sub
