Attribute VB_Name = "Main"
Option Explicit

Private Type Payload
    Value As Long
End Type

Public Sub GoodOptional(Optional first As Long = 1, Optional second As Long = 2)
End Sub

Public Sub GoodUDT(ByRef value As Payload)
End Sub

Public Sub GoodParamArray(ByVal first As Long, ParamArray values() As Variant)
End Sub
