Attribute VB_Name = "Main"
Option Explicit

#If VBA7 Then
Public Function PtrValue(ByVal value As LongPtr) As LongPtr
#Else
Public Function PtrValue(ByVal value As Long) As Long
#End If
    PtrValue = value
End Function
