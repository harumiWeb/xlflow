Attribute VB_Name = "Main"
Option Explicit

Public Sub OracleControlReject()
    Dim value As Long
    value = MissingRequiredArgument(1)
End Sub

Private Function MissingRequiredArgument(ByVal first As Long, ByVal second As Long) As Long
    MissingRequiredArgument = first + second
End Function
