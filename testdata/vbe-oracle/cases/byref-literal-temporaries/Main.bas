Attribute VB_Name = "Main"
Option Explicit

Private Sub TakeInteger(ByRef value As Integer)
End Sub

Private Sub TakeByte(ByRef value As Byte)
End Sub

Private Sub TakeString(ByRef value As String)
End Sub

Private Sub TakeLong(ByRef value As Long)
End Sub

Public Sub ProbeByRefLiteralTemporaries()
    Call TakeInteger(-1)
    Call TakeByte(255)
    Call TakeString(0)
    Call TakeLong("abc")
End Sub
