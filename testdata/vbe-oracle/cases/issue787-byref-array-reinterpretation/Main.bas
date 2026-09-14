Attribute VB_Name = "Main"
Option Explicit

#If VBA7 = 0 Then
Private Enum LongPtr
    [_]
End Enum
Private Enum LONG_PTR
    [_]
End Enum
#End If

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Sub InitSafeArray(ByRef sa As SAFEARRAY_1D, ByVal elemSize As Long)
    sa.cDims = 1
    sa.cbElements = elemSize
End Sub

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub

Public Sub ProbeIssue787()
    Static pa(0) As PointerAccessor
    With pa(0)
        If .sa.cDims = 0 Then
            InitSafeArray .sa, 8
            WritePtrNatively pa, VarPtr(.sa)
        End If
    End With
End Sub
