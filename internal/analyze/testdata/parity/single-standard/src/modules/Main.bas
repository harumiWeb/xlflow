Option Explicit
Public Sub Run(ByVal shouldSet As Boolean)
  Dim target As Object
  If shouldSet Then
    Set target = CreateObject("Scripting.Dictionary")
  End If
  Debug.Print target.Count
  ' xlflow:disable-next-line VBA205
  Range("A1").Value = 1
  Range("A2").Value = 2
End Sub
