Option Explicit
Public Sub Alpha()
  Helpers.Beta
End Sub
Public Sub Validate()
  Dim count As Long
  Helpers.NeedsText count
End Sub
Public Sub Lookup()
  Dim rng As Range
  Dim result As Variant
  result = WorksheetFunction.Match("key", rng)
End Sub
