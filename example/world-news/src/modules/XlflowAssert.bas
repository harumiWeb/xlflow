Attribute VB_Name = "XlflowAssert"
Option Explicit

Public Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If IsObject(expected) Or IsObject(actual) Then
    Err.Raise vbObjectError + 514, "XlflowAssert.AssertEquals", "AssertEquals supports scalar values only. Compare object properties such as Range.Value2."
  End If

  If IsArray(expected) Or IsArray(actual) Then
    Err.Raise vbObjectError + 515, "XlflowAssert.AssertEquals", "AssertEquals supports scalar values only. Array comparison is not supported."
  End If

  If IsNull(expected) Or IsNull(actual) Then
    If IsNull(expected) And IsNull(actual) Then
      Exit Sub
    End If
    RaiseAssertEqualsFailure expected, actual, message
  End If

  If IsEmpty(expected) Or IsEmpty(actual) Then
    If IsEmpty(expected) And IsEmpty(actual) Then
      Exit Sub
    End If
    RaiseAssertEqualsFailure expected, actual, message
  End If

  If expected <> actual Then
    RaiseAssertEqualsFailure expected, actual, message
  End If
End Sub

Private Sub RaiseAssertEqualsFailure(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
  Dim detail As String
  detail = "expected <" & DescribeAssertValue(expected) & "> but got <" & DescribeAssertValue(actual) & ">"
  If Len(message) > 0 Then
    detail = message & ": " & detail
  End If
  Err.Raise vbObjectError + 513, "XlflowAssert.AssertEquals", detail
End Sub

Private Function DescribeAssertValue(ByVal value As Variant) As String
  If IsNull(value) Then
    DescribeAssertValue = "Null"
  ElseIf IsEmpty(value) Then
    DescribeAssertValue = "Empty"
  Else
    DescribeAssertValue = CStr(value)
  End If
End Function
