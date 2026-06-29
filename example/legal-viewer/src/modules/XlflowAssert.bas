Attribute VB_Name = "XlflowAssert"
Option Explicit

' Minimal assertion helpers for workbook-side tests.
' Keep assertions scalar so failures stay easy to read from xlflow JSON and terminal output.
Public Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal Message As String = "")
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
    RaiseAssertFailure Message, "expected <" & DescribeAssertValue(expected) & "> but got <" & DescribeAssertValue(actual) & ">", "XlflowAssert.AssertEquals"
  End If

  If expected <> actual Then
    RaiseAssertFailure Message, "expected <" & DescribeAssertValue(expected) & "> but got <" & DescribeAssertValue(actual) & ">", "XlflowAssert.AssertEquals"
  End If
End Sub

Public Sub AssertNotEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal Message As String = "")
  If IsObject(expected) Or IsObject(actual) Then
    Err.Raise vbObjectError + 514, "XlflowAssert.AssertNotEqual", "AssertNotEqual supports scalar values only."
  End If

  If IsArray(expected) Or IsArray(actual) Then
    Err.Raise vbObjectError + 515, "XlflowAssert.AssertNotEqual", "AssertNotEqual supports scalar values only."
  End If

  If IsNull(expected) And IsNull(actual) Then
    RaiseAssertFailure Message, "expected values to differ, but both are Null", "XlflowAssert.AssertNotEqual"
    Exit Sub
  End If

  If expected = actual Then
    RaiseAssertFailure Message, "expected <" & DescribeAssertValue(expected) & "> to differ from <" & DescribeAssertValue(actual) & ">", "XlflowAssert.AssertNotEqual"
  End If
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal Message As String = "")
  If Not condition Then
    RaiseAssertFailure Message, "expected True but got False", "XlflowAssert.AssertTrue"
  End If
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal Message As String = "")
  If condition Then
    RaiseAssertFailure Message, "expected False but got True", "XlflowAssert.AssertFalse"
  End If
End Sub

Public Sub AssertFail(Optional ByVal Message As String = "")
  RaiseAssertFailure Message, "assertion failed", "XlflowAssert.AssertFail"
End Sub

Public Sub AssertInconclusive(Optional ByVal Message As String = "")
  Dim detail As String
  detail = "inconclusive"
  If Len(Message) > 0 Then
    detail = Message
  End If
  Err.Raise vbObjectError + 516, "XlflowAssert.AssertInconclusive", detail
End Sub

Public Sub AssertIsNothing(ByVal value As Variant, Optional ByVal Message As String = "")
  If Not IsObject(value) Then
    RaiseAssertFailure Message, "expected an object but got a non-object", "XlflowAssert.AssertIsNothing"
    Exit Sub
  End If
  If Not value Is Nothing Then
    RaiseAssertFailure Message, "expected Nothing but got an object reference", "XlflowAssert.AssertIsNothing"
  End If
End Sub

Public Sub AssertIsNotNothing(ByVal value As Variant, Optional ByVal Message As String = "")
  If Not IsObject(value) Then
    RaiseAssertFailure Message, "expected an object but got a non-object", "XlflowAssert.AssertIsNotNothing"
    Exit Sub
  End If
  If value Is Nothing Then
    RaiseAssertFailure Message, "expected an object reference but got Nothing", "XlflowAssert.AssertIsNotNothing"
  End If
End Sub

Private Sub RaiseAssertFailure(ByVal Message As String, ByVal detail As String, ByVal source As String)
  If Len(Message) > 0 Then
    detail = Message & ": " & detail
  End If
  Err.Raise vbObjectError + 513, source, detail
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
