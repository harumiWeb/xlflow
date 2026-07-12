Attribute VB_Name = "XlflowAssert"
Option Explicit

Private Const assertFailureNumber As Long = vbObjectError + 513

''' Asserts that two scalar values are equal.
'''
''' Args:
'''     expected: Expected scalar value.
'''     actual: Actual scalar value.
'''     message: Optional failure message prefix.
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
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " but got " & FormatAssertValue(actual), "XlflowAssert.AssertEquals"
  End If

  If expected <> actual Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " but got " & FormatAssertValue(actual), "XlflowAssert.AssertEquals"
  End If
End Sub

''' Asserts that two scalar values have the same VarType and value.
'''
''' Args:
'''     expected: Expected scalar value.
'''     actual: Actual scalar value.
'''     message: Optional failure message prefix.
Public Sub AssertStrictEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If IsObject(expected) Or IsObject(actual) Then
    RaiseAssertFailure message, "AssertStrictEquals supports scalar values only. Object comparison is not supported.", "XlflowAssert.AssertStrictEquals"
  End If

  If IsArray(expected) Or IsArray(actual) Then
    RaiseAssertFailure message, "AssertStrictEquals supports scalar values only. Array comparison is not supported.", "XlflowAssert.AssertStrictEquals"
  End If

  If VarType(expected) <> VarType(actual) Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " but got " & FormatAssertValue(actual), "XlflowAssert.AssertStrictEquals"
  End If

  If IsNull(expected) Or IsEmpty(expected) Then
    Exit Sub
  End If

  If expected <> actual Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " but got " & FormatAssertValue(actual), "XlflowAssert.AssertStrictEquals"
  End If
End Sub

''' Asserts that two scalar values are different.
'''
''' Args:
'''     expected: Value that should differ from actual.
'''     actual: Value that should differ from expected.
'''     message: Optional failure message prefix.
Public Sub AssertNotEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If IsObject(expected) Or IsObject(actual) Then
    Err.Raise vbObjectError + 514, "XlflowAssert.AssertNotEqual", "AssertNotEqual supports scalar values only."
  End If

  If IsArray(expected) Or IsArray(actual) Then
    Err.Raise vbObjectError + 515, "XlflowAssert.AssertNotEqual", "AssertNotEqual supports scalar values only."
  End If

  If IsNull(expected) And IsNull(actual) Then
    RaiseAssertFailure message, "expected values to differ, but both are Null", "XlflowAssert.AssertNotEqual"
    Exit Sub
  End If

  If IsNull(expected) Or IsNull(actual) Then
    Exit Sub
  End If

  If expected = actual Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " to differ from " & FormatAssertValue(actual), "XlflowAssert.AssertNotEqual"
  End If
End Sub

''' Asserts that a value is Null.
Public Sub AssertNull(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsNull(value) Then
    RaiseAssertFailure message, "expected <Null> but got " & FormatAssertValue(value), "XlflowAssert.AssertNull"
  End If
End Sub

''' Asserts that a value is not Null.
Public Sub AssertNotNull(ByVal value As Variant, Optional ByVal message As String = "")
  If IsNull(value) Then
    RaiseAssertFailure message, "expected a non-Null value but got <Null>", "XlflowAssert.AssertNotNull"
  End If
End Sub

''' Asserts that a value is Empty according to VBA IsEmpty.
Public Sub AssertEmpty(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsEmpty(value) Then
    RaiseAssertFailure message, "expected <Empty> but got " & FormatAssertValue(value), "XlflowAssert.AssertEmpty"
  End If
End Sub

''' Asserts that a value is not Empty according to VBA IsEmpty.
Public Sub AssertNotEmpty(ByVal value As Variant, Optional ByVal message As String = "")
  If IsEmpty(value) Then
    RaiseAssertFailure message, "expected a non-Empty value but got <Empty>", "XlflowAssert.AssertNotEmpty"
  End If
End Sub

''' Asserts that two numeric values are within a non-negative tolerance.
Public Sub AssertNear(ByVal expected As Variant, ByVal actual As Variant, ByVal tolerance As Double, Optional ByVal message As String = "")
  Dim source As String
  source = "XlflowAssert.AssertNear"

  If tolerance < 0 Then
    RaiseAssertFailure message, "tolerance must be non-negative but got " & FormatAssertValue(tolerance), source
  End If

  If Not IsAssertNumeric(expected) Then
    RaiseAssertFailure message, "expected value must be numeric but got " & FormatAssertValue(expected), source
  End If

  If Not IsAssertNumeric(actual) Then
    RaiseAssertFailure message, "actual value must be numeric but got " & FormatAssertValue(actual), source
  End If

  Dim difference As Double
  difference = Abs(CDbl(expected) - CDbl(actual))
  If difference > tolerance Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " +/- " & FormatAssertValue(tolerance) & " but got " & FormatAssertValue(actual) & "; difference was " & FormatAssertValue(difference), source
  End If
End Sub

''' Asserts that a string contains an expected substring using binary comparison.
Public Sub AssertContains(ByVal expectedSubstring As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsAssertString(expectedSubstring) Or Not IsAssertString(actual) Then
    RaiseAssertFailure message, "AssertContains expects string values; expected substring was " & FormatAssertValue(expectedSubstring) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertContains"
  End If

  If InStr(1, CStr(actual), CStr(expectedSubstring), vbBinaryCompare) = 0 Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(actual) & " to contain " & FormatAssertValue(expectedSubstring), "XlflowAssert.AssertContains"
  End If
End Sub

''' Asserts that a string starts with an expected prefix using binary comparison.
Public Sub AssertStartsWith(ByVal prefix As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsAssertString(prefix) Or Not IsAssertString(actual) Then
    RaiseAssertFailure message, "AssertStartsWith expects string values; prefix was " & FormatAssertValue(prefix) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertStartsWith"
  End If

  If Left$(CStr(actual), Len(CStr(prefix))) <> CStr(prefix) Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(actual) & " to start with " & FormatAssertValue(prefix), "XlflowAssert.AssertStartsWith"
  End If
End Sub

''' Asserts that a string ends with an expected suffix using binary comparison.
Public Sub AssertEndsWith(ByVal suffix As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsAssertString(suffix) Or Not IsAssertString(actual) Then
    RaiseAssertFailure message, "AssertEndsWith expects string values; suffix was " & FormatAssertValue(suffix) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertEndsWith"
  End If

  If Right$(CStr(actual), Len(CStr(suffix))) <> CStr(suffix) Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(actual) & " to end with " & FormatAssertValue(suffix), "XlflowAssert.AssertEndsWith"
  End If
End Sub

''' Asserts that a string matches a VBScript.RegExp pattern.
Public Sub AssertMatches(ByVal pattern As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  Dim source As String
  source = "XlflowAssert.AssertMatches"

  If Not IsAssertString(pattern) Or Not IsAssertString(actual) Then
    RaiseAssertFailure message, "AssertMatches expects string values; pattern was " & FormatAssertValue(pattern) & " and actual was " & FormatAssertValue(actual), source
  End If

  Dim regex As Object
  ' xlflow:disable-next-line VB004
  On Error Resume Next
  Set regex = CreateObject("VBScript.RegExp")
  If Err.Number <> 0 Then
    Dim createError As String
    createError = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseAssertFailure message, "VBScript.RegExp is not available: " & createError, source
  End If
  regex.Pattern = CStr(pattern)
  regex.IgnoreCase = False
  regex.Global = False
  regex.MultiLine = False
  If Err.Number <> 0 Then
    Dim patternError As String
    patternError = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseAssertFailure message, "invalid regex pattern " & FormatAssertValue(pattern) & ": " & patternError, source
  End If

  Dim matched As Boolean
  matched = regex.Test(CStr(actual))
  If Err.Number <> 0 Then
    Dim testError As String
    testError = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseAssertFailure message, "invalid regex pattern " & FormatAssertValue(pattern) & ": " & testError, source
  End If
  On Error GoTo 0

  If Not matched Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(actual) & " to match pattern " & FormatAssertValue(pattern), source
  End If
End Sub

''' Asserts that one- or two-dimensional scalar arrays are equal.
Public Sub AssertArrayEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  Dim source As String
  source = "XlflowAssert.AssertArrayEquals"

  If Not IsArray(expected) Or Not IsArray(actual) Then
    RaiseAssertFailure message, "AssertArrayEquals expects arrays; expected was " & FormatAssertValue(expected) & " and actual was " & FormatAssertValue(actual), source
  End If

  Dim expectedDims As Long
  Dim actualDims As Long
  expectedDims = ArrayDimensionCount(expected)
  actualDims = ArrayDimensionCount(actual)
  If expectedDims <> actualDims Then
    RaiseAssertFailure message, "array dimensions differ; expected " & CStr(expectedDims) & " but got " & CStr(actualDims), source
  End If
  If expectedDims < 1 Or expectedDims > 2 Then
    RaiseAssertFailure message, "AssertArrayEquals supports one- and two-dimensional arrays only; got " & CStr(expectedDims) & " dimensions", source
  End If

  Dim dimension As Long
  For dimension = 1 To expectedDims
    If LBound(expected, dimension) <> LBound(actual, dimension) Or UBound(expected, dimension) <> UBound(actual, dimension) Then
      RaiseAssertFailure message, "array bounds differ on dimension " & CStr(dimension) & vbCrLf & "expected <" & CStr(LBound(expected, dimension)) & " To " & CStr(UBound(expected, dimension)) & ">" & vbCrLf & "actual   <" & CStr(LBound(actual, dimension)) & " To " & CStr(UBound(actual, dimension)) & ">", source
    End If
  Next dimension

  Dim i As Long
  Dim j As Long
  If expectedDims = 1 Then
    For i = LBound(expected, 1) To UBound(expected, 1)
      If Not ScalarArrayValuesEqual(expected(i), actual(i)) Then
        RaiseAssertFailure message, "array mismatch at (" & CStr(i) & ")" & vbCrLf & "expected " & FormatAssertValue(expected(i)) & vbCrLf & "actual   " & FormatAssertValue(actual(i)), source
      End If
    Next i
  Else
    For i = LBound(expected, 1) To UBound(expected, 1)
      For j = LBound(expected, 2) To UBound(expected, 2)
        If Not ScalarArrayValuesEqual(expected(i, j), actual(i, j)) Then
          RaiseAssertFailure message, "array mismatch at (" & CStr(i) & ", " & CStr(j) & ")" & vbCrLf & "expected " & FormatAssertValue(expected(i, j)) & vbCrLf & "actual   " & FormatAssertValue(actual(i, j)), source
        End If
      Next j
    Next i
  End If
End Sub

''' Asserts that expected scalar or two-dimensional array values equal Range.Value2.
Public Sub AssertRangeEquals(ByVal expected As Variant, ByVal actualRange As Object, Optional ByVal message As String = "")
  Dim source As String
  source = "XlflowAssert.AssertRangeEquals"

  If actualRange Is Nothing Then
    RaiseAssertFailure message, "actualRange must be an Excel Range object but got <Nothing>", source
  End If

  Dim rowCount As Long
  Dim columnCount As Long
  Dim actualValues As Variant
  ' xlflow:disable-next-line VB004
  On Error Resume Next
  rowCount = CLng(actualRange.Rows.Count)
  columnCount = CLng(actualRange.Columns.Count)
  actualValues = actualRange.Value2
  If Err.Number <> 0 Then
    Dim rangeError As String
    rangeError = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseAssertFailure message, "actualRange must expose Range members Rows, Columns, Cells, and Value2: " & rangeError, source
  End If
  On Error GoTo 0

  If rowCount = 1 And columnCount = 1 Then
    If IsArray(expected) Then
      RaiseAssertFailure message, "scalar expected value is required for a single-cell range but got " & FormatAssertValue(expected), source
    End If
    If Not ScalarArrayValuesEqual(expected, actualValues) Then
      RaiseAssertFailure message, "range mismatch at " & RangeCellLabel(actualRange, 1, 1) & vbCrLf & "expected " & FormatAssertValue(expected) & vbCrLf & "actual   " & FormatAssertValue(actualValues), source
    End If
    Exit Sub
  End If

  If Not IsArray(expected) Then
    RaiseAssertFailure message, "two-dimensional expected array is required for a multi-cell range but got " & FormatAssertValue(expected), source
  End If

  If ArrayDimensionCount(expected) <> 2 Then
    RaiseAssertFailure message, "expected array for a multi-cell range must be two-dimensional", source
  End If

  Dim expectedRows As Long
  Dim expectedColumns As Long
  expectedRows = UBound(expected, 1) - LBound(expected, 1) + 1
  expectedColumns = UBound(expected, 2) - LBound(expected, 2) + 1
  If expectedRows <> rowCount Or expectedColumns <> columnCount Then
    RaiseAssertFailure message, "range size differs; expected <" & CStr(expectedRows) & " x " & CStr(expectedColumns) & "> but got <" & CStr(rowCount) & " x " & CStr(columnCount) & ">", source
  End If

  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim expectedRow As Long
  Dim expectedColumn As Long
  For rowIndex = 1 To rowCount
    expectedRow = LBound(expected, 1) + rowIndex - 1
    For columnIndex = 1 To columnCount
      expectedColumn = LBound(expected, 2) + columnIndex - 1
      If Not ScalarArrayValuesEqual(expected(expectedRow, expectedColumn), actualValues(rowIndex, columnIndex)) Then
        RaiseAssertFailure message, "range mismatch at " & RangeCellLabel(actualRange, rowIndex, columnIndex) & vbCrLf & "expected " & FormatAssertValue(expected(expectedRow, expectedColumn)) & vbCrLf & "actual   " & FormatAssertValue(actualValues(rowIndex, columnIndex)), source
      End If
    Next columnIndex
  Next rowIndex
End Sub

''' Asserts that two object references are the same reference.
Public Sub AssertSame(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsObject(expected) Or Not IsObject(actual) Then
    RaiseAssertFailure message, "AssertSame expects object references; expected was " & FormatAssertValue(expected) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertSame"
  End If

  If Not (expected Is actual) Then
    RaiseAssertFailure message, "expected same object reference but got different references", "XlflowAssert.AssertSame"
  End If
End Sub

''' Asserts that two object references are not the same reference.
Public Sub AssertNotSame(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsObject(expected) Or Not IsObject(actual) Then
    RaiseAssertFailure message, "AssertNotSame expects object references; expected was " & FormatAssertValue(expected) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertNotSame"
  End If

  If expected Is actual Then
    RaiseAssertFailure message, "expected different object references but both were the same", "XlflowAssert.AssertNotSame"
  End If
End Sub

''' Asserts that a condition is True.
'''
''' Args:
'''     condition: Boolean condition to verify.
'''     message: Optional failure message.
Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String = "")
  If Not condition Then
    RaiseAssertFailure message, "expected True but got False", "XlflowAssert.AssertTrue"
  End If
End Sub

''' Asserts that a condition is False.
'''
''' Args:
'''     condition: Boolean condition to verify.
'''     message: Optional failure message.
Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String = "")
  If condition Then
    RaiseAssertFailure message, "expected False but got True", "XlflowAssert.AssertFalse"
  End If
End Sub

''' Fails the current test immediately.
'''
''' Args:
'''     message: Optional failure message.
Public Sub AssertFail(Optional ByVal message As String = "")
  RaiseAssertFailure message, "assertion failed", "XlflowAssert.AssertFail"
End Sub

''' Marks the current test as inconclusive.
'''
''' Args:
'''     message: Optional inconclusive reason.
Public Sub AssertInconclusive(Optional ByVal message As String = "")
  Dim detail As String
  detail = "inconclusive"
  If Len(message) > 0 Then
    detail = message
  End If
  Err.Raise vbObjectError + 516, "XlflowAssert.AssertInconclusive", detail
End Sub

''' Asserts that an object reference is Nothing.
'''
''' Args:
'''     value: Object reference to verify.
'''     message: Optional failure message.
Public Sub AssertIsNothing(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsObject(value) Then
    RaiseAssertFailure message, "expected an object but got a non-object", "XlflowAssert.AssertIsNothing"
    Exit Sub
  End If
  If Not (value Is Nothing) Then
    RaiseAssertFailure message, "expected Nothing but got an object reference", "XlflowAssert.AssertIsNothing"
  End If
End Sub

''' Asserts that an object reference is not Nothing.
'''
''' Args:
'''     value: Object reference to verify.
'''     message: Optional failure message.
Public Sub AssertIsNotNothing(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsObject(value) Then
    RaiseAssertFailure message, "expected an object but got a non-object", "XlflowAssert.AssertIsNotNothing"
    Exit Sub
  End If
  If value Is Nothing Then
    RaiseAssertFailure message, "expected an object reference but got Nothing", "XlflowAssert.AssertIsNotNothing"
  End If
End Sub

Private Sub RaiseAssertFailure(ByVal message As String, ByVal detail As String, ByVal source As String)
  If Len(message) > 0 Then
    detail = message & ": " & detail
  End If
  Err.Raise assertFailureNumber, source, detail
End Sub

Private Function FormatAssertValue(ByVal value As Variant) As String
  If IsObject(value) Then
    If value Is Nothing Then
      FormatAssertValue = "<Nothing>"
    Else
      FormatAssertValue = "<Object: " & TypeName(value) & ">"
    End If
    Exit Function
  End If

  If IsArray(value) Then
    FormatAssertValue = "<Array>"
    Exit Function
  End If

  If IsNull(value) Then
    FormatAssertValue = "<Null>"
  ElseIf IsEmpty(value) Then
    FormatAssertValue = "<Empty>"
  ElseIf VarType(value) = vbString Then
    FormatAssertValue = "<String: """ & EscapeAssertString(CStr(value)) & """>"
  ElseIf VarType(value) = vbBoolean Then
    If CBool(value) Then
      FormatAssertValue = "<Boolean: True>"
    Else
      FormatAssertValue = "<Boolean: False>"
    End If
  ElseIf VarType(value) = vbDate Then
    FormatAssertValue = "<Date: " & Format$(CDate(value), "yyyy-mm-dd hh:nn:ss") & ">"
  Else
    FormatAssertValue = "<" & AssertValueTypeName(value) & ": " & CStr(value) & ">"
  End If
End Function

Private Function EscapeAssertString(ByVal value As String) As String
  EscapeAssertString = Replace(value, """", """""")
  EscapeAssertString = Replace(EscapeAssertString, vbCrLf, "\r\n")
  EscapeAssertString = Replace(EscapeAssertString, vbCr, "\r")
  EscapeAssertString = Replace(EscapeAssertString, vbLf, "\n")
  EscapeAssertString = Replace(EscapeAssertString, vbTab, "\t")
End Function

Private Function AssertValueTypeName(ByVal value As Variant) As String
  Select Case VarType(value)
    Case vbByte
      AssertValueTypeName = "Byte"
    Case vbInteger
      AssertValueTypeName = "Integer"
    Case vbLong
      AssertValueTypeName = "Long"
    Case vbSingle
      AssertValueTypeName = "Single"
    Case vbDouble
      AssertValueTypeName = "Double"
    Case vbCurrency
      AssertValueTypeName = "Currency"
    Case vbDecimal
      AssertValueTypeName = "Decimal"
    Case Else
      AssertValueTypeName = TypeName(value)
  End Select
End Function

Private Function IsAssertNumeric(ByVal value As Variant) As Boolean
  If IsObject(value) Or IsArray(value) Or IsNull(value) Or IsEmpty(value) Then
    IsAssertNumeric = False
  Else
    IsAssertNumeric = IsNumeric(value)
  End If
End Function

Private Function IsAssertString(ByVal value As Variant) As Boolean
  IsAssertString = Not IsObject(value) And Not IsArray(value) And VarType(value) = vbString
End Function

Private Function ScalarArrayValuesEqual(ByVal expected As Variant, ByVal actual As Variant) As Boolean
  If IsObject(expected) Or IsObject(actual) Or IsArray(expected) Or IsArray(actual) Then
    ScalarArrayValuesEqual = False
  ElseIf IsNull(expected) Or IsNull(actual) Then
    ScalarArrayValuesEqual = IsNull(expected) And IsNull(actual)
  Else
    ScalarArrayValuesEqual = (expected = actual)
  End If
End Function

Private Function ArrayDimensionCount(ByVal value As Variant) As Long
  Dim dimension As Long
  Dim lowerBound As Long
  ' xlflow:disable-next-line VB004
  On Error Resume Next
  For dimension = 1 To 60
    Err.Clear
    lowerBound = LBound(value, dimension)
    If Err.Number <> 0 Then
      Exit For
    End If
    If lowerBound > UBound(value, dimension) Then
      Exit For
    End If
  Next dimension
  On Error GoTo 0
  ArrayDimensionCount = dimension - 1
End Function

Private Function RangeCellLabel(ByVal actualRange As Object, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
  Dim address As String
  Dim sheetName As String
  ' xlflow:disable-next-line VB004
  On Error Resume Next
  address = actualRange.Cells(rowIndex, columnIndex).Address(False, False)
  sheetName = actualRange.Worksheet.Name
  If Err.Number <> 0 Then
    Err.Clear
    RangeCellLabel = "row " & CStr(rowIndex) & ", column " & CStr(columnIndex)
  ElseIf Len(sheetName) > 0 Then
    RangeCellLabel = sheetName & "!" & address
  Else
    RangeCellLabel = address
  End If
  On Error GoTo 0
End Function
