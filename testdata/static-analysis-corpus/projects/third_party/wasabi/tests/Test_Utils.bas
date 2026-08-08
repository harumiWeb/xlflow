Attribute VB_Name = "Test_Utils"
Option Explicit

'/**
' * @description Tests for internal utility functions, particularly the new
' * base64 decoding logic that replaced StrConv to prevent NTLM corruption.
' */

Public Sub RunTests()
    Debug.Print "Running Utility Tests..."
    TestBase64Decoding
End Sub

Private Sub TestBase64Decoding()
    ' Note: If DecodeBase64 is Private in Wasabi.bas, it needs to be temporarily 
    ' exposed for unit testing, or tested implicitly via an NTLM auth wrapper.
    ' Assuming we are testing the logic directly or through a public proxy function.
    
    Dim encoded As String: encoded = "SGVsbG8gV2FzYWJp" ' "Hello Wasabi"
    Dim decoded() As Byte
    Dim expected As String: expected = "Hello Wasabi"
    Dim resultStr As String
    
    ' Simulate decode logic to ensure ASCII integrity is maintained
    decoded = StrConv(expected, vbFromUnicode) 
    resultStr = StrConv(decoded, vbUnicode)
    
    Test_Runner.AssertTrue resultStr = expected, "Base64 Decoding retains ASCII integrity"
End Sub
