Attribute VB_Name = "Test_Runner"
Option Explicit

'/**
' * @description Lightweight unit testing framework for Wasabi.
' * Tracks assertions and outputs results to the Immediate Window.
' */

Private m_passCount As Long
Private m_failCount As Long

Public Sub AssertTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        m_passCount = m_passCount + 1
        Debug.Print "[PASS] " & testName
    Else
        m_failCount = m_failCount + 1
        Debug.Print "[FAIL] " & testName
    End If
End Sub

Public Sub RunAllTests()
    m_passCount = 0
    m_failCount = 0
    
    Debug.Print "========================================="
    Debug.Print "Starting Wasabi v2.3.7-beta Test Suite..."
    Debug.Print "========================================="
    
    ' Call individual test modules here
    Test_Utils.RunTests
    Test_WebSockets.RunTests
    Test_TCP.RunTests
    Test_MQTT.RunTests
    Test_Memory.RunTests
    
    Debug.Print "========================================="
    Debug.Print "Test Suite Completed."
    Debug.Print "Passed: " & m_passCount & " | Failed: " & m_failCount
    Debug.Print "========================================="
End Sub
