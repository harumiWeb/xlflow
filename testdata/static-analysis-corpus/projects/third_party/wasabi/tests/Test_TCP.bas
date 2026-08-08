Attribute VB_Name = "Test_TCP"
Option Explicit

'/**
' * @description Unit tests for raw TCP sockets.
' * Focuses on connection stability and boundary conditions.
' */

Public Sub RunTests()
    Debug.Print "Running TCP Tests..."
    TestTcpConnection
    TestTcpInvalidHost
End Sub

Private Sub TestTcpConnection()
    Dim handle As Long
    Dim connected As Boolean
    
    connected = Wasabi.TcpConnect("google.com", 80, handle)
    Test_Runner.AssertTrue connected, "TcpConnect successfully binds to valid host"
    
    If connected Then
        Wasabi.TcpDisconnect handle
    End If
End Sub

Private Sub TestTcpInvalidHost()
    Dim handle As Long
    Dim connected As Boolean
    
    ' Deliberately trying to connect to a non-existent domain to test error handling
    connected = Wasabi.TcpConnect("this.domain.should.not.exist.internal", 80, handle)
    Test_Runner.AssertTrue connected = False, "TcpConnect correctly returns False on invalid host"
    
    Dim errDesc As String
    errDesc = Wasabi.WasabiGetErrorDescription(handle)
    Test_Runner.AssertTrue Len(errDesc) > 0, "Error description is populated on connection failure"
End Sub
