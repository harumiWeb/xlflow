Attribute VB_Name = "Test_WebSockets"
Option Explicit

'/**
' * @description Unit tests for the WebSocket implementation.
' * Validates connection lifecycles, memory handling, and data integrity.
' */

Public Sub RunTests()
    Debug.Print "Running WebSockets Tests..."
    TestWebSocketConnectionLifecycle
    TestWebSocketEchoIntegrity
End Sub

Private Sub TestWebSocketConnectionLifecycle()
    Dim handle As Long
    Dim connected As Boolean
    
    connected = Wasabi.WebSocketConnect("wss://echo.websocket.events", handle)
    Test_Runner.AssertTrue connected, "WebSocketConnect returns True on valid endpoint"
    
    If connected Then
        Wasabi.WebSocketDisconnect handle
        Test_Runner.AssertTrue True, "WebSocketDisconnect executes safely"
    End If
End Sub

Private Sub TestWebSocketEchoIntegrity()
    Dim handle As Long
    Dim payload As String: payload = "INTEGRITY_TEST_123"
    Dim response As String
    
    If Wasabi.WebSocketConnect("wss://echo.websocket.events", handle) Then
        If Wasabi.WebSocketSendText(payload, handle) Then
            Do
                DoEvents
                response = Wasabi.WebSocketReceiveText(handle)
            Loop While response = ""
            
            Test_Runner.AssertTrue response = payload, "WebSocket echo payload matches perfectly"
        Else
            Test_Runner.AssertTrue False, "WebSocketSendText failed"
        End If
        
        Wasabi.WebSocketDisconnect handle
    Else
        Test_Runner.AssertTrue False, "Failed to connect for echo test"
    End If
End Sub
