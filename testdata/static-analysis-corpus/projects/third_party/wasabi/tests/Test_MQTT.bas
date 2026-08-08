Attribute VB_Name = "Test_MQTT"
Option Explicit

'/**
' * @description Unit tests for the MQTT implementation.
' * Validates the handshake, protocol selection, and basic publishing mechanisms.
' */

Public Sub RunTests()
    Debug.Print "Running MQTT Tests..."
    TestMqttConnection
    TestMqttInvalidTopic
End Sub

Private Sub TestMqttConnection()
    Dim handle As Long
    Dim connectedWS As Boolean
    Dim connectedMQTT As Boolean
    
    ' Connect WS with MQTT subprotocol
    connectedWS = Wasabi.WebSocketConnect("wss://test.mosquitto.org:8081", handle, , , "mqtt")
    Test_Runner.AssertTrue connectedWS, "WebSocket transport establishes with 'mqtt' subprotocol"
    
    If connectedWS Then
        connectedMQTT = Wasabi.MqttConnect("WasabiTestRunner_" & Int(Rnd * 1000), , , 60, handle)
        Test_Runner.AssertTrue connectedMQTT, "MQTT Connect handshake completes successfully"
        
        If connectedMQTT Then
            Wasabi.MqttDisconnect handle
        End If
        Wasabi.WebSocketDisconnect handle
    End If
End Sub

Private Sub TestMqttInvalidTopic()
    Dim handle As Long
    ' Assuming we just test if the library handles a null/empty topic gracefully
    Dim publishResult As Boolean
    
    ' Try to publish without a connection
    publishResult = Wasabi.MqttPublish("", "payload", 0, False, , , handle)
    Test_Runner.AssertTrue publishResult = False, "MqttPublish correctly fails on empty topic or invalid state"
End Sub
