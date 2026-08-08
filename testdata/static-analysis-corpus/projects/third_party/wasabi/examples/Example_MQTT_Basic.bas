Attribute VB_Name = "Example_MQTT_Basic"
Option Explicit

'/**
' * @description MQTT over WebSockets example.
' * Connects to a public broker and publishes a message to a test topic.
' */
Public Sub RunMqttExample()
    Dim handle As Long
    Dim url As String: url = "wss://test.mosquitto.org:8081" ' Standard MQTT-WS port
    Dim clientId As String: clientId = "WasabiClient_" & Int(Rnd * 1000)
    
    ' First, establish the WebSocket transport
    ' We specify 'mqtt' as the sub-protocol
    If Wasabi.WebSocketConnect(url, handle, , , "mqtt") Then
        Debug.Print "WebSocket Transport established."
        
        ' Perform MQTT CONNECT handshake
        If Wasabi.MqttConnect(clientId, , , 60, handle) Then
            Debug.Print "MQTT Connected as " & clientId
            
            ' Publish a message to a topic
            If Wasabi.MqttPublish("wasabi/test", "Wasabi MQTT is alive!", 1, False, , , handle) Then
                Debug.Print "Message published to wasabi/test"
            End If
            
            ' Wait briefly for ACKs
            Dim ack As String
            ack = Wasabi.MqttReceive(3000, handle)
            Debug.Print "Broker response: " & ack
            
            ' Disconnect MQTT and Close Socket
            Wasabi.MqttDisconnect handle
            Wasabi.WebSocketDisconnect handle
        End If
    Else
        Debug.Print "WS Transport failed: " & Wasabi.WasabiGetErrorDescription(handle)
    End If
End Sub
