Attribute VB_Name = "Example_MQTT_Subscribe"
Option Explicit

'/**
' * @description MQTT over WebSockets: Subscribing to a topic and listening for incoming messages.
' */
Public Sub RunMqttSubscribeExample()
    Dim handle As Long
    Dim url As String: url = "wss://test.mosquitto.org:8081"
    Dim clientId As String: clientId = "WasabiSub_" & Int(Rnd * 1000)
    Dim topic As String: topic = "wasabi/test/broadcast"
    
    If Wasabi.WebSocketConnect(url, handle, , , "mqtt") Then
        If Wasabi.MqttConnect(clientId, , , 60, handle) Then
            Debug.Print "MQTT Connected. Subscribing to: " & topic
            
            ' Subscribe to the topic with QoS 0
            If Wasabi.MqttSubscribe(topic, 0, handle) Then
                Debug.Print "Subscribed. Listening for 10 seconds..."
                
                Dim endTime As Double
                endTime = Timer + 10 ' Listen for 10 seconds
                
                Dim msg As String
                Do While Timer < endTime
                    DoEvents
                    ' Receive MQTT application messages
                    msg = Wasabi.MqttReceive(100, handle)
                    If msg <> "" Then
                        Debug.Print "Received from broker: " & msg
                    End If
                Loop
                
                Debug.Print "Finished listening."
            End If
            
            Wasabi.MqttDisconnect handle
        End If
        Wasabi.WebSocketDisconnect handle
    End If
End Sub
