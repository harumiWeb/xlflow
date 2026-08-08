Attribute VB_Name = "Example_MQTT_KeepAlive"
Option Explicit

'/**
' * @description Demonstrates manually sending MQTT Ping requests to keep the connection alive.
' * Uses the MqttSendPing function renamed in v2.3.7-beta.
' */
Public Sub RunMqttKeepAliveExample()
    Dim handle As Long
    Dim url As String: url = "wss://test.mosquitto.org:8081"
    
    If Wasabi.WebSocketConnect(url, handle, , , "mqtt") Then
        ' Connect with a 60-second Keep-Alive
        If Wasabi.MqttConnect("WasabiPingTest", , , 60, handle) Then
            Debug.Print "MQTT Connected."
            
            ' In a long-running synchronous process, if you don't publish/receive
            ' frequently, you must send a PINGREQ to prevent the broker from dropping you.
            
            ' Simulating idle time...
            Debug.Print "Idling..."
            
            ' Send a Ping Request
            If Wasabi.MqttSendPing(handle) Then
                Debug.Print "Ping request sent."
                
                ' The broker should respond with a PINGRESP (Ping Response)
                Dim resp As String
                resp = Wasabi.MqttReceive(1000, handle)
                
                ' Note: MqttReceive might swallow the PINGRESP internally and handle it,
                ' or return a specific opcode depending on Wasabi's internal implementation.
                Debug.Print "Ping cycle completed."
            End If
            
            Wasabi.MqttDisconnect handle
        End If
        Wasabi.WebSocketDisconnect handle
    End If
End Sub
