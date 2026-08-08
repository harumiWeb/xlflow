Attribute VB_Name = "Example_WebSocket_Broadcast"
Option Explicit

'/**
' * @description Demonstrates broadcasting a text message to all active WebSocket connections.
' * Uses the newly renamed WebSocketBroadcastText function from v2.3.7-beta.
' */
Public Sub RunWebSocketBroadcastExample()
    Dim handles(0 To 2) As Long
    Dim i As Long
    
    ' Connect to multiple endpoints (simulated here with the same echo server for demonstration)
    For i = 0 To 2
        If Wasabi.WebSocketConnect("wss://echo.websocket.events", handles(i)) Then
            Debug.Print "Connected WS handle " & handles(i)
        End If
    Next i
    
    ' Broadcast a message to ALL connected WebSocket handles simultaneously
    Dim broadcastMsg As String: broadcastMsg = "{""status"": ""maintenance"", ""time"": 5}"
    Wasabi.WebSocketBroadcastText broadcastMsg
    Debug.Print "Broadcasted: " & broadcastMsg
    
    ' Read responses and cleanup
    Dim resp As String
    For i = 0 To 2
        ' Small delay to allow echo to return
        DoEvents
        resp = Wasabi.WebSocketReceiveText(handles(i))
        Debug.Print "Handle " & handles(i) & " received: " & resp
        
        Wasabi.WebSocketDisconnect handles(i)
    Next i
End Sub
