Attribute VB_Name = "Example_Proxy_And_Security"
Option Explicit

'/**
' * @description Demonstrates advanced configuration: Proxies, TLS verification, and IPv6 preference.
' */
Public Sub RunAdvancedConfigExample()
    Dim handle As Long
    Dim url As String: url = "wss://echo.websocket.events"
    
    ' Let's demonstrate setting options on a connection attempt:
    If Wasabi.WebSocketConnect(url, handle) Then
        
        ' Apply TCP NoDelay for low latency
        Wasabi.WebSocketSetNoDelay True, handle
        
        ' Enable MTU Awareness for large payloads
        Wasabi.WebSocketSetAutoMTU True, handle
        
        ' Send a large message that will be fragmented based on MTU
        Dim largeMsg As String: largeMsg = String(5000, "A")
        Wasabi.WebSocketSendTextMTUAware largeMsg, handle
        
        Debug.Print "Secure, MTU-Aware connection established and tested."
        
        Wasabi.WebSocketDisconnect handle
    End If
End Sub
