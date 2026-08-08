Attribute VB_Name = "Example_WebSocket_Sync"
Option Explicit

'/**
' * @description Demonstrates a basic synchronous WebSocket connection.
' * Connects to an echo server, sends a message, and waits for a response.
' */
Public Sub RunSyncWebSocketExample()
    Dim handle As Long
    Dim url As String: url = "wss://echo.websocket.events"
    Dim msg As String: msg = "Hello from Wasabi!"
    Dim response As String
    
    ' Attempt to connect to the WebSocket server
    If Wasabi.WebSocketConnect(url, handle) Then
        Debug.Print "Connected to: " & url & " (Handle: " & handle & ")"
        
        ' Send a text message
        If Wasabi.WebSocketSendText(msg, handle) Then
            Debug.Print "Message sent: " & msg
            
            ' Poll for a response (synchronous wait)
            ' In real apps, you might loop or use a timeout
            Do
                DoEvents
                response = Wasabi.WebSocketReceiveText(handle)
            Loop While response = ""
            
            Debug.Print "Received Echo: " & response
        End If
        
        ' Close the connection gracefully
        Wasabi.WebSocketDisconnect handle
        Debug.Print "Disconnected."
    Else
        Debug.Print "Connection failed."
        Debug.Print "Error: " & Wasabi.WasabiGetErrorDescription(handle)
    End If
End Sub
