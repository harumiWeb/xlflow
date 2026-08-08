Attribute VB_Name = "Example_WebSocket_Binary"
Option Explicit

'/**
' * @description Demonstrates sending and receiving raw binary data over WebSockets.
' * Useful for custom protocols, file transfers, or serialized data (like Protobuf).
' */
Public Sub RunBinaryWebSocketExample()
    Dim handle As Long
    Dim url As String: url = "wss://echo.websocket.events"
    Dim outBuffer() As Byte
    Dim inBuffer() As Byte
    Dim i As Long
    
    ' Prepare a dummy binary payload (0 to 9)
    ReDim outBuffer(0 To 9)
    For i = 0 To 9
        outBuffer(i) = CByte(i)
    Next i
    
    If Wasabi.WebSocketConnect(url, handle) Then
        Debug.Print "Connected to: " & url
        
        ' Send Binary Frame
        If Wasabi.WebSocketSendBinary(outBuffer, handle) Then
            Debug.Print "Sent " & (UBound(outBuffer) + 1) & " bytes of binary data."
            
            ' Wait for echo
            Do
                DoEvents
                inBuffer = Wasabi.WebSocketReceiveBinary(handle)
            Loop While (Not Not inBuffer) = 0 ' Check if array is initialized
            
            Debug.Print "Received " & (UBound(inBuffer) + 1) & " bytes of binary data."
            Debug.Print "First byte: " & inBuffer(0) & ", Last byte: " & inBuffer(UBound(inBuffer))
        End If
        
        Wasabi.WebSocketDisconnect handle
    Else
        Debug.Print "Connection failed."
    End If
End Sub
