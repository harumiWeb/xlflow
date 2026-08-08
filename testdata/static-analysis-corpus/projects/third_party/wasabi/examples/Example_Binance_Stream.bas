Attribute VB_Name = "Example_Binance_Stream"
Option Explicit

'/**
' * @description Real-world asynchronous WebSocket example.
' * Connects to multiple Binance public streams concurrently to receive live crypto trades.
' * Note: Requires a customized cWasabiAsync class to handle the UI/Sheet updates.
' */
Private handlers(0 To 4) As Object ' Should be strongly typed to your Async Class
Private handles(0 To 4) As Long

Public Sub StartCryptoStream()
    Dim i As Long
    Dim urls(0 To 4) As String
    
    urls(0) = "wss://stream.binance.com:9443/ws/btcusdt@trade"
    urls(1) = "wss://stream.binance.com:9443/ws/ethusdt@trade"
    urls(2) = "wss://stream.binance.com:9443/ws/bnbusdt@trade"
    urls(3) = "wss://stream.binance.com:9443/ws/solusdt@trade"
    urls(4) = "wss://stream.binance.com:9443/ws/xrpusdt@trade"
    
    For i = 0 To 4
        ' Initialize your custom handler (assuming cWasabiAsync exists)
        Set handlers(i) = New cWasabiAsync
        
        If Wasabi.WebSocketConnect(urls(i), handles(i)) Then
            Debug.Print "Connected to " & urls(i)
            Wasabi.WasabiUseAsync handlers(i), handles(i)
        Else
            Debug.Print "Failed to connect: " & urls(i)
            Debug.Print Wasabi.WasabiGetErrorDescription(handles(i))
        End If
    Next i
End Sub

Public Sub StopCryptoStream()
    ' Disconnect all active WebSocket connections cleanly
    Wasabi.WebSocketDisconnectAll
    Debug.Print "All streams stopped."
End Sub
