# Wasabi Example Suite

> [!NOTE]
> These examples can run and display results in any ![](../resources/svg/ms-office.svg) **Microsoft Office** program.

This repository section provides a comprehensive collection of production ready examples designed to demonstrate the integration of Wasabi into various architectural paradigms. The modules cover a wide spectrum of use cases, ranging from foundational synchronous connections to complex asynchronous event driven architectures suitable for high frequency data streaming and strict corporate proxy environments.

## Core Capabilities Demonstrated

* **WebSocket Implementations:** Examples covering synchronous blocking requests, raw binary frame handling, and concurrent message broadcasting across multiple active handles.
* **Asynchronous Streaming:** A real world demonstration utilizing the Binance public API to capture live cryptocurrency trades via non blocking callbacks and native window message subclassing.
* **MQTT Telemetry:** Full workflow examples for IoT and messaging scenarios, including connection handshakes, topic subscription, message publishing, and manual Keep Alive pinging.
* **Raw TCP Sockets:** Foundational examples for establishing direct TCP streams and broadcasting raw byte payloads to multiple connected clients.
* **Enterprise and Diagnostics:** Configurations for environments requiring TLS verification, MTU awareness, proxy routing, and deep technical buffer diagnostics.

## Example Index

### WebSockets
* `Example_WebSocket_Sync.bas`: Standard connection polling and text frame exchange.
* `Example_WebSocket_Binary.bas`: Serialized binary payload transmission and memory array handling.
* `Example_WebSocket_Broadcast.bas`: Simultaneous multiplexed broadcasting to multiple endpoints.

### Asynchronous and Streaming
* `Example_Binance_Stream.bas`: High throughput asynchronous data ingestion.
* `cWasabiHandler.cls`: The required class interface mapping for native window message callbacks.

### MQTT Protocol
* `Example_MQTT_Basic.bas`: Handshake initiation and standard QoS publishing.
* `Example_MQTT_Subscribe.bas`: Continuous event listening on designated broker topics.
* `Example_MQTT_KeepAlive.bas`: Connection persistence strategies using manual Ping requests.

### Raw TCP Operations
* `Example_TCP_Basic.bas`: Standard host resolution and raw HTTP GET request formatting.
* `Example_TCP_Broadcast.bas`: Raw byte stream replication across active socket handles.

### Advanced Configuration and Debugging
* `Example_Proxy_And_Security.bas`: MTU awareness mapping and TCP NoDelay configuration.
* `Example_Advanced_Diagnostics.bas`: Internal buffer sizing and technical state extraction for hard to track networking issues.

## How to Use

1. Import `Wasabi.bas` into your VBA or VB6 project.
2. Import the desired example module from this suite.
3. If exploring asynchronous examples, ensure `cWasabiHandler.cls` is also imported and properly instantiated. For example, you can use this:
```vba
Option Explicit

'/**
' * @description Custom handler class for asynchronous WebSocket events.
' * Implements the required interface for WasabiUseAsync.
' */

Public Sub OnConnect(ByVal handle As Long)
    Debug.Print "[Async] Handle " & handle & " connected successfully."
End Sub

Public Sub OnReceive(ByVal handle As Long)
    Dim msg As String
    ' Retrieve messages from the internal queue
    Do
        msg = Wasabi.WebSocketReceiveText(handle)
        If msg <> "" Then
            Debug.Print "[Async Receive] " & msg
        End If
    Loop While msg <> ""
End Sub

Public Sub OnReadyToSend(ByVal handle As Long)
    ' Called when the socket is ready for more data
End Sub

Public Sub OnError(ByVal handle As Long, ByVal errorCode As Long, ByVal eventType As Long)
    Debug.Print "[Async Error] Code: " & errorCode & " on event " & eventType
End Sub

Public Sub OnClose(ByVal handle As Long)
    Debug.Print "[Async] Connection closed for handle " & handle
End Sub

Public Sub OnDisconnect(ByVal handle As Long)
    Debug.Print "[Async] Handler detached from handle " & handle
End Sub
``` 
4. Execute the public subroutines directly from the Immediate Window or bind them to your application UI.
