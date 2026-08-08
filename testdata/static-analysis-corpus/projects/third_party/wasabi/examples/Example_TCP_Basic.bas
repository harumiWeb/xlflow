Attribute VB_Name = "Example_TCP_Basic"
Option Explicit

'/**
' * @description Basic TCP client example.
' * Demonstrates connecting to a raw TCP server, sending text, and receiving data.
' */
Public Sub RunTcpBasicExample()
    Dim handle As Long
    Dim host As String: host = "google.com"
    Dim port As Long: port = 80
    Dim request As String
    Dim response As String
    
    ' Connect via standard TCP
    If Wasabi.TcpConnect(host, port, handle) Then
        Debug.Print "TCP Connected to " & host & ":" & port
        
        ' Send a basic HTTP GET request (TCP raw)
        request = "GET / HTTP/1.1" & vbCrLf & "Host: " & host & vbCrLf & vbCrLf
        If Wasabi.TcpSendText(request, handle) Then
            Debug.Print "Request sent."
            
            ' Wait for response with a 5000ms timeout
            response = Wasabi.TcpReceiveUntil(vbCrLf & vbCrLf, 5000, handle)
            Debug.Print "HTTP Headers received:"
            Debug.Print response
        End If
        
        ' Clean up
        Wasabi.TcpDisconnect handle
    Else
        Debug.Print "TCP Connection failed: " & Wasabi.WasabiGetErrorDescription(handle)
    End If
End Sub
