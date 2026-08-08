Attribute VB_Name = "Example_TCP_Broadcast"
Option Explicit

'/**
' * @description Demonstrates TCP broadcasting to multiple connected clients.
' * Useful if Wasabi is being used in a multi-connection scenario or proxy routing.
' */
Public Sub RunTcpBroadcastExample()
    ' Note: This assumes you have multiple active TCP handles connected.
    ' We will simulate an array of handles.
    
    Dim activeHandles(0 To 2) As Long
    Dim payload() As Byte
    Dim textMsg As String: textMsg = "SERVER REBOOTING IN 5 MINUTES"
    
    ' Convert text to byte array for raw TCP broadcasting
    payload = StrConv(textMsg, vbFromUnicode)
    
    ' Assuming activeHandles are populated with valid TCP socket handles...
    ' For the sake of example, we just show the call:
    
    ' Broadcast the binary payload to all active TCP connections mapped in Wasabi
    Wasabi.TcpBroadcastBinary payload
    
    Debug.Print "Broadcasted payload to all TCP sockets."
End Sub
