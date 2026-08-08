Attribute VB_Name = "Example_Advanced_Diagnostics"
Option Explicit

'/**
' * @description Demonstrates tweaking internal buffer sizes and extracting deep technical details
' * when debugging hard-to-track connection issues.
' */
Public Sub RunDiagnosticsExample()
    Dim handle As Long
    Dim url As String: url = "wss://echo.websocket.events"
    
    ' 1. Buffer Size Adjustment
    ' Sometimes default buffers are too small for massive payloads.
    ' We use WebSocketSetBufferSize (renamed in v2.3.7-beta) before connecting.
    ' E.g., setting Rx and Tx buffers to 128KB
    ' Note: Requires a pre-allocated handle or global setting depending on Wasabi's API structure.
    
    If Wasabi.WebSocketConnect(url, handle) Then
        
        ' Adjust buffer if Wasabi allows post-connect buffer sizing, or just for demonstration
        ' Wasabi.WebSocketSetBufferSize 131072, 131072, handle
        
        Debug.Print "Connected. Deliberately causing an issue or checking stats..."
        
        ' 2. Getting Technical Details
        ' As mentioned in the release notes, if you run into anything, 
        ' you should drop the output of WebSocketGetTechnicalDetails.
        
        Dim techDetails As String
        techDetails = Wasabi.WebSocketGetTechnicalDetails(handle)
        
        Debug.Print "--- TECHNICAL DETAILS ---"
        Debug.Print techDetails
        Debug.Print "-------------------------"
        
        Dim errDesc As String
        errDesc = Wasabi.WasabiGetErrorDescription(handle)
        If errDesc <> "Success" And errDesc <> "" Then
            Debug.Print "Current Error State: " & errDesc
        End If
        
        Wasabi.WebSocketDisconnect handle
    End If
End Sub
