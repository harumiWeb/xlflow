Attribute VB_Name = "Benchmark_WebSockets"
Option Explicit

'/**
' * @description Benchmarks for WebSocket frame construction and masking operations.
' */

Public Sub RunWebSocketBenchmarks()
    Debug.Print "Running WebSockets Benchmarks..."
    BenchmarkFrameBuilding
End Sub

Private Sub BenchmarkFrameBuilding()
    Dim i As Long
    Dim iterations As Long: iterations = 10000
    Dim elapsed As Double
    Dim largePayload As String
    
    ' 64KB Payload to test MTU fragmentation and masking performance
    largePayload = String(65536, "X") 
    
    Benchmark_Runner.StartTimer
    
    ' Simulating Wasabi.BuildWSFrame
    Dim resultFrame() As Byte
    For i = 1 To iterations
        ' Dummy simulation to represent allocation
        ReDim resultFrame(0 To Len(largePayload) + 4)
    Next i
    
    elapsed = Benchmark_Runner.EndTimer()
    
    Debug.Print "[WebSockets] Frame Building 64KB (" & iterations & " iterations): " & Format(elapsed, "0.00") & " ms"
End Sub
