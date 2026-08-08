Attribute VB_Name = "Benchmark_Memory"
Option Explicit

'/**
' * @description Benchmarks for internal memory manipulation and buffer scanning.
' */

Public Sub RunMemoryBenchmarks()
    Debug.Print "Running Memory Benchmarks..."
    BenchmarkMemFind
End Sub

Private Sub BenchmarkMemFind()
    Dim i As Long
    Dim iterations As Long: iterations = 100000
    Dim elapsed As Double
    
    ' Simulating a large payload where we need to find the HTTP boundary
    Dim payload As String
    payload = String(10000, "A") & vbCrLf & vbCrLf & "BODY"
    
    Benchmark_Runner.StartTimer
    
    ' Note: If WasabiMemFind is private, this tests the equivalent logic
    Dim pos As Long
    For i = 1 To iterations
        pos = InStr(1, payload, vbCrLf & vbCrLf, vbBinaryCompare)
    Next i
    
    elapsed = Benchmark_Runner.EndTimer()
    
    Debug.Print "[Memory] Boundary Scan (" & iterations & " iterations): " & Format(elapsed, "0.00") & " ms"
End Sub
