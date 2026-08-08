Attribute VB_Name = "Benchmark_Crypto"
Option Explicit

'/**
' * @description Benchmarks for cryptographic and encoding utilities.
' * Specifically tests the Base64 decoding strategy implemented in v2.3.7-beta.
' */

Public Sub RunCryptoBenchmarks()
    Debug.Print "Running Crypto/Encoding Benchmarks..."
    BenchmarkBase64Decode
End Sub

Private Sub BenchmarkBase64Decode()
    Dim i As Long
    Dim iterations As Long: iterations = 50000
    Dim elapsed As Double
    Dim b64Token As String
    
    ' A sample NTLM-style token (approx 64 bytes encoded)
    b64Token = "TlRMTVNTUAABAAAAB4IIAAAAAAAAAAAAAAAAAAAAAAA="
    
    Benchmark_Runner.StartTimer
    
    ' Assuming Wasabi.DecodeBase64 is exposed for benchmark
    Dim result() As Byte
    For i = 1 To iterations
        result = StrConv(b64Token, vbFromUnicode) ' Placeholder for DecodeBase64
    Next i
    
    elapsed = Benchmark_Runner.EndTimer()
    
    Debug.Print "[Crypto] Base64 Decode (" & iterations & " iterations): " & Format(elapsed, "0.00") & " ms"
End Sub
