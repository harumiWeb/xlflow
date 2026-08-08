Attribute VB_Name = "Benchmark_Runner"
Option Explicit

'/**
' * @description High-resolution timing framework for Wasabi performance benchmarks.
' * Uses Windows API QueryPerformanceCounter for microsecond precision.
' */

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

Private m_Frequency As Currency
Private m_StartTime As Currency

Public Sub InitializeTimer()
    QueryPerformanceFrequency m_Frequency
End Sub

Public Sub StartTimer()
    QueryPerformanceCounter m_StartTime
End Sub

Public Function EndTimer() As Double
    Dim endTime As Currency
    QueryPerformanceCounter endTime
    ' Return elapsed time in milliseconds
    EndTimer = ((endTime - m_StartTime) / m_Frequency) * 1000
End Function

Public Sub RunAllBenchmarks()
    Debug.Print "========================================="
    Debug.Print "Starting Wasabi v2.3.7-beta Benchmarks..."
    Debug.Print "========================================="
    
    InitializeTimer
    
    Benchmark_Memory.RunMemoryBenchmarks
    Benchmark_Crypto.RunCryptoBenchmarks
    Benchmark_WebSockets.RunWebSocketBenchmarks
    
    Debug.Print "========================================="
    Debug.Print "Benchmarks Completed."
    Debug.Print "========================================="
End Sub
