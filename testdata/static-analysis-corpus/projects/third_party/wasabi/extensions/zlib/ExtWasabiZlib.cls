VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtWasabiZlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/**
' * ============================================================================
' * ExtWasabiZlib v1.0.0
' * Copyright (c) 2026 UesleiDev
' *
' * @description Compression handler extension for Wasabi.
' * Implements permessage-deflate (RFC 7692) via zlib.dll.
' * Drop this class into your project and pass an instance to
' * WasabiUseCompression before connecting.
' *
' * Requires zlib1.dll (or zlib.dll) accessible on the system PATH
' * or in the same directory as the host application.
' * ============================================================================
' */

Option Explicit

' ============================================================================
' API DECLARATIONS
' ============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function deflateInit2_ Lib "zlib1.dll" ( _
        ByRef strm As z_stream, _
        ByVal level As Long, _
        ByVal method As Long, _
        ByVal windowBits As Long, _
        ByVal memLevel As Long, _
        ByVal strategy As Long, _
        ByVal version As String, _
        ByVal stream_size As Long) As Long

    Private Declare PtrSafe Function deflate Lib "zlib1.dll" ( _
        ByRef strm As z_stream, _
        ByVal flush As Long) As Long

    Private Declare PtrSafe Function deflateEnd Lib "zlib1.dll" ( _
        ByRef strm As z_stream) As Long

    Private Declare PtrSafe Function inflateInit2_ Lib "zlib1.dll" ( _
        ByRef strm As z_stream, _
        ByVal windowBits As Long, _
        ByVal version As String, _
        ByVal stream_size As Long) As Long

    Private Declare PtrSafe Function inflate Lib "zlib1.dll" ( _
        ByRef strm As z_stream, _
        ByVal flush As Long) As Long

    Private Declare PtrSafe Function inflateEnd Lib "zlib1.dll" ( _
        ByRef strm As z_stream) As Long

    Private Declare PtrSafe Function deflateReset Lib "zlib1.dll" ( _
        ByRef strm As z_stream) As Long

    Private Declare PtrSafe Function inflateReset Lib "zlib1.dll" ( _
        ByRef strm As z_stream) As Long

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef dest As Any, ByRef src As Any, ByVal size As Long)

    Private Declare PtrSafe Sub CopyMemoryFromPtr Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef dest As Any, ByVal src As LongPtr, ByVal size As Long)

    Private Declare PtrSafe Sub CopyMemoryToPtr Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal dest As LongPtr, ByRef src As Any, ByVal size As Long)
#Else
    Private Declare Function deflateInit2_ Lib "zlib1.dll" ( _
        ByRef strm As z_stream, _
        ByVal level As Long, _
        ByVal method As Long, _
        ByVal windowBits As Long, _
        ByVal memLevel As Long, _
        ByVal strategy As Long, _
        ByVal version As String, _
        ByVal stream_size As Long) As Long

    Private Declare Function deflate Lib "zlib1.dll" ( _
        ByRef strm As z_stream, _
        ByVal flush As Long) As Long

    Private Declare Function deflateEnd Lib "zlib1.dll" ( _
        ByRef strm As z_stream) As Long

    Private Declare Function inflateInit2_ Lib "zlib1.dll" ( _
        ByRef strm As z_stream, _
        ByVal windowBits As Long, _
        ByVal version As String, _
        ByVal stream_size As Long) As Long

    Private Declare Function inflate Lib "zlib1.dll" ( _
        ByRef strm As z_stream, _
        ByVal flush As Long) As Long

    Private Declare Function inflateEnd Lib "zlib1.dll" ( _
        ByRef strm As z_stream) As Long

    Private Declare Function deflateReset Lib "zlib1.dll" ( _
        ByRef strm As z_stream) As Long

    Private Declare Function inflateReset Lib "zlib1.dll" ( _
        ByRef strm As z_stream) As Long

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef dest As Any, ByRef src As Any, ByVal size As Long)

    Private Declare Sub CopyMemoryFromPtr Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef dest As Any, ByVal src As Long, ByVal size As Long)

    Private Declare Sub CopyMemoryToPtr Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal dest As Long, ByRef src As Any, ByVal size As Long)
#End If

' ============================================================================
' ZLIB CONSTANTS
' ============================================================================

Private Const ZLIB_VERSION          As String = "1.2.11"
Private Const Z_OK                  As Long = 0
Private Const Z_STREAM_END          As Long = 1
Private Const Z_BUF_ERROR           As Long = -5
Private Const Z_SYNC_FLUSH          As Long = 2
Private Const Z_FINISH              As Long = 4
Private Const Z_DEFLATED            As Long = 8
Private Const Z_DEFAULT_COMPRESSION As Long = -1
Private Const Z_DEFAULT_STRATEGY    As Long = 0
Private Const Z_MEM_DEFAULT         As Long = 8
Private Const CHUNK_SIZE            As Long = 65536

' ============================================================================
' ZLIB STREAM STRUCT
' ============================================================================

' z_stream layout (matches zlib.h for both 32 and 64-bit)
' The struct is padded to work safely across architectures by using
' a fixed-size byte array. Total size: 112 bytes on x64, 56 on x86.
#If Win64 Then
Private Type z_stream
    next_in   As LongPtr
    avail_in  As Long
    total_in  As LongPtr
    next_out  As LongPtr
    avail_out As Long
    total_out As LongPtr
    msg       As LongPtr
    state     As LongPtr
    zalloc    As LongPtr
    zfree     As LongPtr
    opaque    As LongPtr
    data_type As Long
    adler     As LongPtr
    reserved  As LongPtr
End Type
#Else
Private Type z_stream
    next_in   As Long
    avail_in  As Long
    total_in  As Long
    next_out  As Long
    avail_out As Long
    total_out As Long
    msg       As Long
    state     As Long
    zalloc    As Long
    zfree     As Long
    opaque    As Long
    data_type As Long
    adler     As Long
    reserved  As Long
End Type
#End If

' ============================================================================
' STATE
' ============================================================================

Private m_DeflateReady   As Boolean
Private m_InflateReady   As Boolean
Private m_DeflateStream  As z_stream
Private m_InflateStream  As z_stream
Private m_LastError      As String
Private m_CompressionLevel As Long

' ============================================================================
' INIT / TEARDOWN
' ============================================================================

Private Sub Class_Initialize()
    m_CompressionLevel = Z_DEFAULT_COMPRESSION
    m_LastError = ""
End Sub

Private Sub Class_Terminate()
    TeardownStreams
End Sub

Private Sub TeardownStreams()
    If m_DeflateReady Then
        deflateEnd m_DeflateStream
        m_DeflateReady = False
    End If
    If m_InflateReady Then
        inflateEnd m_InflateStream
        m_InflateReady = False
    End If
End Sub

' ============================================================================
' WASABI COMPRESSION HANDLER INTERFACE
' Called automatically by Wasabi on connect/disconnect.
' ============================================================================

Public Sub OnConnect(ByVal handle As Long)
    ' Streams are initialized lazily on first use so window bits
    ' from the server handshake are available by then.
End Sub

Public Sub OnDisconnect(ByVal handle As Long)
    TeardownStreams
End Sub

' ============================================================================
' PUBLIC API
' ============================================================================

'/**
' * @brief Sets the zlib compression level (1 = fastest, 9 = best, -1 = default).
' * Call this before connecting.
' */
Public Sub SetCompressionLevel(ByVal level As Long)
    If level >= -1 And level <= 9 Then
        m_CompressionLevel = level
    End If
End Sub

'/**
' * @brief Returns the last zlib error description, if any.
' */
Public Function GetLastError() As String
    GetLastError = m_LastError
End Function

'/**
' * @brief Compresses a byte array using raw deflate (no zlib or gzip header).
' * Called by Wasabi before sending a WebSocket frame when permessage-deflate
' * is active. windowBits is negative (e.g. -15) as required by RFC 7692.
' * contextTakeover controls whether the deflate stream is reset between frames.
' * @param data Input byte array.
' * @param windowBits Negative window size (-8 to -15).
' * @param contextTakeover True to reuse the stream context across frames.
' * @return Compressed byte array with the trailing 00 00 FF FF bytes stripped,
' *         as required by the permessage-deflate spec.
' */
Public Function Deflate(ByRef data() As Byte, ByVal windowBits As Long, ByVal contextTakeover As Boolean) As Byte()
    Dim result()    As Byte
    Dim outBuf()    As Byte
    Dim outTotal    As Long
    Dim ret         As Long
    Dim chunk()     As Byte
    Dim chunkLen    As Long
    Dim dataLen     As Long

    m_LastError = ""

    dataLen = UBound(data) - LBound(data) + 1
    If dataLen = 0 Then
        Deflate = result
        Exit Function
    End If

    ' Initialize stream on first use or if context takeover is off
    If Not m_DeflateReady Or Not contextTakeover Then
        If m_DeflateReady Then
            If Not contextTakeover Then
                deflateEnd m_DeflateStream
                m_DeflateReady = False
            Else
                deflateReset m_DeflateStream
            End If
        End If

        If Not m_DeflateReady Then
            Dim zs As z_stream
            m_DeflateStream = zs
            ret = deflateInit2_(m_DeflateStream, m_CompressionLevel, Z_DEFLATED, windowBits, Z_MEM_DEFAULT, Z_DEFAULT_STRATEGY, ZLIB_VERSION, LenB(m_DeflateStream))
            If ret <> Z_OK Then
                m_LastError = "deflateInit2_ failed: " & ret
                Deflate = result
                Exit Function
            End If
            m_DeflateReady = True
        End If
    End If

    ' Point stream at input
    ReDim chunk(0 To CHUNK_SIZE - 1)
    ReDim outBuf(0 To CHUNK_SIZE - 1)
    outTotal = 0

    m_DeflateStream.next_in = VarPtr(data(LBound(data)))
    m_DeflateStream.avail_in = dataLen

    Do
        m_DeflateStream.next_out = VarPtr(chunk(0))
        m_DeflateStream.avail_out = CHUNK_SIZE

        ret = deflate(m_DeflateStream, Z_SYNC_FLUSH)

        If ret <> Z_OK And ret <> Z_BUF_ERROR And ret <> Z_STREAM_END Then
            m_LastError = "deflate() failed: " & ret
            Deflate = result
            Exit Function
        End If

        chunkLen = CHUNK_SIZE - m_DeflateStream.avail_out

        If chunkLen > 0 Then
            If outTotal + chunkLen > UBound(outBuf) + 1 Then
                ReDim Preserve outBuf(0 To outTotal + chunkLen + CHUNK_SIZE - 1)
            End If
            CopyMemory outBuf(outTotal), chunk(0), chunkLen
            outTotal = outTotal + chunkLen
        End If

    Loop While m_DeflateStream.avail_out = 0

    ' Strip the trailing 00 00 FF FF sync marker (RFC 7692 section 7.2.1)
    If outTotal >= 4 Then
        If outBuf(outTotal - 4) = &H0 And _
           outBuf(outTotal - 3) = &H0 And _
           outBuf(outTotal - 2) = &HFF And _
           outBuf(outTotal - 1) = &HFF Then
            outTotal = outTotal - 4
        End If
    End If

    If outTotal > 0 Then
        ReDim result(0 To outTotal - 1)
        CopyMemory result(0), outBuf(0), outTotal
    End If

    Deflate = result
End Function

'/**
' * @brief Decompresses a byte array using raw inflate.
' * Called by Wasabi when receiving a compressed WebSocket frame.
' * Appends the 00 00 FF FF tail before inflating, as required by RFC 7692.
' * @param data Compressed byte array (without the sync tail).
' * @param windowBits Negative window size (-8 to -15).
' * @param contextTakeover True to reuse the stream context across frames.
' * @return Decompressed byte array.
' */
Public Function Inflate(ByRef data() As Byte, ByVal windowBits As Long, ByVal contextTakeover As Boolean) As Byte()
    Dim result()    As Byte
    Dim outBuf()    As Byte
    Dim outTotal    As Long
    Dim ret         As Long
    Dim chunk()     As Byte
    Dim chunkLen    As Long
    Dim dataLen     As Long
    Dim inputBuf()  As Byte

    m_LastError = ""

    dataLen = UBound(data) - LBound(data) + 1
    If dataLen = 0 Then
        Inflate = result
        Exit Function
    End If

    ' Re-append the 00 00 FF FF tail stripped during compression
    ReDim inputBuf(0 To dataLen + 3)
    CopyMemory inputBuf(0), data(LBound(data)), dataLen
    inputBuf(dataLen)     = &H0
    inputBuf(dataLen + 1) = &H0
    inputBuf(dataLen + 2) = &HFF
    inputBuf(dataLen + 3) = &HFF

    ' Initialize stream on first use or if context takeover is off
    If Not m_InflateReady Or Not contextTakeover Then
        If m_InflateReady Then
            If Not contextTakeover Then
                inflateEnd m_InflateStream
                m_InflateReady = False
            Else
                inflateReset m_InflateStream
            End If
        End If

        If Not m_InflateReady Then
            Dim zs As z_stream
            m_InflateStream = zs
            ret = inflateInit2_(m_InflateStream, windowBits, ZLIB_VERSION, LenB(m_InflateStream))
            If ret <> Z_OK Then
                m_LastError = "inflateInit2_ failed: " & ret
                Inflate = result
                Exit Function
            End If
            m_InflateReady = True
        End If
    End If

    ReDim chunk(0 To CHUNK_SIZE - 1)
    ReDim outBuf(0 To CHUNK_SIZE - 1)
    outTotal = 0

    m_InflateStream.next_in = VarPtr(inputBuf(0))
    m_InflateStream.avail_in = dataLen + 4

    Do
        m_InflateStream.next_out = VarPtr(chunk(0))
        m_InflateStream.avail_out = CHUNK_SIZE

        ret = inflate(m_InflateStream, Z_SYNC_FLUSH)

        If ret <> Z_OK And ret <> Z_STREAM_END And ret <> Z_BUF_ERROR Then
            m_LastError = "inflate() failed: " & ret
            Inflate = result
            Exit Function
        End If

        chunkLen = CHUNK_SIZE - m_InflateStream.avail_out

        If chunkLen > 0 Then
            If outTotal + chunkLen > UBound(outBuf) + 1 Then
                ReDim Preserve outBuf(0 To outTotal + chunkLen + CHUNK_SIZE - 1)
            End If
            CopyMemory outBuf(outTotal), chunk(0), chunkLen
            outTotal = outTotal + chunkLen
        End If

        If ret = Z_STREAM_END Then Exit Do

    Loop While m_InflateStream.avail_out = 0

    If outTotal > 0 Then
        ReDim result(0 To outTotal - 1)
        CopyMemory result(0), outBuf(0), outTotal
    End If

    Inflate = result
End Function
