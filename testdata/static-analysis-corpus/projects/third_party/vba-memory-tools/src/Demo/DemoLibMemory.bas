Attribute VB_Name = "DemoLibMemory"
Option Explicit
Option Private Module

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
        Private Declare PtrSafe Function FillMemory Lib "/usr/lib/libc.dylib" Alias "memset" (Destination As Any, ByVal Fill As Byte, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
        Private Declare Function FillMemory Lib "/usr/lib/libc.dylib" Alias "memset" (Destination As Any, ByVal Fill As Byte, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
        Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
    #End If
#End If

Sub DemoMain()
    Dim s As String: s = String$(13, "-")
    Debug.Print String(26, "-") & " Speed (seconds)" & String(27, "-")
    Debug.Print AlignCenter("Data Type"), AlignCenter("Iterations"), AlignCenter("By Ref") _
              , AlignCenter("CopyMemory"), AlignCenter("Notes")
    Debug.Print AlignCenter("(To copy)"), AlignCenter("(Count)") _
              , AlignCenter("(This Lib)"), AlignCenter("(DLL Export)")
    Debug.Print s, s, s, s, s
    DemoMemByteSpeed
    DemoMemIntSpeed
    DemoMemLongSpeed
    DemoMemLongPtrSpeed
    DemoMemObjectSpeed
    Debug.Print String(21, "-") & " Redirection " & String(21, "-")
    DemoInstanceRedirection
    Debug.Print String(37, "-") & " MemCopy " & String(37, "-")
    DemoMemCopySpeed
    Debug.Print String(30, "-") & " MemFill " & String(30, "-")
    DemoMemFillSpeed
End Sub

Private Function AlignRight(ByRef s As String, Optional ByVal size As Long = 13) As String
    AlignRight = Right$(Space$(size) & s, size)
End Function
Private Function AlignCenter(ByRef s As String, Optional ByVal size As Long = 13) As String
    Dim i As Long: i = size - Len(s)
    If i < 1 Then
        AlignCenter = s
    Else
        AlignCenter = Space$(i \ 2) & s & Space$(i / 2)
    End If
End Function

Private Sub DemoInstanceRedirection()
    Const loopsCount As Long = 100000
    Dim i As Long
    Dim t As Double
    '
    Debug.Print Format$(loopsCount, "#,##0") & " times"
    t = Timer
    For i = 1 To loopsCount
        Debug.Assert DemoClass.Factory2(i).ID = i
    Next i
    Debug.Print "Public  Init (seconds): " & VBA.Round(Timer - t, 3)
    '
    t = Timer
    For i = 1 To loopsCount
        Debug.Assert DemoClass.Factory(i).ID = i
    Next i
    Debug.Print "Private Init (seconds): " & VBA.Round(Timer - t, 3)
End Sub

Private Sub DemoMemByteSpeed()
    Dim x1 As Byte: x1 = 1
    Dim x2 As Byte: x2 = 2
    Dim i As Long
    Dim t As Double
    Dim slowFactor As Long
    Dim res1 As Double
    Dim res2 As Double
    Const iterations As Long = 1000000
    Dim sp As String: sp = Space$(13)
    '
    t = Timer
    For i = 1 To iterations
        MemByte(VarPtr(x1)) = MemByte(VarPtr(x2))
    Next i
    res1 = Round(Timer - t, 3)
    '
    slowFactor = 10000 'In case API call is too slow
    Do
        t = Timer
        For i = 1 To iterations \ slowFactor
            CopyMemory x1, x2, 1
        Next i
        res2 = Round(Timer - t, 3)
        If res2 < 0.1 Then
            slowFactor = slowFactor \ 10
        Else
            Exit Do
        End If
    Loop Until slowFactor = 0
    If slowFactor = 0 Then slowFactor = 1 'For IIf (Div by Zero)
    '
    Debug.Print TypeName(x1), AlignRight(Format$(iterations, "#,##0")) _
              , AlignRight(Format$(res1, "#,##0.000")) _
              , AlignRight(Format$(res2 * slowFactor, "#,##0.000")) _
              , IIf(slowFactor > 1, "(extrapolated from " _
              & Format$(iterations \ slowFactor, "#,##0") _
              & " iterations that took " & res2 & " seconds)", "")
    DoEvents
End Sub

Private Sub DemoMemIntSpeed()
    Dim x1 As Integer: x1 = 11111
    Dim x2 As Integer: x2 = 22222
    Dim i As Long
    Dim t As Double
    Dim slowFactor As Long
    Dim res1 As Double
    Dim res2 As Double
    Const iterations As Long = 1000000
    '
    t = Timer
    For i = 1 To iterations
        MemInt(VarPtr(x1)) = MemInt(VarPtr(x2))
    Next i
    res1 = Round(Timer - t, 3)
    '
    slowFactor = 10000 'In case API call is too slow
    Do
        t = Timer
        For i = 1 To iterations \ slowFactor
            CopyMemory x1, x2, 2
        Next i
        res2 = Round(Timer - t, 3)
        If res2 < 0.1 Then
            slowFactor = slowFactor \ 10
        Else
            Exit Do
        End If
    Loop Until slowFactor = 0
    If slowFactor = 0 Then slowFactor = 1 'For IIf (Div by Zero)
    '
    Debug.Print TypeName(x1), AlignRight(Format$(iterations, "#,##0")) _
              , AlignRight(Format$(res1, "#,##0.000")) _
              , AlignRight(Format$(res2 * slowFactor, "#,##0.000")) _
              , IIf(slowFactor > 1, "(extrapolated from " _
              & Format$(iterations \ slowFactor, "#,##0") _
              & " iterations that took " & res2 & " seconds)", "")
    DoEvents
End Sub

Private Sub DemoMemLongSpeed()
    Dim x1 As Long: x1 = 111111111
    Dim x2 As Long: x2 = 222222222
    Dim i As Long
    Dim t As Double
    Dim slowFactor As Long
    Dim res1 As Double
    Dim res2 As Double
    Const iterations As Long = 1000000
    '
    t = Timer
    For i = 1 To iterations
        MemLong(VarPtr(x1)) = MemLong(VarPtr(x2))
    Next i
    res1 = Round(Timer - t, 3)
    '
    slowFactor = 10000 'In case API call is too slow
    Do
        t = Timer
        For i = 1 To iterations \ slowFactor
            CopyMemory x1, x2, 4
        Next i
        res2 = Round(Timer - t, 3)
        If res2 < 0.1 Then
            slowFactor = slowFactor \ 10
        Else
            Exit Do
        End If
    Loop Until slowFactor = 0
    If slowFactor = 0 Then slowFactor = 1 'For IIf (Div by Zero)
    '
    Debug.Print TypeName(x1), AlignRight(Format$(iterations, "#,##0")) _
              , AlignRight(Format$(res1, "#,##0.000")) _
              , AlignRight(Format$(res2 * slowFactor, "#,##0.000")) _
              , IIf(slowFactor > 1, "(extrapolated from " _
              & Format$(iterations \ slowFactor, "#,##0") _
              & " iterations that took " & res2 & " seconds)", "")
    DoEvents
End Sub

Private Sub DemoMemLongPtrSpeed()
    #If Win64 Then
        Dim x1 As LongLong: x1 = 111111111111111^
        Dim x2 As LongLong: x2 = 111111111111112^
    #Else
        Dim x1 As Long: x1 = 111111111
        Dim x2 As Long: x2 = 222222222
    #End If
    Dim i As Long
    Dim t As Double
    Dim slowFactor As Long
    Dim res1 As Double
    Dim res2 As Double
    Const iterations As Long = 1000000
    '
    t = Timer
    For i = 1 To iterations
        MemLongPtr(VarPtr(x1)) = MemLongPtr(VarPtr(x2))
    Next i
    res1 = Round(Timer - t, 3)
    '
    slowFactor = 10000 'In case API call is too slow
    Do
        t = Timer
        For i = 1 To iterations \ slowFactor
            CopyMemory x1, x2, PTR_SIZE
        Next i
        res2 = Round(Timer - t, 3)
        If res2 < 0.1 Then
            slowFactor = slowFactor \ 10
        Else
            Exit Do
        End If
    Loop Until slowFactor = 0
    If slowFactor = 0 Then slowFactor = 1 'For IIf (Div by Zero)
    '
    Debug.Print TypeName(x1), AlignRight(Format$(iterations, "#,##0")) _
              , AlignRight(Format$(res1, "#,##0.000")) _
              , AlignRight(Format$(res2 * slowFactor, "#,##0.000")) _
              , IIf(slowFactor > 1, "(extrapolated from " _
              & Format$(iterations \ slowFactor, "#,##0") _
              & " iterations that took " & res2 & " seconds)", "")
    DoEvents
End Sub

Private Sub DemoMemObjectSpeed()
    Dim i As Long
    Dim t As Double
    Dim d As DemoClass: Set d = New DemoClass
    Dim obj As Object
    #If Win64 Then
        Dim ptr As LongLong
    #Else
        Dim ptr As Long
    #End If
    Const iterations As Long = 1000000
    '
    ptr = ObjPtr(d)
    t = Timer
    For i = 1 To iterations
        Set obj = MemObj(ptr)
    Next i
    Debug.Print
    Debug.Print "Dereferenced an Object " & Format$(iterations, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemCopySpeed()
    Dim t As Double
    Dim a1() As Byte
    Dim a2() As Byte
    Dim size As Long
    Dim iterations As Long
    Dim i As Long
    Dim src As LongPtr
    Dim dest As LongPtr
    Dim res1 As Double
    Dim res2 As Double
    Dim slowFactor As Long
    Dim s As String: s = String$(13, "-")
    '
    size = 2
    iterations = 2 ^ 21
    Debug.Print AlignCenter("Size"), AlignCenter("Iterations"), AlignCenter("MemCopy") _
              , AlignCenter("CopyMemory"), AlignCenter("Notes")
    Debug.Print AlignCenter("(Bytes)"), AlignCenter("(Count)"), AlignCenter("(ACCESSOR)") _
              , AlignCenter("(DLL export)")
    Debug.Print s, s, s, s, s
    Do
        ReDim a1(0 To size - 1)
        ReDim a2(0 To size - 1)
        '
        src = VarPtr(a2(0))
        dest = VarPtr(a1(0))
        '
        t = Timer
        For i = 1 To iterations
            MemCopy dest, src, size
        Next i
        res1 = Round(Timer - t, 3)
        '
        slowFactor = 10000 'In case API call is too slow
        Do
            t = Timer
            For i = 1 To iterations \ slowFactor
                CopyMemory ByVal dest, ByVal src, size
            Next i
            res2 = Round(Timer - t, 3)
            If res2 < 0.1 Then
                slowFactor = slowFactor \ 10
            Else
                Exit Do
            End If
        Loop Until slowFactor = 0
        If slowFactor = 0 Then slowFactor = 1 'For IIf (Div by Zero)
        '
        Debug.Print AlignRight(Format$(size, "#,##0")) _
                  , AlignRight(Format$(iterations, "#,##0")) _
                  , AlignRight(Format$(res1, "#,##0.000")) _
                  , AlignRight(Format$(res2 * slowFactor, "#,##0.000")) _
                  , IIf(slowFactor > 1, "(extrapolated from " _
                  & Format$(iterations \ slowFactor, "#,##0") _
                  & " iterations that took " & res2 & " seconds)", "")
        '
        Const maxLong As Long = 2147483647
        If CDbl(size) * 2 > maxLong Then
            If CDbl(size) * 2 - 1 > maxLong Then
                iterations = 2
            Else
                size = CDbl(size) * 2 - 1
            End If
        Else
            size = size * 2
        End If
        iterations = iterations / 1.6
        DoEvents
    Loop Until iterations = 1
End Sub

Private Sub DemoMemFillSpeed()
    Dim t As Double
    Dim a() As Byte
    Dim size As Long
    Dim iterations As Long
    Dim i As Long
    Dim dest As LongPtr
    Dim res1 As Double
    Dim res2 As Double
    Dim slowFactor As Long
    Dim s As String: s = String$(13, "-")
    Const b As Byte = 255
    '
    size = 2
    iterations = 2 ^ 21
    Debug.Print AlignCenter("Size"), AlignCenter("Iterations"), AlignCenter("MemFill") _
              , AlignCenter("FillMemory"), AlignCenter("Notes")
    Debug.Print AlignCenter("(Bytes)"), AlignCenter("(Count)") _
              , AlignCenter("MidB-MemCopy"), AlignCenter("(DLL export)")
    Debug.Print s, s, s, s, s
    Do
        ReDim a(0 To size - 1)
        '
        dest = VarPtr(a(0))
        '
        t = Timer
        For i = 1 To iterations
            MemFill dest, size, b
        Next i
        res1 = Round(Timer - t, 3)
        '
        slowFactor = 10000 'In case API call is too slow
        Do
            t = Timer
            For i = 1 To iterations \ slowFactor
                #If Mac Then
                    FillMemory ByVal dest, b, size
                #Else
                    FillMemory ByVal dest, size, b
                #End If
            Next i
            res2 = Round(Timer - t, 3)
            If res2 < 0.1 Then
                slowFactor = slowFactor \ 10
            Else
                Exit Do
            End If
        Loop Until slowFactor = 0
        If slowFactor = 0 Then slowFactor = 1 'For IIf (Div by Zero)
        '
        Debug.Print AlignRight(Format$(size, "#,##0")) _
                  , AlignRight(Format$(iterations, "#,##0")) _
                  , AlignRight(Format$(res1, "#,##0.000")) _
                  , AlignRight(Format$(res2 * slowFactor, "#,##0.000")) _
                  , IIf(slowFactor > 1, "(extrapolated from " _
                  & Format$(iterations \ slowFactor, "#,##0") _
                  & " iterations that took " & res2 & " seconds)", "")
        '
        Const maxLong As Long = 2147483647
        If CDbl(size) * 2 > maxLong Then
            If CDbl(size) * 2 - 1 > maxLong Then
                iterations = 2
            Else
                size = CDbl(size) * 2 - 1
            End If
        Else
            size = size * 2
        End If
        iterations = iterations / 1.6
        DoEvents
    Loop Until iterations = 1
End Sub

