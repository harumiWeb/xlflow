Attribute VB_Name = "TestLibMemory"
'''=============================================================================
''' VBA MemoryTools
''' -----------------------------------------------
''' https://github.com/cristianbuse/VBA-MemoryTools
''' -----------------------------------------------
''' MIT License
'''
''' Copyright (c) 2020 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit
Option Private Module

Public Sub RunAllTests()
    TestReadByte
    TestWriteByte
    TestReadInteger
    TestWriteInteger
    TestReadLong
    TestWriteLong
    TestReadLongLong
    TestWriteLongLong
    '
    TestReadBoolean
    TestWriteBoolean
    TestReadSingle
    TestWriteSingle
    TestReadCurrency
    TestWriteCurrency
    TestReadDate
    TestWriteDate
    TestReadDouble
    TestWriteDouble
    '
    TestMemCopy
    TestCloneParamArray
    TestStringToIntegers
    TestIntegersToString
    TestEmptyArray
    TestMemFill
    '
    Debug.Print "Finished running tests at " & Now()
End Sub

Private Sub TestReadByte()
    Dim b As Byte
    Dim i As Integer
    Dim l1 As Long, l2 As Long, l3 As Long, l4 As Long
    Dim l As Long
    Dim s As String
    '
    For l1 = 0 To 255
        b = l1
        Debug.Assert MemByte(VarPtr(b)) = b
    Next l1
    '
    For l1 = 0 To 255
        For l2 = 0 To 255
            i = ((l1 Xor &H8000) + l2 * 256) Xor &H8000
            Debug.Assert MemByte(VarPtr(i)) = l1
            Debug.Assert MemByte(VarPtr(i) + 1) = l2
        Next l2
    Next l1
    '
    For l1 = 0 To 255 Step 25
        For l2 = 0 To 255 Step 25
            For l3 = 0 To 255 Step 5
                For l4 = 0 To 255 Step 5
                    l = (((l1 + l2 * 256 + l3 * 256 ^ 2) Xor &H80000000) + l4 * 256 ^ 3) Xor &H80000000
                    Debug.Assert MemByte(VarPtr(l)) = l1
                    Debug.Assert MemByte(VarPtr(l) + 1) = l2
                    Debug.Assert MemByte(VarPtr(l) + 2) = l3
                    Debug.Assert MemByte(VarPtr(l) + 3) = l4
                Next l4
            Next l3
        Next l2
    Next l1
    '
    s = Chr$(66)
    Debug.Assert MemByte(StrPtr(s)) = 66
    Debug.Assert MemByte(StrPtr(s) + 1) = 0
    Debug.Assert MemByte(StrPtr(s) - 4) = 2 'Byte Count in BSTR
End Sub

Private Sub TestWriteByte()
    Dim b As Byte
    Dim i As Integer
    Dim l1 As Long, l2 As Long, l3 As Long, l4 As Long
    Dim l As Long
    Dim s As String
    '
    For l1 = 0 To 255
        MemByte(VarPtr(b)) = l1
        Debug.Assert b = l1
    Next l1
    '
    For l1 = 0 To 255
        For l2 = 0 To 255
            MemByte(VarPtr(i)) = l1
            MemByte(VarPtr(i) + 1) = l2
            Debug.Assert i = ((l1 Xor &H8000) + l2 * 256) Xor &H8000
        Next l2
    Next l1
    '
    For l1 = 0 To 255 Step 25
        For l2 = 0 To 255 Step 25
            For l3 = 0 To 255 Step 5
                For l4 = 0 To 255 Step 5
                    MemByte(VarPtr(l)) = l1
                    MemByte(VarPtr(l) + 1) = l2
                    MemByte(VarPtr(l) + 2) = l3
                    MemByte(VarPtr(l) + 3) = l4
                    Debug.Assert l = (((l1 + l2 * 256 + l3 * 256 ^ 2) Xor &H80000000) + l4 * 256 ^ 3) Xor &H80000000
                Next l4
            Next l3
        Next l2
    Next l1
    '
    s = Space(1)
    MemByte(StrPtr(s)) = 66
    Debug.Assert s = "B"
    MemByte(StrPtr(s) - 4) = 4
    Debug.Assert LenB(s) = 4
    MemByte(StrPtr(s) - 4) = 2
    Debug.Assert LenB(s) = 2
End Sub

Private Sub TestReadInteger()
    Dim i As Integer
    Dim l1 As Long, l2 As Long
    Dim l As Long
    Dim s As String
    '
    For l1 = &H8000 To &H7FFF
        i = l1
        Debug.Assert MemInt(VarPtr(i)) = i
    Next l1
    '
    For l1 = &H8000 To &H7FFF Step 128
        For l2 = &H8000 To &H7FFF Step 128
            l = l1 + IIf(l1 And &H8000, &H10000, 0) + l2 * &H10000
            Debug.Assert MemInt(VarPtr(l)) = l1
            Debug.Assert MemInt(VarPtr(l) + 2) = l2
        Next l2
    Next l1
    '
    s = Chr$(66) & Chr$(65)
    Debug.Assert MemInt(StrPtr(s)) = 66
    Debug.Assert MemInt(StrPtr(s) + 2) = 65
    Debug.Assert MemInt(StrPtr(s) - 4) = 4 'Byte Count in BSTR
End Sub

Private Sub TestWriteInteger()
    Dim i As Integer
    Dim l1 As Long, l2 As Long
    Dim l As Long
    Dim s As String
    '
    For l1 = &H8000 To &H7FFF
        MemInt(VarPtr(i)) = l1
        Debug.Assert i = l1
    Next l1
    '
    For l1 = &H8000 To &H7FFF Step 128
        For l2 = &H8000 To &H7FFF Step 128
            MemInt(VarPtr(l)) = l1
            MemInt(VarPtr(l) + 2) = l2
            Debug.Assert l = l1 + IIf(l1 And &H8000, &H10000, 0) + l2 * &H10000
        Next l2
    Next l1
    '
    s = Space(2)
    MemInt(StrPtr(s)) = 66
    Debug.Assert Mid$(s, 1, 1) = "B"
    MemInt(StrPtr(s) + 2) = 65
    Debug.Assert Mid$(s, 2, 1) = "A"
    MemInt(StrPtr(s) - 4) = 2
    Debug.Assert LenB(s) = 2
    MemInt(StrPtr(s) - 4) = 4
    Debug.Assert LenB(s) = 4
End Sub

Private Sub TestReadLong()
    Dim l As Long
    Dim s As String
    Dim c As Currency
    '
    For l = &H80000000 To &H7FFFFFFF - &H1000 Step &H1000
        Debug.Assert MemLong(VarPtr(l)) = l
        c = l / 10000
        Debug.Assert MemLong(VarPtr(c)) = l
        Debug.Assert MemLong(VarPtr(c) + 4) = IIf(l And &H80000000, -1, 0)
    Next l
    l = &H7FFFFFFF
    Debug.Assert MemLong(VarPtr(l)) = l
    c = l / 10000
    Debug.Assert MemLong(VarPtr(c)) = l
    '
    s = Chr$(65) & Chr$(66)
    Debug.Assert MemLong(StrPtr(s)) = 65 + 66 * 256 ^ 2
    Debug.Assert MemLong(StrPtr(s) + 2) = 66
    Debug.Assert MemLong(StrPtr(s) - 4) = 4 'Byte Count in BSTR
End Sub

Private Sub TestWriteLong()
    Dim l As Long, l1 As Long
    Dim s As String
    Dim c As Currency
    '
    For l1 = &H80000000 To &H7FFFFFFF - &H1000 Step &H1000
        MemLong(VarPtr(l)) = l1
        Debug.Assert l = l1
        '
        MemLong(VarPtr(c)) = l1
        MemLong(VarPtr(c) + 4) = IIf(l And &H80000000, -1, 0)
        Debug.Assert c = l1 / 10000
    Next l1
    l1 = &H7FFFFFFF
    MemLong(VarPtr(l)) = l1
    Debug.Assert l = l1
    MemLong(VarPtr(c)) = l1
    Debug.Assert c = l1 / 10000
    '
    s = Chr$(65) & Chr$(66)
    Debug.Assert MemLong(StrPtr(s)) = 65 + 66 * 256 ^ 2
    Debug.Assert MemLong(StrPtr(s) + 2) = 66
    Debug.Assert MemLong(StrPtr(s) - 4) = 4 'Byte Count in BSTR
    '
    s = Space(2)
    MemLong(StrPtr(s)) = 65 + 66 * 256 ^ 2
    Debug.Assert Mid$(s, 1, 2) = "AB"
    MemLong(StrPtr(s) - 4) = 2
    Debug.Assert LenB(s) = 2
    MemInt(StrPtr(s) - 4) = 4
    Debug.Assert LenB(s) = 4
End Sub

Private Sub TestReadLongLong()
#If Win64 Then
    Dim ll As LongLong
    Dim s As String
    Const loopStep As LongLong = &H1000000000000^
    '
    ll = &H8000000000000000^
    Do
        Debug.Assert MemLongLong(VarPtr(ll)) = ll
        ll = ll + loopStep
    Loop Until ll > &H7FFFFFFFFFFFFFFF^ - loopStep
    '
    s = Chr$(65) & Chr$(66) & Chr$(67) & Chr$(68)
    Debug.Assert MemLongLong(StrPtr(s)) = &H44004300420041^
    Debug.Assert MemLongLong(VarPtr(s)) = StrPtr(s)
#End If
End Sub

Private Sub TestWriteLongLong()
#If Win64 Then
    Dim ll As LongLong, ll2 As LongLong, ptr As LongLong
    Dim s As String, s2 As String
    Const loopStep As LongLong = &H1000000000000^
    '
    ll = &H8000000000000000^
    Do
        MemLongLong(VarPtr(ll2)) = ll
        Debug.Assert ll = ll2
        ll = ll + loopStep
    Loop Until ll > &H7FFFFFFFFFFFFFFF^ - loopStep
    '
    s = Space(4)
    MemLongLong(StrPtr(s)) = &H44004300420041^
    Debug.Assert Mid$(s, 1, 4) = "ABCD"
    '
    s2 = "TEST"
    ptr = StrPtr(s)
    MemLongLong(VarPtr(s)) = StrPtr(s2)
    Debug.Assert Mid$(s, 1, 4) = "TEST"
    MemLongLong(VarPtr(s)) = ptr
    Debug.Assert Mid$(s, 1, 4) = "ABCD"
#End If
End Sub

Private Sub TestReadBoolean()
    Dim b As Boolean
    Dim i As Integer
    '
    b = False
    Debug.Assert MemBool(VarPtr(b)) = b
    b = True
    Debug.Assert MemBool(VarPtr(b)) = b
    '
    i = 0
    Debug.Assert MemBool(VarPtr(i)) = False
    i = -1
    Debug.Assert MemBool(VarPtr(i)) = True
    i = 1
    Debug.Assert MemBool(VarPtr(i)) = 1
    Debug.Assert MemBool(VarPtr(i)) <> True
    Debug.Assert MemBool(VarPtr(i)) <> False
    '
    i = -255
    Debug.Assert MemBool(VarPtr(i)) = -255
    Debug.Assert MemBool(VarPtr(i)) <> True
    Debug.Assert MemBool(VarPtr(i)) <> False
End Sub

Private Sub TestWriteBoolean()
    Dim b As Boolean
    Dim i As Integer
    '
    MemBool(VarPtr(b)) = False
    Debug.Assert b = False
    MemBool(VarPtr(b)) = True
    Debug.Assert b = True
    MemBool(VarPtr(b)) = 0 'The 'newValue' parameter converts to Bool before memory is written
    Debug.Assert b = False
    MemBool(VarPtr(b)) = 5 'The 'newValue' parameter converts to Bool before memory is written
    Debug.Assert b = True
    MemBool(VarPtr(b)) = -5 'The 'newValue' parameter converts to Bool before memory is written
    Debug.Assert b = True
    '
    MemBool(VarPtr(i)) = False
    Debug.Assert i = 0
    MemBool(VarPtr(i)) = True
    Debug.Assert i = -1
    MemBool(VarPtr(i)) = 5 'The 'newValue' parameter converts to Bool before memory is written
    Debug.Assert i = -1
End Sub

Private Sub TestReadSingle()
    Dim s As Single
    Dim v As Variant
    Dim l As Long
    '
    Debug.Assert MemSng(VarPtr(&H7F800000)) = PosInf()
    Debug.Assert MemSng(VarPtr(&HFF800000)) = NegInf()
    Debug.Assert CStr(MemSng(VarPtr(&HFFC00000))) = CStr(SNAN())
    Debug.Assert CStr(MemSng(VarPtr(&H7FC00000))) = CStr(QNAN())
    '
    For Each v In Array(-3.402823E+38, -1.401298E-45, 0, 1.401298E-45, 3.402823E+38)
        s = v
        Debug.Assert MemSng(VarPtr(s)) = s
    Next v
    '
    For l = &H80000000 To &H7FFFFFFF - &H10000 Step &H10000
        If (l And &H7F800000) <> &H7F800000 Then 'Skip INF/NAN
            Debug.Assert MemSng(VarPtr(l)) = LongToSingle(l)
        End If
    Next l
End Sub
Public Function PosInf() As Double
    On Error Resume Next
    PosInf = 1 / 0
    On Error GoTo 0
End Function
Public Function NegInf() As Double
    NegInf = -PosInf
End Function
Public Function SNAN() As Double
    On Error Resume Next
    SNAN = 0 / 0
    On Error GoTo 0
End Function
Public Function QNAN() As Double
    QNAN = -SNAN
End Function
Private Function LongToSingle(ByVal l As Long) As Single
    Dim signBit As Long
    Dim exponentBits As Long
    Dim fractionBits As Single
    Dim i As Long
    '
    signBit = IIf(l And &H80000000, -1, 1)
    For i = 23 To 30
        exponentBits = exponentBits - CBool(l And 2 ^ i) * 2 ^ (i - 23)
    Next i
    For i = 1 To 23
        fractionBits = fractionBits - CBool(l And 2 ^ (23 - i)) * 2 ^ -i
    Next i
    If exponentBits = 0 Then
        If fractionBits <> 0 Then exponentBits = -126
    ElseIf exponentBits = 255 Then
        If fractionBits = 0 Then
            LongToSingle = PosInf()
        Else
            LongToSingle = SNAN()
        End If
        If signBit = -1 Then LongToSingle = -LongToSingle
        Exit Function
    Else
        Const bias As Long = 127
        exponentBits = exponentBits - bias
        fractionBits = fractionBits + 1
    End If
    LongToSingle = signBit * 2 ^ exponentBits * fractionBits
End Function

Private Sub TestWriteSingle()
    Dim s As Single, s2 As Single
    Dim v As Variant
    Dim l As Long
    '
    MemSng(VarPtr(l)) = PosInf()
    Debug.Assert l = &H7F800000
    '
    MemSng(VarPtr(l)) = NegInf()
    Debug.Assert l = &HFF800000
    '
    MemSng(VarPtr(l)) = SNAN()
    Debug.Assert l = &HFFC00000
    '
    MemSng(VarPtr(l)) = QNAN()
    Debug.Assert l = &H7FC00000
    '
    For Each v In Array(-3.402823E+38, -1.401298E-45, 0, 1.401298E-45, 3.402823E+38)
        MemSng(VarPtr(s)) = v
        Debug.Assert s = v
    Next v
    '
    For l = &H80000000 To &H7FFFFFFF - &H10000 Step &H10000
        If (l And &H7F800000) <> &H7F800000 Then 'Skip INF/NAN
            s = LongToSingle(l)
            MemSng(VarPtr(s2)) = s
            Debug.Assert s = s2
        End If
    Next l
End Sub

Private Sub TestReadCurrency()
    Dim c As Currency
    Dim l As Long
    Dim s As String
    '
    For l = 1 To 62
        c = -2 ^ l / 10000
        Debug.Assert MemCur(VarPtr(c)) = c
        c = (2 ^ l - 1) / 10000
        Debug.Assert MemCur(VarPtr(c)) = c
    Next l
    c = CCur("-922337203685477.5808")
    Debug.Assert MemCur(VarPtr(c)) = c
    c = CCur("922337203685477.5807")
    Debug.Assert MemCur(VarPtr(c)) = c
    '
    s = Chr$(65) & Chr$(66) & Chr$(67) & Chr$(68)
    Debug.Assert MemCur(StrPtr(s)) = CCur("1914058618345.8881")
End Sub

Private Sub TestWriteCurrency()
    Dim c As Currency, c2 As Currency
    Dim l As Long
    Dim s As String
    '
    For l = 1 To 62
        c = -2 ^ l / 10000
        MemCur(VarPtr(c2)) = c
        Debug.Assert c = c2
        c = (2 ^ l - 1) / 10000
        MemCur(VarPtr(c2)) = c
        Debug.Assert c = c2
    Next l
    c = CCur("-922337203685477.5808")
    MemCur(VarPtr(c2)) = c
    Debug.Assert c = c2
    c = CCur("922337203685477.5807")
    MemCur(VarPtr(c2)) = c
    Debug.Assert c = c2
    '
    s = Space(4)
    MemCur(StrPtr(s)) = CCur("1914058618345.8881")
    Debug.Assert Mid$(s, 1, 4) = "ABCD"
End Sub

Private Sub TestReadDate()
    Const minDate As Date = #1/1/100#
    Const maxDate As Date = #12/31/9999#
    '
    Dim dt As Date
    Dim i As Long, j As Long
    Dim d As Double
    Dim s As String
    '
    For i = minDate To maxDate Step 200
        d = CDbl(i) 'No time added
        dt = d
        Debug.Assert MemDate(VarPtr(d)) = dt
        For j = 1 To 100
            d = CDbl(i) + Rnd() 'Add some random time (hh:mm:ss)
            dt = d
            Debug.Assert MemDate(VarPtr(d)) = dt
        Next j
    Next i
    '
    d = CDbl(minDate) - 1000 'Invalid date
    dt = MemDate(VarPtr(d))
    Debug.Assert dt + 1000 = minDate
    '
    On Error Resume Next
    s = CStr(dt)
    Debug.Assert Err.Number = 5
    On Error GoTo 0
    '
    d = CDbl(maxDate) + 5000 'Invalid date
    dt = MemDate(VarPtr(d))
    Debug.Assert dt - 5000 = maxDate
    '
    On Error Resume Next
    s = CStr(dt)
    Debug.Assert Err.Number = 5
    On Error GoTo 0
End Sub

Private Sub TestWriteDate()
    Const minDate As Date = #1/1/100#
    Const maxDate As Date = #12/31/9999#
    '
    Dim dt As Date
    Dim i As Long, j As Long
    Dim d As Double
    Dim s As String
    '
    For i = minDate To maxDate Step 200
        d = CDbl(i) 'No time added
        MemDate(VarPtr(dt)) = d
        Debug.Assert dt = d
        For j = 1 To 100
            d = CDbl(i) + Rnd() 'Add some random time (hh:mm:ss)
            MemDate(VarPtr(dt)) = d
            Debug.Assert dt = d
        Next j
    Next i
    '
    d = CDbl(minDate) - 5000 'Invalid date
    MemDate(VarPtr(dt)) = MemDate(VarPtr(d))
    Debug.Assert dt + 5000 = minDate
    '
    On Error Resume Next
    s = CStr(dt)
    Debug.Assert Err.Number = 5 'Invalid date
    On Error GoTo 0
    '
    d = CDbl(maxDate) + 50000
    MemDate(VarPtr(dt)) = MemDate(VarPtr(d))
    Debug.Assert dt - 50000 = maxDate
    '
    On Error Resume Next
    s = CStr(dt)
    Debug.Assert Err.Number = 5
    On Error GoTo 0
End Sub

Private Sub TestReadDouble()
    Dim d As Double
    Dim v As Variant
    '
    For Each v In Array(-1.79769313486231E+308, -4.94065645841247E-324, 0 _
                       , 4.94065645841247E-324, 1.79769313486231E+308)
        d = v
        Debug.Assert MemDbl(VarPtr(d)) = d
    Next v
    '
#If Win64 Then
    Debug.Assert MemDbl(VarPtr(&H7FF0000000000000^)) = PosInf()
    Debug.Assert MemDbl(VarPtr(&HFFF0000000000000^)) = NegInf()
    Debug.Assert CStr(MemDbl(VarPtr(&HFFF8000000000000^))) = CStr(SNAN())
    Debug.Assert CStr(MemDbl(VarPtr(&H7FF8000000000000^))) = CStr(QNAN())
    '
    Dim ll As LongLong
    Const loopStep As LongLong = &H1000000000000^
    '
    ll = &H8000000000000000^
    Do
        If (ll And &H7FF0000000000000^) <> &H7FF0000000000000^ Then 'Skip INF/NAN
            Debug.Assert MemDbl(VarPtr(ll)) = LongLongToDouble(ll)
        End If
        ll = ll + loopStep
    Loop Until ll > &H7FFFFFFFFFFFFFFF^ - loopStep
#End If
End Sub

#If Win64 Then
Private Function LongLongToDouble(ByVal ll As LongLong) As Double
    Dim signBit As Long
    Dim exponentBits As Long
    Dim fractionBits As Double
    Dim i As Long
    '
    signBit = IIf(ll And &H8000000000000000^, -1, 1)
    For i = 52 To 62
        exponentBits = exponentBits - CBool(ll And 2 ^ i) * 2 ^ (i - 52)
    Next i
    For i = 1 To 52
        fractionBits = fractionBits - CBool(ll And 2 ^ (52 - i)) * 2 ^ -i
    Next i
    If exponentBits = 0 Then
        If fractionBits <> 0 Then exponentBits = -1022
    ElseIf exponentBits = 2047 Then
        If fractionBits = 0 Then
            LongLongToDouble = PosInf()
        Else
            LongLongToDouble = SNAN()
        End If
        If signBit = -1 Then LongLongToDouble = -LongLongToDouble
        Exit Function
    Else
        Const bias As Long = 1023
        exponentBits = exponentBits - bias
        fractionBits = fractionBits + 1
    End If
    LongLongToDouble = signBit * 2 ^ exponentBits * fractionBits
End Function
#End If

Private Sub TestWriteDouble()
    Dim d As Double, d2 As Double
    Dim v As Variant
    '
    For Each v In Array(-1.79769313486231E+308, -4.94065645841247E-324, 0 _
                       , 4.94065645841247E-324, 1.79769313486231E+308)
        MemDbl(VarPtr(d)) = v
        Debug.Assert d = v
    Next v
    '
#If Win64 Then
    Dim ll As LongLong
    '
    MemDbl(VarPtr(ll)) = PosInf()
    Debug.Assert ll = &H7FF0000000000000^
    '
    MemDbl(VarPtr(ll)) = NegInf()
    Debug.Assert ll = &HFFF0000000000000^
    '
    MemDbl(VarPtr(ll)) = SNAN()
    Debug.Assert ll = &HFFF8000000000000^
    '
    MemDbl(VarPtr(ll)) = QNAN()
    Debug.Assert ll = &H7FF8000000000000^
    '
    Const loopStep As LongLong = &H1000000000000^
    '
    ll = &H8000000000000000^
    Do
        If (ll And &H7FF0000000000000^) <> &H7FF0000000000000^ Then 'Skip INF/NAN
            d = LongLongToDouble(ll)
            MemDbl(VarPtr(d2)) = d
            Debug.Assert d = d2
        End If
        ll = ll + loopStep
    Loop Until ll > &H7FFFFFFFFFFFFFFF^ - loopStep
#End If
End Sub

Private Sub TestMemCopy()
    Dim arr1() As Byte
    Dim arr2() As Byte
    Dim i As Long, j As Long
    '
    ReDim arr1(0 To 2 ^ 24)
    arr2 = arr1
    '
    For i = LBound(arr2) To UBound(arr2)
        arr2(i) = i Mod 256
    Next i
    '
    For i = LBound(arr2) To 2 ^ 13
        MemCopy VarPtr(arr1(0)), VarPtr(arr2(0)), i + 1
        For j = 0 To i
            Debug.Assert arr1(j) = arr2(j)
            arr1(j) = 0  'Clear for next run
        Next j
    Next i
    '
    For i = 2 ^ 13 To 2 ^ 18 Step 2 ^ 6 - 1
        MemCopy VarPtr(arr1(0)), VarPtr(arr2(0)), i + 1
        For j = 0 To i Step 2 ^ 6 - 1
            Debug.Assert arr1(j) = arr2(j)
            arr1(j) = 0  'Clear for next run
        Next j
    Next i
    '
    For i = 2 ^ 18 To 2 ^ 24 Step 2 ^ 16 - 1
        MemCopy VarPtr(arr1(0)), VarPtr(arr2(0)), i + 1
        For j = 0 To i Step 2 ^ 10 - 1
            Debug.Assert arr1(j) = arr2(j)
            arr1(j) = 0  'Clear for next run
        Next j
    Next i
    '
#If Win64 Then
    ReDim arr1(0 To 2 ^ 31 - 2)
#Else
    ReDim arr1(0 To 2 ^ 27 - 1)
#End If
    arr2 = arr1
    '
    For i = LBound(arr2) To UBound(arr2) - 2 ^ 16 Step 2 ^ 16 - 1
        arr2(i) = i Mod 256
    Next i
    arr2(UBound(arr2) - 1) = 55
    '
    MemCopy VarPtr(arr1(0)), VarPtr(arr2(0)), UBound(arr2)
    For i = LBound(arr2) To UBound(arr2) - 2 ^ 16 Step 2 ^ 16 - 1
        Debug.Assert arr1(i) = arr2(i)
    Next i
    Debug.Assert arr1(UBound(arr1) - 1) = 55
    '
    'Test overlap
    ReDim arr1(0 To 100)
    For i = 0 To 100
        arr1(i) = i
    Next i
    MemCopy VarPtr(arr1(10)), VarPtr(arr1(0)), 50
    For i = 0 To 9
        Debug.Assert arr1(i) = i
    Next i
    For i = 10 To 59
        Debug.Assert arr1(i) = i - 10
    Next i
    For i = 60 To 100
        Debug.Assert arr1(i) = i
    Next i
    '
    For i = 0 To 100
        arr1(i) = i
    Next i
    MemCopy VarPtr(arr1(10)), VarPtr(arr1(13)), 50
    For i = 0 To 9
        Debug.Assert arr1(i) = i
    Next i
    For i = 10 To 59
        Debug.Assert arr1(i) = i + 3
    Next i
    For i = 60 To 100
        Debug.Assert arr1(i) = i
    Next i
    '
    For i = 0 To 100
        arr1(i) = i
    Next i
    MemCopy VarPtr(arr1(10)), VarPtr(arr1(7)), 50
    For i = 0 To 9
        Debug.Assert arr1(i) = i
    Next i
    For i = 10 To 59
        Debug.Assert arr1(i) = i - 3
    Next i
    For i = 60 To 100
        Debug.Assert arr1(i) = i
    Next i
    '
    For i = 0 To 100
        arr1(i) = i
    Next i
    MemCopy VarPtr(arr1(10)), VarPtr(arr1(1)), 50
    For i = 0 To 9
        Debug.Assert arr1(i) = i
    Next i
    For i = 10 To 59
        Debug.Assert arr1(i) = i - 9
    Next i
    For i = 60 To 100
        Debug.Assert arr1(i) = i
    Next i
    '
    For i = 0 To 100
        arr1(i) = i
    Next i
    MemCopy VarPtr(arr1(10)), VarPtr(arr1(15)), 50
    For i = 0 To 9
        Debug.Assert arr1(i) = i
    Next i
    For i = 10 To 59
        Debug.Assert arr1(i) = i + 5
    Next i
    For i = 60 To 100
        Debug.Assert arr1(i) = i
    Next i
    '
    For i = 0 To 100
        arr1(i) = i
    Next i
    MemCopy VarPtr(arr1(10)), VarPtr(arr1(27)), 50
    For i = 0 To 9
        Debug.Assert arr1(i) = i
    Next i
    For i = 10 To 59
        Debug.Assert arr1(i) = i + 17
    Next i
    For i = 60 To 100
        Debug.Assert arr1(i) = i
    Next i
End Sub

Private Sub TestCloneParamArray()
    Dim i1 As Long: i1 = 1
    Dim i2 As Long: i2 = 2
    Dim d As Double: d = 2.25
    Dim s1 As String: s1 = "ABC"
    Dim s2 As String: s2 = "DEF"
    Dim v1 As Variant: v1 = "ABC"
    Dim v2 As Variant: Set v2 = New Collection
    Dim v3 As Variant: v3 = Null
    Dim o1 As Object: Set o1 = New Collection
    Dim o2 As Collection: Set o2 = Nothing
    Dim arr() As Variant
    '
    TestParamArray i1, (i2), 1, d, 2.2, "ABC", s1, (s2), v1, v2, New Collection, v3, o1, o2, Null, Nothing, arr, Array(1, 2, 3)
    '
    Debug.Assert i1 = 2
    Debug.Assert i2 = 2
    Debug.Assert d = 3.14
    Debug.Assert s1 = "GHI"
    Debug.Assert s2 = "DEF"
    Debug.Assert v1 = 777
    Debug.Assert v2 Is Nothing
    Debug.Assert v3 Is Nothing
    Debug.Assert o1 Is Application
    Debug.Assert o2.Count = 1
    Debug.Assert UBound(arr) - LBound(arr) + 1 = 3
    Debug.Assert arr(UBound(arr)) = "ABC"
End Sub
Private Sub TestParamArray(ParamArray args() As Variant)
    Dim arr() As Variant
    CloneParamArray Not Not args, arr
    
    LetSet(arr(0)) = 2
    LetSet(arr(1)) = 3
    LetSet(arr(2)) = 4
    LetSet(arr(3)) = 3.14
    LetSet(arr(4)) = "2.2"
    LetSet(arr(5)) = 2.2
    LetSet(arr(6)) = "GHI"
    LetSet(arr(7)) = "ABC"
    LetSet(arr(8)) = 777
    LetSet(arr(9)) = Nothing
    LetSet(arr(10)) = Null
    LetSet(arr(11)) = Nothing
    LetSet(arr(12)) = Application
    LetSet(arr(13)) = New Collection: arr(13).Add Empty
    LetSet(arr(14)) = Empty
    LetSet(arr(15)) = Array(1, 2, 3)
    LetSet(arr(16)) = Array(1, 2, "ABC")
    LetSet(arr(17)) = Null
End Sub
Private Property Let LetSet(ByRef result As Variant, ByRef v As Variant)
    If IsObject(v) Then Set result = v Else result = v
End Property

Private Sub TestStringToIntegers()
    Dim arr() As Integer
    '
    arr = StringToIntegers("ABC")
    Debug.Assert arr(0) = AscW("A")
    Debug.Assert arr(1) = AscW("B")
    Debug.Assert arr(2) = AscW("C")
    '
    arr = StringToIntegers("ABC", outLowBound:=5)
    Debug.Assert arr(5) = AscW("A")
    Debug.Assert arr(6) = AscW("B")
    Debug.Assert arr(7) = AscW("C")
    '
    arr = StringToIntegers("ABC", 2, , 5)
    Debug.Assert arr(5) = AscW("B")
    Debug.Assert arr(6) = AscW("C")
    On Error Resume Next
    Dim i As Integer: i = arr(7)
    #If TWINBASIC Then
        Debug.Assert Err.Number <> 0
    #Else
        Debug.Assert Err.Number = 9
    #End If
    On Error GoTo 0
    '
    arr = StringToIntegers("ABC", 2, 7, 5)
    Debug.Assert arr(5) = AscW("B")
    Debug.Assert arr(6) = AscW("C")
    '
    arr = StringToIntegers("ABC", 2, 0, 5)
    Debug.Assert LBound(arr) = 5
    Debug.Assert UBound(arr) = 4
    '
    arr = StringToIntegers(vbNullString, 2, 0, 5)
    Debug.Assert LBound(arr) = 5
    Debug.Assert UBound(arr) = 4
    '
    arr = StringToIntegers(vbNullString)
    Debug.Assert UBound(arr) - LBound(arr) + 1 = 0
    '
    arr = StringToIntegers(StrConv("ABC", vbFromUnicode))
    Debug.Assert arr(0) = Asc("A") + Asc("B") * &H100
End Sub

Private Sub TestIntegersToString()
    Dim arr() As Integer
    Dim s As String
    '
    On Error Resume Next
    s = IntegersToString(arr)
    Debug.Assert Err.Number = 5
    On Error GoTo 0
    '
    ReDim arr(0 To 0): arr(0) = Asc("A")
    Debug.Assert IntegersToString(arr) = "A"
    '
    ReDim Preserve arr(0 To 1)
    Debug.Assert IntegersToString(arr) = "A" & vbNullChar
    Debug.Assert IntegersToString(arr, 5) = vbNullString
    Debug.Assert IntegersToString(arr, 1) = vbNullChar
    Debug.Assert IntegersToString(arr, 1, -1) = vbNullChar
    Debug.Assert IntegersToString(arr, 1, 2) = vbNullChar
    Debug.Assert IntegersToString(arr, 1, 0) = vbNullString
End Sub

Private Sub TestEmptyArray()
    Dim arr As Variant
    Dim v As Variant
    Dim i As Long
    '
    For Each v In Array(vbByte, vbInteger, vbLong, vbLongLong, vbCurrency, vbDecimal, vbDouble, vbSingle, vbDate, vbBoolean, vbString, vbObject, vbDataObject, vbVariant)
        For i = 1 To 60
            arr = EmptyArray(i, v)
            Debug.Assert VarType(arr) = vbArray + v
            Debug.Assert GetArrayDimsCount(arr) = i
        Next i
    Next v
    '
    On Error Resume Next
    arr = EmptyArray(61, vbBoolean)
    Debug.Assert Err.Number = 5
    On Error GoTo 0
    '
    On Error Resume Next
    arr = EmptyArray(2, 500)
    Debug.Assert Err.Number = 13
    On Error GoTo 0
End Sub
Private Function GetArrayDimsCount(ByRef arr As Variant) As Long
    Const MAX_DIMENSION As Long = 60 'VB limit
    Dim dimension As Long
    Dim tempBound As Long
    '
    On Error GoTo FinalDimension
    For dimension = 1 To MAX_DIMENSION
        tempBound = LBound(arr, dimension)
    Next dimension
FinalDimension:
    GetArrayDimsCount = dimension - 1
End Function

Private Sub TestMemFill()
    Dim arr() As Byte
    Dim i As Long, j As Long
    Const maxTestSize As Long = 2 ^ 24 ' Can increase but will take longer to run
    Dim size As Long
    Dim stepSize As Long
    Dim b As Byte
    '
    ReDim arr(1 To maxTestSize)
    size = 1
    '
    For i = 1 To 2 ^ 10
        b = i Mod 256
        MemFill VarPtr(arr(1)), i, b
        For j = 1 To i
            Debug.Assert arr(j) = b
        Next j
        For j = i + 1 To i * 2
            Debug.Assert arr(j) = 0
        Next j
    Next i
    '
    stepSize = 2 ^ 12 + 1
    For i = 2 ^ 14 To 2 ^ 18 Step stepSize
        b = i Mod 256
        MemFill VarPtr(arr(1)), i, b
        For j = 1 To i
            Debug.Assert arr(j) = b
        Next j
        For j = i + 1 To i * 2
            Debug.Assert arr(j) = 0
        Next j
    Next i
    '
    b = 5
    MemFill VarPtr(arr(1)), maxTestSize, b
    For i = 1 To maxTestSize
        Debug.Assert arr(i) = b
    Next i
    '
    MemFill VarPtr(arr(1000)), 1000, 7
    For i = 1 To 999
        Debug.Assert arr(i) = b
    Next i
    For i = 1000 To 1999
        Debug.Assert arr(i) = 7
    Next i
    For i = 2000 To maxTestSize
        Debug.Assert arr(i) = b
    Next i
End Sub
