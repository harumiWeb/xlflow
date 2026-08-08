Attribute VB_Name = "BenchTests"
Option Explicit

#If Win64 Then
    Private Const size As Long = 2 ^ 24
#Else
    Private Const size As Long = 2 ^ 22 'To avoid 'out of memory' issues
#End If

'This module uses the excellent 'LibStringTools' module found at:
'https://github.com/guwidoe/VBA-StringTools
'to generate random keys

Sub TimeAddObjects1()
    Dim arr1(1 To size) As Variant, i As Long
    Dim arr2(1 To size) As Variant
    For i = 1 To size
        Set arr1(i) = New Class1
        Set arr2(i) = New Class1
    Next i
    Benchmark arr1, arr2, , "Object (Class1)"
End Sub
Sub TimeAddObjects2()
    Dim arr1(1 To size) As Variant, i As Long
    Dim arr2(1 To size) As Variant
    For i = 1 To size
        Set arr1(i) = New Collection
        Set arr2(i) = New Collection
    Next i
    Benchmark arr1, arr2, , "Object (Collection)"
End Sub
Sub TimeAddNumbers1()
    Dim arr1(1 To size) As Variant, i As Long
    Dim arr2(1 To size) As Variant
    For i = 1 To size
        arr1(i) = i
        arr2(i) = i + size
    Next i
    Benchmark arr1, arr2, , "Number (Long small)"
End Sub
Sub TimeAddNumbers2()
    Dim arr1(1 To size) As Variant, i As Long
    Dim arr2(1 To size) As Variant
    For i = 1 To size
        arr1(i) = &H7FFFFFFF - i
        arr2(i) = arr1(i) - size
    Next i
    Benchmark arr1, arr2, , "Number (Long large)"
End Sub
Sub TimeAddNumbers3()
    Dim arr1(1 To size) As Variant, i As Long
    Dim arr2(1 To size) As Variant
    For i = 1 To size
        arr1(i) = CDbl(i)
        arr2(i) = CDbl(i + size)
    Next i
    Benchmark arr1, arr2, , "Number (Double small ints)"
End Sub
Sub TimeAddNumbers4()
    Dim arr1(1 To size) As Variant, i As Long
    Dim arr2(1 To size) As Variant
    For i = 1 To size
        arr1(i) = CDbl(i + 99999999)
        arr2(i) = arr1(i) + CDbl(size)
    Next i
    Benchmark arr1, arr2, , "Number (Double large ints)"
End Sub
Sub TimeAddNumbers5()
    Dim arr1(1 To size) As Variant, i As Long
    Dim arr2(1 To size) As Variant
    For i = 1 To size
        arr1(i) = 9999999999# * LibStringTools.RndWH
        arr2(i) = 9999999999# * LibStringTools.RndWH
    Next i
    Benchmark arr1, arr2, , "Number (Double fractional)"
End Sub
Sub TimeAddText1()
    TimeAddText textLen:=5, compMode:=vbBinaryCompare, asciiOnly:=False
End Sub
Sub TimeAddText2()
    TimeAddText textLen:=5, compMode:=vbTextCompare, asciiOnly:=False
End Sub
Sub TimeAddText3()
    TimeAddText textLen:=10, stepLen:=2, compMode:=vbBinaryCompare, asciiOnly:=True
End Sub
Sub TimeAddText4()
    TimeAddText textLen:=10, stepLen:=2, compMode:=vbTextCompare, asciiOnly:=True
End Sub
Sub TimeAddText5()
    TimeAddText textLen:=20, stepLen:=3, compMode:=vbBinaryCompare, asciiOnly:=False
End Sub
Sub TimeAddText6()
    TimeAddText textLen:=20, stepLen:=3, compMode:=vbTextCompare, asciiOnly:=False
End Sub
Sub TimeAddText7()
    TimeAddText textLen:=50, stepLen:=10, compMode:=vbBinaryCompare, asciiOnly:=True
End Sub
Sub TimeAddText8()
    TimeAddText textLen:=50, stepLen:=10, compMode:=vbTextCompare, asciiOnly:=True
End Sub
Sub TimeAddMixed()
    Dim arr1(1 To size) As Variant
    Dim arr2(1 To size) As Variant
    Dim i As Long
    Dim d As New Dictionary: d.CompareMode = vbTextCompare 'So Collection does not throw
    '
    On Error Resume Next
    For i = 1 To size
        Select Case i Mod 3
        Case 0
            Set arr1(i) = New Collection
            Set arr2(i) = New Collection
        Case 1
            arr1(i) = 9999999999# * LibStringTools.RndWH
            arr2(i) = 9999999999# * LibStringTools.RndWH
        Case 2
            Do
                arr1(i) = LibStringTools.RandomString(5 + Rnd * 20, 1, &HFFFF&, True)
                d.Add arr1(i), i
                If Err.Number = 0 Then Exit Do
                Err.Clear
            Loop Until Err.Number = 0
            Do
                arr2(i) = LibStringTools.RandomString(5 + Rnd * 20, 1, &HFFFF&, True)
                d.Add arr2(i), i
                If Err.Number = 0 Then Exit Do
                Err.Clear
            Loop Until Err.Number = 0
        End Select
    Next i
    Set d = Nothing
    On Error GoTo 0
    Benchmark arr1, arr2, , "Mixed (binary compare)"
End Sub

Private Sub TimeAddText(ByVal textLen As Long _
                      , Optional ByVal stepLen As Long = 0 _
                      , Optional ByVal compMode As VbCompareMethod = vbBinaryCompare _
                      , Optional ByVal asciiOnly As Boolean = False)
    Dim arr1(1 To size) As Variant
    Dim arr2(1 To size) As Variant
    Dim i As Long
    Dim s1() As String
    Dim s2() As String
    Dim keyType As String
    Dim maxCodepoint As Long
    '
    If textLen < 4 Then textLen = 4
    stepLen = Abs(stepLen)
    If stepLen > textLen Then stepLen = textLen
    If asciiOnly Then
        maxCodepoint = 127
    Else
        maxCodepoint = &HFFFF& 'Could use &H10FFFF as Guido's library supports it but
                                'VBA.Collection sees surrogates as equal characters.
                                'This is also a problem with StrComp in textCompare or
                                'Option Compare Text e.g. ChrW(&HD883) = ChrW(&HD994)
    End If
    '
    'https://github.com/guwidoe/VBA-StringTools --- LibStringTools bas module
    s1 = LibStringTools.RandomStringArray(size, textLen + stepLen, textLen - stepLen, 1, maxCodepoint, True)
    s2 = LibStringTools.RandomStringArray(size, textLen + stepLen, textLen - stepLen, 1, maxCodepoint, True)
    '
    Dim d As New Dictionary: d.CompareMode = vbTextCompare 'So Collection does not throw
    '
    On Error Resume Next
    For i = 1 To size
        Do
            arr1(i) = s1(i - 1)
            d.Add arr1(i), i
            If Err.Number = 0 Then Exit Do
            Err.Clear
            arr1(i) = LibStringTools.RandomString(textLen, 1, maxCodepoint, True)
        Loop
        Do
            arr2(i) = s2(i - 1)
            d.Add arr2(i), i
            If Err.Number = 0 Then Exit Do
            Err.Clear
            arr2(i) = LibStringTools.RandomString(textLen, 1, maxCodepoint, True)
        Loop
    Next i
    On Error GoTo 0
    Set d = Nothing
    '
    keyType = "Text (len "
    If stepLen = 0 Then
        keyType = keyType & textLen
    Else
        keyType = keyType & textLen - stepLen & "-" & textLen + stepLen
    End If
    If compMode = vbBinaryCompare Then
        keyType = keyType & ", binary compare"
    Else
        keyType = keyType & ", text compare"
    End If
    If asciiOnly Then
        keyType = keyType & ", ASCII)"
    Else
        keyType = keyType & ", Unicode)"
    End If
    Benchmark arr1, arr2, compMode, keyType
End Sub
