Attribute VB_Name = "TestDictionary"
Option Explicit
Option Private Module

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private Type JustMyUDT
    dict As Dictionary
End Type

#Const x64StackIssues = ((Mac = 0) And Win64 And (TWINBASIC = 0))
#Const JapanYen = "\" > "}"

Private Const invalidCallErr As Long = 5
Private Const unsupportedKeyErr As Long = invalidCallErr
Private Const keyOrIndexNotFoundErr As Long = 9
Private Const setMissingErr As Long = 450
Private Const duplicatedKeyErr As Long = 457
Private Const failedScriptingErr As Long = 32811

Public Sub RunAllDictionaryTests()
    TestEmptyDictionary
    TestDictionaryAdd
    TestDictionaryAllowDuplicateKeys
    TestDictionaryCompare
    TestDictionaryCount
    TestDictionaryDealloc
    TestDictionaryExists
    TestDictionaryFactory
    TestDictionaryHashVal
    TestDictionaryIndex
    TestDictionaryItem
    TestDictionaryItemAtIndex
    TestDictionaryItems
    TestDictionaryKey
    TestDictionaryKeyAtIndex
    TestDictionaryKeys
    TestDictionaryKeysItems2D
    TestDictionaryLoadFactor
    TestDictionaryNewEnum
    TestDictionaryRemove
    TestDictionaryRemoveAll
    TestDictionarySelf
    TestDictionaryStackFixes
    TestDictionaryTryGetItem
    Debug.Print "Finished running dictionary tests at " & Now()
End Sub

Private Sub TestEmptyDictionary()
    Dim d As New Dictionary
    '
    Debug.Assert d.Count = 0
    Debug.Assert d.CompareMode = vbBinaryCompare
    Debug.Assert Not d.Factory Is d
    Debug.Assert TypeOf d.Factory Is Dictionary
    Debug.Assert d.LoadFactor = 0
    Debug.Assert d.Self Is d
End Sub

Private Sub TestDictionaryAdd()
    Dim d As New Dictionary
    Dim v As Variant
    Dim ptr As LongPtr: ptr = 10
    '
    For Each v In Array(Empty, Null, CInt(1), CLng(2), CSng(3), CDbl(4), CCur(5) _
                      , CDate(6), CStr(7), Nothing, New Collection, CVErr(2042) _
                      , True, False, GetDefaultInterface(d.Factory) _
                      , CByte(9), ptr, vbNullString)
        d.Add v, v
    Next v
    #If Mac = 0 Then
        v = CDec(8)
        d.Add v, v
    #End If
    '
    On Error Resume Next
    For Each v In Array(CLng(3), Array(), 0, -1, Array(1, 2, 3), Empty)
        d.Add v, v
        If IsArray(v) Then
            Debug.Assert Err.Number = unsupportedKeyErr
        Else
            Debug.Assert Err.Number = duplicatedKeyErr
        End If
        Err.Clear
    Next v
    On Error GoTo 0
    '
    For Each v In Array(CLng(3), Array(), 0, -1, Array(1, 2, 3), Empty)
        Debug.Assert Not d.Add(v, v, IgnoreErrors:=True)
    Next v
    '
    Debug.Assert d.Add(11, Array())
    Debug.Assert d.Add(12, New Dictionary)
    Debug.Assert d.Add("Test Add", Nothing)
    Debug.Assert d.Add("test add", New Collection)
    Debug.Assert d.Add("test add" & vbNewLine, 1)
    Debug.Assert d.Add("test add" & vbNullChar, 1)
End Sub
Private Function GetDefaultInterface(ByVal obj As stdole.IUnknown) As Object
    Set GetDefaultInterface = obj
End Function

Private Sub TestDictionaryAllowDuplicateKeys()
    Dim d As New Dictionary
    d.AllowDuplicateKeys = True
    '
    d.Add 1, 2
    d.Add 1, 3
    '
    Debug.Assert d(1) = 2
    Debug.Assert d.Count = 2
    Debug.Assert d.Items()(0) = 2
    Debug.Assert d.Items()(1) = 3
    '
    d.Remove 1
    Debug.Assert d(1) = 3
    '
    d.Add 1, 4
    Debug.Assert d(1) = 3
    Debug.Assert d.Count = 2
    Debug.Assert d.Items()(0) = 3
    Debug.Assert d.Items()(1) = 4
    '
    d.Key(1) = 2
    Debug.Assert d(1) = 4
    Debug.Assert d(2) = 3
    '
    On Error Resume Next
    d.AllowDuplicateKeys = False
    Debug.Assert Err.Number = invalidCallErr
    On Error GoTo 0
    '
    d.RemoveAll
    d.AllowDuplicateKeys = False
End Sub

Private Sub TestDictionaryCompare()
    Dim d As New Dictionary
    d.CompareMode = vbBinaryCompare
    '
    d.Add "Test Add", Nothing
    d.Add "test add", New Collection
    d.Add "Aa", 1
    d.Add "aa", 2
    d.Add "AA", 3
    d.Add "aA", 4
    d.Add "AA" & vbNullChar, 3
    '
    d.CompareMode = vbBinaryCompare
    d.RemoveAll
    d.CompareMode = vbTextCompare
    '
    d.Add "Test Add", Nothing
    On Error Resume Next
    d.Add "test add", New Collection
    Debug.Assert Err.Number = duplicatedKeyErr
    On Error GoTo 0
    '
    'StrComp sees the following 2 Unicode surrogates as equal but they are not
    '   and the dictionary should be able to handle them in text compare mode
    d.Add ChrW$(CInt(&HD883)), 1
    d.Add ChrW$(CInt(&HD994)), 2
    '
    Const lcidRomanian As Long = 1048
    Const lcidCroatian As Long = 1050
    Dim uDZ As String: uDZ = ChrW$(497)
    Dim ldz As String: ldz = ChrW$(499)
    '
    d.RemoveAll
    d.CompareMode = lcidRomanian
    d.Add uDZ, 1
    Debug.Assert d.Exists(ldz)
    '
    d.RemoveAll
    d.CompareMode = lcidCroatian
    d.Add uDZ, 1
#If Mac = 0 Then
    Debug.Assert Not d.Exists(ldz)
#End If
End Sub

Private Sub TestDictionaryCount()
    Dim d As New Dictionary
    Dim i As Long
    '
    Debug.Assert d.Count = 0
    '
    d.Add 1, 1
    Debug.Assert d.Count = 1
    '
    On Error Resume Next
    d.Add 1, 2
    Debug.Assert Err.Number = duplicatedKeyErr
    On Error GoTo 0
    Debug.Assert d.Count = 1
    '
    For i = 1 To 50
        d.Add CStr(i), Array(i)
    Next i
    Debug.Assert d.Count = 51
    '
    d.RemoveAll
    Debug.Assert d.Count = 0
End Sub

Private Sub TestDictionaryDealloc()
    Dim i As Long
    Dim d() As Dictionary
    Const n As Long = 30000
    ReDim d(1 To n) As Dictionary
    '
    For i = 1 To n
#If JapanYen Then
        If i = Fix(n / 2) Then Set Dictionary = d(Fix(n / 4))
#Else
        If i = n \ 2 Then Set Dictionary = d(n \ 4)
#End If
        Set d(i) = New Dictionary
        If i Mod 7 = 0 Then Set d(i) = New Dictionary
    Next i
    For i = 1 To n / 2
        Set d(Rnd() * i + 1) = Nothing
    Next i
End Sub

Private Sub TestDictionaryExists()
    Dim d As New Dictionary
    Dim c As New Collection
    '
    d.Add "key1", 1
    d.Add "key3", Nothing
    d.Add c, Empty
    d.Add 1, 3
    d.Add Empty, Null
    d.Add Null, Null
    d.Add 0, 0
    d.Add #3/31/2024#, "KB"
    d.Add PosInf, 1
    d.Add NegInf, 2
    d.Add SNaN, 3
    d.Add QNaN, 4
    '
    Debug.Assert d.Exists("key1")
    Debug.Assert Not d.Exists("key2")
    Debug.Assert d.Exists("key3")
    Debug.Assert Not d.Exists("test")
    Debug.Assert Not d.Exists(437547)
    Debug.Assert d.Exists(Empty)
    Debug.Assert Not d.Exists("")
    Debug.Assert d.Exists(0)
    Debug.Assert d.Exists(Null)
    Debug.Assert d.Exists(c)
    Debug.Assert Not d.Exists(New Collection)
    Debug.Assert Not d.Exists(Nothing)
    Debug.Assert d.Exists(CDbl(1))
    Debug.Assert d.Exists(CDbl(#3/31/2024#))
    Debug.Assert Not d.Exists(CDbl(#3/30/2024#))
    Debug.Assert d.Exists(PosInf)
    Debug.Assert d.Exists(NegInf)
    Debug.Assert d.Exists(SNaN)
    Debug.Assert d.Exists(QNaN)
    '
    On Error Resume Next
    d.Exists Array()
    Debug.Assert Err.Number = unsupportedKeyErr
    On Error GoTo 0
    '
    Debug.Assert Not d.Exists(Array(), True)
End Sub

'@Description("IEEE754 +inf")
Private Property Get PosInf() As Double
    On Error Resume Next
    PosInf = 1 / 0
    On Error GoTo 0
End Property
'@Description("IEEE754 signaling NaN (sNaN)")
Private Property Get SNaN() As Double
    On Error Resume Next
    SNaN = 0 / 0
    On Error GoTo 0
End Property
'@Description("IEEE754 -inf")
Private Property Get NegInf() As Double
    NegInf = -PosInf
End Property
'@Description("IEEE754 quiet NaN (qNaN)")
Private Property Get QNaN() As Double
    QNaN = -SNaN
End Property

Private Sub TestDictionaryFactory()
    With New Dictionary
        Debug.Assert Not .Factory Is .Self
        Debug.Assert TypeOf .Factory Is Dictionary
        Debug.Assert Not .Factory Is Dictionary
        Debug.Assert Not .Factory Is Dictionary.Factory
        Debug.Assert TypeOf .Factory.Factory.Factory Is Dictionary
    End With
    Debug.Assert Not Dictionary.Factory Is Dictionary.Self
    Debug.Assert TypeOf Dictionary.Factory Is Dictionary
End Sub

Private Function ArrayToCSV(arr As Variant _
    , Optional ByVal delimiter As String = "," _
) As String
    Dim s As String
    Dim v As Variant
    Dim tColl As Collection
    '
    s = "["
    For Each v In arr
        If IsObject(v) Then
            s = s & TypeName(v)
        Else
            Select Case VarType(v)
            Case vbNull
                s = s & "Null"
            Case vbEmpty
                s = s & "Empty"
            Case vbString
                s = s & """" & v & """"
            Case vbError, vbBoolean
                s = s & CStr(v)
            Case vbByte, vbInteger, vbLong, 20 _
               , vbCurrency, vbDecimal, vbDouble, vbSingle
                s = s & v
            Case vbDate
                s = s & CDbl(v)
            Case vbArray To vbArray + vbUserDefinedType
                s = s & ArrayToCSV(v, delimiter)
            Case vbUserDefinedType
                Err.Raise 5, , "User defined types not supported"
            Case vbDataObject
                 s = s & TypeName(v)
            End Select
        End If
        s = s & delimiter
    Next v
    If Len(s) > 1 Then s = Left$(s, Len(s) - Len(delimiter))
    ArrayToCSV = s & "]"
End Function

Private Sub TestDictionaryHashVal()
    Dim d As New Dictionary
    Dim c As New Collection
    Dim i As Long
    '
    Debug.Assert d.HashVal(CDbl(1)) = d.HashVal(CByte(1))
    Debug.Assert d.HashVal(CDbl(-1)) = d.HashVal(True)
    Debug.Assert d.HashVal(CDbl(#3/30/2024#)) = d.HashVal(#3/30/2024#)
    Debug.Assert d.HashVal(Empty) <> d.HashVal(Null)
    Debug.Assert d.HashVal(Empty) <> d.HashVal(Nothing)
    Debug.Assert d.HashVal(Empty) <> d.HashVal(0)
    Debug.Assert d.HashVal(Empty) <> d.HashVal("")
    Debug.Assert d.HashVal(Null) <> d.HashVal(0)
    Debug.Assert d.HashVal(Null) <> d.HashVal("")
    Debug.Assert d.HashVal(0) <> d.HashVal("")
    Debug.Assert d.HashVal(c) = d.HashVal(GetDefaultInterface(c))
    Debug.Assert d.HashVal(c) <> d.HashVal(New Collection)
    Debug.Assert d.HashVal("AA") <> d.HashVal("aa")
    Debug.Assert d.HashVal("AA") <> d.HashVal("AA" & vbNullChar)
    Debug.Assert d.HashVal(1E+300) > 0
    Debug.Assert d.HashVal(9999999999#) > 0
    For i = 1 To 99999
        Debug.Assert d.HashVal(i) > 0
        Debug.Assert d.HashVal(CStr(i)) > 0
        Debug.Assert d.HashVal(CStr(i)) <> d.HashVal(i)
    Next i
    '
    d.CompareMode = vbTextCompare
    Debug.Assert d.HashVal("AA") = d.HashVal("aa")
    Debug.Assert d.HashVal("AAAAAA") = d.HashVal("AAaAAA")
    '
    '
    On Error Resume Next
    i = d.HashVal(Array())
    Debug.Assert Err.Number = unsupportedKeyErr
    On Error GoTo 0
    '
    Debug.Assert d.HashVal(Array(), True) = 0
End Sub

Private Sub TestDictionaryIndex()
    Dim i As Long
    Dim d As New Dictionary
    '
    For i = 0 To 50000
        d.Add i, i
    Next i
    '
    Debug.Assert d.Index(25000) = 25000
    '
    For i = 2001 To 20000
        d.Remove i
    Next i
    '
    Debug.Assert d.Index(25000) = 7000
    Debug.Assert d.Index(30000) = 12000
    '
    For i = 30000 To 45000
        d.Remove i
    Next i
    '
    Debug.Assert d.Index(25000) = 7000
    Debug.Assert d.Index(28000) = 10000
    Debug.Assert d.Index(45001) = 12000
    '
    On Error Resume Next
    d.Index 15000
    Debug.Assert Err.Number = keyOrIndexNotFoundErr
    On Error GoTo 0
    '
    Debug.Assert d.Index(15000, True) = -1
End Sub

Private Sub TestDictionaryItem()
    Dim d As New Dictionary
    Dim c As New Collection
    Dim i As Long
    Dim v As Variant
    '
    For i = 1 To 5
        d.Add i, i
    Next i
    d.Add "coll", c
    d.Add c, Nothing
    d.Add "unk", GetDefaultInterface(c)
    d.Add Empty, Null
    d.Add Null, Empty
    d.Add CVErr(2042), 312
    '
    Debug.Assert d.Item("coll") Is c
    Debug.Assert d(c) Is Nothing 'Default property .Item is optional
    Debug.Assert d("unk") Is c
    Debug.Assert Not d("unk") Is Nothing
    For i = 1 To 5
        Debug.Assert d(i) = i
    Next i
    Debug.Assert IsNull(d(Empty))
    Debug.Assert IsEmpty(d(Null))
    Debug.Assert d(CVErr(2042)) = 312
    '
    d(Empty) = 5
    Debug.Assert d(Empty) = 5
    '
    On Error Resume Next
    d(1) = c
    Debug.Assert Err.Number = setMissingErr
    On Error GoTo 0
    '
    Set d(1) = c
    Debug.Assert d(1) Is c
    '
    d(1) = 5
    Debug.Assert d(1) = 5
    '
    Set d(1) = Nothing
    Debug.Assert d(1) Is Nothing
    '
    d.Item("new") = True 'Adds a new item
    Debug.Assert d("new") = True
    '
    If d.StrictScriptingMode Then
        Debug.Assert IsEmpty(d.Item("test"))
        For i = 100 To 105
            Debug.Assert IsEmpty(d.Item(i))
        Next i
    Else
        On Error Resume Next
        v = d.Item("test")
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
        For i = 100 To 105
            Err.Clear
            v = d.Item(i)
            Debug.Assert Err.Number = keyOrIndexNotFoundErr
        Next i
        On Error GoTo 0
        '
        d.CreateEmptyItemIfMissingKey = True
        Debug.Assert IsEmpty(d.Item("test"))
        For i = 100 To 105
            Debug.Assert IsEmpty(d.Item(i))
        Next i
        '
        d.CreateEmptyItemIfMissingKey = False
        On Error Resume Next
        For i = 106 To 200
            Err.Clear
            v = d.Item(i)
            Debug.Assert Err.Number = keyOrIndexNotFoundErr
        Next i
        On Error GoTo 0
    End If
    '
    On Error Resume Next
    v = d.Item(Array())
    Debug.Assert Err.Number = unsupportedKeyErr
    On Error GoTo 0
    '
    d.Add "Aa", 1
    d.Add "aa", 2
    d.Add "AA", 3
    d.Add "aA", 4
    d.Add "AA" & vbNullChar, 5
    Debug.Assert d("Aa") = 1
    Debug.Assert d("aa") = 2
    Debug.Assert d("AA") = 3
    Debug.Assert d("aA") = 4
    Debug.Assert d("AA" & vbNullChar) = 5
    '
    d.RemoveAll
    d.CompareMode = vbTextCompare
    '
    d.Add "AA", 1
    Debug.Assert d("Aa") = 1
    Debug.Assert d("aa") = 1
    Debug.Assert d("AA") = 1
    Debug.Assert d("aA") = 1
End Sub

Private Sub TestDictionaryItemAtIndex()
    Dim d As New Dictionary
    Dim c As New Collection
    Dim i As Long
    Dim v As Variant
    '
    For i = 1 To 5
        d.Add i, i
    Next i
    d.Add "coll", c
    d.Add c, Nothing
    d.Add "unk", GetDefaultInterface(c)
    d.Add Empty, Null
    d.Add Null, Empty
    d.Add CVErr(2042), 312
    '
    Debug.Assert d.ItemAtIndex(5) Is c
    Debug.Assert d.ItemAtIndex(6) Is Nothing
    Debug.Assert d.ItemAtIndex(7) Is c
    Debug.Assert Not d.ItemAtIndex(7) Is Nothing
    For i = 1 To 5
        Debug.Assert d.ItemAtIndex(i - 1) = i
    Next i
    Debug.Assert IsNull(d.ItemAtIndex(8))
    Debug.Assert IsEmpty(d.ItemAtIndex(9))
    Debug.Assert d.ItemAtIndex(10) = 312
    '
    d.ItemAtIndex(9) = 5
    Debug.Assert d.ItemAtIndex(9) = 5
    '
    On Error Resume Next
    d.ItemAtIndex(0) = c
    Debug.Assert Err.Number = setMissingErr
    On Error GoTo 0
    '
    Set d.ItemAtIndex(0) = c
    Debug.Assert d.ItemAtIndex(0) Is c
    '
    d.ItemAtIndex(0) = 5
    Debug.Assert d.ItemAtIndex(0) = 5
    '
    Set d.ItemAtIndex(0) = Nothing
    Debug.Assert d.ItemAtIndex(0) Is Nothing
    '
    On Error Resume Next
    v = d.ItemAtIndex(50)
    Debug.Assert Err.Number = keyOrIndexNotFoundErr
    For i = 100 To 105
        Err.Clear
        v = d.ItemAtIndex(i)
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    Next i
    On Error GoTo 0
    '
    d.Remove 3
    Debug.Assert d.ItemAtIndex(4) Is c
    Debug.Assert d.ItemAtIndex(5) Is Nothing
    Debug.Assert d.ItemAtIndex(6) Is c
    Debug.Assert d.ItemAtIndex(0) Is Nothing
    Debug.Assert d.ItemAtIndex(3) = 5
    Debug.Assert IsNull(d.ItemAtIndex(7))
    Debug.Assert d.ItemAtIndex(8) = 5
    '
    d.Remove 1
    Debug.Assert d.ItemAtIndex(0) = 2
    Debug.Assert d.ItemAtIndex(2) = 5
    Debug.Assert IsNull(d.ItemAtIndex(6))
    Debug.Assert d.ItemAtIndex(7) = 5
    '
    On Error Resume Next
    v = d.ItemAtIndex(d.Count)
    Debug.Assert Err.Number = keyOrIndexNotFoundErr
    On Error GoTo 0
End Sub

Private Sub TestDictionaryItems()
    Dim d As New Dictionary
    Dim i As Long
    '
    Debug.Assert ArrayToCSV(d.Items) = "[]"
    For i = 1 To 5
        d.Add i, i
    Next i
    Debug.Assert ArrayToCSV(d.Items) = "[1,2,3,4,5]"
    d.Add "coll", New Collection
    d.Add 111, Nothing
    d.Add Empty, Null
    d.Add Null, Empty
    d.Add CVErr(2042), 312
    Debug.Assert ArrayToCSV(d.Items) = "[1,2,3,4,5,Collection,Nothing,Null,Empty,312]"
    '
    For i = 1 To 3
        d.Remove i
    Next i
    d.Remove 111
    d.Remove "coll"
    d.Remove CVErr(2042)
    Debug.Assert ArrayToCSV(d.Items) = "[4,5,Null,Empty]"
    '
    d.RemoveAll
    Debug.Assert ArrayToCSV(d.Items) = "[]"
End Sub

Private Sub TestDictionaryKey()
    Dim d As New Dictionary
    Dim i As Long
    '
    d.Add "oldKey", 555
    d.Add "someKey", 444
    '
    On Error Resume Next
    d.Key("oldKey") = "someKey"
    Debug.Assert Err.Number = duplicatedKeyErr
    Err.Clear
    d.Key("oldKeyX") = "newKey"
    If d.StrictScriptingMode Then
        Debug.Assert Err.Number = failedScriptingErr
    Else
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    End If
    Err.Clear
    d.Key("oldkey") = "newKey"
    If d.StrictScriptingMode Then
        Debug.Assert Err.Number = failedScriptingErr
    Else
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    End If
    Err.Clear
    d.Key(Array()) = "newKey"
    Debug.Assert Err.Number = unsupportedKeyErr
    Err.Clear
    d.Key("oldkey") = Array(1, 2, 3)
    Debug.Assert Err.Number = unsupportedKeyErr
    On Error GoTo 0
    '
    d.Key("oldKey") = "newKey"
    Debug.Assert d("newKey") = 555
    Debug.Assert d("someKey") = 444
    '
    For i = 1 To 10
        d.Add i, i
    Next i
    For i = 1 To 10
        d.Key(i) = i + 10
    Next i
    For i = 1 To 10
        Debug.Assert d(i + 10) = i
    Next i
    '
    d.RemoveAll
    d.CompareMode = vbTextCompare
    d.Add "oldKey", 555
    d.Add "someKey", 444
    '
    On Error Resume Next
    d.Key("oldkey") = "SOMEKey"
    Debug.Assert Err.Number = duplicatedKeyErr
    Err.Clear
    d.Key("oldKeyX") = "newkey"
    If d.StrictScriptingMode Then
        Debug.Assert Err.Number = failedScriptingErr
    Else
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    End If
    Err.Clear
    d.Key("oldkey") = "newKey"
    Debug.Assert Err.Number = 0
    On Error GoTo 0
    '
    d.RemoveAll
    d.Add "oldKey", New Collection
    d.Add "someKey", 444
    '
    On Error Resume Next
    d.Key("oldkey") = "SOMEKey"
    Debug.Assert Err.Number = duplicatedKeyErr
    Err.Clear
    d.Key("oldkey") = "newKey"
    Debug.Assert Err.Number = 0
    On Error GoTo 0
End Sub

Private Sub TestDictionaryKeyAtIndex()
    Dim d As New Dictionary
    Dim c As New Collection
    Dim i As Long
    Dim v As Variant
    '
    For i = 1 To 5
        d.Add i, i
    Next i
    d.Add "coll", c
    d.Add c, Nothing
    d.Add "unk", GetDefaultInterface(c)
    d.Add Empty, Null
    d.Add Null, Empty
    d.Add CVErr(2042), 312
    '
    Debug.Assert d.KeyAtIndex(5) = "coll"
    Debug.Assert d.KeyAtIndex(6) Is c
    Debug.Assert d.KeyAtIndex(7) = "unk"
    For i = 1 To 5
        Debug.Assert d.KeyAtIndex(i - 1) = i
    Next i
    Debug.Assert IsEmpty(d.KeyAtIndex(8))
    Debug.Assert IsNull(d.KeyAtIndex(9))
    Debug.Assert IsError(d.KeyAtIndex(10))
    '
    On Error Resume Next
    v = d.KeyAtIndex(50)
    Debug.Assert Err.Number = keyOrIndexNotFoundErr
    For i = 100 To 105
        Err.Clear
        v = d.KeyAtIndex(i)
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    Next i
    On Error GoTo 0
    '
    d.Remove 3
    Debug.Assert d.KeyAtIndex(4) = "coll"
    Debug.Assert d.KeyAtIndex(5) Is c
    Debug.Assert d.KeyAtIndex(6) = "unk"
    Debug.Assert d.KeyAtIndex(0) = 1
    Debug.Assert d.KeyAtIndex(3) = 5
    Debug.Assert IsEmpty(d.KeyAtIndex(7))
    Debug.Assert IsNull(d.KeyAtIndex(8))
    '
    d.Remove 1
    Debug.Assert d.KeyAtIndex(0) = 2
    Debug.Assert d.KeyAtIndex(2) = 5
    Debug.Assert IsEmpty(d.KeyAtIndex(6))
    Debug.Assert IsNull(d.KeyAtIndex(7))
    '
    On Error Resume Next
    v = d.KeyAtIndex(d.Count)
    Debug.Assert Err.Number = keyOrIndexNotFoundErr
    On Error GoTo 0
End Sub

Private Sub TestDictionaryKeys()
    Dim d As New Dictionary
    Dim i As Long
    '
    Debug.Assert ArrayToCSV(d.Keys) = "[]"
    For i = 1 To 5
        d.Add i, i + 10
    Next i
    Debug.Assert ArrayToCSV(d.Keys) = "[1,2,3,4,5]"
    d.Add "coll", New Collection
    d.Add 111, Nothing
    d.Add Empty, Null
    d.Add Null, Empty
    d.Add CVErr(2042), 312
    Debug.Assert ArrayToCSV(d.Keys) = "[1,2,3,4,5,""coll"",111,Empty,Null," & CStr(CVErr(2042)) & "]"
    '
    For i = 1 To 3
        d.Remove i
    Next i
    d.Remove 111
    d.Remove "coll"
    d.Remove CVErr(2042)
    Debug.Assert ArrayToCSV(d.Keys) = "[4,5,Empty,Null]"
    '
    d.RemoveAll
    Debug.Assert ArrayToCSV(d.Keys) = "[]"
End Sub

Private Sub TestDictionaryKeysItems2D()
    Dim d As New Dictionary
    Dim i As Long
    Dim arr() As Variant
    '
    For i = 1 To 5
        d.Add i, i + 10
    Next i
    '
    arr = d.KeysItems2D
    For i = 1 To 5
        Debug.Assert arr(i - 1, 0) = i
        Debug.Assert arr(i - 1, 1) = i + 10
    Next i
    '
    d.Add "coll", New Collection
    d.Add 111, Nothing
    d.Add Empty, Null
    d.Add Null, Empty
    d.Add CVErr(2042), 312
    '
    arr = d.KeysItems2D
    Debug.Assert arr(5, 0) = "coll"
    Debug.Assert arr(6, 0) = 111
    Debug.Assert IsEmpty(arr(7, 0))
    Debug.Assert IsNull(arr(8, 0))
    Debug.Assert IsError(arr(9, 0))
    Debug.Assert TypeOf arr(5, 1) Is Collection
    Debug.Assert arr(6, 1) Is Nothing
    Debug.Assert IsNull(arr(7, 1))
    Debug.Assert IsEmpty(arr(8, 1))
    Debug.Assert arr(9, 1) = 312
    '
    For i = 1 To 3
        d.Remove i
    Next i
    d.Remove 111
    d.Remove "coll"
    d.Remove CVErr(2042)
    '
    arr = d.KeysItems2D
    Debug.Assert arr(0, 0) = 4
    Debug.Assert arr(1, 0) = 5
    Debug.Assert IsEmpty(arr(2, 0))
    Debug.Assert IsNull(arr(3, 0))
    Debug.Assert arr(0, 1) = 14
    Debug.Assert arr(1, 1) = 15
    Debug.Assert IsNull(arr(2, 1))
    Debug.Assert IsEmpty(arr(3, 1))
    '
    d.RemoveAll
    arr = d.KeysItems2D
    Debug.Assert UBound(arr, 1) - LBound(arr, 1) = -1
    Debug.Assert UBound(arr, 2) - LBound(arr, 2) = 0
End Sub

Private Sub TestDictionaryLoadFactor()
    Dim d As New Dictionary
    Dim i As Long
    Const MAX_LOAD_FACTOR As Single = 0.5
    '
    Debug.Assert d.LoadFactor = 0
    For i = 1 To 9999
        d.Add i, i
        Debug.Assert d.LoadFactor > 0
        Debug.Assert d.LoadFactor <= MAX_LOAD_FACTOR
    Next i
    d.RemoveAll
    Debug.Assert d.LoadFactor = 0
End Sub

Private Sub TestDictionaryNewEnum()
    Dim d As New Dictionary
    Dim i As Long
    Dim v As Variant
    Dim arr() As Variant
    '
    On Error Resume Next
    For Each v In d
        Err.Raise 5
    Next v
    Debug.Assert Err.Number = 0
    On Error GoTo 0
    '
    For i = 1 To 7
        d.Add i, i
    Next i
    arr = Array(1, 2, 3, 4, 5, 6, 7)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
    Next v
    Debug.Assert i = UBound(arr) + 1
    '
    d.Remove 5
    d.Remove 6
    arr = Array(1, 2, 3, 4, 7)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
    Next v
    Debug.Assert i = UBound(arr) + 1
    '
    d.Remove 1
    d.Remove 7
    arr = Array(2, 3, 4)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
    Next v
    Debug.Assert i = UBound(arr) + 1
    '
    d.Remove 2
    d.Remove 4
    '
    For Each v In d
        Debug.Assert v = 3
    Next v
    '
    d.Remove 3
    '
    On Error Resume Next
    For Each v In d
        Err.Raise 5
    Next v
    Debug.Assert Err.Number = 0
    On Error GoTo 0
    '
    For i = 1 To 7
        d.Add i, i
    Next i
    arr = Array(1, 2, 4, 21, 8, 9, 18, 22)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
        If i = 2 Then
            d.Remove 3
            d.Remove 5
            d.Key(6) = 21
            d.Remove 7
        End If
        If i = 3 Then
            Dim j As Long
            'Force Redim
            For j = 8 To 19
                d.Add j, j
            Next j
            For j = 10 To 17
                d.Remove j
            Next j
            d.Key(19) = 22
        End If
    Next v
    Debug.Assert i = UBound(arr) + 1
    '
    arr = Array(1, 4, 21, 8, 18)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
        If i = 1 Then
            d.Remove 1 'Already iterated
            d.Remove 2
            d.Remove 9
            d.Remove 22
        End If
    Next v
    Debug.Assert i = UBound(arr) + 1
    '
    d.RemoveAll
    '
    For i = 1 To 7
        d.Add i, i
    Next i
    '
    d.Remove 4
    arr = Array(1, 2, 3, 5, 6, 7)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
    Next v
    Debug.Assert i = UBound(arr) + 1
    '
    d.Remove 5
    d.Remove 6
    arr = Array(1, 2, 3, 7)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
    Next v
    Debug.Assert i = UBound(arr) + 1
    '
    d.Remove 1
    arr = Array(2, 3, 7)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
    Next v
    Debug.Assert i = UBound(arr) + 1
    '
    d.RemoveAll
    For i = 1 To 7
        d.Add i, i
    Next i
    '
    d.Key(5) = 9
    arr = Array(1, 2, 9, 11)
    '
    i = 0
    For Each v In d
        Debug.Assert v = arr(i)
        i = i + 1
        If i = 2 Then
            d.Remove 3
            d.Remove 4
            d.Remove 7
            d.Key(6) = 11
        End If
    Next v
    Debug.Assert i = UBound(arr) + 1
End Sub

Private Sub TestDictionaryRemove()
    Dim d As New Dictionary
    Dim i As Long
    '
    On Error Resume Next
    d.Remove Empty
    If d.StrictScriptingMode Then
        Debug.Assert Err.Number = failedScriptingErr
    Else
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    End If
    On Error GoTo 0
    '
    For i = 1 To 10
        d.Add CStr(i), i
    Next i
    '
    Debug.Assert d.Exists(CStr(5))
    d.Remove CStr(5)
    Debug.Assert Not d.Exists(CStr(5))
    Debug.Assert ArrayToCSV(d.Items) = "[1,2,3,4,6,7,8,9,10]"
    '
    For i = 2 To 10 Step 2
        d.Remove CStr(i)
    Next i
    Debug.Assert ArrayToCSV(d.Items) = "[1,3,7,9]"
    '
    d.Add CStr(5), 5
    For i = 1 To 10
        Debug.Assert d.Exists(CStr(i)) Xor (i Mod 2 = 0)
    Next i
     Debug.Assert d.Count = 5
    '
    d.RemoveAll
    '
    For i = 1 To 5
        d.Add CVErr(i), i
    Next i
    '
    On Error Resume Next
    d.Remove 1
    If d.StrictScriptingMode Then
        Debug.Assert Err.Number = failedScriptingErr
    Else
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    End If
    Err.Clear
    d.Remove CStr(1)
    If d.StrictScriptingMode Then
        Debug.Assert Err.Number = failedScriptingErr
    Else
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    End If
    On Error GoTo 0
    '
    Debug.Assert Not d.Remove(1, IgnoreErrors:=True)
    Debug.Assert Not d.Remove(CStr(1), IgnoreErrors:=True)
    '
    d.Add Null, Empty
    d.Add Empty, Null
    '
    Debug.Assert d.Exists(Null)
    Debug.Assert d.Remove(Null, True)
    Debug.Assert Not d.Exists(Null)
    '
    On Error Resume Next
    d.Remove Null, False
    If d.StrictScriptingMode Then
        Debug.Assert Err.Number = failedScriptingErr
    Else
        Debug.Assert Err.Number = keyOrIndexNotFoundErr
    End If
    Err.Clear
    d.Remove Array()
    Debug.Assert Err.Number = unsupportedKeyErr
    On Error GoTo 0
    '
    d.Remove CVErr(3)
    Debug.Assert ArrayToCSV(d.Items) = "[1,2,4,5,Null]"
    Debug.Assert ArrayToCSV(d.Keys) = "[" & Join(Array(CStr(CVErr(1)), CStr(CVErr(2)) _
                                                     , CStr(CVErr(4)), CStr(CVErr(5)) _
                                                     , "Empty"), ",") & "]"
    Debug.Assert d.Count = 5
    '
    d.Key(Empty) = CVErr(3)
    For i = 1 To 5
        d.Remove CVErr(i)
        Debug.Assert Not d.Exists(CVErr(i))
        Debug.Assert d.Count = 5 - i
    Next i
End Sub

Private Sub TestDictionaryRemoveAll()
    Dim d As New Dictionary
    Dim i As Long
    '
    d.RemoveAll
    For i = 1 To 500
        d.Add CStr(i), i
    Next i
    Debug.Assert d.Count = 500
    '
    d.RemoveAll
    Debug.Assert d.Count = 0
    '
    d.Add 1, 1
    d.RemoveAll
    Debug.Assert d.Count = 0
End Sub

Private Sub TestDictionarySelf()
    Dim d As New Dictionary
    '
    Debug.Assert d.Self Is d
    Debug.Assert Dictionary.Self Is Dictionary
End Sub

Private Sub TestDictionaryStackFixes()
#If x64StackIssues Then
    TestTerminate1
    TestTerminate2
    TestTerminate3
    TestDefault1
    TestDefault2
    TestNesting1
    TestNesting2
    TestNesting3
    TestForEach1
    TestForEach2
    TestForEach3
#End If
End Sub

#If x64StackIssues Then
'https://github.com/cristianbuse/VBA-FastDictionary/issues/10
Private Sub TestTerminate1()
    Dim v As Variant
    For Each v In ReturnVariantDict.Keys
    Next v
End Sub
Private Function ReturnVariantDict() As Variant
    Set ReturnVariantDict = New Dictionary
End Function

'https://stackoverflow.com/questions/65041832/vba-takes-wrong-branch-at-if-statement-severe-compiler-bug
Private Sub TestTerminate2()
    If Falsee(New Dictionary) Then
        Err.Raise 5, , "This shouldn't happen"
    End If
End Sub
Private Function Falsee(dict As Variant) As Boolean
    Falsee = False
End Function

'https://github.com/cristianbuse/VBA-FastDictionary/issues/15
Private Sub TestTerminate3()
    Dim i As Long
    For i = 1 To 10
        TestUDTCrash
    Next i
End Sub
Private Sub TestUDTCrash()
    Static myUDT As JustMyUDT
    Dim u As JustMyUDT
    myUDT = u 'Forces call to Class_Terminate
    Set myUDT.dict = New Dictionary
End Sub

'https://github.com/cristianbuse/VBA-FastDictionary/issues/16
Private Sub TestDefault1()
    Dim v As Variant: Set v = New Dictionary
    Set v(0) = New Dictionary
    v(0).Add "test", "test"
End Sub

Private Sub TestDefault2()
    Dim d As New Dictionary
    Dim i As Long
    Dim v As Variant
    Dim w As Variant
    Dim o As Object
    Const iterations As Long = 10000
    '
    d.Add 1, 1
    d.Add 2, 2
    For i = 1 To iterations
        v = d(1) 'Early bound
    Next i
    '
    Set o = d
    For i = 1 To iterations
        v = o(1) 'Late bound
    Next i
    '
    Set w = d
    For i = 1 To iterations
        v = w(1) 'Late bound Variant
    Next i
    '
    For i = 1 To iterations
        v = d(1) 'Early bound
        v = o(1) 'Late bound
        v = w(1) 'Late bound Variant
    Next i
End Sub

Private Sub TestNesting1()
    Dim d As Dictionary
    Dim i As Long
    Dim t As Variant
    Const nestingDepth As Long = 5000
    '
    Set d = New Dictionary
    Set t = d
    Do While i < nestingDepth
        t.Add 1, New Dictionary
        t.Add 2, New Dictionary
        i = i + 1
        Set t = t.Item(1) 'To test late-bind on Variant
    Loop
End Sub

Private Sub TestNesting2()
    Dim d As Dictionary
    Dim i As Long
    Dim t As Object
    Const nestingDepth As Long = 5000
    '
    Set d = New Dictionary
    Set t = d
    Do While i < nestingDepth
        t.Add 1, New Dictionary
        t.Add 2, New Dictionary
        i = i + 1
        Set t = t.Item(1) 'To test late-bind on Object
    Loop
End Sub

Private Sub TestNesting3()
    Dim d As Dictionary
    Dim i As Long
    Dim t As Dictionary
    Const nestingDepth As Long = 5000
    '
    Set d = New Dictionary
    Set t = d
    Do While i < nestingDepth
        t.Add 1, New Dictionary
        t.Add 2, New Dictionary
        i = i + 1
        Set t = t.Item(1) 'To test early-bind
    Loop
End Sub

'https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes
Private Sub TestForEach1()
    Dim d As Dictionary: Set d = New Dictionary
    d.Add 1, 1
    d.Add 2, 2
    ShowBug d
End Sub
Private Sub ShowBug(ByRef d As Dictionary)
    Dim ptr0 As LongPtr, ptr1 As LongPtr, ptr2 As LongPtr
    Dim ptr3 As LongPtr, ptr4 As LongPtr, ptr5 As LongPtr
    Dim ptr6 As LongPtr, ptr7 As LongPtr, ptr8 As LongPtr
    Dim v As Variant
    '
    On Error Resume Next
    For Each v In d
    Next v
    On Error GoTo 0
    '
    Debug.Assert (ptr0 = 0) And (ptr1 = 0) And (ptr2 = 0) _
             And (ptr3 = 0) And (ptr4 = 0) And (ptr5 = 0) _
             And (ptr6 = 0) And (ptr7 = 0) And (ptr8 = 0)
End Sub

Private Sub TestForEach2()
    Dim d As New Dictionary
    Dim i As Long
    Dim v As Variant
    '
    For i = 1 To 10000
        For Each v In d 'Late-bound
        Next
    Next i
End Sub

Private Sub TestForEach3()
    Dim d As New Dictionary
    Dim i As Long
    Dim c As New Collection
    Dim v As Variant
    '
    d.Add 1, 1
    d.Add 2, 2
    For i = 1 To 1000
        c.Add d.NewEnum 'Early-bound
    Next i
    For i = 1 To 1000   'Alternate
        c.Add d.NewEnum 'Early-bound
        For Each v In d 'Late-bound
        Next
    Next i
    Set c = Nothing
End Sub
#End If

Private Sub TestDictionaryTryGetItem()
    Dim d As New Dictionary
    Dim v As Variant
    '
    Debug.Assert Not d.TryGetItem("missing", v)
    '
    d.Add "key", 1
    Debug.Assert d.TryGetItem("key", v)
    Debug.Assert v = 1
End Sub
