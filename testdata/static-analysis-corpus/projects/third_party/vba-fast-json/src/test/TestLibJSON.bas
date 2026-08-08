Attribute VB_Name = "TestLibJSON"
'''=============================================================================
''' VBA Fast JSON Parser & Serializer - github.com/cristianbuse/VBA-FastJSON
''' ---
''' MIT License - github.com/cristianbuse/VBA-FastJSON/blob/master/LICENSE
''' Copyright (c) 2024 Ion Cristian Buse
'''=============================================================================
''' Additional to own tests, the below test suite includes many tests from:
''' https://github.com/nst/JSONTestSuite
''' ---
''' MIT License - https://github.com/nst/JSONTestSuite/blob/master/LICENSE
''' Copyright (c) 2016 Nicolas Seriot
''' Many thanks to Nicolas! Must-read his article:
''' Parsing JSON is a Minefield - https://seriot.ch/projects/parsing_json.html
'''=============================================================================

Option Explicit
Option Private Module

#Const Windows = (Mac = 0)
#Const x64 = Win64
Private Const dQuoteB As Byte = &H22

Public Sub RunAllJSONTests()
    RunAllJSONParseTests
    RunAllJSONSerializeTests
End Sub

'Test compliance with RFC 8259: https://www.rfc-editor.org/rfc/rfc8259
'Addtional extensions are also tested. See optional parameters for LibJSON.Parse
Public Sub RunAllJSONParseTests()
    TestParseEmptyInvalid
    TestParseLiteralValid
    TestParseLiteralInvalid
    TestParseWhitespaceValid
    TestParseWhitespaceInvalid
    TestParseArrayValid
    TestParseArrayValidLiteral
    TestParseArrayValidNesting
    TestParseArrayValidWithWhitespaces
    TestParseArrayInvalidComma
    TestParseArrayInvalidMisc
    TestParseArrayInvalidUnclosed
    TestParseObjectValid
    TestParseObjectValidNesting
    TestParseObjectInvalidMisc
    TestParseObjectInvalidUnclosed
    TestParseMiscValid
    TestParseMiscInvalid
    TestParseNumberValid
    TestParseNumberValidLargeExponent
    TestParseNumberInvalid
    TestParseStringValid
    TestParseStringValidEscape
    TestParseStringInvalid
    TestParseStringInvalidEscape
    TestParseStringLoneSurrogates
    TestParseStringInvalidUTF8
    Debug.Print "Finished running parser tests at " & Now()
End Sub

'*******************************************************************************
'Parse Tests
'*******************************************************************************
Private Sub TestParseEmptyInvalid()
    Debug.Assert Not Parse(vbNullString).IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse(" ").IsValid
    Debug.Assert Not Parse(vbNewLine).IsValid
    Debug.Assert Not Parse(vbTab).IsValid
    Debug.Assert Not Parse(vbCr).IsValid
    Debug.Assert Not Parse(vbLf).IsValid
    Debug.Assert Not Parse(vbTab & "  " & vbNewLine & "   ").IsValid
End Sub

Private Sub TestParseLiteralValid()
    Debug.Assert Not Parse("false").Value
    Debug.Assert Parse("true").Value
    Debug.Assert IsNull(Parse("null").Value)
    Debug.Assert Parse("[false,null,true]").Value.Count = 3
End Sub

Private Sub TestParseLiteralInvalid()
    Debug.Assert Not Parse("falsee").IsValid
    Debug.Assert Not Parse("False").IsValid
    Debug.Assert Not Parse("FALSE").IsValid
    Debug.Assert Not Parse("fals").IsValid
    Debug.Assert Not Parse("True").IsValid
    Debug.Assert Not Parse("tRue").IsValid
    Debug.Assert Not Parse("tru").IsValid
    Debug.Assert Not Parse("Null").IsValid
    Debug.Assert Not Parse("nul").IsValid
    Debug.Assert Not Parse("nulll").IsValid
    Debug.Assert Not Parse("undefined").IsValid
    Debug.Assert Not Parse("n/a").IsValid
    Debug.Assert Not Parse("empty").IsValid
    Debug.Assert Not Parse("blank").IsValid
End Sub

Private Sub TestParseWhitespaceValid()
    Debug.Assert Parse("    1").Value = 1
    Debug.Assert Parse("    12").Value = 12
    Debug.Assert Parse(" 123 ").Value = 123
    Debug.Assert Parse("    false").Value = False
    Debug.Assert Parse("    """"").IsValid
    Debug.Assert Parse("    """"").Value = ""
    Debug.Assert Parse("    1  ").Value = 1
    Debug.Assert Parse("    true   ").Value = True
    Debug.Assert Parse("    ""tru""").Value = "tru"
    Debug.Assert Parse(vbNewLine & "1").Value = 1
    Debug.Assert Parse(vbLf & "1").Value = 1
    Debug.Assert Parse(vbCr & "1").Value = 1
    Debug.Assert Parse(vbTab & vbTab & "1").Value = 1
    Debug.Assert Parse("  1  " & vbNewLine & "  " & vbLf & "  " & vbCr & "   " & vbTab).Value = 1
End Sub

Private Sub TestParseWhitespaceInvalid()
    Debug.Assert Not Parse(vbFormFeed & "1").IsValid
    Debug.Assert Not Parse(ChrW$(&H2060) & "0").IsValid   'Word Joiner (WJ)
End Sub

Private Sub TestParseArrayValid()
    Debug.Assert AreEqual(Parse("[]").Value, New Collection)
    Debug.Assert AreEqual(Parse("[""""]").Value, Collection(vbNullString))
    Debug.Assert AreEqual(Parse("[""text"",123,{},[],null]").Value _
                        , Collection("text", 123, New Dictionary, New Collection, Null))
    Debug.Assert AreEqual(Parse("[""" & ChrW$(&H2B7F) & """]").Value _
                        , Collection(ChrW$(&H2B7F))) 'Vertical Tab Key (not same as Line Tabulation / Vertical Tab)
    Debug.Assert AreEqual(Parse("[""abc""]").Value, Collection("abc"))
    Debug.Assert AreEqual(Parse("[ ""abc""]").Value, Collection("abc"))
    Debug.Assert AreEqual(Parse("[ ""abc""]" & vbNewLine).Value, Collection("abc"))
End Sub

Private Sub TestParseArrayValidLiteral()
    Debug.Assert AreEqual(Parse("[false]").Value, Collection(False))
    Debug.Assert AreEqual(Parse("[true]").Value, Collection(True))
    Debug.Assert AreEqual(Parse("[null]").Value, Collection(Null))
    Debug.Assert AreEqual(Parse("[false,null,true]").Value, Collection(False, Null, True))
    Debug.Assert AreEqual(Parse("[null,null,null]").Value, Collection(Null, Null, Null))
End Sub

Private Sub TestParseArrayValidNesting()
    Debug.Assert AreEqual(Parse("[[],[[]]]").Value, Collection(New Collection, Collection(New Collection)))
    '
#If TWINBASIC Then 'The built-in TB Collection leads to 'Out of stack space'
    Const nestingLevel As Long = 100
#Else
    Const nestingLevel As Long = 10000
#End If
    Dim v As Collection
    Dim c As Collection
    Dim i As Long
    '
    Set v = Parse(String$(nestingLevel, "[") & String$(nestingLevel, "]"), maxNestingDepth:=nestingLevel).Value
    Set c = v
    Do While v.Count = 1
        Set v = v.Item(1)
        i = i + 1
    Loop
    Debug.Assert v.Count = 0
    Debug.Assert i + 1 = nestingLevel
    '
    'The Termination stack frames would fill the stack and lead to 'Out of stack space'
    'We can only fix this if we build a wrapper class around Collection or if we write our own Array class
    Do
        Set c = c(1)
        i = i - 1
    Loop Until i < 5000
End Sub

Private Sub TestParseArrayValidWithWhitespaces()
    Debug.Assert AreEqual(Parse("[ [] ]" & vbNewLine).Value, Collection(New Collection))
    Debug.Assert AreEqual(Parse("[" & vbTab & "]" & vbCr).Value, New Collection)
    Debug.Assert AreEqual(Parse(" [] ").Value, New Collection)
    Debug.Assert AreEqual(Parse("[0" & vbNewLine & "]").Value, Collection(0))
    Debug.Assert AreEqual(Parse("   [  231]").Value, Collection(231))
    Debug.Assert AreEqual(Parse("[1] ").Value, Collection(1))
End Sub

Private Sub TestParseArrayInvalidComma()
    Debug.Assert Not Parse("[0 null]").IsValid
    Debug.Assert Not Parse("[0:1]").IsValid
    Debug.Assert Not Parse("[0;1]").IsValid
    Debug.Assert Not Parse("[0],").IsValid
    Debug.Assert Not Parse("[,123]").IsValid
    Debug.Assert Not Parse("[,1,2,3]").IsValid
    Debug.Assert Not Parse("[1,,2]").IsValid
    Debug.Assert Not Parse("[1,,]").IsValid
    Debug.Assert Not Parse("[null,,]").IsValid
    Debug.Assert Not Parse("[""a"",,]").IsValid
    Debug.Assert Not Parse("[0,]").IsValid
    Debug.Assert Not Parse("["""",]").IsValid
    Debug.Assert Not Parse("[0[1]]").IsValid
    Debug.Assert Not Parse("[0[1],]").IsValid
    Debug.Assert Not Parse("[,]").IsValid
    Debug.Assert Not Parse("[ , 0 ]").IsValid
    Debug.Assert Not Parse("[0,").IsValid
    Debug.Assert Not Parse("[""value"",,""anotherValue""]").IsValid
End Sub

Private Sub TestParseArrayInvalidMisc()
    Debug.Assert Not Parse("[*]").IsValid
    Debug.Assert Not Parse("[-]").IsValid
    Debug.Assert Not Parse("[+]").IsValid
    Debug.Assert Not Parse("[:]").IsValid
    Debug.Assert Not Parse("[\]").IsValid
    Debug.Assert Not Parse("[']").IsValid
    Debug.Assert Not Parse("['").IsValid
    Debug.Assert Not Parse("[,").IsValid
    Debug.Assert Not Parse("[""abc""\f]").IsValid
    Debug.Assert Not Parse("[a]").IsValid
    Debug.Assert Not Parse("[" & ChrW$(1234) & "]").IsValid
    Debug.Assert Not Parse("[" & vbFormFeed & "]").IsValid
    Debug.Assert Not Parse(String$(512, "[")).IsValid
    Debug.Assert Not Parse(RepeatString("[{"""":", 512)).IsValid
    Debug.Assert Not Parse("[false,tru").IsValid
    Debug.Assert Not Parse("[false,nul").IsValid
    Debug.Assert Not Parse("[fals,true").IsValid
    Debug.Assert Not Parse("[true,fals").IsValid
    Debug.Assert Not Parse("[1]a").IsValid
    Debug.Assert Not Parse("[""abc]").IsValid
    Debug.Assert Not Parse("[][]").IsValid
    Debug.Assert Not Parse("[""a"", a]").IsValid
    Debug.Assert Not Parse(BytesToString(&H5B, &HFF, &H5D)).IsValid
    Debug.Assert Not Parse(BytesToString(&H5B, &H61, &HE5, &H5D)).IsValid
End Sub

Private Sub TestParseArrayInvalidUnclosed()
    Debug.Assert Not Parse("[").IsValid
    Debug.Assert Not Parse("[[").IsValid
    Debug.Assert Not Parse("]").IsValid
    Debug.Assert Not Parse("]]").IsValid
    Debug.Assert Not Parse("[}").IsValid
    Debug.Assert Not Parse("{]").IsValid
    Debug.Assert Not Parse("[{").IsValid
    Debug.Assert Not Parse("{[").IsValid
    Debug.Assert Not Parse("[{}").IsValid
    Debug.Assert Not Parse("[[]]]").IsValid
    Debug.Assert Not Parse("[""a""]]").IsValid
    Debug.Assert Not Parse("[""a""").IsValid
    Debug.Assert Not Parse("[a").IsValid
    Debug.Assert Not Parse("[""a""").IsValid
    Debug.Assert Not Parse("[1").IsValid
    Debug.Assert Not Parse("[1]]").IsValid
    Debug.Assert Not Parse("]1]").IsValid
    Debug.Assert Not Parse("1]").IsValid
    Debug.Assert Not Parse("[1," & vbNewLine & "2" & vbNewLine & ",3").IsValid
    Debug.Assert Not Parse("[""1""," & vbNewLine & "2" & vbNewLine & ",3").IsValid
End Sub

Private Sub TestParseObjectValid()
    Debug.Assert AreEqual(Parse("{}").Value, New Dictionary)
    Debug.Assert AreEqual(Parse("{"""":0}").Value, Dictionary(vbNullString, 0))
    Debug.Assert AreEqual(Parse("{""key"":""value""}").Value _
                        , Dictionary("key", "value"))
    Debug.Assert AreEqual(Parse("{""key1"":123,""key2"":true}").Value _
                        , Dictionary("key1", 123, "key2", True))
    Debug.Assert AreEqual(Parse("{""key1"":123,""key2"":true}" & vbNewLine).Value _
                        , Dictionary("key1", 123, "key2", True))
    Debug.Assert AreEqual(Parse("{""nestedObj"":{""key"":123}}").Value _
                        , Dictionary("nestedObj", Dictionary("key", 123)))
    Debug.Assert AreEqual(Parse("{""abc"":""def""}").Value _
                        , Dictionary("abc", "def"))
    Debug.Assert AreEqual(Parse("{""abc"":""def"",""ghi"":""jkl""}").Value _
                        , Dictionary("abc", "def", "ghi", "jkl"))
    Dim d As Dictionary
    If IsFastDict Then 'Fast-Dictionary can handle Duplicate Keys
        Set d = New Dictionary
        d.AllowDuplicateKeys = True
        d.Add "a", "b"
        d.Add "a", "c"
        Debug.Assert AreEqual(Parse("{""a"":""b"",""a"":""c""}", allowDuplicatedKeys:=True).Value, d)
        '
        Set d = New Dictionary
        d.AllowDuplicateKeys = True
        d.Add "a", "b"
        d.Add "a", "b"
        Debug.Assert AreEqual(Parse("{""a"":""b"",""a"":""b""}", allowDuplicatedKeys:=True).Value, d)
    End If
    Debug.Assert AreEqual(Parse("{""ke\u0000y"":null}").Value _
                        , Dictionary("ke" & vbNullChar & "y", Null))
    Debug.Assert AreEqual(Parse("{""min"":-1.0e+308,""max"":1.0E+308}").Value _
                        , Dictionary("min", -1E+308, "max", 1E+308))
    Debug.Assert AreEqual(Parse("{""a"":[{""b"": ""abcdefghijklmnopqrstuvwxyz0123456789@;:~>""}]}").Value _
                        , Dictionary("a", Collection(Dictionary("b", "abcdefghijklmnopqrstuvwxyz0123456789@;:~>"))))
    Debug.Assert AreEqual(Parse("{""key"":[]}").Value _
                        , Dictionary("key", New Collection))
    Debug.Assert AreEqual(Parse("{""key"":""\u0418\u043D\u0442\u0435\u0440\u0435\u0441\u043D\u043E, \u043D\u0430\u043B\u0438?""}").Value _
                        , Dictionary("key", ChrW$(&H418) & ChrW$(&H43D) & ChrW$(&H442) & ChrW$(&H435) & ChrW$(&H440) _
                                          & ChrW$(&H435) & ChrW$(&H441) & ChrW$(&H43D) & ChrW$(&H43E) & ", " _
                                          & ChrW$(&H43D) & ChrW$(&H430) & ChrW$(&H43B) & ChrW$(&H438) & "?"))
    Debug.Assert AreEqual(Parse(vbNewLine & "{""key"":[]}" & vbTab & vbNewLine).Value _
                        , Dictionary("key", New Collection))
    Debug.Assert AreEqual(Parse(ChrW$(&HBBEF) & ChrB$(&HBF) & "{}").Value, New Dictionary) 'UTF8 BOM - ignored
    Debug.Assert AreEqual(Parse(ChrW$(&HFEFF) & "{}").Value, New Dictionary) 'UTF16LE BOM - ignored
    Debug.Assert AreEqual(Parse("{""\uDFAA"":0}").Value, Dictionary(ChrW$(&HDFAA), 0)) 'Single surrogate
End Sub

Private Sub TestParseObjectValidNesting()
    Dim nestingLevel As Long
    Dim v As Dictionary
    Dim i As Long
    Dim h As Dictionary
    '
#If TWINBASIC Then 'Fast Dictionary does not manage nesting in TB (only in VB*)
    nestingLevel = 100
#Else
    If IsFastDict() Then
        nestingLevel = 10000
    Else 'Scripting or other
        nestingLevel = 100
    End If
#End If
    Set v = Parse(RepeatString("{""key"":", nestingLevel - 1) & "{" _
                    & String$(nestingLevel, "}"), maxNestingDepth:=nestingLevel).Value
    If IsFastDict() Then Set h = v 'Only Fast-Dictionary can handle deep nesting termination
    Do While v.Count = 1
        Set v = v.Item("key")
        i = i + 1
    Loop
    Debug.Assert v.Count = 0
    Debug.Assert i + 1 = nestingLevel
    '
    Dim w As Object
    Set w = Parse(RepeatString("{""key"":[", nestingLevel) _
                    & RepeatString("]}", nestingLevel), maxNestingDepth:=nestingLevel * 2).Value
    If IsFastDict() Then Set h = w 'Only Fast-Dictionary can handle deep nesting termination
    i = 0
    Do While w.Count = 1
        If TypeOf w Is Collection Then
            Set w = w.Item(1)
        Else
            Set w = w.Item("key")
        End If
        i = i + 1
    Loop
    Debug.Assert w.Count = 0
    Debug.Assert i = nestingLevel * 2 - 1
End Sub

Private Sub TestParseObjectInvalidMisc()
    Debug.Assert Not Parse("{key:""value""}").IsValid
    Debug.Assert Not Parse("{""key"":}").IsValid
    Debug.Assert Not Parse("{""key"":""value"",}").IsValid
    Debug.Assert Not Parse("[{""k1"":1,""k2""}]").IsValid
    Debug.Assert Not Parse("[{""k1"":1 ""k2"":2}]").IsValid
    Debug.Assert Not Parse("[{""k1"":1] ""k2"":2}").IsValid
    Debug.Assert Not Parse("{""a"":""b"",""a"":""c""}", allowDuplicatedKeys:=False).IsValid
    Debug.Assert Not Parse("{""a"":""b"",""a"":""b""}", allowDuplicatedKeys:=False).IsValid
    Debug.Assert Not Parse("{""a"":""b"",""a"":""c""}", allowDuplicatedKeys:=False).IsValid
    Debug.Assert Not Parse("{""a"":""b"",""a"":""b""}", allowDuplicatedKeys:=False).IsValid
    Debug.Assert Not Parse(ChrW$(&HBBEF) & ChrB$(&HBF) & "{}", failIfBOMDetected:=True).IsValid
    Debug.Assert Not Parse(ChrW$(&HFEFF) & "{}", failIfBOMDetected:=True).IsValid
    Debug.Assert Not Parse("{1:0}").IsValid
    Debug.Assert Not Parse("{11111:0}").IsValid
    Debug.Assert Not Parse("{11111e11111:0}").IsValid
    Debug.Assert Not Parse("{[: ""abc""}").IsValid
    Debug.Assert Not Parse("{""key"", null}").IsValid
    Debug.Assert Not Parse("{""key""::null}").IsValid
    Debug.Assert Not Parse(BytesToString(&H7B, &HF0, &H9F, &H87, &HA8, &HF0, &H9F, &H87, &HAD, &H7D)).IsValid '{Emoji}
    Debug.Assert Not Parse("{""a"":""a"" 123}").IsValid
    Debug.Assert Not Parse("{key: 'value'}").IsValid
    Debug.Assert Not Parse(BytesToString(&H7B, &H22, &HB9, &H22, &H3A, &H22, &H30, &H22, &H2C, &H7D)).IsValid
    Debug.Assert Not Parse("{""a"" b}").IsValid
    Debug.Assert Not Parse("{:""b""}").IsValid
    Debug.Assert Not Parse("{""a"" ""b""}").IsValid
    Debug.Assert Not Parse("{1:1}").IsValid
    Debug.Assert Not Parse("{null:null,null:null}").IsValid
    Debug.Assert Not Parse("{""id"":0,,,,,}").IsValid
    Debug.Assert Not Parse("{""id"":0,,,,,}", ignoreTrailingComma:=True).IsValid
    Debug.Assert Not Parse("{'a':0}").IsValid
    Debug.Assert Not Parse("{""id"":0,}").IsValid
    Debug.Assert Not Parse("{""a"":""b""}/**/").IsValid  'Comments
    Debug.Assert Not Parse("{""a"":""b""}/**//").IsValid
    Debug.Assert Not Parse("{""a"":""b""}//").IsValid
    Debug.Assert Not Parse("{""a"":""b""}/").IsValid
    Debug.Assert Not Parse("{""a"":1,,""b"":2}").IsValid
    Debug.Assert Not Parse("{a: ""b""}").IsValid
    Debug.Assert Not Parse("{ ""k"" : ""v"", ""s"" }").IsValid
    Debug.Assert Not Parse("{""a"":""b""}#").IsValid
End Sub

Private Sub TestParseObjectInvalidUnclosed()
    Debug.Assert Not Parse("{").IsValid
    Debug.Assert Not Parse("[}").IsValid
    Debug.Assert Not Parse("{]").IsValid
    Debug.Assert Not Parse("{[").IsValid
    Debug.Assert Not Parse("{"":").IsValid
    Debug.Assert Not Parse("{}}").IsValid
    Debug.Assert Not Parse("[0,{,1]").IsValid
    Debug.Assert Not Parse("{""a"":").IsValid
    Debug.Assert Not Parse("{""a""").IsValid
    Debug.Assert Not Parse("{""a"":""a").IsValid
End Sub

Private Sub TestParseMiscValid()
    Debug.Assert AreEqual(Parse("[{""k1"":true}, {""k2"":null}]").Value _
                        , Collection(Dictionary("k1", True), Dictionary("k2", Null)))
    Debug.Assert AreEqual(Parse("{""arr"":[1,2,3],""obj"":{""k1"":1,""K1"":2}}").Value _
                        , Dictionary("arr", Collection(1, 2, 3), "obj", Dictionary("k1", 1, "K1", 2)))
    Debug.Assert AreEqual(Parse(BytesToString(&H0, &H5B, &H0, &H22, &H0, &HE9, &H0, &H22, &H0, &H5D)).Value _
                        , Collection(ChrW$(&HE9))) 'UTF16BE
    Debug.Assert AreEqual(Parse(BytesToString(&HFF, &HFE, &H5B, &H0, &H22, &H0, &HE9, &H0, &H22, &H0, &H5D, &H0)).Value _
                        , Collection(ChrW$(&HE9))) 'UTF16LE with UTF16LE BOM
    Debug.Assert AreEqual(Parse(BytesToString(&HEF, &HBB, &HBF, &H5B, &H0, &H22, &H0, &HE9, &H0, &H22, &H0, &H5D, &H0)).Value _
                        , Collection(ChrW$(&HE9))) 'UTF16LE with UTF8 BOM
    Debug.Assert AreEqual(Parse("[0,]", ignoreTrailingComma:=True).Value, Collection(0))
    Debug.Assert AreEqual(Parse("["""",]", ignoreTrailingComma:=True).Value, Collection(vbNullString))
    Debug.Assert AreEqual(Parse("{""key"":""value"",}", ignoreTrailingComma:=True).Value, Dictionary("key", "value"))
End Sub

Private Sub TestParseMiscInvalid()
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("<.>").IsValid
    Debug.Assert Not Parse("[<null>]").IsValid
    Debug.Assert Not Parse(BytesToString(&H61, &HC3, &HA5)).IsValid
    Debug.Assert Not Parse("[True]").IsValid
    Debug.Assert Not Parse("{""x"": true,").IsValid
    Debug.Assert Not Parse(BytesToString(&HEF, &HBB, &H7B, &H7D)).IsValid 'Incomplete BOM
    Debug.Assert Not Parse(BytesToString(&HEF, &HBB, &HBF, &H7B, &H7D), failIfBOMDetected:=True).IsValid 'Complete BOM
    Debug.Assert Not Parse(BytesToString(&HEF, &HBB, &HBF)).IsValid 'Complete BOM with no data
    Debug.Assert Not Parse(Chr$(&HE5)).IsValid
    Debug.Assert Not Parse("[").IsValid
    Debug.Assert Not Parse("{").IsValid
    Debug.Assert Not Parse("[}").IsValid
    Debug.Assert Not Parse("{[").IsValid
    Debug.Assert Not Parse("[" & vbNullChar & "]").IsValid
    Debug.Assert Not Parse(StrConv("[" & vbNullChar & "]", vbFromUnicode)).IsValid
    Debug.Assert Not Parse("{}}").IsValid
    Debug.Assert Not Parse("{"""":").IsValid
    Debug.Assert Not Parse("{""a"":/*comment*/""b""}").IsValid
    Debug.Assert Not Parse("{""a"": true} ""x""").IsValid
    Debug.Assert Not Parse("{,").IsValid
    Debug.Assert Not Parse("{""a").IsValid
    Debug.Assert Not Parse("{'a'").IsValid
    Debug.Assert Not Parse("[""\{[""\{[""\{[""\{").IsValid
    Debug.Assert Not Parse(ChrW$(&H17D)).IsValid
    Debug.Assert Not Parse("*").IsValid
    Debug.Assert Not Parse("{""a"":""b""}#{}").IsValid
    Debug.Assert Not Parse(BytesToString(&H5B, &HE2, &H81, &HA0, &H5D)).IsValid
    Debug.Assert Not Parse("[\u000A""""]").IsValid
    Debug.Assert Not Parse("{""abc"":1").IsValid
    Debug.Assert Not Parse(BytesToString(&HC3, &HA5)).IsValid
    Debug.Assert Not Parse(BytesToString(&H5B, &HE2, &H81, &HA0, &H5D)).IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
    Debug.Assert Not Parse("").IsValid
End Sub

Private Sub TestParseNumberValid()
    Debug.Assert Parse("0").Value = 0
    Debug.Assert Parse("-0").Value = 0
    Debug.Assert Parse("1").Value = 1
    Debug.Assert Parse("-1").Value = -1
    Debug.Assert Parse("12").Value = 12
    Debug.Assert Parse("-12").Value = -12
    Debug.Assert Parse("123").Value = 123
    Debug.Assert Parse(" 3").Value = 3
    Debug.Assert Parse("-1.2").Value = -1.2
    Debug.Assert Parse("0E0").Value = 0
    Debug.Assert Parse("0E1").Value = 0
    Debug.Assert Parse("0E-1").Value = 0
    Debug.Assert Parse("0E+1").Value = 0
    Debug.Assert Parse("12E4").Value = 120000
    Debug.Assert Parse("1E25").Value = 1E+25
    Debug.Assert Parse("1E-2").Value = 0.01
    Debug.Assert Parse("1E+2").Value = 100
    Debug.Assert Parse("123e45").Value = 1.23E+47
    Debug.Assert Parse("123.45e67").Value = 1.2345E+69
    Debug.Assert Parse("123.4567").Value = 123.4567
End Sub

Private Sub TestParseNumberValidLargeExponent()
    Debug.Assert Parse("[1.23456789E-999]").Value(1) = 0
    Debug.Assert Parse("[4.94065645841247E-324]").Value(1) = 4.94065645841247E-324
    Debug.Assert Parse("[1.79769313486231E308]").Value(1) = 1.79769313486231E+308
    Debug.Assert Parse("[-1.79769313486231E308]").Value(1) = -1.79769313486231E+308
    Debug.Assert Parse("[-4.94065645841247E-324]").Value(1) = -4.94065645841247E-324
    Debug.Assert Parse("-0.000000000000000000000000000000000000000000000000000001").Value = -1E-54
#If Windows Then
    Debug.Assert Parse("[-922337203685477.5807]").Value(1) = -922337203685477.5807@
    Debug.Assert Parse("[-79228162514264337593543950335]").Value(1) = CDec("-79228162514264337593543950335")
    Debug.Assert Parse("[79228162514264337593543950335]").Value(1) = CDec("79228162514264337593543950335")
    Debug.Assert Parse("[-7.9228162514264337593543950335]").Value(1) = CDec(ToLocaleDot("-7.9228162514264337593543950335"))
    Debug.Assert Parse("[7.9228162514264337593543950335]").Value(1) = CDec(ToLocaleDot("7.9228162514264337593543950335"))
#Else
    Debug.Assert Parse("[-922337203685477]").Value(1) = -922337203685477@
#End If
    Debug.Assert Parse("[-0.0000000000000000000000000001]").Value(1) = -1E-28
    Debug.Assert Parse("[0.1E000100]").Value(1) = 1E+99
    Debug.Assert Parse("[0.1E+00100]").Value(1) = 1E+99
    Debug.Assert Parse("-1234567890123456789012345678901234567890").Value - -1.23456789012346E+39 < 1E+25
    Debug.Assert Parse("1234567890123456789012345678901234567890").Value - 1.23456789012346E+39 < 1E+25
End Sub

Private Sub TestParseNumberInvalidLargeExponent()
    Debug.Assert Not Parse("0.1e0123456789").IsValid
    Debug.Assert Not Parse("-1e+1000").IsValid
    Debug.Assert Not Parse("1.1e+10000000000000").IsValid
    Debug.Assert Not Parse("1.1e-10000000000000").IsValid
    Debug.Assert Not Parse("1e-10000000000000").IsValid
    Debug.Assert Not Parse("-1123456e1000").IsValid
    Debug.Assert Not Parse("1123456e1000").IsValid
End Sub

Private Sub TestParseNumberInvalid()
    Debug.Assert Not Parse("123" & vbNullChar).IsValid
    Debug.Assert Not Parse("0123").IsValid  'Leading zeros are not allowed
    Debug.Assert Not Parse("00123").IsValid
    Debug.Assert Not Parse("0.").IsValid    'Point must be followed by at least a digit
    Debug.Assert Not Parse("1E").IsValid    'Exponent must be followed by at least a digit
    Debug.Assert Not Parse("00E0").IsValid
    Debug.Assert Not Parse("1E-").IsValid
    Debug.Assert Not Parse("1E+").IsValid
    Debug.Assert Not Parse("inf").IsValid
    Debug.Assert Not Parse("+inf").IsValid
    Debug.Assert Not Parse("Inf").IsValid
    Debug.Assert Not Parse("+Inf").IsValid
    Debug.Assert Not Parse("infinity").IsValid
    Debug.Assert Not Parse("Infinity").IsValid
    Debug.Assert Not Parse("-inf").IsValid
    Debug.Assert Not Parse("-Inf").IsValid
    Debug.Assert Not Parse("-infinity").IsValid
    Debug.Assert Not Parse("-Infinity").IsValid
    Debug.Assert Not Parse("NaN").IsValid
    Debug.Assert Not Parse("nan").IsValid
    Debug.Assert Not Parse("-NaN").IsValid
    Debug.Assert Not Parse("-nan").IsValid
    Debug.Assert Not Parse("-nan(ind)").IsValid
    Debug.Assert Not Parse("0.1E+001E00").IsValid
    Debug.Assert Not Parse(".-1").IsValid
    Debug.Assert Not Parse(".1e-2").IsValid
    Debug.Assert Not Parse(".12").IsValid
    Debug.Assert Not Parse("+1").IsValid
    Debug.Assert Not Parse("+12").IsValid
    Debug.Assert Not Parse("++1").IsValid
    Debug.Assert Not Parse("--1").IsValid
    Debug.Assert Not Parse("0.1.2").IsValid
    Debug.Assert Not Parse("0.1.").IsValid
    Debug.Assert Not Parse("0.1E").IsValid
    Debug.Assert Not Parse("0.1e+").IsValid
    Debug.Assert Not Parse("0.1e-").IsValid
    Debug.Assert Not Parse("0.E1").IsValid
    Debug.Assert Not Parse("0.e").IsValid
    Debug.Assert Not Parse("0.e+").IsValid
    Debug.Assert Not Parse("0.E+").IsValid
    Debug.Assert Not Parse("0.e-").IsValid
    Debug.Assert Not Parse("0e").IsValid
    Debug.Assert Not Parse("0e+").IsValid
    Debug.Assert Not Parse("0e-").IsValid
    Debug.Assert Not Parse("-01").IsValid
    Debug.Assert Not Parse("-001").IsValid
    Debug.Assert Not Parse("-00.1").IsValid
    Debug.Assert Not Parse("-1.0.").IsValid
    Debug.Assert Not Parse("1.0e").IsValid
    Debug.Assert Not Parse("-1.0e").IsValid
    Debug.Assert Not Parse("1.0e-").IsValid
    Debug.Assert Not Parse("1.0e+").IsValid
    Debug.Assert Not Parse("1 0.0").IsValid
    Debug.Assert Not Parse("1Ee2").IsValid
    Debug.Assert Not Parse("1.").IsValid
    Debug.Assert Not Parse("-1.").IsValid
    Debug.Assert Not Parse("+1.").IsValid
    Debug.Assert Not Parse("1.e2").IsValid
    Debug.Assert Not Parse("1.e+2").IsValid
    Debug.Assert Not Parse("1.e-2").IsValid
    Debug.Assert Not Parse("2.e+").IsValid
    Debug.Assert Not Parse("0e+-1").IsValid
    Debug.Assert Not Parse("0e-+1").IsValid
    Debug.Assert Not Parse("1+2").IsValid
    Debug.Assert Not Parse("0x0").IsValid
    Debug.Assert Not Parse("0x1").IsValid
    Debug.Assert Not Parse("0x00").IsValid
    Debug.Assert Not Parse("0x12").IsValid
    Debug.Assert Not Parse("1.2a").IsValid
    Debug.Assert Not Parse("-1.2a").IsValid
    Debug.Assert Not Parse("112a").IsValid
    Debug.Assert Not Parse(StrConv("112", vbFromUnicode) & ChrW$(&HE55D)).IsValid
    Debug.Assert Not Parse(StrConv("1E2", vbFromUnicode) & ChrW$(&HE55D)).IsValid
    Debug.Assert Not Parse(StrConv("1", vbFromUnicode) & ChrW$(&HE55D)).IsValid
    Debug.Assert Not Parse("-abc").IsValid
    Debug.Assert Not Parse("- 7").IsValid
    Debug.Assert Not Parse("-07").IsValid
    Debug.Assert Not Parse("-.7").IsValid
    Debug.Assert Not Parse("-0x").IsValid
    Debug.Assert Not Parse("1ex").IsValid
    Debug.Assert Not Parse("1e" & ChrW$(&HE55D)).IsValid
    Debug.Assert Not Parse(ChrW$(-239)).IsValid 'U+FF11 (Fullwidth Digit One Unicode Character)
    Debug.Assert Not Parse("1a3").IsValid
    Debug.Assert Not Parse("1.2a3").IsValid
    Debug.Assert Not Parse("1.2a-3").IsValid
    Debug.Assert Not Parse("1.23456789a-999").IsValid
    Debug.Assert Not Parse("1.23456789x-999").IsValid
    Debug.Assert Not Parse("1.23456789h-999").IsValid
    Debug.Assert Not Parse(BytesToString(&H31, &H65, &HE5)).IsValid
End Sub

Private Sub TestParseStringValid()
    Debug.Assert Parse("""""").Value = vbNullString
    Debug.Assert Parse("""abc""").Value = "abc"
    Debug.Assert Parse("""abc """).Value = "abc "
    Debug.Assert Parse("""abc""" & vbNewLine).Value = "abc"
    Debug.Assert Parse("""abc""" & vbNewLine).Value = "abc"
    Debug.Assert Parse(""" """).Value = " "
    Debug.Assert Parse("""a/*b*/c/*d//e""").Value = "a/*b*/c/*d//e" 'Comments are treated as normal text while inside strings
    Debug.Assert Parse(BytesToString(dQuoteB, &HF4, &H8F, &HBF, &HBF, dQuoteB)).Value = ChrW$(&HDBFF) & ChrW$(&HDFFF) 'Non UTF8 character U+10FFFF
    Debug.Assert Parse(BytesToString(dQuoteB, &HEF, &HBF, &HBF, dQuoteB)).Value = ChrW$(&HFFFF)                       'Non UTF8 character U+FFFF
    Debug.Assert Parse(BytesToString(dQuoteB, &H2C, dQuoteB)).Value = ","
    Debug.Assert Parse(BytesToString(dQuoteB, &HCF, &H80, dQuoteB)).Value = ChrW$(&H3C0) 'Math PI
    Debug.Assert Parse(BytesToString(dQuoteB, &HF0, &H9B, &HBF, &HBF, dQuoteB)).Value = ChrW$(&HD82F) & ChrW$(&HDFFF) 'Reserved character U+1BFFF
    Debug.Assert Parse(BytesToString(dQuoteB, &HE2, &H80, &HA8, dQuoteB)).Value = ChrW$(&H2028) 'Line Separator (LS)
    Debug.Assert Parse(BytesToString(dQuoteB, &HE2, &H80, &HA9, dQuoteB)).Value = ChrW$(&H2029) 'Paragraph Separator (PS)
    Debug.Assert Parse(BytesToString(dQuoteB, &H7F, dQuoteB)).Value = Chr$(&H7F) 'DEL
    Debug.Assert Parse(BytesToString(dQuoteB, &H61, &H7F, &H61, dQuoteB)).Value = "a" & Chr$(&H7F) & "a"
    Debug.Assert Parse(BytesToString(dQuoteB, &HE2, &H8D, &H82, &HE3, &H88, &HB4, &HE2, &H8D, &H82, dQuoteB)).Value = ChrW$(&H2342) & ChrW$(&H3234) & ChrW$(&H2342)
    If IsFastDict() Then
        With New Dictionary
            .AllowDuplicateKeys = True
            .CompareMode = vbTextCompare
            .Add "key", 1
            .Add "KEY", 2
            Debug.Assert AreEqual(Parse("{""key"":1,""KEY"":2}", allowDuplicatedKeys:=True, keyCompareMode:=vbTextCompare).Value, .Self)
        End With
    End If
End Sub

Private Sub TestParseStringValidEscape()
    Debug.Assert Parse("""value\nwith\nnewlines""").Value = "value" & vbLf & "with" & vbLf & "newlines"
    Debug.Assert Parse("""value with \""escaped quotes\""""").Value = "value with ""escaped quotes"""
    Debug.Assert Parse("""value\twith\ttabs""").Value = "value" & vbTab & "with" & vbTab & "tabs"
    Debug.Assert Parse("""value with \\backslashes\\""").Value = "value with \backslashes\"
    Debug.Assert Parse("""a\\b""").Value = Parse("""a\u005Cb""").Value
    Debug.Assert Parse("""\u0060\u012a\u12AB""").Value = "`" & ChrW$(&H12A) & ChrW$(&H12AB)
    Debug.Assert Parse("""\uD801\udc37""").Value = ChrW$(&HD801) & ChrW$(&HDC37) 'Surrogate pair
    Debug.Assert Parse("""\ud83d\ude39\ud83d\udc8d""").Value = ChrW$(&HD83D) & ChrW$(&HDE39) & ChrW$(&HD83D) & ChrW$(&HDC8D) 'Surrogate pairs
    Debug.Assert Parse("""\""\\\/\b\f\n\r\t""").Value = """\/" & vbBack & vbFormFeed & vbLf & vbCr & vbTab
    Debug.Assert Parse("""\\u0000""").Value = "\u0000"
    Debug.Assert Parse("""\""""").Value = """"
    Debug.Assert Parse("""\\a""").Value = "\a"
    Debug.Assert Parse("""\\n""").Value = "\n"
    Debug.Assert Parse("""\u0012""").Value = Chr$(&H12) 'Device Control 2
    Debug.Assert Parse(StrConv("""\uFFFF""", vbFromUnicode)).Value = ChrW$(&HFFFF) 'Escaped non-character
    Debug.Assert Parse("""\uDBFF\uDFFF""").Value = ChrW$(&HDBFF) & ChrW$(&HDFFF) 'Last surrogate pair
    Debug.Assert Parse("""new\u00A0line""").Value = "new" & ChrW$(&HA0) & "line" 'No-Break Space (nbsp)
    Debug.Assert Parse("""\u0000""").Value = vbNullChar
    Debug.Assert Parse("""\u002c""").Value = ","
    Debug.Assert Parse("""\uD834\uDd1e""").Value = ChrW$(&HD834) & ChrW$(&HDD1E)
    Debug.Assert Parse("""\u0821""").Value = ChrW$(&H821)
    Debug.Assert Parse("""\u0123""").Value = ChrW$(&H123)
    Debug.Assert Parse("""\u0061\u30af\u30EA\u30b9""").Value = ChrW$(&H61) & ChrW$(&H30AF) & ChrW$(&H30EA) & ChrW$(&H30B9)
    Debug.Assert Parse("""new\u000Aline""").Value = "new" & vbLf & "line"
    Debug.Assert Parse("""\u0022""").Value = """"
    Debug.Assert Parse("""\u005C""").Value = "\"
    Debug.Assert Parse("""\u200B""").Value = ChrW$(&H200B)
    Debug.Assert Parse("""\u2064""").Value = ChrW$(&H2064)
    Debug.Assert Parse("""\uFDD0""").Value = ChrW$(&HFDD0)
    Debug.Assert Parse("""\uFFFE""").Value = ChrW$(&HFFFE)
End Sub

Private Sub TestParseStringInvalid()
    Debug.Assert Not Parse("{""key"":1,""KEY"":2}", keyCompareMode:=vbTextCompare).IsValid
    Debug.Assert Not Parse("""""a").IsValid
    Debug.Assert Not Parse("""\uD800\""").IsValid
    Debug.Assert Not Parse("""\uD800\u""").IsValid
    Debug.Assert Not Parse("""\uD800\u1""").IsValid
    Debug.Assert Not Parse("""\uD800\u1x""").IsValid
    Debug.Assert Not Parse(ChrW$(&H17D)).IsValid
    Debug.Assert Not Parse("""a" & vbNullChar & "a""").IsValid
    Debug.Assert Not Parse("""\" & vbNullChar & """").IsValid
    Debug.Assert Not Parse("""\" & vbTab & """").IsValid
    Debug.Assert Not Parse("""\" & vbLf & """").IsValid
    Debug.Assert Not Parse("""\" & vbCr & """").IsValid
    Debug.Assert Not Parse("""\" & vbFormFeed & """").IsValid
    Debug.Assert Not Parse("""" & vbNewLine & """").IsValid
    Debug.Assert Not Parse("""" & vbTab & """").IsValid
    Debug.Assert Not Parse("""\x00""").IsValid
    Debug.Assert Not Parse("""\\\""").IsValid
    Debug.Assert Not Parse("""\?""").IsValid
    Debug.Assert Not Parse(BytesToString(dQuoteB, &H5C, &HF0, &H9F, &H8C, &H80, dQuoteB)).IsValid '\Emoji
    Debug.Assert Not Parse("""\""").IsValid
    Debug.Assert Not Parse("""\u00a""").IsValid
    Debug.Assert Not Parse("""\uD834\uDd""").IsValid
    Debug.Assert Not Parse("""\uD800\uD800\x""").IsValid
    Debug.Assert Not Parse("""\z""").IsValid
    Debug.Assert Not Parse("""\ugggg""").IsValid
    Debug.Assert Not Parse("""\" & Chr$(&HE5) & """").IsValid
    Debug.Assert Not Parse("""\u" & Chr$(&HE5) & """").IsValid
    Debug.Assert Not Parse("\u0020""asd""").IsValid
    Debug.Assert Not Parse("\n").IsValid 'No quotes
    Debug.Assert Not Parse("'""'").IsValid
    Debug.Assert Not Parse("abc").IsValid
    Debug.Assert Not Parse("""\").IsValid
    Debug.Assert Not Parse("""\UA66D""").IsValid
    Debug.Assert Not Parse(BytesToString(dQuoteB, &H5C, &HE5, dQuoteB)).IsValid
End Sub

'https://datatracker.ietf.org/doc/html/rfc8259#section-7
'\ must be escaped
'" must be escaped
'Control characters U+0000 through U+001F must be escaped
Private Sub TestParseStringInvalidEscape()
    Debug.Assert Not Parse("""\""").IsValid
    Debug.Assert Not Parse("""a\""").IsValid
    Debug.Assert Not Parse("""""""").IsValid
    Debug.Assert Not Parse("""---""---""").IsValid
    Dim i As Long
    For i = &H0 To &H1F
        Debug.Assert Not Parse("""" & Chr$(i) & """").IsValid
    Next i
End Sub

Private Sub TestParseStringLoneSurrogates()
    'Allow
    Debug.Assert Parse("""\uDFAA""", failIfLoneSurrogate:=False).Value = ChrW$(&HDFAA) 'Lone 2nd surrogate
    Debug.Assert Parse("""\uDADA""", failIfLoneSurrogate:=False).Value = ChrW$(&HDADA) 'Lone 1st surrogate
    Debug.Assert Parse("""\uD888\u1234""", failIfLoneSurrogate:=False).Value = ChrW$(&HD888) & ChrW$(&H1234) 'Invalid 2nd surrogate
    Debug.Assert Parse("""\uD800\n""", failIfLoneSurrogate:=False).Value = ChrW$(&HD800) & vbLf
    Debug.Assert Parse("""\uDd1ea""", failIfLoneSurrogate:=False).Value = ChrW$(&HDD1E) & "a"
    Debug.Assert Parse("""\uD800\uD800\n""", failIfLoneSurrogate:=False).Value = ChrW$(&HD800) & ChrW$(&HD800) & vbLf
    Debug.Assert Parse("""\uD800""", failIfLoneSurrogate:=False).Value = ChrW$(&HD800)
    Debug.Assert Parse("""\uD800abc""", failIfLoneSurrogate:=False).Value = ChrW$(&HD800) & "abc"
    Debug.Assert Parse("""\uDd1e\uD834""", failIfLoneSurrogate:=False).Value = ChrW$(&HDD1E) & ChrW$(&HD834) 'Inverted surrogates
    '
    'Fail
    Debug.Assert Not Parse("""\uDFAA""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uDADA""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD888\u1234""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD800\n""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uDd1ea""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD800\uD800\n""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD800""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uD800abc""", failIfLoneSurrogate:=True).IsValid
    Debug.Assert Not Parse("""\uDd1e\uD834""", failIfLoneSurrogate:=True).IsValid
End Sub

Private Sub TestParseStringInvalidUTF8()
    'Allow - replace each bad byte with with 0xfffd default character
    Debug.Assert Parse(BytesToString(dQuoteB, &HFF, dQuoteB), failIfInvalidByteSequence:=False).Value = ChrW$(&HFFFD) 'Default character replaces invalid byte sequence
    Debug.Assert Parse(BytesToString(dQuoteB, &HFF, &HFF, &HFF, dQuoteB)).Value = String$(3, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(dQuoteB, &H81, dQuoteB)).Value = ChrW$(&HFFFD)
    Debug.Assert Parse(BytesToString(dQuoteB, &HC0, &HAF, dQuoteB)).Value = String$(2, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(dQuoteB, &HFC, &H83, &HBF, &HBF, &HBF, &HBF, dQuoteB)).Value = String$(6, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(dQuoteB, &HFC, &H80, &H80, &H80, &H80, &H80, dQuoteB)).Value = String$(6, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(dQuoteB, &HE0, &HFF, dQuoteB)).Value = String$(2, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(dQuoteB, &HE6, &H97, &HA5, &HD1, &H88, &HFA, dQuoteB)).Value = ChrW$(&H65E5) & ChrW$(&H448) & ChrW$(&HFFFD)
#If Windows Then
    Debug.Assert Parse(BytesToString(dQuoteB, &HF4, &HBF, &HBF, &HBF, dQuoteB)).Value = String$(3, ChrW$(&HFFFD))
    Debug.Assert Parse(BytesToString(dQuoteB, &HED, &HA0, &H80, dQuoteB)).Value = String$(2, ChrW$(&HFFFD)) 'Single surrogate 0xD800
#Else
    Debug.Assert Parse(BytesToString(dQuoteB, &HF4, &HBF, &HBF, &HBF, dQuoteB)).Value = ChrW$(&HFFFD)
    Debug.Assert Parse(BytesToString(dQuoteB, &HED, &HA0, &H80, dQuoteB)).Value = ChrW$(&HFFFD)
#End If
    '
    'Fail
    Debug.Assert Not Parse(BytesToString(dQuoteB, &HFF, dQuoteB), failIfInvalidByteSequence:=True).IsValid
    
    Debug.Assert Not Parse(BytesToString(dQuoteB, &H5B, &H22, &H81, &H22, &H5D, dQuoteB), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(dQuoteB, &H81, dQuoteB), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(dQuoteB, &HC0, &HAF, dQuoteB), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(dQuoteB, &HFC, &H83, &HBF, &HBF, &HBF, &HBF, dQuoteB), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(dQuoteB, &HFC, &H80, &H80, &H80, &H80, &H80, dQuoteB), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(dQuoteB, &HE0, &HFF, dQuoteB), failIfInvalidByteSequence:=True).IsValid
#If Windows Then
    Debug.Assert Not Parse(BytesToString(dQuoteB, &HF4, &HBF, &HBF, &HBF, dQuoteB), failIfInvalidByteSequence:=True).IsValid
    Debug.Assert Not Parse(BytesToString(dQuoteB, &HED, &HA0, &H80, dQuoteB), failIfInvalidByteSequence:=True).IsValid
#Else
    Debug.Assert Parse(BytesToString(dQuoteB, &HF4, &HBF, &HBF, &HBF, dQuoteB), failIfInvalidByteSequence:=True).Value = ChrW$(&HFFFD)
    Debug.Assert Parse(BytesToString(dQuoteB, &HED, &HA0, &H80, dQuoteB), failIfInvalidByteSequence:=True).Value = ChrW$(&HFFFD)
#End If
End Sub

Public Sub RunAllJSONSerializeTests()
    TestSerializePrimitive
    TestSerializeNested
    TestSerializeIndent
    TestSerializeEscaped
    TestSerializeMisc
    TestSerializeSortKeys
    TestSerializeNonTextKeys
    TestSerializeCircularRef
    TestSerializeCodePage
    TestSerializeMultiDimensionalArrays
    Debug.Print "Finished running serializer tests at " & Now()
End Sub

Private Sub TestSerializePrimitive()
    Debug.Assert Serialize(Null) = "null"
    Debug.Assert Serialize(False) = "false"
    Debug.Assert Serialize(True) = "true"
    Debug.Assert Serialize(Array(False, Null, True, Empty)) = "[false,null,true,null]"
    Debug.Assert Serialize(CByte(123)) = "123"
    Debug.Assert Serialize(CInt(123)) = "123"
    Debug.Assert Serialize(CLng(123)) = "123"
#If x64 Then
    Debug.Assert Serialize(CLngLng(123)) = "123"
#End If
    Debug.Assert Serialize(CSng(123)) = "123"
    Debug.Assert Serialize(CDbl(123)) = "123"
    Debug.Assert Serialize(CCur(123)) = "123"
#If Windows Then
    Debug.Assert Serialize(CDec(123)) = "123"
#End If
    Debug.Assert Serialize(1E+300) = "1E+300"
    Debug.Assert Serialize(0.12345) = "0.12345"
    Debug.Assert Serialize(CSng(-1.2)) = "-1.2"
    Debug.Assert Serialize(CDbl(-1.2)) = "-1.2"
    Debug.Assert Serialize(CCur(-1.2)) = "-1.2"
#If Windows Then
    Debug.Assert Serialize(CDec(-1.2)) = "-1.2"
#End If
    Debug.Assert Serialize(-1.2E+25) = "-1.2E+25"
    Debug.Assert Serialize(4.94065645841247E-324) = "4.94065645841247E-324"
    Debug.Assert Serialize(1.79769313486231E+308) = "1.79769313486231E+308"
    Debug.Assert Serialize("abc") = """abc"""
    Debug.Assert Serialize("abc\""") = """abc\\\"""""
    Debug.Assert Serialize("") = """"""
    Debug.Assert Serialize("\uD800") = """\\uD800"""
    Debug.Assert Serialize(0) = "0"
    Debug.Assert Serialize(#4/4/2025#) = """2025-04-04 00:00:00"""
    Debug.Assert Serialize(#4/4/2025#, formatDateISO:=True) = """2025-04-04T00:00:00Z"""
    Debug.Assert Serialize(#4/4/2025# + 250.631 / 86400, formatDateISO:=True) = """2025-04-04T00:04:10.631Z"""
    Debug.Assert Serialize(CVErr(123)) = """Error 123"""
End Sub

Private Sub TestSerializeNested()
    Dim s As String
    Dim dict As Dictionary
    '
    s = "[[],[],1,2,3,[[]]]"
    Debug.Assert Serialize(Parse(s).Value) = s
    Debug.Assert Serialize(Array(Array(), Array(), 1, 2, 3, Array(Array()))) = s
    Debug.Assert Serialize(Collection(Collection(), Collection(), 1, 2, 3, Collection(Collection()))) = s
    '
    Debug.Assert Serialize(Dictionary("1", 2, "3", 4)) = "{""1"":2,""3"":4}"
    '
    s = "{""k"":{""1"":2,""3"":{""1"":2,""3"":{}}}}"
    Debug.Assert Serialize(Parse(s).Value) = s
    '
    s = "{""k"":{""1"":2,""3"":{""1"":[[],[],null,false,true,[[[]]]],""3"":{}}}}"
    Debug.Assert Serialize(Parse(s).Value, 0) = s
    '
    s = "{""k"":{""1"":[[],[],1,null,"""",3,[[[]]]],""3"":{""1"":4,""3"":{}}}}"
    Debug.Assert Serialize(Parse(s).Value) = s
    '
    Dim nestingLevel As Long
    Dim i As Long
    '
#If TWINBASIC Then 'Fast Dictionary does not manage nesting in TB (only in VB*)
                   'The built-in TB Collection leads to 'Out of stack space' in x32
    nestingLevel = 100
#Else
    If IsFastDict() Then
        nestingLevel = 10000
    Else 'Scripting or other
        nestingLevel = 100
    End If
#End If
    s = RepeatString("{""key"":[", nestingLevel) & RepeatString("]}", nestingLevel)
    Set dict = Parse(s, maxNestingDepth:=nestingLevel * 2).Value
    Debug.Assert Serialize(dict) = s
    '
    s = "[{},{},{},[],[{}],[{},[]]]"
    Debug.Assert Serialize(Parse(s).Value) = s
End Sub
    
Private Sub TestSerializeIndent()
    Dim coll As Collection
    Dim dict As Dictionary
    Dim s As String
    '
    Set coll = Collection(Collection(), Collection(), 1, 2, 3, Collection(Collection()))
    Debug.Assert Serialize(coll, 0) = "[[],[],1,2,3,[[]]]"
    Debug.Assert Serialize(coll, 2) = Join(Array("[" _
                                               , "  []," _
                                               , "  []," _
                                               , "  1," _
                                               , "  2," _
                                               , "  3," _
                                               , "  [" _
                                               , "    []" _
                                               , "  ]" _
                                               , "]") _
                                         , vbNewLine)
    Debug.Assert Serialize(coll, 4) = Join(Array("[" _
                                               , "    []," _
                                               , "    []," _
                                               , "    1," _
                                               , "    2," _
                                               , "    3," _
                                               , "    [" _
                                               , "        []" _
                                               , "    ]" _
                                               , "]" _
                                         ), vbNewLine)
    '
    s = "{""k"":{""1"":[[],[],1,null,"""",3,[[[]]]],""3"":{""1"":4,""3"":{}}}}"
    Debug.Assert Serialize(Parse(s).Value, 2) = Join(Array("{" _
                                                        , "  ""k"": {" _
                                                        , "    ""1"": [" _
                                                        , "      []," _
                                                        , "      []," _
                                                        , "      1," _
                                                        , "      null," _
                                                        , "      """"," _
                                                        , "      3," _
                                                        , "      [" _
                                                        , "        [" _
                                                        , "          []" _
                                                        , "        ]" _
                                                        , "      ]" _
                                                        , "    ]," _
                                                        , "    ""3"": {" _
                                                        , "      ""1"": 4," _
                                                        , "      ""3"": {}" _
                                                        , "    }" _
                                                        , "  }" _
                                                        , "}" _
                                                   ), vbNewLine)
    '
    s = "[{""_id"":""67ef9f1f7469f832882a1e2c"",""friends"":[{""id"":0,""name"":""Wendi Perkins""}" _
      & "],""favoriteFruit"":""banana""}]"
    Debug.Assert Serialize(Parse(s).Value, 1) = Join(Array("[" _
                                                         , " {" _
                                                         , "  ""_id"": ""67ef9f1f7469f832882a1e2c""," _
                                                         , "  ""friends"": [" _
                                                         , "   {" _
                                                         , "    ""id"": 0," _
                                                         , "    ""name"": ""Wendi Perkins""" _
                                                         , "   }" _
                                                         , "  ]," _
                                                         , "  ""favoriteFruit"": ""banana""" _
                                                         , " }" _
                                                         , "]" _
                                                    ), vbNewLine)
    Debug.Assert Serialize(Parse(s).Value, 0) = s
End Sub

Private Sub TestSerializeEscaped()
    Debug.Assert Serialize("\") = """\\"""
    Debug.Assert Serialize("/") = """/"""
    Debug.Assert Serialize("""") = """\"""""
    '
    Dim i As Long
    Dim res As String
    '
    For i = 0 To 31 'Control Characters
        res = Serialize(Chr$(i))
        Select Case i
            Case Asc(vbBack):     Debug.Assert res = """\b"""
            Case Asc(vbTab):      Debug.Assert res = """\t"""
            Case Asc(vbCr):       Debug.Assert res = """\r"""
            Case Asc(vbFormFeed): Debug.Assert res = """\f"""
            Case Asc(vbLf):       Debug.Assert res = """\n"""
            Case Is < 16:         Debug.Assert res = """\u000" & Hex(i) & """"
            Case Else:            Debug.Assert res = """\u00" & Hex(i) & """"
        End Select
    Next i
    '
    Dim s As String
    '
    s = ChrW(2353) & ChrW(235) & ChrW(-23533) & ChrW(23533)
    Debug.Assert Serialize(s, escapeNonASCII:=True) = """\u0931\u00EB\uA413\u5BED"""
    Debug.Assert Serialize(s, escapeNonASCII:=False) = """" & s & """"
End Sub

Private Sub TestSerializeMisc()
    Debug.Assert Serialize(Array(Empty, Nothing, Err _
                               , PosInf, NegInf, SNaN, QNaN)) = "[null,null,null,null,null,null,null]"
End Sub

Private Sub TestSerializeSortKeys()
    Dim d As Dictionary
    Set d = Dictionary("d", 1, 7, 2, "a", 3, "b", 4, New Collection, 5)
    '
    Debug.Assert Serialize(d) = "{""d"":1,""a"":3,""b"":4}"
    Debug.Assert Serialize(d, sortKeys:=True) = "{""a"":3,""b"":4,""d"":1}"
    Debug.Assert Serialize(d, sortKeys:=True, forceKeysToText:=True) = "{""7"":2,""a"":3,""b"":4,""d"":1}"
End Sub
    
Private Sub TestSerializeNonTextKeys()
    Dim s As String
    Dim dict As Dictionary
    '
    s = "{""k"":{""1"":[[],[],1,2,3,[[[]]]],""3"":{""1"":4,""3"":{}}}}"
    Set dict = Parse(s).Value
    Debug.Assert Serialize(dict) = s
    '
    dict("k").Key("1") = 1
    Debug.Assert Serialize(dict, forceKeysToText:=True) = s
    Debug.Assert Serialize(dict, forceKeysToText:=False) = "{""k"":{""3"":{""1"":4,""3"":{}}}}"
    '
    dict("k").Key("3") = 3
    Debug.Assert Serialize(dict, forceKeysToText:=True) = s
    Debug.Assert Serialize(dict, forceKeysToText:=False) = "{""k"":{}}"
    Debug.Assert Serialize(dict, failIfNonTextKeys:=True) = vbNullString
    Debug.Assert Serialize(dict, forceKeysToText:=True, failIfNonTextKeys:=True) = s
    '
    Set dict = Dictionary(#4/4/2025#, 1, False, 2, True, 3, Null, 4, Nothing, 5, Collection(1, 2, 3), 6)
    Debug.Assert Serialize(dict) = "{}"
    Debug.Assert Serialize(dict, forceKeysToText:=True) = "{""2025-04-04 00:00:00"":1,""false"":2,""true"":3}"
    Debug.Assert Serialize(dict, forceKeysToText:=True, formatDateISO:=True) = "{""2025-04-04T00:00:00Z"":1,""false"":2,""true"":3}"
    Debug.Assert Serialize(dict, failIfNonTextKeys:=True) = vbNullString
    Debug.Assert Serialize(dict, forceKeysToText:=True, failIfNonTextKeys:=True) = vbNullString
    '
    Set dict = Dictionary(CVErr(123), 123, 123.123, 123.123)
    Debug.Assert Serialize(dict, forceKeysToText:=True) = "{""Error 123"":123,""123.123"":123.123}"
End Sub
    
Private Sub TestSerializeCircularRef()
    Dim s As String
    Dim dict As Dictionary
    '
    s = "{""k"":{""1"":[[],[],1,2,3,[[[]]]],""3"":{""1"":4,""3"":{}}}}"
    Set dict = Parse(s).Value
    Debug.Assert Serialize(dict) = s
    '
    Set dict("k") = dict
    Debug.Assert Serialize(dict) = "{""k"":null}"
    '
    Set dict = Parse(s).Value
    Set dict("k")("1") = dict
    Debug.Assert Serialize(dict) = "{""k"":{""1"":null,""3"":{""1"":4,""3"":{}}}}"
    Set dict("k")("3") = dict
    Debug.Assert Serialize(dict) = "{""k"":{""1"":null,""3"":null}}"
    dict.Add "n", dict("k")
    Debug.Assert Serialize(dict) = "{""k"":{""1"":null,""3"":null},""n"":{""1"":null,""3"":null}}"
    Set dict("k")("1") = New Collection
    dict.Remove "n"
    Debug.Assert Serialize(dict) = "{""k"":{""1"":[],""3"":null}}"
    Set dict("k")("3") = dict("k")("1")
    Debug.Assert Serialize(dict) = "{""k"":{""1"":[],""3"":[]}}"
    dict("k")("1").Add dict("k")("1")
    Debug.Assert Serialize(dict) = "{""k"":{""1"":[null],""3"":[null]}}"
    dict("k")("1").Add dict
    Debug.Assert Serialize(dict) = "{""k"":{""1"":[null,null],""3"":[null,null]}}"
    '
    Debug.Assert Serialize(dict, failIfCircularRef:=True) = vbNullString
    Serialize dict, failIfCircularRef:=True, outError:=s
    Debug.Assert LenB(s) > 0 'outError = "Circular reference detected"
End Sub
    
Private Sub TestSerializeCodePage()
    Dim s As String
    Dim dict As Dictionary
    '
    s = "{""k"":{""1"":[[],[],1,2,3,[[[]]]],""3"":{""1"":4,""3"":{}}}}"
    Set dict = Parse(s).Value
    Debug.Assert Serialize(dict) = s
    Debug.Assert Serialize(dict, jpCode:=jpCodeUTF8) = StrConv(s, vbFromUnicode)
    '
    Debug.Assert Serialize(ChrW$(&HFFFD)) = """\uFFFD"""
    Debug.Assert Serialize(ChrW$(&HFFFD), jpCode:=jpCodeUTF8) = StrConv("""\uFFFD""", vbFromUnicode)
    '
    Debug.Assert Serialize(ChrW$(&HFF)) = """\u00FF"""
    Debug.Assert Serialize(ChrW$(&HFF), jpCode:=jpCodeUTF8) = StrConv("""\u00FF""", vbFromUnicode)
    Debug.Assert Serialize(ChrW$(&HFF), escapeNonASCII:=False) = """" & ChrW$(&HFF) & """"
    Debug.Assert Serialize(ChrW$(&HFF), escapeNonASCII:=False, jpCode:=jpCodeUTF8) = ChrW$(&HC322) & ChrW$(&H22BF) '0x22 being "
    '
    'Escaped lone surrogates
    Debug.Assert Serialize(ChrW$(&HDADA)) = """\uDADA""" 'Lone 1st surrogate
    Debug.Assert Serialize(ChrW$(&HDFAA)) = """\uDFAA""" 'Lone 2nd surrogate
    Debug.Assert Serialize(ChrW$(&HD888) & ChrW$(&H1234)) = """\uD888\u1234"""
    '
    'failIfInvalidCharacter does nothing if no conversion
    Debug.Assert Serialize(ChrW$(&HDADA), failIfInvalidCharacter:=True) = """\uDADA"""
    Debug.Assert Serialize(ChrW$(&HDFAA), failIfInvalidCharacter:=True) = """\uDFAA"""
    Debug.Assert Serialize(ChrW$(&HD888) & ChrW$(&H1234), failIfInvalidCharacter:=True) = """\uD888\u1234"""
    '
    'failIfInvalidCharacter does nothing even after conversion because nonAscii is escaped
    Debug.Assert Serialize(ChrW$(&HDADA), failIfInvalidCharacter:=True, jpCode:=jpCodeUTF8) = StrConv("""\uDADA""", vbFromUnicode)
    Debug.Assert Serialize(ChrW$(&HDFAA), failIfInvalidCharacter:=True, jpCode:=jpCodeUTF8) = StrConv("""\uDFAA""", vbFromUnicode)
    Debug.Assert Serialize(ChrW$(&HD888) & ChrW$(&H1234), failIfInvalidCharacter:=True, jpCode:=jpCodeUTF8) = StrConv("""\uD888\u1234""", vbFromUnicode)
    '
    'Unescaped lone surrogates - failIfInvalidCharacter does nothing if no conversion
    Debug.Assert Serialize(ChrW$(&HDADA), escapeNonASCII:=False, failIfInvalidCharacter:=True) = """" & ChrW$(&HDADA) & """"
    Debug.Assert Serialize(ChrW$(&HDFAA), escapeNonASCII:=False, failIfInvalidCharacter:=True) = """" & ChrW$(&HDFAA) & """"
    Debug.Assert Serialize(ChrW$(&HD888) & ChrW$(&H1234), escapeNonASCII:=False, failIfInvalidCharacter:=True) = """" & ChrW$(&HD888) & ChrW$(&H1234) & """"
    '
    'Fails
    Debug.Assert Serialize(ChrW$(&HDADA), escapeNonASCII:=False, failIfInvalidCharacter:=True, jpCode:=jpCodeUTF8) = vbNullString
    Debug.Assert Serialize(ChrW$(&HDFAA), escapeNonASCII:=False, failIfInvalidCharacter:=True, jpCode:=jpCodeUTF8) = vbNullString
    Debug.Assert Serialize(ChrW$(&HD888) & ChrW$(&H1234), escapeNonASCII:=False, failIfInvalidCharacter:=True, jpCode:=jpCodeUTF8) = vbNullString
    '
    'Uses replacement 0xFFFD
    Debug.Assert Serialize(ChrW$(&HDADA), escapeNonASCII:=False, jpCode:=jpCodeUTF8) = ChrW$(&HEF22) & ChrW$(&HBDBF) & ChrB(&H22) '0x22 being "
    Debug.Assert Serialize(ChrW$(&HDFAA), escapeNonASCII:=False, jpCode:=jpCodeUTF8) = ChrW$(&HEF22) & ChrW$(&HBDBF) & ChrB(&H22)
    Debug.Assert Serialize(ChrW$(&HD888) & ChrW$(&H1234), escapeNonASCII:=False, jpCode:=jpCodeUTF8) = ChrW$(&HEF22) & ChrW$(&HBDBF) & ChrW$(&H88E1) & ChrW$(&H22B4)
End Sub

Private Sub TestSerializeMultiDimensionalArrays()
    Dim arr2D(1 To 3, 1 To 2) As Long
    Dim arr3D(1 To 2, 1 To 2, 1 To 4) As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    '
    'Populate a 2 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr2D, 1) To UBound(arr2D, 1)
        For j = LBound(arr2D, 2) To UBound(arr2D, 2)
            n = n + 1
            arr2D(i, j) = n
        Next j
    Next i
    '
    'Populate a 3 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr3D, 1) To UBound(arr3D, 1)
        For j = LBound(arr3D, 2) To UBound(arr3D, 2)
            For k = LBound(arr3D, 3) To UBound(arr3D, 3)
                n = n + 1
                arr3D(i, j, k) = n
            Next k
        Next j
    Next i
    '
    Debug.Assert Serialize(arr2D) = "[[1,2],[3,4],[5,6]]"
    Debug.Assert Serialize(arr3D) = "[[[1,2,3,4],[5,6,7,8]],[[9,10,11,12],[13,14,15,16]]]"
End Sub

'*******************************************************************************
'Utilities
'*******************************************************************************
Private Function AreEqual(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean
    On Error GoTo ErrorHandler
    'No need to check for vbDataObject
    If IsObject(v1) Then
        If Not IsObject(v2) Then Exit Function
        '
        If TypeOf v1 Is Collection Then
            If Not TypeOf v2 Is Collection Then Exit Function
            If v1.Count <> v2.Count Then Exit Function
            '
            Dim i As Long
            For i = 1 To v1.Count
                If Not AreEqual(v1(i), v2(i)) Then Exit Function
            Next i
            AreEqual = True
        ElseIf TypeOf v2 Is Dictionary Then
            If Not TypeOf v2 Is Dictionary Then Exit Function
            If v1.Count <> v2.Count Then Exit Function
            '
            Dim k As Variant
            Dim adk As Boolean
            '
            If IsFastDict() Then adk = v1.AllowDuplicateKeys Else adk = False
            If adk Then
                Dim k1() As Variant: k1 = v1.Keys
                Dim k2() As Variant: k2 = v2.Keys
                Dim i1() As Variant: i1 = v1.Items
                Dim i2() As Variant: i2 = v2.Items
                '
                For i = 0 To v1.Count - 1
                    If Not AreEqual(k1(i), k2(i)) Then Exit Function
                    If Not AreEqual(i1(i), i2(i)) Then Exit Function
                Next i
            Else
                For Each k In v1
                    If Not AreEqual(v1(k), v2(k)) Then Exit Function
                Next k
            End If
            AreEqual = True
        Else
            AreEqual = (TypeName(v1) = TypeName(v2))
        End If
        Exit Function
    End If
    If IsObject(v2) Then Exit Function
    '
    Dim vt1 As VbVarType: vt1 = VarType(v1)
    Dim vt2 As VbVarType: vt2 = VarType(v2)
    '
    If vt1 < vt2 Then 'Allow number comparison across types
        #If (x32) Or Mac Then
            Const vbLongLong = 20
        #End If
        '
        If vt1 <= vbNull Or vt1 > vbLongLong Then Exit Function
        If vt2 <= vbNull Or vt2 > vbLongLong Then Exit Function
        If vt1 = vbString Or vt1 = vbError Then Exit Function
        If vt2 = vbString Or vt2 = vbError Then Exit Function
    End If
    If IsNull(v1) Then
        AreEqual = True
    Else
        AreEqual = (v1 = v2)
    End If
ErrorHandler:
End Function
Private Function Collection(ParamArray Values() As Variant) As Collection
    Dim v As Variant
    Dim coll As Collection
    '
    Set coll = New Collection
    For Each v In Values
        coll.Add v
    Next v
    Set Collection = coll
End Function
Private Function Dictionary(ParamArray Values() As Variant) As Dictionary
    Dim i As Long
    Dim dict As Dictionary
    '
    Set dict = New Dictionary
    For i = 0 To UBound(Values) Step 2
        dict.Add Values(i), Values(i + 1)
    Next i
    Set Dictionary = dict
End Function
Private Function RepeatString(ByRef str As String _
                            , ByVal repeatTimes As Long) As String
    If repeatTimes <= 0 Or LenB(str) = 0 Then Exit Function
    '
    Dim newLength As Long: newLength = LenB(str) * repeatTimes
    RepeatString = Space$((newLength + 1) \ 2)
    If newLength Mod 2 = 1 Then RepeatString = MidB$(RepeatString, 2)
    '
    MidB$(RepeatString, 1) = str
    If repeatTimes > 1 Then MidB$(RepeatString, LenB(str) + 1) = RepeatString
End Function
Private Function IsFastDict() As Boolean
    Dim o As Object:     Set o = New Dictionary
    On Error Resume Next
    Dim s As Single:     s = o.LoadFactor
    Dim b As Boolean:    b = o.AllowDuplicateKeys
    Dim d As Dictionary: Set d = o.Self.Factory
    IsFastDict = (Err.Number = 0)
    On Error GoTo 0
End Function
Private Property Get PosInf() As Double 'IEEE754 +inf
    On Error Resume Next
    PosInf = 1 / 0
    On Error GoTo 0
End Property
Private Property Get SNaN() As Double 'IEEE754 signaling NaN (sNaN)
    On Error Resume Next
    SNaN = 0 / 0
    On Error GoTo 0
End Property
Private Property Get NegInf() As Double 'IEEE754 -inf
    NegInf = -PosInf
End Property
Private Property Get QNaN() As Double 'IEEE754 quiet NaN (qNaN)
    QNaN = -SNaN
End Property
'Useful for building UTF8 strings
Private Function BytesToString(ParamArray bytes() As Variant) As String
    Dim i As Long
    Dim b() As Byte: ReDim b(0 To UBound(bytes))
    For i = 0 To UBound(bytes)
        b(i) = bytes(i)
    Next i
    BytesToString = b
End Function
Private Function ToLocaleDot(ByVal s As String) As String
    Dim chDot As String: chDot = Mid$(CStr(1.1), 2, 1)
    ToLocaleDot = Replace(s, ".", chDot)
End Function
