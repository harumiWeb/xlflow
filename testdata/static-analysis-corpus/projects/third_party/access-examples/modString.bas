Attribute VB_Name = "modString"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              String Manipulation/Parsing
'
' Ver.  Date            Author              Details
' 1.00  01 JAN 2001     Microsoft           Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Function LowerCC()

' Converts the current control to lower case
  Screen.ActiveControl = LCase(Screen.ActiveControl)
  
End Function

Public Sub ParseCSZ(ByVal s As String, City As String, State As String, Zip As String)
'
' Parses address "New York NY 00123" into separate fields.
' Supports the following formats:
'   New York NY 12345-9876
'   Pierre, North Dakota 45678-7654
'   San Diego, CA, 98765-4321
'
' Words are extracted in the following order if no commas are found to delimit the values:
'   Zip, State, City
'
Dim P As Integer
'
' Check for comma after city name
'
  P = InStr(s, ",")
  If P > 0 Then
    City = Trim$(Left$(s, P - 1))
    s = Trim$(Mid$(s, P + 1))
'
'   Check for comma after state
'
    P = InStr(s, ",")
    If P > 0 Then
      State = Trim$(Left$(s, P - 1))
      Zip = Trim$(Mid$(s, P + 1))
    Else                           ' No comma between state and zip
      Zip = CutLastWord(s, s)
      State = s
    End If
  Else                             ' No commas between city, state, or zip
    Zip = CutLastWord(s, s)
    State = CutLastWord(s, s)
    City = s
  End If
'
' Clean up any dangling commas
'
  If Right$(State, 1) = "," Then
    State = Left$(State, Len(State) - 1)
  End If
  If Right$(City, 1) = "," Then
    City = Left$(City, Len(City) - 1)
  End If
End Sub

Public Sub ParseName(ByVal s As String, TITLE As String, FName As String, MName As String, LName As String, Pedigree As String, Degree As String)
'
' Parses name "Mr. Bill A. Jones III, PhD" into separate fields.
' Words are extracted in the following order: Title, Degree, Pedigree, LName, FName, MName
' Assumes Pedigree is not preceded by a comma, or else it will end up with the Degree(s).
'
Dim Word As String, P As Integer, Found As Integer
Const Titles = "Mr.Mrs.Ms.Dr.Mme.Mssr.Mister,Miss,Doctor,Sir,Lord,Lady,Madam,Mayor,President"
Const Pedigrees = "Jr.Sr.III,IV,VIII,IX,XIII"
  TITLE = ""
  FName = ""
  MName = ""
  LName = ""
  Pedigree = ""
  Degree = ""
'
' Get Title
'
 ' Word = CutWord(S, S)
  If InStr(Titles, Word) Then
    TITLE = Word
  Else
    s = Word & " " & s
  End If
'
' Get Degree
'
  P = InStr(s, ",")
  If P > 0 Then
    Degree = Trim$(Mid$(s, P + 1))
    s = Trim$(Left$(s, P - 1))
  End If
'
' Get Pedigree
'
  Word = CutLastWord(s, s)
  If InStr(Pedigrees, Word) Then
    Pedigree = Word
  Else
    s = s & " " & Word
  End If
'
' Get Last Name
'
  LName = CutLastWord(s, s)
'
' Get First Name
'
'  fName = CutWord(S, S)
'
' Get Middle Name(s)
'
  MName = Trim(s)
End Sub

Public Function Proper(x)
'  Capitalize first letter of every word in a field.
'  Use in an event procedure in AfterUpdate of control;
'  for example, [Last Name] = Proper([Last Name]).
'  Names such as O'Brien and Wilson-Smythe are properly capitalized,
'  but MacDonald is changed to Macdonald, and van Buren to Van Buren.
'  Note: For this function to work correctly, you must specify
'  Option Compare Database in the Declarations section of this module.
'
'  See Also: StrConv Function in the Microsoft Access 97 online Help.

Dim Temp$, c$, OldC$, i As Integer
  If IsNull(x) Then
    Exit Function
  Else
    Temp$ = CStr(LCase(x))
    '  Initialize OldC$ to a single space because first
    '  letter needs to be capitalized but has no preceding letter.
    OldC$ = " "
    For i = 1 To Len(Temp$)
      c$ = Mid$(Temp$, i, 1)
      If c$ >= "a" And c$ <= "z" And (OldC$ < "a" Or OldC$ > "z") Then
        Mid$(Temp$, i, 1) = UCase$(c$)
      End If
      OldC$ = c$
    Next i
    Proper = Temp$
  End If
End Function

Public Function ProperEx(var As Variant) As Variant
' Purpose: Convert the case of var so that the first letter of each word capitalized.

   Dim strV As String, intChar As Integer, i As Integer
   Dim fWasSpace As Integer    'Flag: was previous char a space?
   
   If IsNull(var) Then Exit Function
   strV = var
   fWasSpace = True              'Initialize to capitalize first letter.
   For i = 1 To Len(strV)
      intChar = Asc(Mid$(strV, i, 1))
      Select Case intChar
      Case 65 To 90              ' A to Z
         If Not fWasSpace Then Mid$(strV, i, 1) = Chr$(intChar Or &H20)
      Case 97 To 122             ' a to z
         If fWasSpace Then Mid$(strV, i, 1) = Chr$(intChar And &HDF)
      End Select
      fWasSpace = (intChar = 32)
   Next
   'Proper = strV
End Function

Public Function ProperCC()
' Applies the Proper function to the current control
  Screen.ActiveControl = Proper(Screen.ActiveControl)
End Function

Public Function ProperWord(n)
'
' Assumes N contains a single word
' N: can be null
'
  ProperWord = UCase(Left(Trim(n), 1)) & LCase(Mid(Trim(n), 2))
End Function

Public Function StrToHex(s As Variant) As Variant
'
' Converts a string to a series of hexadecimal digits.
' Useful if you want a true ASCII sort in your query.
'
' StrToHex(Chr(9) & "A~") returns "09417E"
'
Dim Temp As String, i As Integer
  If VarType(s) <> 8 Then
    StrToHex = s
  Else
    Temp = ""
    For i = 1 To Len(s)
      Temp = Temp & Format(Hex(Asc(Mid(s, i, 1))), "00")
    Next i
    StrToHex = Temp
  End If
End Function

Public Sub TestParseName()
'
'  Use this procedure to test the ParseName and ParseCSZ procedures
'
  Dim n As String, t As String, f As String, M As String, L As String, P As String, d As String
  n = "Dr. James George William Joyce-Brothers IV, MS, PhD"
  ParseName n, t, f, M, L, P, d
  Debug.Print t, f, M, L, P, d
  n = "New York NY 45678-9876"
  ParseCSZ n, t, f, M
  Debug.Print t, f, M
  
End Sub

Public Function UpperCC()

' Converts the current control to upper case
  Screen.ActiveControl = UCase(Screen.ActiveControl)
  
End Function

Public Function CountCSVWords(ByVal pvalTemp As Variant) As Integer
'---------------------------------------------------------------------------
'Purpose:           Counts words in a string separated by commas.
'---------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim intWordCnt  As Integer
Dim intPosition As Integer
    
    If VarType(pvalTemp) <> 8 Or Len(pvalTemp) = 0 Then
        CountCSVWords = 0
        Exit Function
    End If
    intWordCnt = 1
    intPosition = InStr(pvalTemp, ",")
    Do While intPosition > 0
        intWordCnt = intWordCnt + 1
        intPosition = InStr(intPosition + 1, pvalTemp, ",")
    Loop
    CountCSVWords = intWordCnt

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("CountCSVWords", "modString", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function CountWords(ByVal pvarTemp As Variant) As Integer
'---------------------------------------------------------------------------
'Purpose:           Counts words in a string separated by 1 or more spaces
'---------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim intWordCnt  As Integer
Dim intLoop     As Integer
Dim blnSpace    As Boolean

    If VarType(pvarTemp) <> 8 Or Len(Trim(pvarTemp)) = 0 Then
        CountWords = 0
        Exit Function
    End If
    intWordCnt = 0
    blnSpace = True
    For intLoop = 1 To Len(pvarTemp)
        If Mid(pvarTemp, intLoop, 1) = " " Then
            blnSpace = True
        Else
            If blnSpace Then
                blnSpace = False
                intWordCnt = intWordCnt + 1
            End If
        End If
    Next intLoop
    CountWords = intWordCnt

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("CountWords", "modString", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function CutFirstWord( _
ByVal pvarWord As Variant, _
ByRef pvarRemainder As Variant, _
Optional ByVal pstrSearch As String = " ") As Variant
'---------------------------------------------------------------------------
'Purpose:           CutFirstWord:returns the first word in pvarWord.  varRemainder:returns the rest
'---------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim varTemp     As Variant
Dim intP        As Integer

    varTemp = Trim(pvarWord)
    intP = InStr(varTemp, pstrSearch)
    If intP = 0 Then
        CutFirstWord = varTemp
        pvarRemainder = Null
    Else
        CutFirstWord = Left(varTemp, intP - 1)
        pvarRemainder = Trim(Mid(varTemp, intP + 1))
    End If

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("CutFirstWord", "modString", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function CutLastWord( _
ByVal pvarWord As Variant, _
ByRef pvarRemainder As Variant, _
Optional ByVal pstrSearch As String = " ") As Variant
'---------------------------------------------------------------------------
'Purpose:           CutLastWord:returns the first word in pvarWord.  varRemainder:returns the rest
'---------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim varTemp     As Variant
Dim intP        As Integer
Dim intX        As Integer

    varTemp = Trim(pvarWord)
    intP = 1
    For intX = Len(varTemp) To 1 Step -1
        If Mid(varTemp, intX, 1) = pstrSearch Then
            intP = intX + 1
            Exit For
        End If
    Next intX
    If intP = 1 Then
        CutLastWord = varTemp
        pvarRemainder = Null
    Else
        CutLastWord = Mid(varTemp, intP)
        pvarRemainder = Trim(Left(varTemp, intP - 1))
    End If

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("CutLastWord", "modString", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function ReplaceStr( _
ByVal pvarTextIn As Variant, _
ByVal pvarSearchStr As String, _
ByVal pvarReplacement As String, _
Optional ByVal pintCompMode As Integer = vbBinaryCompare)
'---------------------------------------------------------------------------
'Purpose:           Replaces the pvarSearchStr string with pvarReplacement string in the pvarTextIn string.
'                   Uses pintCompMode to determine comparison mode as an option
'---------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim strWorkText As String
Dim intPointer As Integer

    If IsNull(pvarTextIn) Then
        ReplaceStr = Null
    Else
        strWorkText = pvarTextIn
        intPointer = InStr(1, strWorkText, pvarSearchStr, pintCompMode)
        Do While intPointer > 0
            strWorkText = Left(strWorkText, intPointer - 1) & pvarReplacement & Mid(strWorkText, intPointer + Len(pvarSearchStr))
            intPointer = InStr(intPointer + Len(pvarReplacement), strWorkText, pvarSearchStr, pintCompMode)
        Loop
        ReplaceStr = strWorkText
    End If

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("ReplaceStr", "modString", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function GetCSVWord(ByVal pvarTemp As Variant, ByVal pintIndex As Integer) As Variant
'---------------------------------------------------------------------------
'Purpose:           Returns the <Indx>th word from a comma-separated string.
'                   For example, GetCSVWord("Nancy, Bob", 2) returns Bob.
'---------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim intWordCnt  As Integer
Dim intCount    As Integer
Dim intSPos     As Integer
Dim intEPos     As Integer

    intWordCnt = CountCSVWords(pvarTemp)
    If pintIndex < 1 Or pintIndex > intWordCnt Then
        GetCSVWord = Null
        Exit Function
    End If
    
    intCount = 1
    intSPos = 1
    
    For intCount = 2 To pintIndex
        intSPos = InStr(intSPos, pvarTemp, ",") + 1
    Next intCount
    
    intEPos = InStr(intSPos, pvarTemp, ",") - 1
    If intEPos <= 0 Then intEPos = Len(pvarTemp)
    GetCSVWord = Mid(pvarTemp, intSPos, intEPos - intSPos + 1)

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("ReplaceStr", "modString", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function GetWord(ByVal pvarTemp As Variant, ByVal pintIndex As Integer) As Variant
'---------------------------------------------------------------------------
'Purpose:           Extracts a word in text where words are separated by 1 or more spaces.
'---------------------------------------------------------------------------
On Error GoTo ErrTrap
Dim intLoop     As Integer
Dim intWordCnt  As Integer
Dim intCount    As Integer
Dim intSPos     As Integer
Dim intEPos     As Integer
Dim blnSpace    As Boolean

    intWordCnt = CountWords(pvarTemp)
    If pintIndex < 1 Or pintIndex > intWordCnt Then
        GetWord = Null
        Exit Function
    End If
    intCount = 0
    blnSpace = True
    For intLoop = 1 To Len(pvarTemp)
        If Mid(pvarTemp, intLoop, 1) = " " Then
            blnSpace = True
        Else
            If blnSpace Then
                blnSpace = False
                intCount = intCount + 1
                If intCount = pintIndex Then
                    intSPos = intLoop
                    Exit For
                End If
            End If
        End If
    Next intLoop
    intEPos = InStr(intSPos, pvarTemp, " ") - 1
    If intEPos <= 0 Then intEPos = Len(pvarTemp)
    GetWord = Mid(pvarTemp, intSPos, intEPos - intSPos + 1)

ExitProcedure:
    On Error Resume Next
    
ErrTrap:
    Select Case Err.Number
        Case Is <> 0
            Call ErrorMsg("GetWord", "modString", Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Function

Public Function Like2(ByVal Text As String, ByVal Mask As String) As Integer
'
' This function does simple pattern matching.
' It allows the following wildcards:
'   # (digit)
'   ? (any character)
'   @ (alpha)
'
Dim Match As Integer, i As Integer, c As String * 1, mc As String * 1
  If Len(Text) <> Len(Mask) Then
    Match = False
  Else
    Match = True
    For i = 1 To Len(Mask)
      c = Mid(Text, i, 1)
      mc = Mid(Mask, i, 1)
      Select Case mc
        Case "#"  ' Match digit
          If c < "0" Or c > "9" Then
            Match = False
            Exit For
          End If
        Case "@"  ' Match A-Z
          If Not (c >= "A" And c <= "Z") And Not (c >= "a" And c <= "z") Then
            Match = False
            Exit For
          End If
        Case "?"  ' Match anything
        Case Else ' Exact match
          If c <> mc Then
            Match = False
            Exit For
          End If
      End Select
    Next i
  End If
  Like2 = Match
  
End Function

Public Function LPad(s, ByVal c As String, n As Integer) As String
'
' Adds character C to the left of S to make it right-justified
'
  If Len(c) = 0 Then c = " "
  If n < 1 Then
    LPad = ""
  Else
    LPad = Right$(String$(n, Left$(c, 1)) & s, n)
  End If
  
End Function

Public Function ParseItemsToArray(ByVal s As String, a() As String, ByVal Delim As String, ByVal Compare As Integer) As Integer
'
' Function parses S, using delimiter Delim, and copies each element into array A().
' The function returns the number of items copied.
'
' Compare:
'   0 = Binary comparison     - can search for Tabs (chr$(9))
'   1 = Text comparison       - can't search for Tabs
'   2 = Database comparison   - can't search for Tabs
'
' Calling convention:
'   ReDim Items(20)
'   ItemCount = ParseItemsToArray("A,B,C",Items(),Delim)
'
Dim P As Integer, i As Integer
'
' Check for valid delimiter
'
  If Delim = "" Then
    ParseItemsToArray = -1
    Exit Function
  End If
'
' Copy Items
'
  i = 0
  P = InStr(1, s, Delim, Compare)
  Do While P > 0
    a(LBound(a) + i) = Left$(s, P - 1)
    i = i + 1
    s = Mid$(s, P + 1)
    P = InStr(1, s, Delim, Compare)
  Loop
'
' Copy Last Item
'
  a(LBound(a) + i) = s
  i = i + 1
'
  ParseItemsToArray = i
  
End Function

Public Function RPad( _
ByVal pvarTemp As Variant, _
ByVal pstrC As String, _
ByVal pintN As Integer) As String
'
' Adds character pstrC to the right of pvarTemp to make it left-justified.
'
    If Len(pstrC) = 0 Then pstrC = " "
    If pintN < 1 Then
        RPad = ""
    Else
        RPad = Left(pvarTemp & String(pintN, Left(pstrC, 1)), pintN)
    End If
  
End Function

Public Sub TestParseItems()
'
' Use this procedure to test the ParseItemsToArray procedure
'
Dim ItemCount As Integer
Dim i As Integer
ReDim ItemArray(1 To 20) As String

    ItemCount = ParseItemsToArray("A,B,C", ItemArray(), ",", 0)
    For i = LBound(ItemArray) To LBound(ItemArray) + ItemCount - 1
        Debug.Print ItemArray(i)
    Next i
  
End Sub

Public Function ParseArticle( _
strOldTitle As String, _
Optional varKeepArticle As Variant) As String
'
' Removes articles (a, an, the) from the beginning of a string.
' If you specify TRUE for the varKeepArticle argument, the article is
' moved to the end of the string.
' ParseArticle("The Beatles") returns "Beatles."
' ParseArticle("The Beatles", True) returns "Beatles, The."
'
On Error GoTo ErrTrap
Dim intLength As Integer
Dim strArticle As String

    If IsMissing(varKeepArticle) Then
       varKeepArticle = False
    End If
    intLength = Len(strOldTitle)
    strArticle = ""
     
    ' Check Value for preceding article ("a", "an", or "the").
    If Left(strOldTitle, 2) = "a " Then
       strArticle = ", " & Left(strOldTitle, 1)
       strOldTitle = Right(strOldTitle, intLength - 2)
    
    ElseIf Left(strOldTitle, 3) = "an " Then
       strArticle = ", " & Left(strOldTitle, 2)
       strOldTitle = Right(strOldTitle, intLength - 3)
    ElseIf Left(strOldTitle, 4) = "the " Then
       strArticle = ", " & Left(strOldTitle, 3)
       strOldTitle = Right(strOldTitle, intLength - 4)
    End If
       
    ' If varKeepArticle is TRUE, then add the article string to the end.
    If varKeepArticle Then
       ParseArticle = strOldTitle & strArticle
    Else
       ParseArticle = strOldTitle
    End If

ExitProcedure:
    On Error Resume Next
    Exit Function

ErrTrap:
    ParseArticle = "#Error"

End Function
