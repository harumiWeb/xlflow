Attribute VB_Name = "Benchmarking"
Option Explicit
Option Private Module

Private Enum Operation
    opAdd = 0
    opExistsTrue = 1
    opExistsFalse = 2
    opItemGet = 3
    opItemLet = 4
    opKeyLet = 5
    opNewEnum = 6
    opRemove = 7
End Enum

Public Sub Benchmark(ByRef keysToAdd() As Variant _
                   , ByRef keysMissing() As Variant _
                   , Optional ByVal compMode As VbCompareMethod = vbBinaryCompare _
                   , Optional ByVal keyType As String)
    Dim headers() As Variant: headers = Array("Iterations", "Operation" _
                                            , "VBA-Dictionary", "VBA.Collection" _
                                            , "Scripting.Dictionary", "cHashD (16384)" _
                                            , "cHashD (10% load)", "cHashD (38.5% load)" _
                                            , "Dictionary", "Dictionary (predict)")
    '
    Dim vdict As VBA_Dictionary:         Set vdict = New VBA_Dictionary
    Dim cdict As cHashD:                 Set cdict = New cHashD
    Dim coll As Collection:              Set coll = New Collection
#If Mac = 0 Then
    Dim sdict As Scripting.Dictionary:   Set sdict = New Scripting.Dictionary
#End If
    Dim ndict As Dictionary:             Set ndict = New Dictionary
    '
    vdict.CompareMode = compMode
    cdict.StringCompareMode = compMode
#If Mac = 0 Then
    sdict.CompareMode = compMode
#End If
    ndict.CompareMode = compMode
    '
    Dim rng As Range
    Dim iterations As Long
    Dim itemLevel As Long
    Dim i As Long, j As Long, k As Long
    Dim elapsed() As Variant
    Dim prevElapsed() As Variant
    Dim arrOps() As Variant
    Dim arrRes() As Range
    Dim res As Variant
    Dim b As Boolean
    '
    arrOps = Array("Add", "Exists (True)", "Exists (False)", "Item (Get)" _
                 , "Item (Let)", "Key (Let)", "For Each", "Remove")
    '
    ReDim arrRes(opAdd To opRemove)
    ThisWorkbook.Names("KeyType").Value = keyType
    ThisWorkbook.Names("VBInfo").Value = VBInfo
    
    For k = opAdd To opRemove
        With ThisWorkbook.Worksheets(arrOps(k))
            Set arrRes(k) = .Names("Results").RefersToRange
            arrRes(k).ClearContents
        End With
    Next k
    DoEvents
    '
    Dictionary.Add 1, 1: Dictionary.Remove 1 'Just initialize text hasher
    iterations = 1
    itemLevel = 1
    ReDim elapsed(LBound(headers) To UBound(headers), opAdd To opRemove)
    Do Until iterations > UBound(keysToAdd)
        prevElapsed = elapsed
        For k = opAdd To opRemove
            elapsed(0, k) = iterations
            elapsed(1, k) = arrOps(k)
        Next k
        For j = LBound(headers) + 2 To UBound(headers)
            For k = opAdd To opRemove
                elapsed(j, k) = Empty
            Next k
        Next j
        For j = 2 To UBound(headers)
            Const threeSecondsU As Long = 3 * 10 ^ 6
            Const thirtySecondsU As Long = threeSecondsU * 10
            If IsNumeric(prevElapsed(j, opAdd)) Then
                If prevElapsed(j, opAdd) > threeSecondsU Then
                    For k = opAdd To opRemove
                        If IsNumeric(prevElapsed(j, k)) Then
                            elapsed(j, k) = "'Add' too slow"
                        Else
                            elapsed(j, k) = prevElapsed(j, k)
                        End If
                    Next k
                Else
                    Select Case headers(j)
                    Case "VBA-Dictionary"
                        elapsed(j, opAdd) = AccurateTimerUs
                        For i = 1 To iterations
                            vdict.Add keysToAdd(i), i
                        Next i
                        elapsed(j, opAdd) = Round(AccurateTimerUs - elapsed(j, opAdd), 0)
                        '
                        elapsed(j, opExistsTrue) = AccurateTimerUs
                        For i = 1 To iterations
                            b = vdict.Exists(keysToAdd(i))
                        Next i
                        elapsed(j, opExistsTrue) = Round(AccurateTimerUs - elapsed(j, opExistsTrue), 0)
                        '
                        elapsed(j, opExistsFalse) = AccurateTimerUs
                        For i = 1 To iterations
                            b = vdict.Exists(keysMissing(i))
                        Next i
                        elapsed(j, opExistsFalse) = Round(AccurateTimerUs - elapsed(j, opExistsFalse), 0)
                        '
                        elapsed(j, opItemGet) = AccurateTimerUs
                        For i = 1 To iterations
                            res = vdict.Item(keysToAdd(i))
                        Next i
                        elapsed(j, opItemGet) = Round(AccurateTimerUs - elapsed(j, opItemGet), 0)
                        '
                        If IsNumeric(prevElapsed(j, opItemLet)) Then
                            If prevElapsed(j, opItemLet) > thirtySecondsU Then
                                elapsed(j, opItemLet) = "'Item(Let)' slow"
                            Else
                                #If Mac Then
                                    On Error Resume Next 'To avoid implementation bugs
                                #End If
                                elapsed(j, opItemLet) = AccurateTimerUs
                                For i = 1 To iterations
                                    vdict.Item(keysToAdd(i)) = i
                                Next i
                                elapsed(j, opItemLet) = Round(AccurateTimerUs - elapsed(j, opItemLet), 0)
                                #If Mac Then
                                    On Error GoTo 0
                                #End If
                            End If
                        Else
                            elapsed(j, opItemLet) = prevElapsed(j, opItemLet)
                        End If
                        '
                        If IsNumeric(prevElapsed(j, opKeyLet)) Then
                            If prevElapsed(j, opKeyLet) > thirtySecondsU Then
                                elapsed(j, opKeyLet) = "'Key(Let)' slow"
                            Else
                                #If Mac Then
                                    On Error Resume Next 'To avoid implementation bugs
                                #End If
                                elapsed(j, opKeyLet) = AccurateTimerUs
                                For i = 1 To iterations
                                    vdict.Key(keysToAdd(i)) = keysMissing(i)
                                Next i
                                elapsed(j, opKeyLet) = Round(AccurateTimerUs - elapsed(j, opKeyLet), 0)
                                #If Mac Then
                                    On Error GoTo 0
                                #End If
                            End If
                        Else
                            elapsed(j, opKeyLet) = prevElapsed(j, opKeyLet)
                        End If
                        '
                        elapsed(j, opNewEnum) = "not supported"
                        '
                        If IsNumeric(prevElapsed(j, opRemove)) Then
                            If prevElapsed(j, opRemove) > thirtySecondsU Then
                                elapsed(j, opRemove) = "'Remove' slow"
                            Else
                                #If Mac Then
                                    On Error Resume Next 'To avoid implementation bugs
                                #End If
                                elapsed(j, opRemove) = AccurateTimerUs
                                For i = 1 To iterations
                                    vdict.Remove keysMissing(i)
                                Next i
                                elapsed(j, opRemove) = Round(AccurateTimerUs - elapsed(j, opRemove), 0)
                                #If Mac Then
                                    On Error GoTo 0
                                #End If
                            End If
                        Else
                            elapsed(j, opRemove) = prevElapsed(j, opRemove)
                        End If
                        '
                        Set vdict = New VBA_Dictionary
                    Case "VBA.Collection"
                        elapsed(j, opAdd) = AccurateTimerUs
                        For i = 1 To iterations
                            If IsObject(keysToAdd(i)) Then
                                coll.Add i, CStr(ObjPtr(keysToAdd(i)))
                            ElseIf VarType(keysToAdd(i)) <> vbString Then
                                coll.Add i, CStr(keysToAdd(i))
                            Else
                                coll.Add i, keysToAdd(i)
                            End If
                        Next i
                        elapsed(j, opAdd) = Round(AccurateTimerUs - elapsed(j, opAdd), 0)
                        '
                        elapsed(j, opExistsTrue) = AccurateTimerUs
                        For i = 1 To iterations
                            On Error Resume Next
                            If IsObject(keysToAdd(i)) Then
                                coll.Item CStr(ObjPtr(keysToAdd(i)))
                            ElseIf VarType(keysToAdd(i)) <> vbString Then
                                coll.Item CStr(keysToAdd(i))
                            Else
                                coll.Item keysToAdd(i)
                            End If
                            b = (Err.Number = 0)
                            On Error GoTo 0
                        Next i
                        elapsed(j, opExistsTrue) = Round(AccurateTimerUs - elapsed(j, opExistsTrue), 0)
                        '
                        elapsed(j, opExistsFalse) = AccurateTimerUs
                        For i = 1 To iterations
                            On Error Resume Next
                            If IsObject(keysToAdd(i)) Then
                                coll.Item CStr(ObjPtr(keysMissing(i)))
                            ElseIf VarType(keysToAdd(i)) <> vbString Then
                                coll.Item CStr(keysMissing(i))
                            Else
                                coll.Item keysMissing(i)
                            End If
                            b = (Err.Number = 0)
                            On Error GoTo 0
                        Next i
                        elapsed(j, opExistsFalse) = Round(AccurateTimerUs - elapsed(j, opExistsFalse), 0)
                        '
                        On Error Resume Next 'Collection has issues with Unicode surrogates
                        elapsed(j, opItemGet) = AccurateTimerUs
                        For i = 1 To iterations
                            If IsObject(keysToAdd(i)) Then
                                res = coll.Item(CStr(ObjPtr(keysToAdd(i))))
                            ElseIf VarType(keysToAdd(i)) <> vbString Then
                                res = coll.Item(CStr(keysToAdd(i)))
                            Else
                                res = coll.Item(keysToAdd(i))
                            End If
                        Next i
                        elapsed(j, opItemGet) = Round(AccurateTimerUs - elapsed(j, opItemGet), 0)
                        '
                        elapsed(j, opItemLet) = "not supported"
                        elapsed(j, opKeyLet) = "not supported"
                        elapsed(j, opNewEnum) = "not supported"
                        '
                        elapsed(j, opRemove) = AccurateTimerUs
                        For i = 1 To iterations
                            If IsObject(keysToAdd(i)) Then
                                coll.Remove CStr(ObjPtr(keysToAdd(i)))
                            ElseIf VarType(keysToAdd(i)) <> vbString Then
                                coll.Remove CStr(keysToAdd(i))
                            Else
                                coll.Remove keysToAdd(i)
                            End If
                        Next i
                        elapsed(j, opRemove) = Round(AccurateTimerUs - elapsed(j, opRemove), 0)
                        On Error GoTo 0
                        '
                        Set coll = New Collection
                    Case "Scripting.Dictionary"
                    #If Mac Then
                        For k = opAdd To opRemove
                            elapsed(j, k) = "not supported"
                        Next k
                    #Else
                        elapsed(j, opAdd) = AccurateTimerUs
                        For i = 1 To iterations
                            sdict.Add keysToAdd(i), i
                        Next i
                        elapsed(j, opAdd) = Round(AccurateTimerUs - elapsed(j, opAdd), 0)
                        '
                        elapsed(j, opExistsTrue) = AccurateTimerUs
                        For i = 1 To iterations
                            b = sdict.Exists(keysToAdd(i))
                        Next i
                        elapsed(j, opExistsTrue) = Round(AccurateTimerUs - elapsed(j, opExistsTrue), 0)
                        '
                        elapsed(j, opExistsFalse) = AccurateTimerUs
                        For i = 1 To iterations
                            b = sdict.Exists(keysMissing(i))
                        Next i
                        elapsed(j, opExistsFalse) = Round(AccurateTimerUs - elapsed(j, opExistsFalse), 0)
                        '
                        elapsed(j, opItemGet) = AccurateTimerUs
                        For i = 1 To iterations
                            res = sdict.Item(keysToAdd(i))
                        Next i
                        elapsed(j, opItemGet) = Round(AccurateTimerUs - elapsed(j, opItemGet), 0)
                        '
                        elapsed(j, opItemLet) = AccurateTimerUs
                        For i = 1 To iterations
                            sdict.Item(keysToAdd(i)) = i
                        Next i
                        elapsed(j, opItemLet) = Round(AccurateTimerUs - elapsed(j, opItemLet), 0)
                        '
                        If IsNumeric(prevElapsed(j, opKeyLet)) Then
                            If prevElapsed(j, opKeyLet) > thirtySecondsU Then
                                elapsed(j, opKeyLet) = "'Key(Let)' slow"
                            Else
                                elapsed(j, opKeyLet) = AccurateTimerUs
                                For i = 1 To iterations
                                    sdict.Key(keysToAdd(i)) = keysMissing(i)
                                Next i
                                elapsed(j, opKeyLet) = Round(AccurateTimerUs - elapsed(j, opKeyLet), 0)
                            End If
                        Else
                            elapsed(j, opKeyLet) = prevElapsed(j, opKeyLet)
                        End If
                        '
                        elapsed(j, opNewEnum) = AccurateTimerUs
                        For Each res In sdict
                        Next res
                        elapsed(j, opNewEnum) = Round(AccurateTimerUs - elapsed(j, opNewEnum), 0)
                        '
                        If IsNumeric(prevElapsed(j, opRemove)) Then
                            If prevElapsed(j, opRemove) > thirtySecondsU Then
                                elapsed(j, opRemove) = "'Remove' slow"
                            Else
                                elapsed(j, opRemove) = AccurateTimerUs
                                For i = 1 To iterations
                                    sdict.Remove keysMissing(i)
                                Next i
                                elapsed(j, opRemove) = Round(AccurateTimerUs - elapsed(j, opRemove), 0)
                            End If
                        Else
                            elapsed(j, opRemove) = prevElapsed(j, opRemove)
                        End If
                        '
                        Set sdict = New Scripting.Dictionary
                    #End If
                    Case "cHashD (16384)"
                        cdict.ReInit 16 'To avoid timing unwanted deallocation
                        '
                        elapsed(j, opAdd) = AccurateTimerUs
                        cdict.ReInit 16384
                        For i = 1 To iterations
                            cdict.Add keysToAdd(i), i
                        Next i
                        elapsed(j, opAdd) = Round(AccurateTimerUs - elapsed(j, opAdd), 0)
                        '
                        elapsed(j, opExistsTrue) = AccurateTimerUs
                        For i = 1 To iterations
                            b = cdict.Exists(keysToAdd(i))
                        Next i
                        elapsed(j, opExistsTrue) = Round(AccurateTimerUs - elapsed(j, opExistsTrue), 0)
                        '
                        elapsed(j, opExistsFalse) = AccurateTimerUs
                        For i = 1 To iterations
                            b = cdict.Exists(keysMissing(i))
                        Next i
                        elapsed(j, opExistsFalse) = Round(AccurateTimerUs - elapsed(j, opExistsFalse), 0)
                        '
                        elapsed(j, opItemGet) = AccurateTimerUs
                        For i = 1 To iterations
                            res = cdict.Item(keysToAdd(i))
                        Next i
                        elapsed(j, opItemGet) = Round(AccurateTimerUs - elapsed(j, opItemGet), 0)
                        '
                        elapsed(j, opItemLet) = AccurateTimerUs
                        For i = 1 To iterations
                            cdict.Item(keysToAdd(i)) = i
                        Next i
                        elapsed(j, opItemLet) = Round(AccurateTimerUs - elapsed(j, opItemLet), 0)
                        '
                        elapsed(j, opKeyLet) = "not supported"
                        elapsed(j, opNewEnum) = "not supported"
                        '
                        elapsed(j, opRemove) = AccurateTimerUs
                        For i = 1 To iterations
                            cdict.Remove keysToAdd(i)
                        Next i
                        elapsed(j, opRemove) = Round(AccurateTimerUs - elapsed(j, opRemove), 0)
                        '
                        Set cdict = New cHashD
                    End Select
                End If
            Else
                For k = opAdd To opRemove
                    elapsed(j, k) = prevElapsed(j, k)
                Next k
            End If
            Select Case headers(j) 'We run these regardless of how long they take
            Case "cHashD (10% load)"
                cdict.ReInit 16 'To avoid timing unwanted deallocation
                '
                elapsed(j, opAdd) = AccurateTimerUs
                cdict.ReInit 2 ^ (-VBA.Int(-Log(iterations * 1.2) / Log(2)))
                For i = 1 To iterations
                    cdict.Add keysToAdd(i), i
                Next i
                elapsed(j, opAdd) = Round(AccurateTimerUs - elapsed(j, opAdd), 0)
                '
                elapsed(j, opExistsTrue) = AccurateTimerUs
                For i = 1 To iterations
                    b = cdict.Exists(keysToAdd(i))
                Next i
                elapsed(j, opExistsTrue) = Round(AccurateTimerUs - elapsed(j, opExistsTrue), 0)
                '
                elapsed(j, opExistsFalse) = AccurateTimerUs
                For i = 1 To iterations
                    b = cdict.Exists(keysMissing(i))
                Next i
                elapsed(j, opExistsFalse) = Round(AccurateTimerUs - elapsed(j, opExistsFalse), 0)
                '
                elapsed(j, opItemGet) = AccurateTimerUs
                For i = 1 To iterations
                    res = cdict.Item(keysToAdd(i))
                Next i
                elapsed(j, opItemGet) = Round(AccurateTimerUs - elapsed(j, opItemGet), 0)
                '
                elapsed(j, opItemLet) = AccurateTimerUs
                For i = 1 To iterations
                    cdict.Item(keysToAdd(i)) = i
                Next i
                elapsed(j, opItemLet) = Round(AccurateTimerUs - elapsed(j, opItemLet), 0)
                '
                elapsed(j, opKeyLet) = "not supported"
                elapsed(j, opNewEnum) = "not supported"
                '
                elapsed(j, opRemove) = AccurateTimerUs
                For i = 1 To iterations
                    cdict.Remove keysToAdd(i)
                Next i
                elapsed(j, opRemove) = Round(AccurateTimerUs - elapsed(j, opRemove), 0)
                '
                Set cdict = New cHashD
            Case "cHashD (38.5% load)"
                cdict.ReInit 16 'To avoid timing unwanted deallocation
                '
                elapsed(j, opAdd) = AccurateTimerUs
                cdict.ReInit 2 ^ (-VBA.Int(-Log(iterations * 0.3) / Log(2)))
                For i = 1 To iterations
                    cdict.Add keysToAdd(i), i
                Next i
                elapsed(j, opAdd) = Round(AccurateTimerUs - elapsed(j, opAdd), 0)
                '
                elapsed(j, opExistsTrue) = AccurateTimerUs
                For i = 1 To iterations
                    b = cdict.Exists(keysToAdd(i))
                Next i
                elapsed(j, opExistsTrue) = Round(AccurateTimerUs - elapsed(j, opExistsTrue), 0)
                '
                elapsed(j, opExistsFalse) = AccurateTimerUs
                For i = 1 To iterations
                    b = cdict.Exists(keysMissing(i))
                Next i
                elapsed(j, opExistsFalse) = Round(AccurateTimerUs - elapsed(j, opExistsFalse), 0)
                '
                elapsed(j, opItemGet) = AccurateTimerUs
                For i = 1 To iterations
                    res = cdict.Item(keysToAdd(i))
                Next i
                elapsed(j, opItemGet) = Round(AccurateTimerUs - elapsed(j, opItemGet), 0)
                '
                elapsed(j, opItemLet) = AccurateTimerUs
                For i = 1 To iterations
                    cdict.Item(keysToAdd(i)) = i
                Next i
                elapsed(j, opItemLet) = Round(AccurateTimerUs - elapsed(j, opItemLet), 0)
                '
                elapsed(j, opKeyLet) = "not supported"
                elapsed(j, opNewEnum) = "not supported"
                '
                elapsed(j, opRemove) = AccurateTimerUs
                For i = 1 To iterations
                    cdict.Remove keysToAdd(i)
                Next i
                elapsed(j, opRemove) = Round(AccurateTimerUs - elapsed(j, opRemove), 0)
                '
                Set cdict = New cHashD
            Case "Dictionary"
                elapsed(j, opAdd) = AccurateTimerUs
                For i = 1 To iterations
                    ndict.Add keysToAdd(i), i
                Next i
                elapsed(j, opAdd) = Round(AccurateTimerUs - elapsed(j, opAdd), 0)
                '
                elapsed(j, opExistsTrue) = AccurateTimerUs
                For i = 1 To iterations
                    b = ndict.Exists(keysToAdd(i))
                Next i
                elapsed(j, opExistsTrue) = Round(AccurateTimerUs - elapsed(j, opExistsTrue), 0)
                '
                elapsed(j, opExistsFalse) = AccurateTimerUs
                For i = 1 To iterations
                    b = ndict.Exists(keysMissing(i))
                Next i
                elapsed(j, opExistsFalse) = Round(AccurateTimerUs - elapsed(j, opExistsFalse), 0)
                '
                elapsed(j, opItemGet) = AccurateTimerUs
                For i = 1 To iterations
                    res = ndict.Item(keysToAdd(i))
                Next i
                elapsed(j, opItemGet) = Round(AccurateTimerUs - elapsed(j, opItemGet), 0)
                '
                elapsed(j, opItemLet) = AccurateTimerUs
                For i = 1 To iterations
                    ndict.Item(keysToAdd(i)) = i
                Next i
                elapsed(j, opItemLet) = Round(AccurateTimerUs - elapsed(j, opItemLet), 0)
                '
                elapsed(j, opKeyLet) = AccurateTimerUs
                For i = 1 To iterations
                    ndict.Key(keysToAdd(i)) = keysMissing(i)
                Next i
                elapsed(j, opKeyLet) = Round(AccurateTimerUs - elapsed(j, opKeyLet), 0)
                '
                elapsed(j, opNewEnum) = AccurateTimerUs
                For Each res In ndict
                Next res
                elapsed(j, opNewEnum) = Round(AccurateTimerUs - elapsed(j, opNewEnum), 0)
                '
                elapsed(j, opRemove) = AccurateTimerUs
                For i = 1 To iterations
                    ndict.Remove keysMissing(i)
                Next i
                elapsed(j, opRemove) = Round(AccurateTimerUs - elapsed(j, opRemove), 0)
                '
                Set ndict = New Dictionary
            Case "Dictionary (predict)" 'Only Add is of interest
                elapsed(j, opAdd) = AccurateTimerUs
                ndict.PredictCount iterations
                For i = 1 To iterations
                    ndict.Add keysToAdd(i), i
                Next i
                elapsed(j, opAdd) = Round(AccurateTimerUs - elapsed(j, opAdd), 0)
                '
                Set ndict = New Dictionary
            End Select
            For k = opAdd To opRemove
                With arrRes(k)
                    If j = 2 Then .Cells(itemLevel, 1).Value2 = iterations
                    If j <= .Columns.Count Then .Cells(itemLevel, j).Value2 = elapsed(j, k)
                End With
            Next k
            DoEvents
        Next j
        iterations = iterations * 2
        itemLevel = itemLevel + 1
    Loop
    #If Win64 = 0 Then
        For k = opAdd To opRemove
            With arrRes(k)
                For i = itemLevel To .Rows.Count
                    .Cells(i, 1).Value2 = 2 ^ (i - 1)
                    For j = 2 To .Columns.Count
                        .Cells(i, j).Value2 = "out of memory"
                        DoEvents
                    Next j
                Next i
            End With
        Next k
    #End If
    ThisWorkbook.Save  'To avoid freezing when deallocating test data
End Sub

Public Function VBInfo() As String
    Dim res(0 To 2) As String
    #If Mac Then
        res(0) = "Mac"
    #Else
        res(0) = "Win"
    #End If
    #If VBA7 Then
        res(1) = "VBA7"
    #Else
        res(1) = "VBA6"
    #End If
    #If Win64 Then
        res(2) = "x64"
    #Else
        res(2) = "x32"
    #End If
    VBInfo = Join(res, " ")
End Function

Public Sub ExportBenchResults()
    #If Mac Then
        Const PATH_SEPARATOR = "/"
    #Else
        Const PATH_SEPARATOR = "\"
    #End If
    Dim pic As Picture
    Dim folderPath As String
    Dim temp As String
    Dim picPath As String
    '
    folderPath = BrowseForFolder(, "Choose path for saving screenshots")
    If LenB(folderPath) = 0 Then Exit Sub
    folderPath = folderPath & PATH_SEPARATOR
    '
    temp = ThisWorkbook.Names("KeyType").Value
    temp = Mid$(temp, 3, Len(temp) - 3)
    temp = Join(Array(vbNullString, Replace(temp, ",", ""), VBInfo), "_")
    temp = LCase$(Replace(temp, " ", "_"))
    '
    Camera.Activate
    ActiveWindow.Zoom = 100
    On Error Resume Next
    For Each pic In Camera.UsedRange.Worksheet.Pictures
        pic.CopyPicture
        With Camera.ChartObjects.Add(0, 0, pic.Width, pic.Height)
            .Select
            With .Chart
                .Paste
                picPath = folderPath & pic.Name & temp & ".png"
                Kill picPath
                .Export picPath
            End With
            .Delete
        End With
        DoEvents
    Next pic
    On Error GoTo 0
    ActiveWindow.Zoom = 55
End Sub

'*******************************************************************************
'Returns a folder path by using a FolderPicker FileDialog
'*******************************************************************************
Private Function BrowseForFolder(Optional ByRef initialPath As String _
                               , Optional ByRef dialogTitle As String) As String
#If Mac Then
    'If user has not accesss [initialPath] previously, will be prompted by
    'Mac OS to Grant permission to directory
    If LenB(initialPath) > 0 Then
        If Not Right(initialPath, 1) = Application.PathSeparator Then
            initialPath = initialPath & Application.PathSeparator
        End If
        Dir initialPath, Attributes:=vbDirectory
    End If
    Dim retPath
    If LenB(dialogTitle) = 0 Then dialogTitle = "Choose Foldler"
    retPath = MacScript("choose folder with prompt """ & dialogTitle & """ as string")
    If Len(retPath) > 0 Then
        retPath = MacScript("POSIX path of """ & retPath & """")
        If LenB(retPath) > 0 Then
            BrowseForFolder = retPath
        End If
    End If
#Else
    'In case reference to Microsoft Office X.XX Object Library is missing
    Const dialogTypeFolderPicker As Long = 4 'msoFileDialogFolderPicker
    Const actionButton As Long = -1
    '
    With Application.FileDialog(dialogTypeFolderPicker)
        If LenB(dialogTitle) > 0 Then .Title = dialogTitle
        If LenB(initialPath) > 0 Then .InitialFileName = initialPath
        If LenB(.InitialFileName) = 0 Then
            Dim app As Object: Set app = Application 'Needs to be late-binded
            Select Case Application.Name
                Case "Microsoft Excel": .InitialFileName = app.ThisWorkbook.Path
                Case "Microsoft Word":  .InitialFileName = app.ThisDocument.Path
            End Select
        End If
        If .Show = actionButton Then
            .InitialFileName = .SelectedItems.Item(1)
            BrowseForFolder = .InitialFileName
        End If
    End With
#End If
End Function

Private Sub FormatColors()
    Dim rng As Range
    Dim temp As Range
    Dim ws As Worksheet
    Dim res As Name
    Dim coll As New Collection
    Const maxMultiplier As Long = 10
    '
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set res = Nothing
        Set res = ws.Names("Results")
        On Error GoTo 0
        If Not res Is Nothing Then
            Set rng = res.RefersToRange
            coll.Add rng.Offset(0, 1).Resize(rng.Rows.Count, rng.Columns.Count - 1)
            ws.Cells.FormatConditions.Delete
        End If
    Next ws
    '
    For Each temp In coll
        For Each rng In temp.Rows
            rng.FormatConditions.AddColorScale ColorScaleType:=3
            With rng.FormatConditions(1)
                With .ColorScaleCriteria(1)
                    .Type = xlConditionValueLowestValue
                    .FormatColor.Color = 8109667
                    .FormatColor.TintAndShade = 0
                End With
                With .ColorScaleCriteria(2)
                    .Type = xlConditionValuePercentile
                    .Value = 50
                    .FormatColor.Color = 8711167
                    .FormatColor.TintAndShade = 0
                End With
                With .ColorScaleCriteria(3)
                    Dim x As ColorScaleCriteria
                    .Type = xlConditionValueFormula
                    .Value = "=" & rng.Cells(1, rng.Columns.Count).Address & "*" & maxMultiplier
                    .FormatColor.Color = 7039480
                    .FormatColor.TintAndShade = 0
                End With
            End With
        Next rng
    Next temp
End Sub
