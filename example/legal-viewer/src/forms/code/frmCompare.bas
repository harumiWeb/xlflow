Option Explicit

Private Const FORM_WIDTH_POINTS As Double = 1010
Private Const FORM_HEIGHT_POINTS As Double = 680

Private Sub UserForm_Initialize()
    ConfigureFormWindow
    ConfigureControls
    modUiState.ApplyDefaultUserFormFont Me
    RefreshComparisonSourceLists
End Sub

Private Sub UserForm_Activate()
    modUiState.ApplyDefaultUserFormFont Me
    modWindowPlacement.CenterUserFormOnExcelMonitor Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCompare_Click()
    CompareSelectedBodyUnitsForForm
End Sub

Private Sub lstLeftLaws_Click()
    RefreshBodyListForSide Me.lstLeftLaws, Me.lstLeftBodies, True
    CompareSelectedBodyUnitsForForm
End Sub

Private Sub lstRightLaws_Click()
    RefreshBodyListForSide Me.lstRightLaws, Me.lstRightBodies, False
    CompareSelectedBodyUnitsForForm
End Sub

Private Sub lstLeftBodies_Click()
    CompareSelectedBodyUnitsForForm
End Sub

Private Sub lstRightBodies_Click()
    CompareSelectedBodyUnitsForForm
End Sub

Private Sub rdoTableMarkdown_Click()
    CompareSelectedBodyUnitsForForm
End Sub

Private Sub rdoTableAscii_Click()
    CompareSelectedBodyUnitsForForm
End Sub

Public Function RefreshComparisonSourceLists() As String
    On Error GoTo ErrHandler

    Dim leftCount As Long
    Dim rightCount As Long
    leftCount = RefreshLawListForSide(Me.lstLeftLaws, Me.lstLeftBodies, True)
    rightCount = RefreshLawListForSide(Me.lstRightLaws, Me.lstRightBodies, False)

    CompareSelectedBodyUnitsForForm
    Me.lblStatus.caption = "比較用法令 " & CStr(leftCount) & " / " & CStr(rightCount) & " 件"
    RefreshComparisonSourceLists = "ok:" & CStr(leftCount) & ":" & CStr(rightCount)
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "CompareForm", "frmCompare.RefreshComparisonSourceLists", Err.description, "", "", ""
    Me.lblStatus.caption = "比較対象の読み込み失敗: " & Err.description
    RefreshComparisonSourceLists = "error:" & Err.description
End Function

Public Function CompareSelectedBodyUnitsForForm() As String
    On Error GoTo ErrHandler

    Dim leftBodyRow As Long
    Dim rightBodyRow As Long
    leftBodyRow = SelectedBodyUnitStoreRow(Me.lstLeftBodies)
    rightBodyRow = SelectedBodyUnitStoreRow(Me.lstRightBodies)

    If leftBodyRow <= 0 Or rightBodyRow <= 0 Then
        Me.txtDiff.text = ""
        Me.lblStatus.caption = "左右の本文単位を選んでください"
        CompareSelectedBodyUnitsForForm = "empty"
        Exit Function
    End If

    Dim diffText As String
    diffText = modCompare.CompareBodyUnitRows(leftBodyRow, rightBodyRow, TableDisplayModeForForm())
    Me.txtDiff.text = diffText
    Me.lblStatus.caption = "比較しました: 左 " & CStr(leftBodyRow) & " / 右 " & CStr(rightBodyRow)
    CompareSelectedBodyUnitsForForm = "ok:" & CStr(Len(diffText))
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "CompareForm", "frmCompare.CompareSelectedBodyUnitsForForm", Err.description, "", "", ""
    Me.txtDiff.text = "比較失敗: " & Err.description
    Me.lblStatus.caption = "比較失敗: " & Err.description
    CompareSelectedBodyUnitsForForm = "error:" & Err.description
End Function

Public Function CompareFormSmoke() As String
    On Error GoTo ErrHandler

    Dim smokeAddedRow As Long
    smokeAddedRow = modLawNavigator.PrepareBodyNavigatorSmokeData()
    If smokeAddedRow <= 0 Then
        Err.Raise vbObjectError + 7701, "frmCompare.CompareFormSmoke", "Smoke law row was not prepared."
    End If

    Load frmCompare
    If Me.cmdCompare.caption <> "比較" Then
        Err.Raise vbObjectError + 7702, "frmCompare.CompareFormSmoke", "Compare button caption was not initialized."
    End If
    If Me.cmdClose.caption <> "閉じる" Then
        Err.Raise vbObjectError + 7703, "frmCompare.CompareFormSmoke", "Close button caption was not initialized."
    End If
    If Me.lblStatus.Top > 650 Or Me.lblStatus.Height < 18 Then
        Err.Raise vbObjectError + 7704, "frmCompare.CompareFormSmoke", "Status label was not moved up enough."
    End If

    SelectAddedLawRowByStoreRow Me.lstLeftLaws, smokeAddedRow
    RefreshBodyListForSide Me.lstLeftLaws, Me.lstLeftBodies, True
    SelectAddedLawRowByStoreRow Me.lstRightLaws, smokeAddedRow
    RefreshBodyListForSide Me.lstRightLaws, Me.lstRightBodies, False

    If Me.lstLeftBodies.ListCount < 4 Or Me.lstRightBodies.ListCount < 4 Then
        Err.Raise vbObjectError + 7705, "frmCompare.CompareFormSmoke", "Compare body lists were not populated."
    End If

    SelectBodyRowByStoreRow Me.lstLeftBodies, 2
    SelectBodyRowByStoreRow Me.lstRightBodies, 4
    Dim resultText As String
    resultText = CompareSelectedBodyUnitsForForm()
    If Left$(resultText, 3) <> "ok:" Then
        Err.Raise vbObjectError + 7706, "frmCompare.CompareFormSmoke", "Comparison failed."
    End If
    If Len(Trim$(Me.txtDiff.text)) = 0 Or InStr(1, Me.txtDiff.text, "- ", vbBinaryCompare) = 0 Or InStr(1, Me.txtDiff.text, "+ ", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7707, "frmCompare.CompareFormSmoke", "Diff text was not generated."
    End If

    CompareFormSmoke = "ok:" & CStr(Len(Me.txtDiff.text))
    Unload frmCompare
    modTempCache.DeleteAllTempCaches ThisWorkbook
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "CompareForm", "frmCompare.CompareFormSmoke", Err.description, "", "", ""
    CompareFormSmoke = "error:" & CStr(Err.Number) & ":" & Err.description
    On Error GoTo CleanupFail
    Unload frmCompare
    modTempCache.DeleteAllTempCaches ThisWorkbook
    GoTo CleanupDone

CleanupFail:
    Resume CleanupDone

CleanupDone:
End Function

Private Sub ConfigureFormWindow()
    Me.caption = "条文比較"
    Me.Width = FORM_WIDTH_POINTS
    Me.Height = FORM_HEIGHT_POINTS
    modUiState.ApplyDefaultUserFormFont Me
End Sub

Private Sub ConfigureControls()
    modAddedLawStore.ConfigureAddedLawListBox Me.lstLeftLaws
    modAddedLawStore.ConfigureAddedLawListBox Me.lstRightLaws
    modLawNavigator.ConfigureBodyNavListBox Me.lstLeftBodies
    modLawNavigator.ConfigureBodyNavListBox Me.lstRightBodies

    Me.rdoTableMarkdown.GroupName = "compare_table_display"
    Me.rdoTableAscii.GroupName = "compare_table_display"
    Me.rdoTableMarkdown.value = True
    Me.rdoTableAscii.value = False

    ConfigureDiffTextBox Me.txtDiff
End Sub

Private Sub ConfigureDiffTextBox(ByVal targetTextBox As Object)
    targetTextBox.MultiLine = True
    targetTextBox.WordWrap = False
    targetTextBox.ScrollBars = 3
    targetTextBox.Locked = True
    modUiState.ApplyDefaultControlFont targetTextBox
End Sub

Private Function RefreshLawListForSide(ByVal targetLawListBox As Object, ByVal targetBodyListBox As Object, ByVal isLeftSide As Boolean) As Long
    On Error GoTo ErrHandler

    Dim selectedStoreRow As Long
    selectedStoreRow = SelectedAddedLawStoreRow(targetLawListBox)

    Dim itemCount As Long
    itemCount = modAddedLawStore.PopulateAddedLawListBox(targetLawListBox)
    If itemCount > 0 Then
        SelectAddedLawRowByStoreRow targetLawListBox, selectedStoreRow
        If targetLawListBox.listIndex < 0 Then
            targetLawListBox.listIndex = 0
        End If
    End If

    RefreshBodyListForSide targetLawListBox, targetBodyListBox, isLeftSide
    RefreshLawListForSide = itemCount
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "CompareForm", "frmCompare.RefreshLawListForSide", Err.description, "", "", ""
    RefreshLawListForSide = 0
End Function

Private Sub RefreshBodyListForSide(ByVal targetLawListBox As Object, ByVal targetBodyListBox As Object, ByVal isLeftSide As Boolean)
    On Error GoTo ErrHandler

    Dim lawRow As Long
    lawRow = SelectedAddedLawStoreRow(targetLawListBox)
    targetBodyListBox.Clear
    If lawRow <= 0 Then
        Exit Sub
    End If

    Dim selectedBodyRow As Long
    selectedBodyRow = SelectedBodyUnitStoreRow(targetBodyListBox)

    Dim itemCount As Long
    itemCount = modLawNavigator.PopulateBodyNavListBox(targetBodyListBox, lawRow)
    If itemCount > 0 Then
        SelectBodyRowByStoreRow targetBodyListBox, selectedBodyRow
        If targetBodyListBox.listIndex < 0 Then
            targetBodyListBox.listIndex = 0
        End If
    End If
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "CompareForm", "frmCompare.RefreshBodyListForSide", Err.description, "", "", IIf(isLeftSide, "left", "right")
End Sub

Private Function SelectedAddedLawStoreRow(ByVal targetListBox As Object) As Long
    On Error GoTo ErrHandler

    If targetListBox.listIndex < 0 Then Exit Function
    SelectedAddedLawStoreRow = CLng(Val(targetListBox.List(targetListBox.listIndex, 4)))
    Exit Function

ErrHandler:
    SelectedAddedLawStoreRow = 0
End Function

Private Function SelectedBodyUnitStoreRow(ByVal targetListBox As Object) As Long
    On Error GoTo ErrHandler

    If targetListBox.listIndex < 0 Then Exit Function
    SelectedBodyUnitStoreRow = CLng(Val(targetListBox.List(targetListBox.listIndex, 3)))
    Exit Function

ErrHandler:
    SelectedBodyUnitStoreRow = 0
End Function

Private Sub SelectAddedLawRowByStoreRow(ByVal targetListBox As Object, ByVal storeRow As Long)
    Dim listIndex As Long
    If storeRow <= 0 Then
        targetListBox.listIndex = -1
        Exit Sub
    End If

    For listIndex = 0 To targetListBox.ListCount - 1
        If CLng(Val(targetListBox.List(listIndex, 4))) = storeRow Then
            targetListBox.listIndex = listIndex
            Exit Sub
        End If
    Next listIndex
End Sub

Private Sub SelectBodyRowByStoreRow(ByVal targetListBox As Object, ByVal storeRow As Long)
    Dim listIndex As Long
    If storeRow <= 0 Then
        targetListBox.listIndex = -1
        Exit Sub
    End If

    For listIndex = 0 To targetListBox.ListCount - 1
        If CLng(Val(targetListBox.List(listIndex, 3))) = storeRow Then
            targetListBox.listIndex = listIndex
            Exit Sub
        End If
    Next listIndex
End Sub

Private Function TableDisplayModeForForm() As String
    If Me.rdoTableAscii.value Then
        TableDisplayModeForForm = modLawNavigator.TableDisplayModeAscii()
    Else
        TableDisplayModeForForm = modLawNavigator.TableDisplayModeMarkdown()
    End If
End Function
