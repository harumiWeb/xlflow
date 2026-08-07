VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_003_Buttons_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
On Error Resume Next

    DoCmd.Close

End Sub

Private Sub cmdRecordAdd_Click()
On Error GoTo ErrTrap

    DoCmd.GoToRecord , , acNewRec

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2501  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdRecordAdd_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdRecordDelete_Click()
On Error GoTo ErrTrap

    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2105  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdRecordDelete_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdRecordCopy_Click()
On Error GoTo ErrTrap

    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 2, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 5, , acMenuVer70 'Paste Append

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2046  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdRecordCopy_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdRecordPrint_Click()
On Error GoTo ErrTrap

    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.PrintOut acSelection

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2501  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdRecordPrint_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdRecordSave_Click()
On Error GoTo ErrTrap

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2501  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdRecordSave_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdRecordUndo_Click()
On Error GoTo ErrTrap

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2046  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdRecordUndo_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdMailReport_Click()
On Error GoTo ErrTrap
Dim stDocName As String

    stDocName = "rpt Bad Actors by Catalog Profile Subreport"
    DoCmd.SendObject acReport, stDocName

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2501  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdMailReport_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFindNext_Click()
On Error GoTo ErrTrap

    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2143  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFindNext_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFindRecord_Click()
On Error GoTo ErrTrap

    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2137  ' action not available at this time
            Resume ExitProcedure
        Case Is = 2501  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFindRecord_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFindFirstRecord_Click()
On Error GoTo ErrTrap

    DoCmd.GoToRecord , , acFirst

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2143 ' did not specify record
            Resume ExitProcedure
        Case Is = 2501  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFindFirstRecord_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFindLastRecord_Click()
On Error GoTo ErrTrap

    DoCmd.GoToRecord , , acLast

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2501  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFindLastRecord_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFindNextRecord_Click()
On Error GoTo ErrTrap

    DoCmd.GoToRecord , , acNext

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2105  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFindNextRecord_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFindPreviousRecord_Click()
On Error GoTo ErrTrap

    DoCmd.GoToRecord , , acPrevious

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2105  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFindPreviousRecord_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFilterQBF_Click()
On Error GoTo ErrTrap

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 2, , acMenuVer70

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2046  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFilterQBF_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFilterEdit_Click()
On Error GoTo ErrTrap

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 0, 2, acMenuVer70

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2046  ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFilterEdit_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub

Private Sub cmdFilterBySelection_Click()
On Error GoTo ErrTrap

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 2, , acMenuVer70

ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case Is = 2046 ' action was canceled by user
            Resume ExitProcedure
        Case Is <> 0
            Call ErrorMsg("cmdFilterBySelection_Click", Me.Name, Err.Number, Err.Description)
            Resume ExitProcedure
        Case Else
            Resume ExitProcedure
    End Select
    
End Sub
