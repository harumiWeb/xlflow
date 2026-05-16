Attribute VB_Name = "GameMain"
Option Explicit

Public Sub StartInvaderGame()
  Dim openForm As Object

  For Each openForm In VBA.UserForms
    If TypeName(openForm) = "frmInvader" Then
      Unload openForm
      Exit For
    End If
  Next openForm

  Load frmInvader
  frmInvader.Show vbModeless
End Sub

Public Sub StopInvaderGame()
  Dim openForm As Object

  For Each openForm In VBA.UserForms
    If TypeName(openForm) = "frmInvader" Then
      openForm.RequestStop
      Unload openForm
      Exit For
    End If
  Next openForm
End Sub
