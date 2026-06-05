VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MazeChaseForm 
   Caption         =   "Maze Chase"
   ClientHeight    =   7032
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6180
   OleObjectBlob   =   "MazeChaseForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MazeChaseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    MazeChaseRenderer.AttachForm Me
    MazeChaseGame.InitializeGame
    MazeChaseRenderer.BuildScene
End Sub

Private Sub UserForm_Activate()
    MazeChaseInput.BindKeys
    MazeChaseLoop.StartLoop
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    MazeChaseGame.QueueDirectionByKey CLng(KeyCode)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    MazeChaseLoop.StopLoop
    MazeChaseInput.UnbindKeys
    MazeChaseRenderer.DetachForm
End Sub

Private Sub UserForm_Terminate()
    MazeChaseLoop.StopLoop
    MazeChaseInput.UnbindKeys
    MazeChaseRenderer.DetachForm
End Sub


