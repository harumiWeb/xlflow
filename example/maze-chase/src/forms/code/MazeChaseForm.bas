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
