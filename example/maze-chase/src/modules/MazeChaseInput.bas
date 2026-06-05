Attribute VB_Name = "MazeChaseInput"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

Private mLeftWasPressed As Boolean
Private mUpWasPressed As Boolean
Private mRightWasPressed As Boolean
Private mDownWasPressed As Boolean

Public Sub BindKeys()
    ResetKeyStateTracking
    Application.OnKey "{LEFT}", "MazeChaseInput.LeftKey"
    Application.OnKey "{UP}", "MazeChaseInput.UpKey"
    Application.OnKey "{RIGHT}", "MazeChaseInput.RightKey"
    Application.OnKey "{DOWN}", "MazeChaseInput.DownKey"
End Sub

Public Sub UnbindKeys()
    Application.OnKey "{LEFT}"
    Application.OnKey "{UP}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{DOWN}"
    ResetKeyStateTracking
End Sub

Public Sub LeftKey()
    MazeChaseGame.QueueDirectionByKey vbKeyLeft
End Sub

Public Sub UpKey()
    MazeChaseGame.QueueDirectionByKey vbKeyUp
End Sub

Public Sub RightKey()
    MazeChaseGame.QueueDirectionByKey vbKeyRight
End Sub

Public Sub DownKey()
    MazeChaseGame.QueueDirectionByKey vbKeyDown
End Sub

Public Sub PollKeys()
    Dim leftPressed As Boolean
    Dim upPressed As Boolean
    Dim rightPressed As Boolean
    Dim downPressed As Boolean

    leftPressed = IsPressed(vbKeyLeft)
    upPressed = IsPressed(vbKeyUp)
    rightPressed = IsPressed(vbKeyRight)
    downPressed = IsPressed(vbKeyDown)

    If leftPressed And Not mLeftWasPressed Then
        MazeChaseGame.QueueDirectionByKey vbKeyLeft
    ElseIf upPressed And Not mUpWasPressed Then
        MazeChaseGame.QueueDirectionByKey vbKeyUp
    ElseIf rightPressed And Not mRightWasPressed Then
        MazeChaseGame.QueueDirectionByKey vbKeyRight
    ElseIf downPressed And Not mDownWasPressed Then
        MazeChaseGame.QueueDirectionByKey vbKeyDown
    End If

    mLeftWasPressed = leftPressed
    mUpWasPressed = upPressed
    mRightWasPressed = rightPressed
    mDownWasPressed = downPressed
End Sub

Private Function IsPressed(ByVal virtualKey As Long) As Boolean
    IsPressed = ((GetAsyncKeyState(virtualKey) And &H8000) <> 0)
End Function

Private Sub ResetKeyStateTracking()
    mLeftWasPressed = False
    mUpWasPressed = False
    mRightWasPressed = False
    mDownWasPressed = False
End Sub
