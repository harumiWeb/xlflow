Attribute VB_Name = "MazeChaseLoop"
Option Explicit

Private Const TimerIntervalMs As Long = 180

#If VBA7 Then
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal uIDEvent As LongPtr) As Long
Private mTimerId As LongPtr
#Else
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal uIDEvent As Long) As Long
Private mTimerId As Long
#End If

Private mTickInProgress As Boolean

Public Sub StartLoop()
    If mTimerId <> 0 Then
        Exit Sub
    End If

    #If VBA7 Then
    mTimerId = SetTimer(0, 0, TimerIntervalMs, AddressOf MazeChaseTimerProc)
    #Else
    mTimerId = SetTimer(0, 0, TimerIntervalMs, AddressOf MazeChaseTimerProc)
    #End If
End Sub

Public Sub StopLoop()
    If mTimerId = 0 Then
        Exit Sub
    End If

    KillTimer 0, mTimerId
    mTimerId = 0
End Sub

#If VBA7 Then
Public Sub MazeChaseTimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    On Error GoTo FailSafe

    If mTickInProgress Then
        Exit Sub
    End If

    If Not MazeChaseRenderer.HasForm() Then
        StopLoop
        Exit Sub
    End If

    mTickInProgress = True
    MazeChaseInput.PollKeys
    MazeChaseGame.TickGame
    MazeChaseRenderer.RenderFrame

    If MazeChaseGame.IsFinished() Then
        StopLoop
    End If

    SafeExit:
    mTickInProgress = False
    Exit Sub

    FailSafe:
    StopLoop
    Resume SafeExit
End Sub
#Else
Public Sub MazeChaseTimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error GoTo FailSafe

    If mTickInProgress Then
        Exit Sub
    End If

    If Not MazeChaseRenderer.HasForm() Then
        StopLoop
        Exit Sub
    End If

    mTickInProgress = True
    MazeChaseInput.PollKeys
    MazeChaseGame.TickGame
    MazeChaseRenderer.RenderFrame

    If MazeChaseGame.IsFinished() Then
        StopLoop
    End If

    SafeExit:
    mTickInProgress = False
    Exit Sub

    FailSafe:
    StopLoop
    Resume SafeExit
End Sub
#End If
