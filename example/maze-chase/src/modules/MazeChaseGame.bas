Attribute VB_Name = "MazeChaseGame"
Option Explicit

Private Const MazeDirectionNone As Long = 0
Private Const MazeDirectionLeft As Long = 1
Private Const MazeDirectionUp As Long = 2
Private Const MazeDirectionRight As Long = 3
Private Const MazeDirectionDown As Long = 4

Private Const BoardRows As Long = 17
Private Const BoardCols As Long = 17
Private Const CellSizePoints As Long = 20
Private Const BoardLeftPoints As Long = 16
Private Const BoardTopPoints As Long = 42
Private Const HudHeightPoints As Long = 24
Private Const FooterHeightPoints As Long = 24
Private Const DefaultLives As Long = 3
Private Const DotScore As Long = 10
Private Const PowerScore As Long = 50
Private Const GhostScore As Long = 200
Private Const FrightenedDurationTicks As Long = 45

Private mWalls(1 To BoardRows, 1 To BoardCols) As Boolean
Private mDots(1 To BoardRows, 1 To BoardCols) As Boolean
Private mPowers(1 To BoardRows, 1 To BoardCols) As Boolean
Private mGhostRow(1 To 4) As Long
Private mGhostCol(1 To 4) As Long
Private mGhostSpawnRow(1 To 4) As Long
Private mGhostSpawnCol(1 To 4) As Long
Private mGhostDirection(1 To 4) As Long
Private mPlayerRowValue As Long
Private mPlayerColValue As Long
Private mPlayerSpawnRow As Long
Private mPlayerSpawnCol As Long
Private mPlayerDirectionValue As Long
Private mRequestedDirection As Long
Private mPlayerMouthOpen As Boolean
Private mScore As Long
Private mLives As Long
Private mRemainingPellets As Long
Private mFrightenedTicks As Long
Private mGameOver As Boolean
Private mPlayerWon As Boolean
Private mStatusText As String
Private mConsumedRow As Long
Private mConsumedCol As Long
Private mConsumedPower As Boolean
Private mPowerBoostReady As Boolean

Public Sub InitializeGame()
    BuildBoard
    mScore = 0
    mLives = DefaultLives
    mGameOver = False
    mPlayerWon = False
    mRequestedDirection = MazeDirectionNone
    mPlayerMouthOpen = True
    mPowerBoostReady = False
    mStatusText = "Collect every dot"
    ResetActors
    ClearConsumedMarker
End Sub

Public Sub QueueDirectionByKey(ByVal KeyCode As Long)
    Select Case KeyCode
    Case vbKeyLeft
        mRequestedDirection = MazeDirectionLeft
    Case vbKeyUp
        mRequestedDirection = MazeDirectionUp
    Case vbKeyRight
        mRequestedDirection = MazeDirectionRight
    Case vbKeyDown
        mRequestedDirection = MazeDirectionDown
    End Select
End Sub

Public Sub TickGame()
    ClearConsumedMarker

    If mGameOver Then
        Exit Sub
    End If

    AdvancePlayer
    HandleCurrentTile
    ResolveGhostCollisions
    ApplyPowerBoostMovement
    AdvanceGhosts
    ResolveGhostCollisions

    If mGameOver Then
        Exit Sub
    End If

    If mRemainingPellets = 0 Then
        mPlayerWon = True
        mGameOver = True
        mStatusText = "You win"
    ElseIf mFrightenedTicks > 0 Then
        mFrightenedTicks = mFrightenedTicks - 1
        If mFrightenedTicks = 0 Then
            mPowerBoostReady = False
            mStatusText = "Collect every dot"
        End If
    End If
End Sub

Public Function Rows() As Long
    Rows = BoardRows
End Function

Public Function Cols() As Long
    Cols = BoardCols
End Function

Public Function CellSize() As Long
    CellSize = CellSizePoints
End Function

Public Function BoardLeft() As Long
    BoardLeft = BoardLeftPoints
End Function

Public Function BoardTop() As Long
    BoardTop = BoardTopPoints
End Function

Public Function FormWidth() As Long
    FormWidth = BoardLeftPoints * 2 + BoardCols * CellSizePoints + 12
End Function

Public Function FormHeight() As Long
    FormHeight = BoardTopPoints + BoardRows * CellSizePoints + FooterHeightPoints + 30
End Function

Public Function scoreValue() As Long
    scoreValue = mScore
End Function

Public Function LivesValue() As Long
    LivesValue = mLives
End Function

Public Function StatusText() As String
    StatusText = mStatusText
End Function

Public Function PlayerRow() As Long
    PlayerRow = mPlayerRowValue
End Function

Public Function PlayerCol() As Long
    PlayerCol = mPlayerColValue
End Function

Public Function PlayerDirection() As Long
    PlayerDirection = mPlayerDirectionValue
End Function

Public Function PlayerMouthOpen() As Boolean
    PlayerMouthOpen = mPlayerMouthOpen
End Function

Public Function GhostCount() As Long
    GhostCount = 4
End Function

Public Function GhostRow(ByVal index As Long) As Long
    GhostRow = mGhostRow(index)
End Function

Public Function GhostCol(ByVal index As Long) As Long
    GhostCol = mGhostCol(index)
End Function

Public Function GhostHomeRow(ByVal index As Long) As Long
    GhostHomeRow = mGhostSpawnRow(index)
End Function

Public Function GhostHomeCol(ByVal index As Long) As Long
    GhostHomeCol = mGhostSpawnCol(index)
End Function

Public Function IsWallAt(ByVal rowIndex As Long, ByVal colIndex As Long) As Boolean
    IsWallAt = mWalls(rowIndex, colIndex)
End Function

Public Function HasDotAt(ByVal rowIndex As Long, ByVal colIndex As Long) As Boolean
    HasDotAt = mDots(rowIndex, colIndex)
End Function

Public Function HasPowerAt(ByVal rowIndex As Long, ByVal colIndex As Long) As Boolean
    HasPowerAt = mPowers(rowIndex, colIndex)
End Function

Public Function IsGhostFrightened() As Boolean
    IsGhostFrightened = (mFrightenedTicks > 0)
End Function

Public Function FrightenedTicksRemaining() As Long
    FrightenedTicksRemaining = mFrightenedTicks
End Function

Public Function IsPowerBoostActive() As Boolean
    IsPowerBoostActive = (mFrightenedTicks > 0)
End Function

Public Function IsFinished() As Boolean
    IsFinished = mGameOver
End Function

Public Function HasPlayerWon() As Boolean
    HasPlayerWon = mPlayerWon
End Function

Public Function HasConsumedTile() As Boolean
    HasConsumedTile = (mConsumedRow > 0)
End Function

Public Function ConsumedRow() As Long
    ConsumedRow = mConsumedRow
End Function

Public Function ConsumedCol() As Long
    ConsumedCol = mConsumedCol
End Function

Public Function ConsumedWasPower() As Boolean
    ConsumedWasPower = mConsumedPower
End Function

Public Sub AcknowledgeConsumedTile()
    ClearConsumedMarker
End Sub

Public Sub DebugSetPlayerPosition(ByVal rowIndex As Long, ByVal colIndex As Long)
    mPlayerRowValue = rowIndex
    mPlayerColValue = colIndex
    mPlayerDirectionValue = MazeDirectionNone
    mRequestedDirection = MazeDirectionNone
End Sub

Public Sub DebugSetPlayerDirection(ByVal directionValue As Long)
    mPlayerDirectionValue = directionValue
End Sub

Public Sub DebugSetPlayerMouthOpen(ByVal mouthOpenValue As Boolean)
    mPlayerMouthOpen = mouthOpenValue
End Sub

Public Sub DebugSetRequestedDirection(ByVal directionValue As Long)
    mRequestedDirection = directionValue
End Sub

Public Sub DebugSetGhostPosition(ByVal index As Long, ByVal rowIndex As Long, ByVal colIndex As Long)
    mGhostRow(index) = rowIndex
    mGhostCol(index) = colIndex
End Sub

Public Sub DebugSetGhostDirection(ByVal index As Long, ByVal directionValue As Long)
    mGhostDirection(index) = directionValue
End Sub

Public Sub DebugConsumeCurrentTile()
    HandleCurrentTile
End Sub

Public Sub DebugResolveCollisions()
    ResolveGhostCollisions
End Sub

Public Sub DebugClearPellets()
    Dim rowIndex As Long
    Dim colIndex As Long

    For rowIndex = 1 To BoardRows
        For colIndex = 1 To BoardCols
            mDots(rowIndex, colIndex) = False
            mPowers(rowIndex, colIndex) = False
        Next colIndex
    Next rowIndex

    mRemainingPellets = 0
End Sub

Private Sub BuildBoard()
    Dim layout As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim cellValue As String
    Dim ghostIndex As Long
    Dim rowText As String

    layout = MazeLayoutRows()
    mRemainingPellets = 0
    ghostIndex = 0

    For rowIndex = 1 To BoardRows
        rowText = CStr(layout(rowIndex - 1))
        For colIndex = 1 To BoardCols
            cellValue = Mid$(rowText, colIndex, 1)
            mWalls(rowIndex, colIndex) = (cellValue = "#")
            mDots(rowIndex, colIndex) = (cellValue = ".")
            mPowers(rowIndex, colIndex) = (cellValue = "o")

            If mDots(rowIndex, colIndex) Or mPowers(rowIndex, colIndex) Then
                mRemainingPellets = mRemainingPellets + 1
            End If

            Select Case cellValue
            Case "P"
                mPlayerSpawnRow = rowIndex
                mPlayerSpawnCol = colIndex
            Case "A", "B", "C", "D"
                ghostIndex = ghostIndex + 1
                mGhostSpawnRow(ghostIndex) = rowIndex
                mGhostSpawnCol(ghostIndex) = colIndex
            End Select
        Next colIndex
    Next rowIndex
End Sub

Private Sub ResetActors()
    Dim ghostIndex As Long

    mPlayerRowValue = mPlayerSpawnRow
    mPlayerColValue = mPlayerSpawnCol
    mPlayerDirectionValue = MazeDirectionNone
    mRequestedDirection = MazeDirectionNone
    mPlayerMouthOpen = True
    mFrightenedTicks = 0
    mPowerBoostReady = False

    For ghostIndex = 1 To 4
        mGhostRow(ghostIndex) = mGhostSpawnRow(ghostIndex)
        mGhostCol(ghostIndex) = mGhostSpawnCol(ghostIndex)
        mGhostDirection(ghostIndex) = MazeDirectionLeft
    Next ghostIndex
End Sub

Private Sub AdvancePlayer()
    Dim previousRow As Long
    Dim previousCol As Long

    If CanMove(mPlayerRowValue, mPlayerColValue, mRequestedDirection) Then
        mPlayerDirectionValue = mRequestedDirection
    ElseIf Not CanMove(mPlayerRowValue, mPlayerColValue, mPlayerDirectionValue) Then
        mPlayerDirectionValue = MazeDirectionNone
    End If

    If mPlayerDirectionValue <> MazeDirectionNone Then
        previousRow = mPlayerRowValue
        previousCol = mPlayerColValue
        MoveActor mPlayerRowValue, mPlayerColValue, mPlayerDirectionValue
        If previousRow <> mPlayerRowValue Or previousCol <> mPlayerColValue Then
            mPlayerMouthOpen = Not mPlayerMouthOpen
        End If
    End If
End Sub

Private Sub AdvanceGhosts()
    Dim ghostIndex As Long
    Dim nextDirection As Long

    For ghostIndex = 1 To 4
        nextDirection = ChooseGhostDirection(ghostIndex)
        mGhostDirection(ghostIndex) = nextDirection
        If nextDirection <> MazeDirectionNone Then
            MoveActor mGhostRow(ghostIndex), mGhostCol(ghostIndex), nextDirection
        End If
    Next ghostIndex
End Sub

Private Function ChooseGhostDirection(ByVal ghostIndex As Long) As Long
    Dim candidate(1 To 4) As Long
    Dim candidateCount As Long
    Dim directionValue As Long
    Dim bestScore As Long
    Dim scoreValue As Long
    Dim bestDirection As Long
    Dim testRow As Long
    Dim testCol As Long
    Dim reverseDirection As Long
    Dim firstChoice As Boolean

    reverseDirection = OppositeDirection(mGhostDirection(ghostIndex))

    For directionValue = MazeDirectionLeft To MazeDirectionDown
        If CanMove(mGhostRow(ghostIndex), mGhostCol(ghostIndex), directionValue) Then
            candidateCount = candidateCount + 1
            candidate(candidateCount) = directionValue
        End If
    Next directionValue

    If candidateCount = 0 Then
        ChooseGhostDirection = MazeDirectionNone
        Exit Function
    End If

    bestDirection = candidate(1)
    firstChoice = True

    For directionValue = 1 To candidateCount
        If Not (candidateCount > 1 And candidate(directionValue) = reverseDirection) Then
            testRow = NextRow(mGhostRow(ghostIndex), candidate(directionValue))
            testCol = NextCol(mGhostCol(ghostIndex), candidate(directionValue))
            scoreValue = Abs(testRow - mPlayerRowValue) + Abs(testCol - mPlayerColValue)

            If mFrightenedTicks > 0 Then
                If firstChoice Or scoreValue > bestScore Then
                    bestScore = scoreValue
                    bestDirection = candidate(directionValue)
                    firstChoice = False
                End If
            ElseIf firstChoice Or scoreValue < bestScore Then
                bestScore = scoreValue
                bestDirection = candidate(directionValue)
                firstChoice = False
            End If
        End If
    Next directionValue

    ChooseGhostDirection = bestDirection
End Function

Private Sub HandleCurrentTile()
    If mDots(mPlayerRowValue, mPlayerColValue) Then
        mDots(mPlayerRowValue, mPlayerColValue) = False
        mRemainingPellets = mRemainingPellets - 1
        mScore = mScore + DotScore
        mStatusText = "Keep moving"
        RegisterConsumedTile mPlayerRowValue, mPlayerColValue, False
    ElseIf mPowers(mPlayerRowValue, mPlayerColValue) Then
        mPowers(mPlayerRowValue, mPlayerColValue) = False
        mRemainingPellets = mRemainingPellets - 1
        mScore = mScore + PowerScore
        mFrightenedTicks = FrightenedDurationTicks
        mPowerBoostReady = True
        mStatusText = "Power up"
        RegisterConsumedTile mPlayerRowValue, mPlayerColValue, True
    End If
End Sub

Private Sub ApplyPowerBoostMovement()
    If mGameOver Then
        Exit Sub
    End If

    If mFrightenedTicks <= 0 Then
        mPowerBoostReady = False
        Exit Sub
    End If

    If Not mPowerBoostReady Then
        mPowerBoostReady = True
        Exit Sub
    End If

    mPowerBoostReady = False
    AdvancePlayer
    HandleCurrentTile
    ResolveGhostCollisions
End Sub

Private Sub ResolveGhostCollisions()
    Dim ghostIndex As Long

    For ghostIndex = 1 To 4
        If mGhostRow(ghostIndex) = mPlayerRowValue And mGhostCol(ghostIndex) = mPlayerColValue Then
            If mFrightenedTicks > 0 Then
                mScore = mScore + GhostScore
                mGhostRow(ghostIndex) = mGhostSpawnRow(ghostIndex)
                mGhostCol(ghostIndex) = mGhostSpawnCol(ghostIndex)
                mGhostDirection(ghostIndex) = MazeDirectionLeft
                mStatusText = "Ghost eaten"
            Else
                mLives = mLives - 1
                If mLives <= 0 Then
                    mGameOver = True
                    mPlayerWon = False
                    mStatusText = "Game over"
                Else
                    mStatusText = "Life lost"
                    ResetActors
                End If
                Exit Sub
            End If
        End If
    Next ghostIndex
End Sub

Private Sub MoveActor(ByRef rowValue As Long, ByRef colValue As Long, ByVal directionValue As Long)
    rowValue = NextRow(rowValue, directionValue)
    colValue = NextCol(colValue, directionValue)
End Sub

Private Function CanMove(ByVal rowIndex As Long, ByVal colIndex As Long, ByVal directionValue As Long) As Boolean
    Dim targetRow As Long
    Dim targetCol As Long

    If directionValue = MazeDirectionNone Then
        CanMove = False
        Exit Function
    End If

    targetRow = NextRow(rowIndex, directionValue)
    targetCol = NextCol(colIndex, directionValue)

    If targetRow < 1 Or targetRow > BoardRows Or targetCol < 1 Or targetCol > BoardCols Then
        CanMove = False
        Exit Function
    End If

    CanMove = Not mWalls(targetRow, targetCol)
End Function

Private Function NextRow(ByVal rowIndex As Long, ByVal directionValue As Long) As Long
    Select Case directionValue
    Case MazeDirectionUp
        NextRow = rowIndex - 1
    Case MazeDirectionDown
        NextRow = rowIndex + 1
    Case Else
        NextRow = rowIndex
    End Select
End Function

Private Function NextCol(ByVal colIndex As Long, ByVal directionValue As Long) As Long
    Select Case directionValue
    Case MazeDirectionLeft
        NextCol = colIndex - 1
    Case MazeDirectionRight
        NextCol = colIndex + 1
    Case Else
        NextCol = colIndex
    End Select
End Function

Private Function OppositeDirection(ByVal directionValue As Long) As Long
    Select Case directionValue
    Case MazeDirectionLeft
        OppositeDirection = MazeDirectionRight
    Case MazeDirectionUp
        OppositeDirection = MazeDirectionDown
    Case MazeDirectionRight
        OppositeDirection = MazeDirectionLeft
    Case MazeDirectionDown
        OppositeDirection = MazeDirectionUp
    Case Else
        OppositeDirection = MazeDirectionNone
    End Select
End Function

Private Sub RegisterConsumedTile(ByVal rowIndex As Long, ByVal colIndex As Long, ByVal wasPower As Boolean)
    mConsumedRow = rowIndex
    mConsumedCol = colIndex
    mConsumedPower = wasPower
End Sub

Private Sub ClearConsumedMarker()
    mConsumedRow = 0
    mConsumedCol = 0
    mConsumedPower = False
End Sub

Private Function MazeLayoutRows() As Variant
    MazeLayoutRows = Array( _
    "#################", _
    "#o....#...#....o#", _
    "#.##.#.#.#.#.##.#", _
    "#...............#", _
    "#.###.#####.###.#", _
    "#...#...#...#...#", _
    "###.#.#.#.#.#.###", _
    "#...#...ABC...#.#", _
    "#.###.#.....#.###", _
    "#.....#..D..#...#", _
    "#.###.###.###...#", _
    "#...............#", _
    "#.##.#.###.#.##.#", _
    "#o....#...#....o#", _
    "#.###.#.#.#.###.#", _
    "#......P........#", _
    "#################")
End Function
