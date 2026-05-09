Attribute VB_Name = "App"
Option Explicit

Private Const GAME_SHEET_NAME As String = "Tetris"
Private Const BOARD_ROWS As Long = 20
Private Const BOARD_COLUMNS As Long = 10
Private Const PIECE_COUNT As Long = 7
Private Const BLOCK_TEXT As String = "[]"
Private Const TICK_SECONDS As Double = 1#

Private gWorkbook As Workbook
Private gGameSheet As Worksheet
Private gBoard(1 To BOARD_ROWS, 1 To BOARD_COLUMNS) As Long
Private gShapeMaps(1 To PIECE_COUNT, 0 To 3) As String
Private gPieceColors(1 To PIECE_COUNT) As Long
Private gCurrentPiece As Long
Private gCurrentRotation As Long
Private gCurrentRow As Long
Private gCurrentColumn As Long
Private gNextTickTime As Date
Private gScore As Long
Private gLinesCleared As Long
Private gTickScheduled As Boolean
Private gIsRunning As Boolean
Private gInitialized As Boolean

Public Sub RunCore(ByVal wb As Workbook)
    EnsureInitialized

    Set gWorkbook = wb
    Set gGameSheet = ResolveGameSheet(wb)

    ConfigureSurface gGameSheet
    StartGame
End Sub

Public Sub Tick()
    If Not gIsRunning Then Exit Sub

    gTickScheduled = False

    If Not TryMovePiece(1, 0, gCurrentRotation) Then
        LockCurrentPiece
        ClearCompletedLines
        SpawnNextPiece
    End If

    RenderGame

    If gIsRunning Then
      ScheduleNextTick
    End If
End Sub

Public Sub MoveLeftKey()
    HandleMove 0, -1
End Sub

Public Sub MoveRightKey()
    HandleMove 0, 1
End Sub

Public Sub SoftDropKey()
    If Not gIsRunning Then Exit Sub

    If TryMovePiece(1, 0, gCurrentRotation) Then
        RenderGame
    Else
        LockCurrentPiece
        ClearCompletedLines
        SpawnNextPiece
        RenderGame
    End If
End Sub

Public Sub RotateKey()
    Dim candidateRotation As Long
    candidateRotation = (gCurrentRotation + 1) Mod 4

    AttemptRotation candidateRotation
End Sub

Public Sub HardDropKey()
    If Not gIsRunning Then Exit Sub

    Do While TryMovePiece(1, 0, gCurrentRotation)
    Loop

    LockCurrentPiece
    ClearCompletedLines
    SpawnNextPiece
    RenderGame
End Sub

Public Sub RestartKey()
    StartGame
End Sub

Public Sub HandleWorkbookBeforeClose()
    StopGameLoop
End Sub

Private Sub EnsureInitialized()
    If gInitialized Then Exit Sub

    Randomize

    gShapeMaps(1, 0) = "....XXXX........"
    gShapeMaps(1, 1) = "..X...X...X...X."
    gShapeMaps(1, 2) = "........XXXX...."
    gShapeMaps(1, 3) = ".X...X...X...X.."

    gShapeMaps(2, 0) = "X...XXX........."
    gShapeMaps(2, 1) = ".XX..X...X......"
    gShapeMaps(2, 2) = "........XXX...X."
    gShapeMaps(2, 3) = ".X...X..XX......"

    gShapeMaps(3, 0) = "..X.XXX........."
    gShapeMaps(3, 1) = ".X...X...XX....."
    gShapeMaps(3, 2) = "........XXX.X..."
    gShapeMaps(3, 3) = "XX...X...X......"

    gShapeMaps(4, 0) = ".XX..XX........."
    gShapeMaps(4, 1) = ".XX..XX........."
    gShapeMaps(4, 2) = ".XX..XX........."
    gShapeMaps(4, 3) = ".XX..XX........."

    gShapeMaps(5, 0) = ".XX.XX.........."
    gShapeMaps(5, 1) = ".X..XX..X......."
    gShapeMaps(5, 2) = "........XX..XX.."
    gShapeMaps(5, 3) = "...X..XX..X....."

    gShapeMaps(6, 0) = ".X..XXX........."
    gShapeMaps(6, 1) = ".X...XX..X......"
    gShapeMaps(6, 2) = "........XXX..X.."
    gShapeMaps(6, 3) = ".X..XX...X......"

    gShapeMaps(7, 0) = "XX...XX........."
    gShapeMaps(7, 1) = "..X..XX..X......"
    gShapeMaps(7, 2) = "........XX...XX."
    gShapeMaps(7, 3) = ".X..XX..X......."

    gPieceColors(1) = RGB(0, 176, 240)
    gPieceColors(2) = RGB(0, 112, 192)
    gPieceColors(3) = RGB(237, 125, 49)
    gPieceColors(4) = RGB(255, 192, 0)
    gPieceColors(5) = RGB(112, 173, 71)
    gPieceColors(6) = RGB(165, 165, 165)
    gPieceColors(7) = RGB(255, 0, 0)

    gInitialized = True
End Sub

Private Function ResolveGameSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, GAME_SHEET_NAME, vbTextCompare) = 0 Then
            Set ResolveGameSheet = ws
            Exit Function
        End If
    Next ws

    If wb.Worksheets.Count > 0 Then
        Set ws = wb.Worksheets(1)
        ws.Name = GAME_SHEET_NAME
        Set ResolveGameSheet = ws
        Exit Function
    End If

    Set ResolveGameSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ResolveGameSheet.Name = GAME_SHEET_NAME
End Function

Private Sub ConfigureSurface(ByVal ws As Worksheet)
    Dim boardRange As Range
    Dim infoRange As Range
    Dim rowIndex As Long
    Dim columnIndex As Long

    Set boardRange = ws.Range("B3:K22")
    Set infoRange = ws.Range("M3:P10")

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(15, 15, 18)
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 11
    ws.Cells.Font.Color = RGB(255, 255, 255)

    ws.Range("B1").Value2 = "TETRIS"
    ws.Range("B1").Font.Size = 20
    ws.Range("B1").Font.Bold = True

    ws.Range("M3").Value2 = "Controls"
    ws.Range("M4").Value2 = "Left / Right"
    ws.Range("N4").Value2 = "Move"
    ws.Range("M5").Value2 = "Up"
    ws.Range("N5").Value2 = "Rotate"
    ws.Range("M6").Value2 = "Down"
    ws.Range("N6").Value2 = "Soft drop"
    ws.Range("M7").Value2 = "Space"
    ws.Range("N7").Value2 = "Hard drop"
    ws.Range("M8").Value2 = "R"
    ws.Range("N8").Value2 = "Restart"

    ws.Range("M10").Value2 = "Score"
    ws.Range("M11").Value2 = "Lines"

    boardRange.Borders.LineStyle = xlContinuous
    boardRange.Borders.Color = RGB(64, 64, 64)
    boardRange.HorizontalAlignment = xlCenter
    boardRange.VerticalAlignment = xlCenter

    infoRange.Font.Bold = False
    ws.Range("M3:M11").Font.Bold = True

    For columnIndex = 2 To 11
        ws.Columns(columnIndex).ColumnWidth = 3.2
    Next columnIndex

    For rowIndex = 3 To 22
        ws.Rows(rowIndex).RowHeight = 20
    Next rowIndex
End Sub

Private Sub StartGame()
    Dim rowIndex As Long
    Dim columnIndex As Long

    StopGameLoop

    For rowIndex = 1 To BOARD_ROWS
        For columnIndex = 1 To BOARD_COLUMNS
            gBoard(rowIndex, columnIndex) = 0
        Next columnIndex
    Next rowIndex

    gScore = 0
    gLinesCleared = 0
    gIsRunning = True

    BindKeys
    SpawnNextPiece
    RenderGame

    If gIsRunning Then
        ScheduleNextTick
    End If
End Sub

Private Sub StopGameLoop()
    If gTickScheduled Then
        CancelScheduledTick
    End If

    gIsRunning = False
    UnbindKeys
End Sub

Private Sub BindKeys()
    Application.OnKey "{LEFT}", "App.MoveLeftKey"
    Application.OnKey "{RIGHT}", "App.MoveRightKey"
    Application.OnKey "{DOWN}", "App.SoftDropKey"
    Application.OnKey "{UP}", "App.RotateKey"
    Application.OnKey " ", "App.HardDropKey"
    Application.OnKey "r", "App.RestartKey"
    Application.OnKey "R", "App.RestartKey"
End Sub

Private Sub UnbindKeys()
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{UP}"
    Application.OnKey " "
    Application.OnKey "r"
    Application.OnKey "R"
End Sub

Private Sub ScheduleNextTick()
    gNextTickTime = Now + (TICK_SECONDS / 86400#)
    gTickScheduled = True
    Application.OnTime EarliestTime:=gNextTickTime, Procedure:="App.Tick", Schedule:=True
End Sub

Private Sub CancelScheduledTick()
    On Error GoTo CancelFailed

    Application.OnTime EarliestTime:=gNextTickTime, Procedure:="App.Tick", Schedule:=False

CancelDone:
    gTickScheduled = False
    Exit Sub
CancelFailed:
    Resume CancelDone
End Sub

Private Sub SpawnNextPiece()
    gCurrentPiece = RandomPieceIndex()
    gCurrentRotation = 0
    gCurrentRow = 1
    gCurrentColumn = 4

    If PieceCollides(gCurrentRow, gCurrentColumn, gCurrentRotation) Then
        gIsRunning = False
    End If
End Sub

Private Function RandomPieceIndex() As Long
    RandomPieceIndex = Int(Rnd * PIECE_COUNT) + 1
End Function

Private Function TryMovePiece(ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal nextRotation As Long) As Boolean
    Dim candidateRow As Long
    Dim candidateColumn As Long

    candidateRow = gCurrentRow + rowOffset
    candidateColumn = gCurrentColumn + columnOffset

    If PieceCollides(candidateRow, candidateColumn, nextRotation) Then
      Exit Function
    End If

    gCurrentRow = candidateRow
    gCurrentColumn = candidateColumn
    gCurrentRotation = nextRotation
    TryMovePiece = True
End Function

Private Sub AttemptRotation(ByVal candidateRotation As Long)
    If Not gIsRunning Then Exit Sub

    If TryMovePiece(0, 0, candidateRotation) Then
      RenderGame
      Exit Sub
    End If

    If TryMovePiece(0, -1, candidateRotation) Then
      RenderGame
      Exit Sub
    End If

    If TryMovePiece(0, 1, candidateRotation) Then
      RenderGame
      Exit Sub
    End If

    If TryMovePiece(0, -2, candidateRotation) Then
      RenderGame
      Exit Sub
    End If

    If TryMovePiece(0, 2, candidateRotation) Then
      RenderGame
      Exit Sub
    End If
End Sub

Private Function PieceCollides(ByVal targetRow As Long, ByVal targetColumn As Long, ByVal targetRotation As Long) As Boolean
    Dim localRow As Long
    Dim localColumn As Long
    Dim boardRow As Long
    Dim boardColumn As Long

    For localRow = 1 To 4
        For localColumn = 1 To 4
            If PieceCellFilled(gCurrentPiece, targetRotation, localRow, localColumn) Then
                boardRow = targetRow + localRow - 1
                boardColumn = targetColumn + localColumn - 1
                If boardColumn < 1 Or boardColumn > BOARD_COLUMNS Then
                    PieceCollides = True
                    Exit Function
                End If
                If boardRow > BOARD_ROWS Then
                    PieceCollides = True
                    Exit Function
                End If
                If boardRow >= 1 Then
                    If gBoard(boardRow, boardColumn) <> 0 Then
                        PieceCollides = True
                        Exit Function
                    End If
                End If
            End If
        Next localColumn
    Next localRow
End Function

Private Sub LockCurrentPiece()
    Dim localRow As Long
    Dim localColumn As Long
    Dim boardRow As Long
    Dim boardColumn As Long

    For localRow = 1 To 4
        For localColumn = 1 To 4
            If PieceCellFilled(gCurrentPiece, gCurrentRotation, localRow, localColumn) Then
                boardRow = gCurrentRow + localRow - 1
                boardColumn = gCurrentColumn + localColumn - 1
                If boardRow >= 1 And boardRow <= BOARD_ROWS Then
                    gBoard(boardRow, boardColumn) = gCurrentPiece
                End If
            End If
        Next localColumn
    Next localRow
End Sub

Private Sub ClearCompletedLines()
    Dim rowIndex As Long
    Dim cleared As Long

    rowIndex = BOARD_ROWS

    Do While rowIndex >= 1
        If RowIsComplete(rowIndex) Then
            CollapseBoard rowIndex
            cleared = cleared + 1
        Else
            rowIndex = rowIndex - 1
        End If
    Loop

    If cleared > 0 Then
        gLinesCleared = gLinesCleared + cleared
        gScore = gScore + (100 * cleared * cleared)
    End If
End Sub

Private Function RowIsComplete(ByVal rowIndex As Long) As Boolean
    Dim columnIndex As Long

    RowIsComplete = True

    For columnIndex = 1 To BOARD_COLUMNS
        If gBoard(rowIndex, columnIndex) = 0 Then
            RowIsComplete = False
            Exit Function
        End If
    Next columnIndex
End Function

Private Sub CollapseBoard(ByVal clearedRow As Long)
    Dim rowIndex As Long
    Dim columnIndex As Long

    For rowIndex = clearedRow To 2 Step -1
        For columnIndex = 1 To BOARD_COLUMNS
            gBoard(rowIndex, columnIndex) = gBoard(rowIndex - 1, columnIndex)
        Next columnIndex
    Next rowIndex

    For columnIndex = 1 To BOARD_COLUMNS
        gBoard(1, columnIndex) = 0
    Next columnIndex
End Sub

Private Sub RenderGame()
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim cell As Range
    Dim boardValue As Long
    Dim screenUpdatingWasEnabled As Boolean

    If gGameSheet Is Nothing Then Exit Sub

    screenUpdatingWasEnabled = Application.ScreenUpdating
    On Error GoTo CleanFail

    Application.ScreenUpdating = False

    For rowIndex = 1 To BOARD_ROWS
        For columnIndex = 1 To BOARD_COLUMNS
            Set cell = gGameSheet.Cells(rowIndex + 2, columnIndex + 1)
            boardValue = EffectiveBoardValue(rowIndex, columnIndex)
            If boardValue = 0 Then
                cell.Value2 = vbNullString
                cell.Interior.Color = RGB(28, 28, 34)
            Else
                cell.Value2 = BLOCK_TEXT
                cell.Interior.Color = gPieceColors(boardValue)
            End If
        Next columnIndex
    Next rowIndex

    gGameSheet.Range("N10").Value2 = gScore
    gGameSheet.Range("N11").Value2 = gLinesCleared

    If gIsRunning Then
        gGameSheet.Range("B24").Value2 = "Playing"
    Else
        gGameSheet.Range("B24").Value2 = "Game Over - press R to restart"
    End If

CleanExit:
    Application.ScreenUpdating = screenUpdatingWasEnabled
    Exit Sub

CleanFail:
    Application.ScreenUpdating = screenUpdatingWasEnabled
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

Private Function EffectiveBoardValue(ByVal rowIndex As Long, ByVal columnIndex As Long) As Long
    If CellBelongsToCurrentPiece(rowIndex, columnIndex) Then
        EffectiveBoardValue = gCurrentPiece
    Else
        EffectiveBoardValue = gBoard(rowIndex, columnIndex)
    End If
End Function

Private Function CellBelongsToCurrentPiece(ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
    Dim localRow As Long
    Dim localColumn As Long

    If Not gIsRunning Then Exit Function

    For localRow = 1 To 4
        For localColumn = 1 To 4
            If PieceCellFilled(gCurrentPiece, gCurrentRotation, localRow, localColumn) Then
                If gCurrentRow + localRow - 1 = rowIndex And gCurrentColumn + localColumn - 1 = columnIndex Then
                    CellBelongsToCurrentPiece = True
                    Exit Function
                End If
            End If
        Next localColumn
    Next localRow
End Function

Private Function PieceCellFilled(ByVal pieceIndex As Long, ByVal rotationIndex As Long, ByVal localRow As Long, ByVal localColumn As Long) As Boolean
    Dim position As Long

    position = ((localRow - 1) * 4) + localColumn
    PieceCellFilled = Mid$(gShapeMaps(pieceIndex, rotationIndex), position, 1) = "X"
End Function

Private Sub HandleMove(ByVal rowOffset As Long, ByVal columnOffset As Long)
    If Not gIsRunning Then Exit Sub

    If TryMovePiece(rowOffset, columnOffset, gCurrentRotation) Then
        RenderGame
        Exit Sub
    End If
End Sub
