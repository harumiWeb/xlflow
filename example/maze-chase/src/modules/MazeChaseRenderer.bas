Attribute VB_Name = "MazeChaseRenderer"
Option Explicit

Private Const DynamicPrefix As String = "mc_"
Private Const SpriteGridSize As Long = 8

Private mForm As Object
Private mWallControls() As Object
Private mDotControls() As Object
Private mPowerControls() As Object
Private mGhostControls(1 To 4) As Object
Private mPlayerControl As Object
Private mScoreCaption As Object
Private mScoreValue As Object
Private mLivesCaption As Object
Private mLivesValue As Object
Private mStatusValue As Object
Private mSceneBuilt As Boolean
Private mLastScore As Long
Private mLastLives As Long
Private mLastStatus As String
Private mLastPlayerDirection As Long
Private mLastPlayerMouthOpen As Boolean

Public Sub AttachForm(ByVal formInstance As Object)
    Set mForm = formInstance
    mSceneBuilt = False
    ReDim mWallControls(1 To MazeChaseGame.Rows(), 1 To MazeChaseGame.Cols())
    ReDim mDotControls(1 To MazeChaseGame.Rows(), 1 To MazeChaseGame.Cols())
    ReDim mPowerControls(1 To MazeChaseGame.Rows(), 1 To MazeChaseGame.Cols())
End Sub

Public Sub DetachForm()
    Set mForm = Nothing
    Set mPlayerControl = Nothing
    Set mScoreCaption = Nothing
    Set mScoreValue = Nothing
    Set mLivesCaption = Nothing
    Set mLivesValue = Nothing
    Set mStatusValue = Nothing
    mSceneBuilt = False
End Sub

Public Function HasForm() As Boolean
    HasForm = Not mForm Is Nothing
End Function

Public Sub BuildScene()
    Dim rowIndex As Long
    Dim colIndex As Long

    If mForm Is Nothing Then
        Exit Sub
    End If

    ClearDynamicControls
    ConfigureForm
    CreateHud

    For rowIndex = 1 To MazeChaseGame.Rows()
        For colIndex = 1 To MazeChaseGame.Cols()
            If MazeChaseGame.IsWallAt(rowIndex, colIndex) Then
                Set mWallControls(rowIndex, colIndex) = CreateBlockControl("wall_" & rowIndex & "_" & colIndex, TileLeft(colIndex), TileTop(rowIndex), MazeChaseGame.CellSize(), MazeChaseGame.CellSize(), RGB(40, 50, 220), "")
            ElseIf MazeChaseGame.HasPowerAt(rowIndex, colIndex) Then
                Set mPowerControls(rowIndex, colIndex) = CreatePowerControl("power_" & rowIndex & "_" & colIndex, TileLeft(colIndex), TileTop(rowIndex), MazeChaseGame.CellSize(), MazeChaseGame.CellSize())
            ElseIf MazeChaseGame.HasDotAt(rowIndex, colIndex) Then
                Set mDotControls(rowIndex, colIndex) = CreateBlockControl("dot_" & rowIndex & "_" & colIndex, TileLeft(colIndex) + 7, TileTop(rowIndex) + 7, 5, 5, RGB(255, 220, 190), "")
            End If
        Next colIndex
    Next rowIndex

    Set mPlayerControl = CreatePlayerSprite("player")
    Set mGhostControls(1) = CreateGhostSprite("ghost_1", GhostColor(1))
    Set mGhostControls(2) = CreateGhostSprite("ghost_2", GhostColor(2))
    Set mGhostControls(3) = CreateGhostSprite("ghost_3", GhostColor(3))
    Set mGhostControls(4) = CreateGhostSprite("ghost_4", GhostColor(4))

    mLastScore = -1
    mLastLives = -1
    mLastStatus = vbNullString
    mLastPlayerDirection = -1
    mLastPlayerMouthOpen = Not MazeChaseGame.PlayerMouthOpen()
    mSceneBuilt = True
    RenderFrame
End Sub

Public Sub RenderFrame()
    Dim ghostIndex As Long
    Dim consumedRowIndex As Long
    Dim consumedColIndex As Long

    If Not mSceneBuilt Then
        Exit Sub
    End If

    UpdatePlayerSprite
    UpdateActorPosition mPlayerControl, MazeChaseGame.PlayerRow(), MazeChaseGame.PlayerCol()

    For ghostIndex = 1 To MazeChaseGame.GhostCount()
        If MazeChaseGame.IsGhostFrightened() Then
            TintGhostSprite mGhostControls(ghostIndex), RGB(80, 120, 255)
        Else
            TintGhostSprite mGhostControls(ghostIndex), GhostColor(ghostIndex)
        End If
        UpdateActorPosition mGhostControls(ghostIndex), MazeChaseGame.GhostRow(ghostIndex), MazeChaseGame.GhostCol(ghostIndex)
    Next ghostIndex

    If MazeChaseGame.HasConsumedTile() Then
        consumedRowIndex = MazeChaseGame.ConsumedRow()
        consumedColIndex = MazeChaseGame.ConsumedCol()
        HidePelletControl consumedRowIndex, consumedColIndex
        MazeChaseGame.AcknowledgeConsumedTile
    End If

    SyncPelletVisibility
    UpdateHud
    mForm.Repaint
End Sub

Private Sub ConfigureForm()
    With mForm
        .Caption = "Maze Chase"
        .Width = MazeChaseGame.FormWidth()
        .Height = MazeChaseGame.FormHeight()
        .BackColor = RGB(0, 0, 0)
    End With
End Sub

Private Sub CreateHud()
    Set mScoreCaption = CreateTextControl("score_caption", 20, 8, 70, 14, "1UP", RGB(255, 255, 255), 12, True)
    Set mScoreValue = CreateTextControl("score_value", 20, 22, 84, 16, "0", RGB(255, 255, 255), 14, True)
    Set mLivesCaption = CreateTextControl("lives_caption", 215, 8, 56, 14, "LIVES", RGB(255, 255, 255), 12, True)
    Set mLivesValue = CreateTextControl("lives_value", 284, 8, 80, 14, "3", RGB(255, 240, 40), 12, True)
    Set mStatusValue = CreateTextControl("status_value", 14, MazeChaseGame.BoardTop() + MazeChaseGame.Rows() * MazeChaseGame.CellSize() + 4, 340, 18, "Collect every dot", RGB(255, 255, 255), 11, False)
End Sub

Private Function CreateBlockControl(ByVal suffix As String, ByVal leftValue As Long, ByVal topValue As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal backColorValue As Long, ByVal captionValue As String) As Object
    Dim labelControl As Object

    Set labelControl = mForm.Controls.Add("Forms.Label.1", DynamicPrefix & suffix, True)
    With labelControl
        .Caption = captionValue
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
        .BackStyle = 1
        .BackColor = backColorValue
        .BorderStyle = 0
        .SpecialEffect = 0
    End With

    Set CreateBlockControl = labelControl
End Function

Private Function CreateTextControl(ByVal suffix As String, ByVal leftValue As Long, ByVal topValue As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal captionValue As String, ByVal foreColorValue As Long, ByVal fontSizeValue As Long, ByVal boldValue As Boolean) As Object
    Dim labelControl As Object

    Set labelControl = mForm.Controls.Add("Forms.Label.1", DynamicPrefix & suffix, True)
    With labelControl
        .Caption = captionValue
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
        .BackStyle = 0
        .ForeColor = foreColorValue
        .BorderStyle = 0
        .SpecialEffect = 0
        .Font.Size = fontSizeValue
        .Font.Bold = boldValue
    End With

    Set CreateTextControl = labelControl
End Function

Private Function CreatePowerControl(ByVal suffix As String, ByVal leftValue As Long, ByVal topValue As Long, ByVal widthValue As Long, ByVal heightValue As Long) As Object
    Dim frameControl As Object

    Set frameControl = mForm.Controls.Add("Forms.Frame.1", DynamicPrefix & suffix, True)
    With frameControl
        .Caption = ""
        .Left = leftValue + 2
        .Top = topValue + 2
        .Width = widthValue - 4
        .Height = heightValue - 4
        .BackColor = RGB(0, 0, 0)
        .SpecialEffect = 0
    End With

    AddPowerPart frameControl, 4, 1, 4, 2, RGB(90, 200, 90), "leaf"
    AddPowerPart frameControl, 8, 1, 4, 2, RGB(90, 200, 90), "leaf"
    AddPowerPart frameControl, 6, 3, 1, 4, RGB(70, 180, 70), "stem"
    AddPowerPart frameControl, 9, 3, 1, 4, RGB(70, 180, 70), "stem"
    AddPowerPart frameControl, 2, 8, 5, 5, RGB(220, 40, 50), "fruit"
    AddPowerPart frameControl, 9, 8, 5, 5, RGB(220, 40, 50), "fruit"
    AddPowerPart frameControl, 3, 9, 1, 1, RGB(255, 190, 190), "shine"
    AddPowerPart frameControl, 10, 9, 1, 1, RGB(255, 190, 190), "shine"

    Set CreatePowerControl = frameControl
End Function

Private Sub AddPowerPart(ByVal parentControl As Object, ByVal leftValue As Long, ByVal topValue As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal backColorValue As Long, ByVal tagValue As String)
    Dim partControl As Object

    Set partControl = parentControl.Controls.Add("Forms.Label.1", parentControl.name & "_" & tagValue & "_" & CStr(parentControl.Controls.count), True)
    With partControl
        .Caption = ""
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
        .BackStyle = 1
        .BackColor = backColorValue
        .BorderStyle = 0
        .SpecialEffect = 0
    End With
End Sub

Private Function CreatePlayerSprite(ByVal suffix As String) As Object
    Dim spriteContainer As Object

    Set spriteContainer = CreateSpriteContainer(suffix)
    DrawPlayerSprite spriteContainer
    Set CreatePlayerSprite = spriteContainer
End Function

Private Function CreateGhostSprite(ByVal suffix As String, ByVal bodyColor As Long) As Object
    Dim spriteContainer As Object

    Set spriteContainer = CreateSpriteContainer(suffix)
    AddSpritePixels spriteContainer, GhostBodyPattern(), bodyColor, "body"
    AddSpritePixels spriteContainer, GhostEyePattern(), RGB(255, 255, 255), "eye"
    AddSpritePixels spriteContainer, GhostPupilPattern(), RGB(40, 90, 255), "pupil"
    Set CreateGhostSprite = spriteContainer
End Function

Private Function CreateSpriteContainer(ByVal suffix As String) As Object
    Dim frameControl As Object

    Set frameControl = mForm.Controls.Add("Forms.Frame.1", DynamicPrefix & suffix, True)
    With frameControl
        .Caption = ""
        .Left = 0
        .Top = 0
        .Width = SpriteSize()
        .Height = SpriteSize()
        .BackColor = RGB(0, 0, 0)
        .SpecialEffect = 0
    End With

    Set CreateSpriteContainer = frameControl
End Function

Private Sub AddSpritePixels(ByVal spriteContainer As Object, ByVal patternRows As Variant, ByVal colorValue As Long, ByVal roleName As String)
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim pixelControl As Object
    Dim patternRow As String
    Dim pixelName As String

    For rowIndex = LBound(patternRows) To UBound(patternRows)
        patternRow = CStr(patternRows(rowIndex))
        For colIndex = 1 To Len(patternRow)
            If Mid$(patternRow, colIndex, 1) = "1" Then
                pixelName = "px_" & Right$(spriteContainer.name, 8) & "_" & CStr(rowIndex) & "_" & CStr(colIndex)
                Set pixelControl = spriteContainer.Controls.Add("Forms.Label.1", pixelName, True)
                With pixelControl
                    .Caption = ""
                    .Left = (colIndex - 1) * PixelSize()
                    .Top = rowIndex * PixelSize()
                    .Width = PixelSize()
                    .Height = PixelSize()
                    .BackStyle = 1
                    .BackColor = colorValue
                    .BorderStyle = 0
                    .SpecialEffect = 0
                    .Tag = roleName
                End With
            End If
        Next colIndex
    Next rowIndex
End Sub

Private Sub DrawPlayerSprite(ByVal spriteContainer As Object)
    ClearSpritePixels spriteContainer
    AddSpritePixels spriteContainer, PlayerPatternForState(MazeChaseGame.PlayerDirection(), MazeChaseGame.PlayerMouthOpen()), RGB(255, 232, 40), "body"
End Sub

Private Sub UpdatePlayerSprite()
    If mPlayerControl Is Nothing Then
        Exit Sub
    End If

    If mLastPlayerDirection <> MazeChaseGame.PlayerDirection() Or mLastPlayerMouthOpen <> MazeChaseGame.PlayerMouthOpen() Then
        DrawPlayerSprite mPlayerControl
        mLastPlayerDirection = MazeChaseGame.PlayerDirection()
        mLastPlayerMouthOpen = MazeChaseGame.PlayerMouthOpen()
    End If
End Sub

Private Sub UpdateActorPosition(ByVal actorControl As Object, ByVal rowIndex As Long, ByVal colIndex As Long)
    With actorControl
        .Left = TileLeft(colIndex) + 2
        .Top = TileTop(rowIndex) + 2
    End With
End Sub

Private Sub TintGhostSprite(ByVal spriteContainer As Object, ByVal bodyColor As Long)
    Dim controlIndex As Long
    Dim pixelControl As Object

    For controlIndex = 0 To spriteContainer.Controls.count - 1
        Set pixelControl = spriteContainer.Controls.Item(controlIndex)
        If pixelControl.Tag = "body" Then
            pixelControl.BackColor = bodyColor
        End If
    Next controlIndex
End Sub

Private Sub ClearSpritePixels(ByVal spriteContainer As Object)
    Dim controlIndex As Long

    For controlIndex = spriteContainer.Controls.count - 1 To 0 Step -1
        spriteContainer.Controls.Remove controlIndex
    Next controlIndex
End Sub

Private Sub HidePelletControl(ByVal rowIndex As Long, ByVal colIndex As Long)
    If Not mDotControls(rowIndex, colIndex) Is Nothing Then
        mDotControls(rowIndex, colIndex).Visible = False
    End If

    If Not mPowerControls(rowIndex, colIndex) Is Nothing Then
        mPowerControls(rowIndex, colIndex).Visible = False
    End If
End Sub

Private Sub SyncPelletVisibility()
    Dim rowIndex As Long
    Dim colIndex As Long

    For rowIndex = 1 To MazeChaseGame.Rows()
        For colIndex = 1 To MazeChaseGame.Cols()
            If Not mDotControls(rowIndex, colIndex) Is Nothing Then
                mDotControls(rowIndex, colIndex).Visible = MazeChaseGame.HasDotAt(rowIndex, colIndex)
            End If

            If Not mPowerControls(rowIndex, colIndex) Is Nothing Then
                mPowerControls(rowIndex, colIndex).Visible = MazeChaseGame.HasPowerAt(rowIndex, colIndex)
            End If
        Next colIndex
    Next rowIndex
End Sub

Private Sub UpdateHud()
    If mLastScore <> MazeChaseGame.scoreValue() Then
        mScoreValue.Caption = CStr(MazeChaseGame.scoreValue())
        mLastScore = MazeChaseGame.scoreValue()
    End If

    If mLastLives <> MazeChaseGame.LivesValue() Then
        mLivesValue.Caption = String$(MazeChaseGame.LivesValue(), "C")
        mLastLives = MazeChaseGame.LivesValue()
    End If

    If mLastStatus <> MazeChaseGame.StatusText() Then
        mStatusValue.Caption = MazeChaseGame.StatusText()
        mLastStatus = MazeChaseGame.StatusText()
    End If
End Sub

Private Sub ClearDynamicControls()
    Dim controlIndex As Long
    Dim controlName As String

    If mForm Is Nothing Then
        Exit Sub
    End If

    For controlIndex = mForm.Controls.count - 1 To 0 Step -1
        controlName = CStr(mForm.Controls.Item(controlIndex).name)
        If Left$(controlName, Len(DynamicPrefix)) = DynamicPrefix Then
            mForm.Controls.Remove controlName
        End If
    Next controlIndex
End Sub

Private Function TileLeft(ByVal colIndex As Long) As Long
    TileLeft = MazeChaseGame.BoardLeft() + (colIndex - 1) * MazeChaseGame.CellSize()
End Function

Private Function TileTop(ByVal rowIndex As Long) As Long
    TileTop = MazeChaseGame.BoardTop() + (rowIndex - 1) * MazeChaseGame.CellSize()
End Function

Private Function GhostColor(ByVal index As Long) As Long
    Select Case index
    Case 1
        GhostColor = RGB(255, 60, 60)
    Case 2
        GhostColor = RGB(255, 160, 200)
    Case 3
        GhostColor = RGB(60, 220, 255)
    Case Else
        GhostColor = RGB(255, 180, 40)
    End Select
End Function

Private Function PixelSize() As Long
    PixelSize = 2
End Function

Private Function SpriteSize() As Long
    SpriteSize = SpriteGridSize * PixelSize()
End Function

Private Function PlayerPattern() As Variant
    PlayerPattern = PlayerPatternForState(4, True)
End Function

Private Function PlayerPatternForState(ByVal directionValue As Long, ByVal mouthOpenValue As Boolean) As Variant
    Select Case directionValue
    Case 1
        If mouthOpenValue Then
            PlayerPatternForState = Array( _
            "00111100", _
            "01111110", _
            "00011111", _
            "00001111", _
            "00001111", _
            "00011111", _
            "01111110", _
            "00111100")
        Else
            PlayerPatternForState = Array( _
            "00111100", _
            "01111110", _
            "11111111", _
            "11111111", _
            "11111111", _
            "11111111", _
            "01111110", _
            "00111100")
        End If
    Case 2
        If mouthOpenValue Then
            PlayerPatternForState = Array( _
            "00011000", _
            "00111100", _
            "01111110", _
            "11111111", _
            "11111111", _
            "01100110", _
            "00100100", _
            "00000000")
        Else
            PlayerPatternForState = Array( _
            "00111100", _
            "01111110", _
            "11111111", _
            "11111111", _
            "11111111", _
            "11111111", _
            "01111110", _
            "00111100")
        End If
    Case 3, 0
        If mouthOpenValue Then
            PlayerPatternForState = Array( _
            "00111100", _
            "01111110", _
            "11111000", _
            "11110000", _
            "11110000", _
            "11111000", _
            "01111110", _
            "00111100")
        Else
            PlayerPatternForState = Array( _
            "00111100", _
            "01111110", _
            "11111111", _
            "11111111", _
            "11111111", _
            "11111111", _
            "01111110", _
            "00111100")
        End If
    Case Else
        If mouthOpenValue Then
            PlayerPatternForState = Array( _
            "00000000", _
            "00100100", _
            "01100110", _
            "11111111", _
            "11111111", _
            "01111110", _
            "00111100", _
            "00011000")
        Else
            PlayerPatternForState = Array( _
            "00111100", _
            "01111110", _
            "11111111", _
            "11111111", _
            "11111111", _
            "11111111", _
            "01111110", _
            "00111100")
        End If
    End Select
End Function

Private Function GhostBodyPattern() As Variant
    GhostBodyPattern = Array( _
    "00111100", _
    "01111110", _
    "11111111", _
    "11111111", _
    "11111111", _
    "11111111", _
    "11011011", _
    "10000001")
End Function

Private Function GhostEyePattern() As Variant
    GhostEyePattern = Array( _
    "00000000", _
    "00000000", _
    "00100100", _
    "01100110", _
    "00000000", _
    "00000000", _
    "00000000", _
    "00000000")
End Function

Private Function GhostPupilPattern() As Variant
    GhostPupilPattern = Array( _
    "00000000", _
    "00000000", _
    "00000000", _
    "00100100", _
    "00000000", _
    "00000000", _
    "00000000", _
    "00000000")
End Function
