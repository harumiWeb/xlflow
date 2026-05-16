Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
  Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

Private Const PLAYER_SPRITE_PIXEL_SIZE As Long = 4
Private Const ENEMY_SPRITE_PIXEL_SIZE As Long = 3
Private Const PLAYER_FLASH_INTERVAL_SECONDS As Double = 0.08
Private Const LIFE_ICON_LEFT As Long = 124
Private Const LIFE_ICON_TOP As Long = 248
Private Const LIFE_ICON_SPACING As Long = 18

Private mPlayer As MSForms.Label
Private mPlayerPixels() As MSForms.Label
Private mPlayerBullets() As MSForms.Label
Private mEnemyBullet As MSForms.Label
Private mEnemyLabels() As MSForms.Label
Private mEnemyPixels() As MSForms.Label
Private mEnemyAlive() As Boolean
Private mEnemyType() As String ' "SQUID"/"CRAB"/"OCTOPUS"
Private mLifeIcons() As MSForms.Label
' UFO用
Private mUfoLabel As MSForms.Label
Private mUfoPixels() As MSForms.Label
Private mUfoActive As Boolean
Private mUfoDirection As Long '1=右, -1=左
Private mUfoScore As Long
Private mUfoNextSpawnTick As Double

Private mSceneBuilt As Boolean
Private mGameRunning As Boolean
Private mLoopActive As Boolean
Private mGameEnded As Boolean
Private mPrevEscapeDown As Boolean
Private mEnemyDirection As Long
Private mLives As Long
Private mPlayerShotClock As Double
Private mEnemyShotClock As Double
Private mPlayerInvulnerable As Boolean
Private mPlayerInvulnerableClock As Double
Private mScore As Long

Private Sub UserForm_Initialize()
  Randomize
  ConfigureUi
  EnsureSceneBuilt
  ResetGame
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  StopGameLoop True
End Sub

Private Sub UserForm_Terminate()
  mGameRunning = False
  mLoopActive = False
End Sub

Private Sub cmdReset_Click()
  StopGameLoop False
  ResetGame
End Sub

Private Sub cmdStart_Click()
  If mLoopActive Then Exit Sub
  If mGameEnded Then ResetGame
  StartGameLoop
End Sub

Private Sub cmdStop_Click()
  StopGameLoop False
End Sub

Public Sub RequestStop()
  StopGameLoop True
End Sub

Private Sub ConfigureUi()
  Dim hudTop As Long
  Dim buttonTop As Long
  Dim buttonWidth As Long
  Dim buttonGap As Long
  Dim rightEdge As Long

  Me.Caption = "Space Invader"
  Me.backColor = RGB(18, 18, 18)
  Me.Width = GameAreaWidth + 60
  Me.Height = GameAreaHeight + 110

  With Me.fraGameArea
    .Left = 12
    .Top = 12
    .Width = GameAreaWidth
    .Height = GameAreaHeight
  End With

  hudTop = Me.fraGameArea.Top + Me.fraGameArea.Height + 10
  buttonTop = hudTop - 2
  buttonWidth = 44
  buttonGap = 6
  rightEdge = Me.fraGameArea.Left + Me.fraGameArea.Width

  With Me.fraGameArea
    .Caption = vbNullString
    .backColor = RGB(0, 0, 0)
    .SpecialEffect = fmSpecialEffectFlat
    .BorderStyle = fmBorderStyleSingle
  End With

  Me.lblScoreTitle.Left = 12
  Me.lblScoreTitle.Top = hudTop
  Me.lblScoreTitle.Caption = "Score"

  Me.lblScoreValue.Left = Me.lblScoreTitle.Left + Me.lblScoreTitle.Width + 4
  Me.lblScoreValue.Top = hudTop

  Me.lblStatus.Left = 12
  Me.lblStatus.Top = hudTop + 22
  Me.lblStatus.Width = 240

  Me.cmdReset.Width = buttonWidth
  Me.cmdStop.Width = buttonWidth
  Me.cmdStart.Width = buttonWidth
  Me.cmdReset.Top = buttonTop
  Me.cmdStop.Top = buttonTop
  Me.cmdStart.Top = buttonTop
  Me.cmdReset.Left = rightEdge - Me.cmdReset.Width
  Me.cmdStop.Left = Me.cmdReset.Left - buttonGap - Me.cmdStop.Width
  Me.cmdStart.Left = Me.cmdStop.Left - buttonGap - Me.cmdStart.Width

  StyleHudLabel Me.lblScoreTitle, RGB(230, 230, 230)
  StyleHudLabel Me.lblScoreValue, RGB(255, 255, 128)
  StyleHudLabel Me.lblStatus, RGB(160, 220, 255)

  StyleButton Me.cmdStart
  StyleButton Me.cmdStop
  StyleButton Me.cmdReset
End Sub

Private Sub StyleHudLabel(ByVal target As MSForms.Label, ByVal foreColor As Long)
  target.BackStyle = fmBackStyleTransparent
  target.foreColor = foreColor
End Sub

Private Sub StyleButton(ByVal target As MSForms.CommandButton)
  target.backColor = RGB(36, 36, 36)
  target.foreColor = RGB(240, 240, 240)
End Sub

Private Sub EnsureSceneBuilt()
  If mSceneBuilt Then Exit Sub

  Set mPlayer = CreateHitboxLabel("lblPlayerHitbox", PlayerWidth, PlayerHeight)
  BuildPlayerSprite
  BuildPlayerBulletPool

  Set mEnemyBullet = CreateBlockLabel("lblEnemyBullet", RGB(255, 180, 90), EnemyBulletWidth, EnemyBulletHeight)
  mEnemyBullet.Visible = False

  BuildEnemyFleet
  BuildLifeIcons
  BuildUfoSprite
  mSceneBuilt = True
End Sub

' UFOのドット絵Label群を生成
Private Sub BuildUfoSprite()
  Dim pattern As Variant, pixelCount As Long, pixelIndex As Long
  Dim rowIndex As Long, columnIndex As Long, rowPattern As String
  pattern = GetUfoPattern()
  pixelCount = CountSpritePixels(pattern)
  Set mUfoLabel = CreateHitboxLabel("lblUfoHitbox", UfoWidth, UfoHeight)
  mUfoLabel.Visible = False
  ReDim mUfoPixels(1 To pixelCount)
  pixelIndex = 0
  For rowIndex = LBound(pattern) To UBound(pattern)
    rowPattern = CStr(pattern(rowIndex))
    For columnIndex = 1 To Len(rowPattern)
      If Mid$(rowPattern, columnIndex, 1) = "1" Then
        pixelIndex = pixelIndex + 1
        Set mUfoPixels(pixelIndex) = CreateSpritePixel("lblUfoPx" & CStr(pixelIndex), RGB(255, 80, 255), 4)
      End If
    Next columnIndex
  Next rowIndex
End Sub

Private Function CreateHitboxLabel(ByVal controlName As String, ByVal widthValue As Long, ByVal heightValue As Long) As MSForms.Label
  Dim hitbox As MSForms.Label

  Set hitbox = CreateBlockLabel(controlName, RGB(0, 0, 0), widthValue, heightValue)
  hitbox.Visible = False
  Set CreateHitboxLabel = hitbox
End Function

Private Function CreateBlockLabel(ByVal controlName As String, ByVal backColor As Long, ByVal widthValue As Long, ByVal heightValue As Long) As MSForms.Label
  Dim block As MSForms.Label

  Set block = Me.fraGameArea.Controls.Add("Forms.Label.1", controlName, True)
  With block
    .Caption = vbNullString
    .Width = widthValue
    .Height = heightValue
    .BackStyle = fmBackStyleOpaque
    .backColor = backColor
    .BorderStyle = fmBorderStyleSingle
    .SpecialEffect = fmSpecialEffectFlat
    .Visible = True
  End With

  Set CreateBlockLabel = block
End Function

Private Function CreateSpritePixel(ByVal controlName As String, ByVal backColor As Long, ByVal pixelSize As Long) As MSForms.Label
  Dim pixel As MSForms.Label

  Set pixel = Me.fraGameArea.Controls.Add("Forms.Label.1", controlName, True)
  With pixel
    .Caption = vbNullString
    .Width = pixelSize
    .Height = pixelSize
    .BackStyle = fmBackStyleOpaque
    .backColor = backColor
    .BorderStyle = fmBorderStyleNone
    .SpecialEffect = fmSpecialEffectFlat
    .Visible = True
  End With

  Set CreateSpritePixel = pixel
End Function

Private Sub BuildPlayerSprite()
  Dim pattern As Variant
  Dim pixelCount As Long
  Dim pixelIndex As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim rowPattern As String

  pattern = GetPlayerShipPattern()
  pixelCount = CountSpritePixels(pattern)
  ReDim mPlayerPixels(1 To pixelCount)

  pixelIndex = 0
  For rowIndex = LBound(pattern) To UBound(pattern)
    rowPattern = CStr(pattern(rowIndex))
    For columnIndex = 1 To Len(rowPattern)
      If Mid$(rowPattern, columnIndex, 1) = "1" Then
        pixelIndex = pixelIndex + 1
        Set mPlayerPixels(pixelIndex) = CreateSpritePixel("lblPlayerPx" & CStr(pixelIndex), RGB(0, 232, 160), PLAYER_SPRITE_PIXEL_SIZE)
      End If
    Next columnIndex
  Next rowIndex
End Sub

Private Sub BuildPlayerBulletPool()
  Dim bulletIndex As Long

  ReDim mPlayerBullets(1 To PlayerBulletCount)
  For bulletIndex = LBound(mPlayerBullets) To UBound(mPlayerBullets)
    Set mPlayerBullets(bulletIndex) = CreateBlockLabel("lblBullet" & CStr(bulletIndex), RGB(255, 250, 170), BulletWidth, BulletHeight)
    mPlayerBullets(bulletIndex).Visible = False
  Next bulletIndex
End Sub

Private Sub BuildEnemyFleet()
  Dim pattern As Variant
  Dim pixelCount As Long
  Dim enemyIndex As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim pixelIndex As Long
  Dim spriteRow As Long
  Dim spriteColumn As Long
  Dim rowPattern As String
  Dim enemyColor As Long
  Dim enemyType As String

  ReDim mEnemyLabels(1 To EnemyRows * EnemyCols)
  ReDim mEnemyPixels(1 To EnemyRows * EnemyCols, 1 To 60) '最大ピクセル数
  ReDim mEnemyAlive(1 To EnemyRows * EnemyCols)
  ReDim mEnemyType(1 To EnemyRows * EnemyCols)

  enemyIndex = 0
  For rowIndex = 1 To EnemyRows
    ' 行ごとに敵種を割り当て
    Select Case rowIndex
      Case 1: pattern = GetEnemySquidPattern(): enemyColor = RGB(80, 255, 80): enemyType = "SQUID" 'イカ(緑)
      Case 2: pattern = GetEnemyCrabPattern(): enemyColor = RGB(90, 230, 255): enemyType = "CRAB" 'カニ(水色)
      Case Else: pattern = GetEnemyOctopusPattern(): enemyColor = RGB(255, 255, 80): enemyType = "OCTOPUS" 'タコ(黄)
    End Select
    pixelCount = CountSpritePixels(pattern)
    For columnIndex = 1 To EnemyCols
      enemyIndex = enemyIndex + 1
      Set mEnemyLabels(enemyIndex) = CreateHitboxLabel("lblEnemyHitbox" & CStr(enemyIndex), EnemyWidth, EnemyHeight)
      mEnemyType(enemyIndex) = enemyType
      pixelIndex = 0
      For spriteRow = LBound(pattern) To UBound(pattern)
        rowPattern = CStr(pattern(spriteRow))
        For spriteColumn = 1 To Len(rowPattern)
          If Mid$(rowPattern, spriteColumn, 1) = "1" Then
            pixelIndex = pixelIndex + 1
            Set mEnemyPixels(enemyIndex, pixelIndex) = CreateSpritePixel("lblEnemyPx" & CStr(enemyIndex) & "_" & CStr(pixelIndex), enemyColor, ENEMY_SPRITE_PIXEL_SIZE)
          End If
        Next spriteColumn
      Next spriteRow
    Next columnIndex
  Next rowIndex
End Sub

Private Sub BuildLifeIcons()
  Dim iconIndex As Long
  Dim iconLabel As MSForms.Label
  Dim firstIconLeft As Long
  Dim iconsTop As Long

  ReDim mLifeIcons(1 To StartingLives)
  firstIconLeft = Me.fraGameArea.Left + ((Me.fraGameArea.Width - (StartingLives * 14) - ((StartingLives - 1) * LIFE_ICON_SPACING)) \ 2)
  iconsTop = Me.fraGameArea.Top + Me.fraGameArea.Height + 10

  For iconIndex = LBound(mLifeIcons) To UBound(mLifeIcons)
    Set iconLabel = Me.Controls.Add("Forms.Label.1", "lblLifeIcon" & CStr(iconIndex), True)
    With iconLabel
      .Caption = ChrW$(&H2665)
      .Left = firstIconLeft + ((iconIndex - 1) * LIFE_ICON_SPACING)
      .Top = iconsTop
      .Width = 14
      .Height = 14
      .BackStyle = fmBackStyleTransparent
      .foreColor = RGB(255, 92, 92)
      .Font.Size = 12
      .Font.Bold = True
      .Visible = True
    End With
    Set mLifeIcons(iconIndex) = iconLabel
  Next iconIndex
End Sub

Private Sub ResetGame()
  Dim enemyIndex As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim startLeft As Long
  Dim startTop As Long

  EnsureSceneBuilt

  mGameRunning = False
  mGameEnded = False
  mEnemyDirection = 1
  mPrevEscapeDown = False
  mLives = StartingLives
  mPlayerShotClock = Timer
  mEnemyShotClock = Timer
  mPlayerInvulnerable = False
  mScore = 0

  ResetUfo

  startLeft = 18
  startTop = 18
  enemyIndex = 0

  For rowIndex = 1 To EnemyRows
    For columnIndex = 1 To EnemyCols
      enemyIndex = enemyIndex + 1
      With mEnemyLabels(enemyIndex)
        .Left = startLeft + ((columnIndex - 1) * (EnemyWidth + EnemyGapX))
        .Top = startTop + ((rowIndex - 1) * (EnemyHeight + EnemyGapY))
      End With
      mEnemyAlive(enemyIndex) = True
      SyncEnemySprite enemyIndex
      SetEnemySpriteVisible enemyIndex, True
    Next columnIndex
  Next rowIndex

  ResetPlayerPosition
  ResetPlayerBullets

  mEnemyBullet.Visible = False
  mEnemyBullet.Left = 0
  mEnemyBullet.Top = 0

  UpdateScoreDisplay
  UpdateLivesDisplay
  UpdateStatus "Ready - Hold E to fire"
End Sub

' UFO状態初期化
Private Sub ResetUfo()
  mUfoActive = False
  mUfoLabel.Visible = False
  mUfoNextSpawnTick = Timer + UfoSpawnIntervalMin + Rnd() * (UfoSpawnIntervalMax - UfoSpawnIntervalMin)
End Sub

Private Sub StartGameLoop()
  On Error GoTo ErrHandler

  If mLoopActive Then Exit Sub

  mGameRunning = True
  mLoopActive = True
  UpdateStatus "Running"

  Do While mGameRunning And Me.Visible
    ProcessInput
    UpdatePlayerBullets
    UpdateEnemies
    MaybeFireEnemyBullet
    UpdateEnemyBullet
    UpdatePlayerInvulnerability
    UpdateUfo
    EvaluateGameState
    WaitForNextFrame FrameIntervalSeconds
  Loop

  mLoopActive = False
  If Not mGameEnded And Me.Visible Then
    UpdateStatus "Stopped"
  End If
  Exit Sub

ErrHandler:
  mGameRunning = False
  mLoopActive = False
  Debug.Print "frmInvader error " & Err.Number & ": " & Err.Description
  If Me.Visible Then
    UpdateStatus "Error: " & Err.Description
  End If
End Sub

' UFOの出現・移動・当たり判定・得点処理
Private Sub UpdateUfo()
  Dim i As Long
  Dim pattern As Variant
  Dim pixelIndex As Long
  Dim rowPattern As String
  Dim col As Long
  Dim bulletIndex As Long

  If Not mUfoActive Then
    If Timer >= mUfoNextSpawnTick Then
      mUfoActive = True
      mUfoLabel.Visible = True
      mUfoDirection = IIf(Rnd() < 0.5, 1, -1)

      If mUfoDirection = 1 Then
        mUfoLabel.Left = -UfoWidth
      Else
        mUfoLabel.Left = GameAreaWidth
      End If

      mUfoLabel.Top = 8
      mUfoScore = UfoScoreMin + Int(Rnd() * (UfoScoreMax - UfoScoreMin + 1))
      pattern = GetUfoPattern()
      pixelIndex = 0

      For i = LBound(pattern) To UBound(pattern)
        rowPattern = CStr(pattern(i))
        For col = 1 To Len(rowPattern)
          If Mid$(rowPattern, col, 1) = "1" Then
            pixelIndex = pixelIndex + 1
            mUfoPixels(pixelIndex).Left = mUfoLabel.Left + ((col - 1) * 4)
            mUfoPixels(pixelIndex).Top = mUfoLabel.Top + ((i - LBound(pattern)) * 4)
            mUfoPixels(pixelIndex).Visible = True
          End If
        Next col
      Next i
    End If
    Exit Sub
  End If

  mUfoLabel.Left = mUfoLabel.Left + mUfoDirection * UfoSpeed
  pattern = GetUfoPattern()
  pixelIndex = 0

  For i = LBound(pattern) To UBound(pattern)
    rowPattern = CStr(pattern(i))
    For col = 1 To Len(rowPattern)
      If Mid$(rowPattern, col, 1) = "1" Then
        pixelIndex = pixelIndex + 1
        mUfoPixels(pixelIndex).Left = mUfoLabel.Left + ((col - 1) * 4)
        mUfoPixels(pixelIndex).Top = mUfoLabel.Top + ((i - LBound(pattern)) * 4)
        mUfoPixels(pixelIndex).Visible = True
      End If
    Next col
  Next i

  If mUfoDirection = 1 And mUfoLabel.Left > GameAreaWidth Then
    mUfoActive = False
    mUfoLabel.Visible = False
    For i = 1 To UBound(mUfoPixels)
      mUfoPixels(i).Visible = False
    Next i
    mUfoNextSpawnTick = Timer + UfoSpawnIntervalMin + Rnd() * (UfoSpawnIntervalMax - UfoSpawnIntervalMin)
    Exit Sub
  ElseIf mUfoDirection = -1 And mUfoLabel.Left + mUfoLabel.Width < 0 Then
    mUfoActive = False
    mUfoLabel.Visible = False
    For i = 1 To UBound(mUfoPixels)
      mUfoPixels(i).Visible = False
    Next i
    mUfoNextSpawnTick = Timer + UfoSpawnIntervalMin + Rnd() * (UfoSpawnIntervalMax - UfoSpawnIntervalMin)
    Exit Sub
  End If

  For bulletIndex = LBound(mPlayerBullets) To UBound(mPlayerBullets)
    If mPlayerBullets(bulletIndex).Visible Then
      If LabelsOverlap(mPlayerBullets(bulletIndex), mUfoLabel) Then
        mPlayerBullets(bulletIndex).Visible = False
        mScore = mScore + mUfoScore
        UpdateScoreDisplay
        UpdateStatus "UFO!! +" & mUfoScore
        mUfoActive = False
        mUfoLabel.Visible = False
        For i = 1 To UBound(mUfoPixels)
          mUfoPixels(i).Visible = False
        Next i
        mUfoNextSpawnTick = Timer + UfoSpawnIntervalMin + Rnd() * (UfoSpawnIntervalMax - UfoSpawnIntervalMin)
        Exit Sub
      End If
    End If
  Next bulletIndex
End Sub

Private Sub StopGameLoop(ByVal preserveStatus As Boolean)
  mGameRunning = False

  If Not preserveStatus And Not mGameEnded Then
    UpdateStatus "Stopped"
  End If
End Sub

Private Sub ProcessInput()
  Dim leftDown As Boolean
  Dim rightDown As Boolean
  Dim fireDown As Boolean
  Dim escapeDown As Boolean

  leftDown = IsVirtualKeyDown(vbKeyLeft)
  rightDown = IsVirtualKeyDown(vbKeyRight)

  If leftDown Xor rightDown Then
    If leftDown Then
      MovePlayer -PlayerSpeed
    Else
      MovePlayer PlayerSpeed
    End If
  End If

  fireDown = IsVirtualKeyDown(Asc("E"))
  If fireDown Then
    FireBullet
  End If

  escapeDown = IsVirtualKeyDown(vbKeyEscape)
  If escapeDown And Not mPrevEscapeDown Then
    StopGameLoop False
  End If
  mPrevEscapeDown = escapeDown
End Sub

Private Function IsVirtualKeyDown(ByVal keyCode As Long) As Boolean
  IsVirtualKeyDown = (GetAsyncKeyState(keyCode) And &H8000) <> 0
End Function

Private Sub MovePlayer(ByVal deltaX As Long)
  Dim nextLeft As Long

  nextLeft = mPlayer.Left + deltaX
  If nextLeft < 0 Then nextLeft = 0
  If nextLeft > GameAreaWidth - mPlayer.Width Then
    nextLeft = GameAreaWidth - mPlayer.Width
  End If

  mPlayer.Left = nextLeft
  SyncPlayerSprite
End Sub

Private Sub FireBullet()
  Dim bulletIndex As Long

  If SecondsElapsed(mPlayerShotClock, Timer) < PlayerFireIntervalSeconds Then Exit Sub

  bulletIndex = FindAvailablePlayerBulletIndex()
  If bulletIndex = 0 Then Exit Sub

  With mPlayerBullets(bulletIndex)
    .Left = mPlayer.Left + ((mPlayer.Width - .Width) \ 2)
    .Top = mPlayer.Top - .Height - 2
    .Visible = True
  End With

  mPlayerShotClock = Timer
End Sub

Private Function FindAvailablePlayerBulletIndex() As Long
  Dim bulletIndex As Long

  For bulletIndex = LBound(mPlayerBullets) To UBound(mPlayerBullets)
    If Not mPlayerBullets(bulletIndex).Visible Then
      FindAvailablePlayerBulletIndex = bulletIndex
      Exit Function
    End If
  Next bulletIndex
End Function

Private Sub ResetPlayerPosition()
  mPlayer.Left = (GameAreaWidth - PlayerWidth) \ 2
  mPlayer.Top = GameAreaHeight - PlayerHeight - 14
  SyncPlayerSprite
End Sub

Private Sub ResetPlayerBullets()
  Dim bulletIndex As Long

  For bulletIndex = LBound(mPlayerBullets) To UBound(mPlayerBullets)
    mPlayerBullets(bulletIndex).Visible = False
    mPlayerBullets(bulletIndex).Left = 0
    mPlayerBullets(bulletIndex).Top = 0
  Next bulletIndex
End Sub

Private Sub MaybeFireEnemyBullet()
  Dim shooterIndex As Long

  If mEnemyBullet.Visible Then Exit Sub
  If SecondsElapsed(mEnemyShotClock, Timer) < EnemyFireIntervalSeconds Then Exit Sub

  shooterIndex = ChooseEnemyShooterIndex()
  If shooterIndex = 0 Then Exit Sub

  mEnemyBullet.Left = mEnemyLabels(shooterIndex).Left + ((mEnemyLabels(shooterIndex).Width - mEnemyBullet.Width) \ 2)
  mEnemyBullet.Top = mEnemyLabels(shooterIndex).Top + mEnemyLabels(shooterIndex).Height + 2
  mEnemyBullet.Visible = True
  mEnemyShotClock = Timer
End Sub

Private Function ChooseEnemyShooterIndex() As Long
  Dim columnChoices() As Long
  Dim columnCount As Long
  Dim columnIndex As Long
  Dim rowIndex As Long
  Dim enemyIndex As Long

  ReDim columnChoices(1 To EnemyCols)
  columnCount = 0

  For columnIndex = 1 To EnemyCols
    For rowIndex = EnemyRows To 1 Step -1
      enemyIndex = ((rowIndex - 1) * EnemyCols) + columnIndex
      If mEnemyAlive(enemyIndex) Then
        columnCount = columnCount + 1
        columnChoices(columnCount) = enemyIndex
        Exit For
      End If
    Next rowIndex
  Next columnIndex

  If columnCount = 0 Then Exit Function

  ChooseEnemyShooterIndex = columnChoices(Int(Rnd() * columnCount) + 1)
End Function

Private Sub UpdatePlayerBullets()
  Dim bulletIndex As Long
  Dim enemyIndex As Long

  For bulletIndex = LBound(mPlayerBullets) To UBound(mPlayerBullets)
    If mPlayerBullets(bulletIndex).Visible Then
      mPlayerBullets(bulletIndex).Top = mPlayerBullets(bulletIndex).Top - BulletSpeed
      If mPlayerBullets(bulletIndex).Top + mPlayerBullets(bulletIndex).Height < 0 Then
        mPlayerBullets(bulletIndex).Visible = False
      Else
        For enemyIndex = LBound(mEnemyLabels) To UBound(mEnemyLabels)
          If mEnemyAlive(enemyIndex) Then
            If LabelsOverlap(mPlayerBullets(bulletIndex), mEnemyLabels(enemyIndex)) Then
              mEnemyAlive(enemyIndex) = False
              mPlayerBullets(bulletIndex).Visible = False
              SetEnemySpriteVisible enemyIndex, False
              Select Case mEnemyType(enemyIndex)
                Case "SQUID": mScore = mScore + 30
                Case "CRAB": mScore = mScore + 20
                Case Else: mScore = mScore + 10
              End Select
              UpdateScoreDisplay
              UpdateStatus "Enemy hit"
              Exit For
            End If
          End If
        Next enemyIndex
      End If
    End If
  Next bulletIndex
End Sub

Private Sub UpdateEnemyBullet()
  If Not mEnemyBullet.Visible Then Exit Sub

  mEnemyBullet.Top = mEnemyBullet.Top + EnemyBulletSpeed
  If mEnemyBullet.Top > GameAreaHeight Then
    mEnemyBullet.Visible = False
    Exit Sub
  End If

  If Not mPlayerInvulnerable Then
    If LabelsOverlap(mEnemyBullet, mPlayer) Then
      HandlePlayerHit
    End If
  End If
End Sub

Private Sub HandlePlayerHit()
  mEnemyBullet.Visible = False
  mLives = mLives - 1
  UpdateLivesDisplay

  If mLives <= 0 Then
    mGameEnded = True
    mGameRunning = False
    SetPlayerSpriteVisible True
    UpdateStatus "Game over"
    Exit Sub
  End If

  ResetPlayerPosition
  ResetPlayerBullets
  ActivatePlayerInvulnerability
  mEnemyShotClock = Timer
  UpdateStatus "Hit! Lives left: " & CStr(mLives)
End Sub

Private Sub ActivatePlayerInvulnerability()
  mPlayerInvulnerable = True
  mPlayerInvulnerableClock = Timer
End Sub

Private Sub UpdatePlayerInvulnerability()
  Dim elapsedSeconds As Double
  Dim shouldShow As Boolean

  If Not mPlayerInvulnerable Then
    SetPlayerSpriteVisible True
    Exit Sub
  End If

  elapsedSeconds = SecondsElapsed(mPlayerInvulnerableClock, Timer)
  If elapsedSeconds >= PlayerInvulnerableSeconds Then
    mPlayerInvulnerable = False
    SetPlayerSpriteVisible True
    Exit Sub
  End If

  shouldShow = (Int(elapsedSeconds / PLAYER_FLASH_INTERVAL_SECONDS) Mod 2) = 0
  SetPlayerSpriteVisible shouldShow
End Sub

Private Sub UpdateEnemies()
  Dim enemyIndex As Long
  Dim nextLeft As Long
  Dim hitWall As Boolean

  hitWall = False
  For enemyIndex = LBound(mEnemyLabels) To UBound(mEnemyLabels)
    If mEnemyAlive(enemyIndex) Then
      nextLeft = mEnemyLabels(enemyIndex).Left + (mEnemyDirection * EnemySpeed)
      If nextLeft <= 0 Or nextLeft + mEnemyLabels(enemyIndex).Width >= GameAreaWidth Then
        hitWall = True
        Exit For
      End If
    End If
  Next enemyIndex

  If hitWall Then
    mEnemyDirection = -mEnemyDirection
  End If

  For enemyIndex = LBound(mEnemyLabels) To UBound(mEnemyLabels)
    If mEnemyAlive(enemyIndex) Then
      If hitWall Then
        mEnemyLabels(enemyIndex).Top = mEnemyLabels(enemyIndex).Top + EnemyDropDistance
      End If
      mEnemyLabels(enemyIndex).Left = mEnemyLabels(enemyIndex).Left + (mEnemyDirection * EnemySpeed)
      SyncEnemySprite enemyIndex
    End If
  Next enemyIndex
End Sub

Private Sub EvaluateGameState()
  Dim enemyIndex As Long

  If AllEnemiesDefeated Then
    mGameEnded = True
    mGameRunning = False
    UpdateStatus "Clear"
    Exit Sub
  End If

  For enemyIndex = LBound(mEnemyLabels) To UBound(mEnemyLabels)
    If mEnemyAlive(enemyIndex) Then
      If mEnemyLabels(enemyIndex).Top + mEnemyLabels(enemyIndex).Height >= mPlayer.Top - 4 Then
        mGameEnded = True
        mGameRunning = False
        UpdateStatus "Game over"
        Exit Sub
      End If
    End If
  Next enemyIndex
End Sub

Private Function AllEnemiesDefeated() As Boolean
  Dim enemyIndex As Long

  For enemyIndex = LBound(mEnemyAlive) To UBound(mEnemyAlive)
    If mEnemyAlive(enemyIndex) Then Exit Function
  Next enemyIndex

  AllEnemiesDefeated = True
End Function

Private Function LabelsOverlap(ByVal firstLabel As MSForms.Label, ByVal secondLabel As MSForms.Label) As Boolean
  LabelsOverlap = firstLabel.Left < secondLabel.Left + secondLabel.Width _
    And firstLabel.Left + firstLabel.Width > secondLabel.Left _
    And firstLabel.Top < secondLabel.Top + secondLabel.Height _
    And firstLabel.Top + firstLabel.Height > secondLabel.Top
End Function

Private Function PlayerPattern() As Variant
  PlayerPattern = GetPlayerShipPattern()
End Function

Private Function EnemyPattern(ByVal enemyIndex As Long) As Variant
  Select Case mEnemyType(enemyIndex)
    Case "SQUID"
      EnemyPattern = GetEnemySquidPattern()
    Case "CRAB"
      EnemyPattern = GetEnemyCrabPattern()
    Case Else
      EnemyPattern = GetEnemyOctopusPattern()
  End Select
End Function

Private Sub SyncPlayerSprite()
  PositionSpritePixels mPlayerPixels, PlayerPattern(), PLAYER_SPRITE_PIXEL_SIZE, mPlayer.Left, mPlayer.Top, True
End Sub

Private Sub SyncEnemySprite(ByVal enemyIndex As Long)
  PositionSpritePixelsForEnemy enemyIndex, EnemyPattern(enemyIndex), ENEMY_SPRITE_PIXEL_SIZE, mEnemyLabels(enemyIndex).Left, mEnemyLabels(enemyIndex).Top, mEnemyAlive(enemyIndex)
End Sub

Private Sub PositionSpritePixels(ByRef spritePixels() As MSForms.Label, ByVal pattern As Variant, ByVal pixelSize As Long, ByVal leftValue As Long, ByVal topValue As Long, ByVal isVisible As Boolean)
  Dim pixelIndex As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim rowPattern As String

  pixelIndex = 0
  For rowIndex = LBound(pattern) To UBound(pattern)
    rowPattern = CStr(pattern(rowIndex))
    For columnIndex = 1 To Len(rowPattern)
      If Mid$(rowPattern, columnIndex, 1) = "1" Then
        pixelIndex = pixelIndex + 1
        spritePixels(pixelIndex).Left = leftValue + ((columnIndex - 1) * pixelSize)
        spritePixels(pixelIndex).Top = topValue + ((rowIndex - LBound(pattern)) * pixelSize)
        spritePixels(pixelIndex).Visible = isVisible
      End If
    Next columnIndex
  Next rowIndex
End Sub

Private Sub PositionSpritePixelsForEnemy(ByVal enemyIndex As Long, ByVal pattern As Variant, ByVal pixelSize As Long, ByVal leftValue As Long, ByVal topValue As Long, ByVal isVisible As Boolean)
  Dim pixelIndex As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim rowPattern As String

  pixelIndex = 0
  For rowIndex = LBound(pattern) To UBound(pattern)
    rowPattern = CStr(pattern(rowIndex))
    For columnIndex = 1 To Len(rowPattern)
      If Mid$(rowPattern, columnIndex, 1) = "1" Then
        pixelIndex = pixelIndex + 1
        mEnemyPixels(enemyIndex, pixelIndex).Left = leftValue + ((columnIndex - 1) * pixelSize)
        mEnemyPixels(enemyIndex, pixelIndex).Top = topValue + ((rowIndex - LBound(pattern)) * pixelSize)
        mEnemyPixels(enemyIndex, pixelIndex).Visible = isVisible
      End If
    Next columnIndex
  Next rowIndex
End Sub

Private Sub SetEnemySpriteVisible(ByVal enemyIndex As Long, ByVal isVisible As Boolean)
  Dim pixelIndex As Long

  For pixelIndex = LBound(mEnemyPixels, 2) To UBound(mEnemyPixels, 2)
    If Not mEnemyPixels(enemyIndex, pixelIndex) Is Nothing Then
      mEnemyPixels(enemyIndex, pixelIndex).Visible = isVisible
    End If
  Next pixelIndex
End Sub

Private Sub SetPlayerSpriteVisible(ByVal isVisible As Boolean)
  Dim pixelIndex As Long

  For pixelIndex = LBound(mPlayerPixels) To UBound(mPlayerPixels)
    mPlayerPixels(pixelIndex).Visible = isVisible
  Next pixelIndex
End Sub

Private Function CountSpritePixels(ByVal pattern As Variant) As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim rowPattern As String

  For rowIndex = LBound(pattern) To UBound(pattern)
    rowPattern = CStr(pattern(rowIndex))
    For columnIndex = 1 To Len(rowPattern)
      If Mid$(rowPattern, columnIndex, 1) = "1" Then
        CountSpritePixels = CountSpritePixels + 1
      End If
    Next columnIndex
  Next rowIndex
End Function

'' 旧パターンは廃止（GameConstantsのGetterを使う）

Private Function EnemyColorForRow(ByVal rowIndex As Long) As Long
  Select Case rowIndex Mod 3
    Case 1
      EnemyColorForRow = RGB(255, 92, 92)
    Case 2
      EnemyColorForRow = RGB(90, 230, 255)
    Case Else
      EnemyColorForRow = RGB(215, 90, 255)
  End Select
End Function

Private Sub UpdateScoreDisplay()
  Me.lblScoreValue.Caption = Format$(mScore, "00000")
End Sub

Private Sub UpdateLivesDisplay()
  Dim iconIndex As Long

  For iconIndex = LBound(mLifeIcons) To UBound(mLifeIcons)
    mLifeIcons(iconIndex).Visible = iconIndex <= mLives
  Next iconIndex
End Sub

Private Sub UpdateStatus(ByVal message As String)
  Me.lblStatus.Caption = message
End Sub

Private Sub WaitForNextFrame(ByVal frameSeconds As Double)
  Dim startTick As Double

  startTick = Timer
  Do While mGameRunning And Me.Visible
    DoEvents
    If SecondsElapsed(startTick, Timer) >= frameSeconds Then Exit Do
  Loop
End Sub

Private Function SecondsElapsed(ByVal startTick As Double, ByVal endTick As Double) As Double
  If endTick >= startTick Then
    SecondsElapsed = endTick - startTick
  Else
    SecondsElapsed = (86400# - startTick) + endTick
  End If
End Function


