Attribute VB_Name = "TestMazeChase"
'@Folder("Tests")
Option Explicit

Public Sub BeforeEach()
    MazeChaseGame.InitializeGame
End Sub

'@Tag("smoke")
Public Sub Test_MazeChase_InitialState()
    XlflowAssert.AssertEquals 0, MazeChaseGame.scoreValue(), "score should start at zero"
    XlflowAssert.AssertEquals 3, MazeChaseGame.LivesValue(), "lives should start at three"
    XlflowAssert.AssertFalse MazeChaseGame.IsFinished(), "game should start active"
End Sub

Public Sub Test_MazeChase_PlayerCannotMoveIntoWall()
    MazeChaseGame.DebugSetPlayerPosition 4, 2
    MazeChaseGame.DebugSetRequestedDirection 1
    MazeChaseGame.TickGame

    XlflowAssert.AssertEquals 4, MazeChaseGame.PlayerRow(), "player row should stay fixed"
    XlflowAssert.AssertEquals 2, MazeChaseGame.PlayerCol(), "player should not move into wall"
End Sub

Public Sub Test_MazeChase_PlayerMovementUpdatesDirectionAndAnimation()
    MazeChaseGame.DebugSetPlayerPosition 16, 8
    MazeChaseGame.DebugSetRequestedDirection 3
    MazeChaseGame.DebugSetPlayerMouthOpen True
    MazeChaseGame.TickGame

    XlflowAssert.AssertEquals 16, MazeChaseGame.PlayerRow(), "player should stay on same row when moving right"
    XlflowAssert.AssertEquals 9, MazeChaseGame.PlayerCol(), "player should advance one cell to the right"
    XlflowAssert.AssertEquals 3, MazeChaseGame.PlayerDirection(), "player direction should face right"
    XlflowAssert.AssertFalse MazeChaseGame.PlayerMouthOpen(), "mouth should toggle after movement"
End Sub

Public Sub Test_MazeChase_ConsumeDotAddsScore()
    MazeChaseGame.DebugSetPlayerPosition 4, 3
    MazeChaseGame.DebugConsumeCurrentTile

    XlflowAssert.AssertEquals 10, MazeChaseGame.scoreValue(), "dot should add ten points"
    XlflowAssert.AssertFalse MazeChaseGame.HasDotAt(4, 3), "dot should be removed"
End Sub

Public Sub Test_MazeChase_ConsumePowerPelletEnablesFrightened()
    MazeChaseGame.DebugSetPlayerPosition 2, 2
    MazeChaseGame.DebugConsumeCurrentTile

    XlflowAssert.AssertEquals 50, MazeChaseGame.scoreValue(), "power pellet should add fifty points"
    XlflowAssert.AssertTrue MazeChaseGame.IsGhostFrightened(), "ghosts should become frightened"
    XlflowAssert.AssertTrue MazeChaseGame.IsPowerBoostActive(), "power pellet should enable speed boost"
    XlflowAssert.AssertFalse MazeChaseGame.HasPowerAt(2, 2), "power pellet should be removed"
End Sub

Public Sub Test_MazeChase_PowerBoostMovesPlayerFarther()
    MazeChaseGame.DebugSetPlayerPosition 2, 2
    MazeChaseGame.DebugConsumeCurrentTile
    MazeChaseGame.DebugSetPlayerPosition 16, 8
    MazeChaseGame.DebugSetRequestedDirection 3
    MazeChaseGame.TickGame

    XlflowAssert.AssertEquals 16, MazeChaseGame.PlayerRow(), "boost should keep player on same row when corridor is clear"
    XlflowAssert.AssertEquals 10, MazeChaseGame.PlayerCol(), "boosted movement should advance two cells in one tick"
End Sub

Public Sub Test_MazeChase_GhostCollisionCostsLife()
    MazeChaseGame.DebugSetPlayerPosition 4, 3
    MazeChaseGame.DebugSetGhostPosition 1, 4, 3
    MazeChaseGame.DebugResolveCollisions

    XlflowAssert.AssertEquals 2, MazeChaseGame.LivesValue(), "collision should remove one life"
    XlflowAssert.AssertFalse MazeChaseGame.IsFinished(), "game should continue while lives remain"
End Sub

Public Sub Test_MazeChase_FrightenedCollisionResetsGhost()
    MazeChaseGame.DebugSetPlayerPosition 2, 2
    MazeChaseGame.DebugConsumeCurrentTile
    MazeChaseGame.DebugSetPlayerPosition 4, 3
    MazeChaseGame.DebugSetGhostPosition 1, 4, 3
    MazeChaseGame.DebugResolveCollisions

    XlflowAssert.AssertEquals 250, MazeChaseGame.scoreValue(), "power pellet plus ghost should add score"
    XlflowAssert.AssertEquals MazeChaseGame.GhostHomeRow(1), MazeChaseGame.GhostRow(1), "ghost should reset to home row"
    XlflowAssert.AssertEquals MazeChaseGame.GhostHomeCol(1), MazeChaseGame.GhostCol(1), "ghost should reset to home col"
End Sub

Public Sub Test_MazeChase_ClearingPelletsWinsGame()
    MazeChaseGame.DebugClearPellets
    MazeChaseGame.TickGame

    XlflowAssert.AssertTrue MazeChaseGame.IsFinished(), "game should end when pellets are gone"
    XlflowAssert.AssertTrue MazeChaseGame.HasPlayerWon(), "clearing pellets should be a win"
End Sub
