Attribute VB_Name = "GameConstants"
Option Explicit

Private Const P_GAME_AREA_WIDTH As Long = 420
Private Const P_GAME_AREA_HEIGHT As Long = 300
Private Const P_STARTING_LIVES As Long = 3

Private Const P_PLAYER_WIDTH As Long = 24
Private Const P_PLAYER_HEIGHT As Long = 12
Private Const P_PLAYER_SPEED As Long = 6
Private Const P_PLAYER_BULLET_COUNT As Long = 3
Private Const P_PLAYER_FIRE_INTERVAL_SECONDS As Double = 0.12
Private Const P_PLAYER_INVULNERABLE_SECONDS As Double = 1.2

Private Const P_BULLET_WIDTH As Long = 4
Private Const P_BULLET_HEIGHT As Long = 12
Private Const P_BULLET_SPEED As Long = 10
Private Const P_ENEMY_BULLET_WIDTH As Long = 4
Private Const P_ENEMY_BULLET_HEIGHT As Long = 12
Private Const P_ENEMY_BULLET_SPEED As Long = 7

Private Const P_ENEMY_WIDTH As Long = 20
Private Const P_ENEMY_HEIGHT As Long = 12
Private Const P_ENEMY_SPEED As Long = 4
Private Const P_ENEMY_ROWS As Long = 3
Private Const P_ENEMY_COLS As Long = 7
Private Const P_ENEMY_GAP_X As Long = 20
Private Const P_ENEMY_GAP_Y As Long = 16
Private Const P_ENEMY_DROP_DISTANCE As Long = 12

Private Const P_FRAME_INTERVAL_SECONDS As Double = 0.05
Private Const P_SCORE_PER_ENEMY As Long = 10
Private Const P_ENEMY_FIRE_INTERVAL_SECONDS As Double = 0.75
Private Const P_UFO_WIDTH As Long = 32
Private Const P_UFO_HEIGHT As Long = 14
Private Const P_UFO_SPEED As Long = 6
Private Const P_UFO_SPAWN_INTERVAL_MIN As Double = 8#
Private Const P_UFO_SPAWN_INTERVAL_MAX As Double = 18#
Private Const P_UFO_SCORE_MIN As Long = 50
Private Const P_UFO_SCORE_MAX As Long = 150

Public Function GameAreaWidth() As Long
        GameAreaWidth = P_GAME_AREA_WIDTH
End Function

Public Function GameAreaHeight() As Long
        GameAreaHeight = P_GAME_AREA_HEIGHT
End Function

Public Function StartingLives() As Long
        StartingLives = P_STARTING_LIVES
End Function

Public Function PlayerWidth() As Long
        PlayerWidth = P_PLAYER_WIDTH
End Function

Public Function PlayerHeight() As Long
        PlayerHeight = P_PLAYER_HEIGHT
End Function

Public Function PlayerSpeed() As Long
        PlayerSpeed = P_PLAYER_SPEED
End Function

Public Function PlayerBulletCount() As Long
        PlayerBulletCount = P_PLAYER_BULLET_COUNT
End Function

Public Function PlayerFireIntervalSeconds() As Double
        PlayerFireIntervalSeconds = P_PLAYER_FIRE_INTERVAL_SECONDS
End Function

Public Function PlayerInvulnerableSeconds() As Double
        PlayerInvulnerableSeconds = P_PLAYER_INVULNERABLE_SECONDS
End Function

Public Function BulletWidth() As Long
        BulletWidth = P_BULLET_WIDTH
End Function

Public Function BulletHeight() As Long
        BulletHeight = P_BULLET_HEIGHT
End Function

Public Function BulletSpeed() As Long
        BulletSpeed = P_BULLET_SPEED
End Function

Public Function EnemyBulletWidth() As Long
        EnemyBulletWidth = P_ENEMY_BULLET_WIDTH
End Function

Public Function EnemyBulletHeight() As Long
        EnemyBulletHeight = P_ENEMY_BULLET_HEIGHT
End Function

Public Function EnemyBulletSpeed() As Long
        EnemyBulletSpeed = P_ENEMY_BULLET_SPEED
End Function

Public Function EnemyWidth() As Long
        EnemyWidth = P_ENEMY_WIDTH
End Function

Public Function EnemyHeight() As Long
        EnemyHeight = P_ENEMY_HEIGHT
End Function

Public Function EnemySpeed() As Long
        EnemySpeed = P_ENEMY_SPEED
End Function

Public Function EnemyRows() As Long
        EnemyRows = P_ENEMY_ROWS
End Function

Public Function EnemyCols() As Long
        EnemyCols = P_ENEMY_COLS
End Function

Public Function EnemyGapX() As Long
        EnemyGapX = P_ENEMY_GAP_X
End Function

Public Function EnemyGapY() As Long
        EnemyGapY = P_ENEMY_GAP_Y
End Function

Public Function EnemyDropDistance() As Long
        EnemyDropDistance = P_ENEMY_DROP_DISTANCE
End Function

Public Function FrameIntervalSeconds() As Double
        FrameIntervalSeconds = P_FRAME_INTERVAL_SECONDS
End Function

Public Function ScorePerEnemy() As Long
        ScorePerEnemy = P_SCORE_PER_ENEMY
End Function

Public Function EnemyFireIntervalSeconds() As Double
        EnemyFireIntervalSeconds = P_ENEMY_FIRE_INTERVAL_SECONDS
End Function

' UFOドット絵（16x7）
Private Function UfoPattern() As Variant
        UfoPattern = Array( _
                "0000011111100000", _
                "0001111111110000", _
                "0011111111111000", _
                "0110110110111100", _
                "1111111111111110", _
                "0011000000001100", _
                "0110000000000110")
End Function

' イカ（30点, 緑, 8x8）
Private Function EnemySquidPattern() As Variant
        EnemySquidPattern = Array( _
                "00011000", _
                "00111100", _
                "01111110", _
                "11011011", _
                "11111111", _
                "00100100", _
                "01000010", _
                "10100101")
End Function

' カニ（20点, 水色, 11x8）
Private Function EnemyCrabPattern() As Variant
        EnemyCrabPattern = Array( _
                "00110001100", _
                "01111111110", _
                "11111111111", _
                "11011011011", _
                "11111111111", _
                "00100100100", _
                "01000000010", _
                "10000000001")
End Function

' タコ（10点, 黄, 8x8）
Private Function EnemyOctopusPattern() As Variant
        EnemyOctopusPattern = Array( _
                "00111100", _
                "01111110", _
                "11111111", _
                "11011011", _
                "11111111", _
                "00100100", _
                "01011010", _
                "10000001")
End Function

' プレイヤー自機（11x8, 緑）
Private Function PlayerShipPattern() As Variant
        PlayerShipPattern = Array( _
                "00000100000", _
                "00001110000", _
                "00011111000", _
                "00111111100", _
                "01111111110", _
                "11111111111", _
                "00110001100", _
                "01000000010")
End Function

' UFO定数Getter
Public Function UfoWidth() As Long: UfoWidth = P_UFO_WIDTH: End Function
Public Function UfoHeight() As Long: UfoHeight = P_UFO_HEIGHT: End Function
Public Function UfoSpeed() As Long: UfoSpeed = P_UFO_SPEED: End Function
Public Function UfoSpawnIntervalMin() As Double: UfoSpawnIntervalMin = P_UFO_SPAWN_INTERVAL_MIN: End Function
Public Function UfoSpawnIntervalMax() As Double: UfoSpawnIntervalMax = P_UFO_SPAWN_INTERVAL_MAX: End Function
Public Function UfoScoreMin() As Long: UfoScoreMin = P_UFO_SCORE_MIN: End Function
Public Function UfoScoreMax() As Long: UfoScoreMax = P_UFO_SCORE_MAX: End Function

' ドット絵パターンGetter
Public Function GetUfoPattern() As Variant: GetUfoPattern = UfoPattern(): End Function
Public Function GetEnemySquidPattern() As Variant: GetEnemySquidPattern = EnemySquidPattern(): End Function
Public Function GetEnemyCrabPattern() As Variant: GetEnemyCrabPattern = EnemyCrabPattern(): End Function
Public Function GetEnemyOctopusPattern() As Variant: GetEnemyOctopusPattern = EnemyOctopusPattern(): End Function
Public Function GetPlayerShipPattern() As Variant: GetPlayerShipPattern = PlayerShipPattern(): End Function
