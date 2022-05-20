Attribute VB_Name = "modGlobals"
Public Declare Function TickCount Lib "kernel32" Alias "GetTickCount" () As Long

' Frame rate
Public Const FRAME_RATE As Integer = 1000 / 60

' Field height and width
Public Const FIELD_WIDTH As Integer = 6
Public Const FIELD_HEIGHT As Integer = 12

' Height and width of the block sprites
Public Const BLOCK_WIDTH As Integer = 32
Public Const BLOCK_HEIGHT As Integer = 32

' Number of blocks
Public Const NUM_BLOCKS As Integer = 5

' Screen height and width
Public Const SCREEN_WIDTH As Integer = 640
Public Const SCREEN_HEIGHT As Integer = 480
Public b_Fullscreen As Boolean

' The current stack height (as it's advancing)
Public i_StackHeight As Integer, i_StackSpeed As Integer, i_Speed As Integer, i_SpeedChange As Integer

' Block state lengths
Public Const PREFALL_LENGTH As Integer = 10
Public Const FLASH_LENGTH As Integer = 45
Public Const POP_DELAY_LENGTH = 15

Public i_PrefallLength As Integer
Public i_FlashLength As Integer
Public i_PopDelay As Integer

'Public Const

' ENUMS
' ------------------------------------

' Game state
Public Enum GAME_STATE
    GS_GAMEOVER = 0
    GS_COUNTDOWN
    GS_NORMAL
    GS_PAUSED
    GS_MENU
    GS_TIMEUP
End Enum

Public Enum CharacterState
    CS_NORMAL = 2
    CS_COMBO
    CS_CHAIN
    CS_BIGCHAIN
    CS_DIE
End Enum

' Block state
Public Enum BLOCK_STATE
    BS_EMPTY = 0
    BS_NORMAL
    BS_SWAPPING
    BS_PREFALL
    BS_FALLING
    BS_FLASHING
    BS_CLEARING
    BS_RESERVED
    BS_UPDATEREQUIRED
End Enum

' Game modes
Public Enum GAME_MODE
    GM_ENDLESS = 1
    GM_TIMETRIAL = 2
End Enum
    

' TYPES
' ------------------------------------

' Animation state
Public Type ANIM_STATE

    ' Current animation
    i_CurrentAnimation As Integer
    
    ' Current frame
    i_CurrentFrame As Integer
    
    ' Max frame
    i_MaxFrame As Integer
    
    ' Loop?
    b_Looping As Boolean

End Type

' Key states
Public b_KeyLeft As Boolean, b_KeyRight As Boolean, b_KeyUp As Boolean, b_KeyDown As Boolean
Public b_KeySwap As Boolean, b_KeyL As Boolean, b_KeyAdvance As Boolean, b_KeyEnter As Boolean

' Chain counter and new line flag
Public i_Chain As Long, b_ChainFrame As Boolean, b_NewLine As Boolean, i_ChainID As Long

' Frame rate stuff
Public i_LastTick As Long, i_StackCounter As Integer

' Key press delay
Public i_KeyCounter As Integer
Public b_KeyPressed As Boolean

' Pause
Public b_Pause As Boolean, b_Swap As Boolean

' The card and score handlers
Public g_Cards As New clsCard, g_Score As New clsScore

' The sound and text handler
Public g_Text As New clsText, g_Sound As New clsSound, g_Music As New clsMusic

' Sound constants
Public Const MAX_SOUNDS As Integer = 32
Public Const S_POP1_CHAIN0 As Integer = 1
Public Const S_POP2_CHAIN0 As Integer = 2
Public Const S_POP3_CHAIN0 As Integer = 3
Public Const S_POP4_CHAIN0 As Integer = 4
Public Const S_CURSOR_MOVE As Integer = 17
Public Const S_SWAP_BLOCK As Integer = 18
Public Const S_BLOCK_FALL As Integer = 19
Public Const S_MUSIC_NORMAL_INTRO As Integer = 20
Public Const S_MUSIC_NORMAL_LOOP As Integer = 21
Public Const S_MUSIC_MENU_INTRO As Integer = 22
Public Const S_MUSIC_MENU_LOOP As Integer = 23
Public Const S_MUSIC_PANIC_INTRO As Integer = 24
Public Const S_MUSIC_PANIC_LOOP As Integer = 25
Public Const S_CHARACTER_CHEER = 26
Public Const S_COUNTDOWN = 27

Public b_Advance As Boolean, i_DisplayFormat As Long

' The stop timer
Public i_StopTime As Long, i_StopStart As Long

' The vertex format
Public Const VERTEX_FVF = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

' The vertex type
Public Type Vertex
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    u As Single
    v As Single
End Type

' The texture with the main graphics
Public g_MainImg As New clsSurface, g_Particles As New clsParticles
Public g_Character As New clsCharacter, i_NumCleared As Integer, g_Menu As New clsMenu
