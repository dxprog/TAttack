VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The DirectX and DirectDraw objects
Private g_Dx As New DirectX8
Private g_Graphics As New clsDirect3D

' The blocks object
Private g_Blocks As New clsBlocks

' The handle to the window
Private g_hWnd As Long

' The game state
Public i_GameState As GAME_STATE

' The cursor position
Private i_CursorX As Integer, i_CursorY As Integer

' Local textures
Private g_Bg As Direct3DSurface8
Private g_Lose As New clsSurface

' When the game started
Private i_GameStart As Long, i_Length As Long

' The countdown counter
Private i_CountDown As Long

' The game mode
Private i_Mode As GAME_MODE

' Fade stuff
Public b_Fade As Boolean, i_FadeLevel As Integer, i_FadeInc As Integer, i_FadeTo As GAME_STATE

Public Function InitGame(ByVal h_Wnd As Long) As Boolean
    
    Dim b_RetVal As Boolean

    ' Save the window handle for later use
    g_hWnd = h_Wnd

    ' Load the saved settings
    g_Menu.LoadDefaults

    ' Initialize DirectDraw
    b_RetVal = g_Graphics.Initialize(g_hWnd, g_Dx, SCREEN_WIDTH, SCREEN_HEIGHT, 16, b_Fullscreen)
    If b_RetVal = False Then
        MsgBox "Unable to initialize Direct3D!", vbCritical Or vbOKOnly, "ERROR"
        Exit Function
    End If

    ' Load the textures
    Set g_Bg = g_Graphics.g_Device.CreateImageSurface(640, 480, i_DisplayFormat)
    g_Graphics.g_D3dx.LoadSurfaceFromFile g_Bg, ByVal 0, ByVal 0, App.Path & "\images\field.png", ByVal 0, D3DX_DEFAULT, 0, ByVal 0
    g_MainImg.CreateSurface g_Graphics.g_Device, App.Path & "\images\blocks.png", 256, 256
    g_Lose.CreateSurface g_Graphics.g_Device, App.Path & "\images\end_game.png", 256, 256
    
    g_Character.Initialize g_Graphics
    g_Text.Initialize g_Graphics
    g_Menu.Init g_Graphics

    ' Initialize sound
    g_Sound.Initialize g_Dx, g_hWnd

    ' Load the sounds
    g_Sound.LoadSound App.Path & "\sound\cursor_move.wav", S_CURSOR_MOVE
    g_Sound.LoadSound App.Path & "\sound\swap_block.wav", S_SWAP_BLOCK
    g_Sound.LoadSound App.Path & "\sound\block_fall.wav", S_BLOCK_FALL
    g_Sound.LoadSound App.Path & "\sound\char_special.wav", S_CHARACTER_CHEER
    g_Sound.LoadSound App.Path & "\sound\countdown.wav", S_COUNTDOWN
    
    ' Load the pop sounds
    For i = 1 To 4
        g_Sound.LoadSound App.Path & "\sound\pop" & i & "_chain0.wav", i
        g_Sound.LoadSound App.Path & "\sound\pop" & i & "_chain1.wav", 4 + i
        g_Sound.LoadSound App.Path & "\sound\pop" & i & "_chain2.wav", 8 + i
        g_Sound.LoadSound App.Path & "\sound\pop" & i & "_chain3.wav", 12 + i
    Next i

    ' Load the music
    g_Music.InitMusic

    ' Set the game state to menu
    i_GameState = GS_MENU

    ' Everything went good so return true
    InitGame = True

End Function

Public Sub NewGame(ByVal i_GameSpeed As Integer, ByVal i_GameMode As Integer, ByVal i_GameLevel As Integer)

    ' Generate a new set of blocks
    g_Blocks.NewGame

    ' Set the cursor position
    i_CursorX = FIELD_WIDTH - 1
    i_CursorY = 0

    ' Reset the stack
    i_StackHeight = 0

    ' Set the game state to countdown
    'i_GameState = GS_COUNTDOWN
    i_CountDown = TickCount() + 6000
    
    ' Set the difficulty level
    i_PopDelay = POP_DELAY_LENGTH / i_GameLevel
    i_PrefallLength = PREFALL_LENGTH / i_GameLevel
    i_FlashLength = FLASH_LENGTH / i_GameLevel
    
    ' Set the score to zip
    g_Score.i_Score = 0
    
    ' Set the game mode
    i_Mode = i_GameMode
    
    ' Stop any currently playing music
    g_Music.StopMusic
    
    ' Set the game speed to 1
    i_Speed = i_GameSpeed
    i_StackSpeed = 60 - (i_Speed * 2)

End Sub

Public Sub GameLoop()

    Dim i_X As Integer, i_Y As Integer, i_Count As Long, i_Alpha As Integer, i_Minor As Long, i_Major As Integer
    Dim i_StackY As Integer
    Static b_CursorMoved As Boolean, i_LastSec As Integer, i_BounceCounter As Integer, i_VelBounce As Integer, i_BounceY As Integer, b_StopBounce As Boolean

    ' Check to see if we're due for another frame update
    If TickCount >= i_LastTick + FRAME_RATE Then
        
        ' Clear the screen and start the D3D scene
        g_Graphics.Clear
        g_Graphics.StartRender
    
        ' Set the new tickcount
        i_LastTick = TickCount
        
        ' Blit the background
        g_Graphics.RenderSurface g_Bg
        
        ' Figure out what we need to do according to the current game state
        Select Case i_GameState
        Case GS_COUNTDOWN
        
            ' Get the current countdown
            i_Count = i_CountDown - TickCount
            
            ' Move the cursor down if need be
            If b_CursorMoved = False And b_Fade = False Then
                If i_CursorY <> (FIELD_HEIGHT / 2) + 1 Then
                    i_CursorY = i_CursorY + 1
                Else
                    If i_CursorX <> FIELD_WIDTH / 2 Then
                        i_CursorX = i_CursorX - 1
                    Else
                        b_CursorMoved = True
                    End If
                End If
            Else
                ' Handle keyboard stuff
                HandleKeyboard
            End If
            
            ' Figure up the block alpha for the fade in
            If b_Fade = False Then
                i_Alpha = (((i_Count / 1000) * 255) - 255) * -1
                If i_Alpha > 255 Or i_Count < 4050 Then i_Alpha = 255
            Else
                i_Alpha = 0
            End If
            
            ' Draw the blocks
            g_Blocks.DrawBlocks False, i_Alpha
            
            ' Draw the character
            g_Character.TimeStep
            
            ' Convert the counter number to seconds and display it
            i_Count = Int(i_Count / 1000)
            If i_Count <> i_LastSec Then
                If b_Fade = False Then _
                    g_Sound.PlaySound S_COUNTDOWN
                i_LastSec = i_Count
            End If
            g_Text.DrawText i_Count, (SCREEN_WIDTH - 12) / 2, ((SCREEN_HEIGHT - 16) / 2) - 50
            
            ' Blit the cursor
            i_X = ((SCREEN_WIDTH - (FIELD_WIDTH * BLOCK_WIDTH)) / 2) + ((i_CursorX - 1) * BLOCK_WIDTH) - 1
            i_Y = ((SCREEN_HEIGHT - (FIELD_HEIGHT * BLOCK_HEIGHT)) / 2) + ((i_CursorY - 1) * BLOCK_HEIGHT) - 1
            g_MainImg.Blit i_X, i_Y - i_StackHeight, 66, 34, 174, 14
            
            ' If we've counted all the way down start the game
            If i_Count < 1 Then
                i_GameState = GS_NORMAL
                g_Music.StartMusic S_MUSIC_NORMAL_INTRO
                i_GameStart = TickCount
                b_CursorMoved = False
                b_StopBounce = False
            End If
        
        Case GS_MENU
        
            g_Menu.HandleMenu
        
        Case GS_NORMAL
        
            ' Increment the game time
            i_GameTime = i_GameTime + 1
        
            ' Draw the score and time
            g_Text.DrawText g_Score.i_Score, 475, 82
            DisplayTime
        
            ' If we've got stop time going, draw that and update it too
            If i_StopTime > 0 Then
            
                ' Get the current amount of time left
                i_Count = (i_StopStart + i_StopTime) - TickCount
            
                ' If the count's up set the timer back to 0
                If i_Count <= 0 Then _
                    i_StopTime = 0
            
                ' Stop any kind of stack advancement
                b_Advance = False
            
                ' Figure up stuff for the timer
                i_Major = Int(i_Count / 1000)
            
                ' Draw the stop time
                g_Text.DrawText "Stop: " & i_Major, 445, 141
                
            Else
            
                g_Text.DrawText i_Speed, 475, 141
                i_StopStart = 0
                i_StopTime = 0
            
            End If
        
            ' Set that there hasn't been a chain this frame
            b_ChainFrame = False
            i_NumCleared = 0
            b_Swap = False
        
            ' Handle any key presses
            HandleKeyboard
            
            ' Let the blocks do their thing
            g_Blocks.HandleBlocks
        
            ' If the blocks advanced a line move the cursor up one
            If b_NewLine Then
                If i_CursorY > 1 Then _
                    i_CursorY = i_CursorY - 1
                ' Set new line back to false
                b_NewLine = False
            End If
        
            ' Draw everything
            If i_StackY = 1 Then
                g_Blocks.DrawBlocks True
            Else
                g_Blocks.DrawBlocks
            End If
        
            ' Handle and draw any particles
            g_Particles.HandleParticles
        
            ' Update and draw chain/combo cards
            g_Cards.HandleCards
            g_Character.TimeStep
        
            ' Blit the cursor
            i_X = ((SCREEN_WIDTH - (FIELD_WIDTH * BLOCK_WIDTH)) / 2) + ((i_CursorX - 1) * BLOCK_WIDTH) - 1
            i_Y = ((SCREEN_HEIGHT - (FIELD_HEIGHT * BLOCK_HEIGHT)) / 2) + ((i_CursorY - 1) * BLOCK_HEIGHT) - 1
            g_MainImg.Blit i_X, i_Y - i_StackHeight, 66, 34, 174, 14
        
        Case GS_PAUSED
        
            ' Show the pause menu
            g_Menu.PauseMenu
        
        Case GS_LOSE
            
            ' Add to the velocity
            If Not b_StopBounce Then
                i_VelBounce = i_VelBounce + 1
                i_BounceY = i_BounceY + i_VelBounce
            End If
            
            ' Draw the score and time
            g_Text.DrawText g_Score.i_Score, 475, 82
            DisplayTime True
            
            ' Draw the jiggling blocks
            g_Blocks.DrawBlocks True
            
            ' If we've hit the bounce point reverse the velocity at half the rate
            If i_BounceY >= 35 And b_StopBounce = False Then
                i_VelBounce = (i_VelBounce / 1.25) * -1
                 
                ' If the velocity is 0 display the play again menu
                If i_VelBounce = 0 Or i_VelBounce = 1 Then _
                    b_StopBounce = True
            ElseIf i_BounceY >= 35 And b_StopBounce = True Then
            
                ' Draw the menu
                g_Menu.GameOver
            
            End If
                
            ' Draw the thing
            g_Character.TimeStep
            g_Lose.Blit (SCREEN_WIDTH - 145) / 2, i_BounceY, 145, 123, 0, 0
            g_Lose.Blit (SCREEN_WIDTH - 177) / 2, i_BounceY + 38, 177, 46, 0, 126
        
        Case GS_TIMEUP
        
            ' Add to the velocity
            If Not b_StopBounce Then
                i_VelBounce = i_VelBounce + 1
                i_BounceY = i_BounceY + i_VelBounce
            End If
            
            ' Draw the score and time
            g_Text.DrawText g_Score.i_Score, 475, 82
            DisplayTime True
            
            ' Draw the jiggling blocks
            g_Blocks.DrawBlocks True
            
            ' If we've hit the bounce point reverse the velocity at half the rate
            If i_BounceY >= 35 And b_StopBounce = False Then
                i_VelBounce = (i_VelBounce / 1.25) * -1
                 
                ' If the velocity is 0 display the play again menu
                If i_VelBounce = 0 Or i_VelBounce = 1 Then _
                    b_StopBounce = True
                    
            ElseIf i_BounceY >= 35 And b_StopBounce = True Then
            
                ' Draw the menu
                g_Menu.GameOver
            
            End If
                
            ' Draw the thing
            g_Character.TimeStep
            g_Lose.Blit (SCREEN_WIDTH - 145) / 2, i_BounceY, 145, 123, 0, 0
            g_Lose.Blit (SCREEN_WIDTH - 177) / 2, i_BounceY + 38, 145, 46, 0, 184
        
        End Select
    
        ' If there's a fade in progress blit it
        If b_Fade Then
            i_FadeLevel = i_FadeLevel + i_FadeInc
            If i_FadeLevel >= 255 Then
                i_FadeLevel = 255
                i_FadeInc = -i_FadeInc
                i_GameState = i_FadeTo
            ElseIf i_FadeLevel <= 0 Then
                i_FadeLevel = 0
                b_Fade = False
            End If
            g_Lose.Blit 0, 0, SCREEN_WIDTH, SCREEN_HEIGHT, 165, 0, i_FadeLevel, 32, 32
        End If
        
        ' Flip it all to the screen
        g_Graphics.Flip
    
    End If

End Sub

Private Sub HandleKeyboard()

    ' See how we need to handle the keys
    Select Case i_GameState
    Case GS_NORMAL
    
        ' If the key counter is at zero or a key was actually pressed this frame
        ' we can take directional requestes
        If i_KeyCounter = 0 Or b_KeyPressed = True Then
        
            ' Handle left key...
            If b_KeyLeft = True And i_CursorX > 1 Then
                i_CursorX = i_CursorX - 1
                g_Sound.PlaySound S_CURSOR_MOVE
            End If
            
            ' ...right key...
            If b_KeyRight = True And i_CursorX < (FIELD_WIDTH - 1) Then
                i_CursorX = i_CursorX + 1
                g_Sound.PlaySound S_CURSOR_MOVE
            End If
            
            ' ...up key...
            If b_KeyUp = True And i_CursorY > 2 Then
                i_CursorY = i_CursorY - 1
                g_Sound.PlaySound S_CURSOR_MOVE
            End If
            
            ' ...and the down key
            If b_KeyDown = True And i_CursorY < FIELD_HEIGHT Then
                i_CursorY = i_CursorY + 1
                g_Sound.PlaySound S_CURSOR_MOVE
            End If
            
            ' Set key pressed back to false
            b_KeyPressed = False
        
        End If
    
        ' If the key counter isn't zero, count it down
        If i_KeyCounter > 0 Then _
            i_KeyCounter = i_KeyCounter - 1
    
        ' If the swap button was pressed swap the blocks and set it back to false
        If b_KeySwap = True Then
            g_Blocks.Swap i_CursorX, i_CursorY
            b_KeySwap = False
        End If
        
        ' If the enter key was pressed pause the thing
        If b_KeyEnter = True Then
            i_GameState = GS_PAUSED
            b_KeyEnter = False
        End If
        
        ' If the advance button was pressed advance the stack
        If b_KeyAdvance Then
            i_StopTime = 0
            b_Advance = True
        End If
    
    Case GS_COUNTDOWN
    
        ' If the key counter is at zero or a key was actually pressed this frame
        ' we can take directional requestes
        If i_KeyCounter = 0 Or b_KeyPressed = True Then
        
            ' Handle left key...
            If b_KeyLeft = True And i_CursorX > 1 Then
                i_CursorX = i_CursorX - 1
                g_Sound.PlaySound S_CURSOR_MOVE
            End If
            
            ' ...right key...
            If b_KeyRight = True And i_CursorX < (FIELD_WIDTH - 1) Then
                i_CursorX = i_CursorX + 1
                g_Sound.PlaySound S_CURSOR_MOVE
            End If
            
            ' ...up key...
            If b_KeyUp = True And i_CursorY > 2 Then
                i_CursorY = i_CursorY - 1
                g_Sound.PlaySound S_CURSOR_MOVE
            End If
            
            ' ...and the down key
            If b_KeyDown = True And i_CursorY < FIELD_HEIGHT Then
                i_CursorY = i_CursorY + 1
                g_Sound.PlaySound S_CURSOR_MOVE
            End If
            
            ' Set key pressed back to false
            b_KeyPressed = False
        
        End If
    
        ' If the key counter isn't zero, count it down
        If i_KeyCounter > 0 Then _
            i_KeyCounter = i_KeyCounter - 1
    
    End Select

End Sub

Private Sub DisplayTime(Optional b_NoInc As Boolean)
    
    Dim i_Min As Integer, i_Sec As Integer, s_Time As String, i_Small As Integer
    Static i_LastSec As Integer
   
    ' Only update the time if we weren't told not to
    If Not b_NoInc Then
        Select Case i_Mode
        Case GM_ENDLESS
            i_Length = TickCount - i_GameStart
        Case GM_TIMETRIAL
            i_Length = (i_GameStart + 120000) - TickCount
        End Select
    End If

    ' Get the seconds from the new number
    i_Sec = Int(i_Length / 1000)
    i_Min = Int(i_Sec / 60)
    i_Sec = i_Sec - (i_Min * 60)

    ' If we're doing time trial and hit 0:00 end the game
    If i_Sec = 0 And i_Min = 0 And Not b_NoInc And i_Mode = GM_TIMETRIAL Then
        i_GameState = GS_TIMEUP
        Exit Sub
    ElseIf i_Sec < 15 And i_Min = 0 And Not b_Inc And i_Mode = GM_TIMETRIAL And i_LastSec <> i_Sec Then
        g_Sound.PlaySound S_COUNTDOWN
    End If

    ' Add a leading zero to the second if necessary
    If i_Sec < 10 Then
        s_Time = i_Min & ":0" & i_Sec
    Else
        s_Time = i_Min & ":" & i_Sec
    End If
    i_LastSec = i_Sec

    ' Draw the time
    g_Text.DrawText s_Time, 100, 82

End Sub

Public Sub ChangeVideoMode()

    ' Reset the video mode
    g_Graphics.ChangeVideoMode SCREEN_WIDTH, SCREEN_HEIGHT, 16, b_Fullscreen

    ' If we're not fullscreen any more, center the form
    If b_Fullscreen = False Then
        frmMain.Left = (Screen.Width - frmMain.Width) / 2
        frmMain.Top = (Screen.Height - frmMain.Height) / 2
    End If

    ' Reload all the surfaces
    Set g_Bg = g_Graphics.g_Device.CreateImageSurface(640, 480, i_DisplayFormat)
    g_Graphics.g_D3dx.LoadSurfaceFromFile g_Bg, ByVal 0, ByVal 0, App.Path & "\images\field.png", ByVal 0, D3DX_DEFAULT, 0, ByVal 0
    g_MainImg.CreateSurface g_Graphics.g_Device, App.Path & "\images\blocks.png", 256, 256
    
    g_Character.Initialize g_Graphics, False
    g_Text.Initialize g_Graphics
    g_Menu.Init g_Graphics

End Sub
