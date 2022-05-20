VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The surfaces
Private g_Bg As New clsSurface
Private g_Labels As New clsSurface

' The game options
Private i_Selected As Integer
Private i_GameLevel As Integer
Private i_Speed As Integer
Private i_GameType As Integer

' Whether to continue doing menu stuff
Private b_Finished As Boolean, b_MusicChange As Boolean

Public Sub Init(ByVal g_Graphics As clsDirect3D)

    ' Load the surfaces
    g_Bg.CreateSurface g_Graphics.g_Device, App.Path & "\images\menu.png", 256, 256
    g_Labels.CreateSurface g_Graphics.g_Device, App.Path & "\images\menu_labels.png", 256, 256
    
    ' Set the generic game stuff
    If i_GameSpeed = 0 Then _
        i_GameLevel = 1
    
    If i_Speed = 0 Then _
        i_Speed = 1
    
    If i_GameType = 0 Then _
        i_GameType = 1
    b_MusicChange = True

End Sub

Public Sub LoadDefaults()

    ' Save the menu settings
    Open App.Path & "\settings.dat" For Binary As #1
    Get #1, , b_Fullscreen
    Get #1, , b_Music
    Get #1, , i_GameLevel
    Get #1, , i_Speed
    Get #1, , i_GameType
    Close #1

End Sub

Public Sub HandleMenu()

    Static i_BgX As Integer, i_BgY As Integer

    ' Update the background coordinates
    i_BgX = i_BgX + 1
    i_BgY = i_BgY + 1
    If i_BgX > 32 Then
        i_BgX = 0
        i_BgY = 0
    End If
    
    ' Do keyboard stuff
    HandleKeyboard
    
    ' Blit the background
    For i = -1 To 21
        For j = -1 To 16
        
            g_Bg.Blit (i * 32) - i_BgX, (j * 32) + i_BgY, 32, 32, 0, 0
        
        Next j
    Next i

    ' Blit the header
    g_Bg.Blit (SCREEN_WIDTH - 512) / 2, 10, 18, 72, 40, 0
    g_Bg.Blit ((SCREEN_WIDTH - 512) / 2) + 18, 10, 124, 72, 52, 0, , 10, 72
    g_Bg.Blit (SCREEN_WIDTH - 228) / 2, 10, 228, 72, 0, 184
    g_Bg.Blit ((SCREEN_WIDTH - 228) / 2) + 227, 10, 125, 72, 52, 0, , 10, 72
    g_Bg.Blit 558, 10, 18, 72, 61, 0
    
    ' Blit the main section thingy
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2, 90, 25, 25, 91, 0
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2 + 25, 90, 488, 25, 116, 0, , 25, 25
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2 + 513, 90, 25, 25, 141, 0
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2, 115, 25, 330, 91, 25, , 25, 25
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2 + 25, 115, 488, 330, 114, 25, , 25, 25
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2 + 513, 115, 25, 330, 141, 25, , 25, 25
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2, 445, 25, 25, 91, 50
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2 + 25, 445, 488, 25, 116, 50, , 25, 25
    g_Bg.Blit (SCREEN_WIDTH - 538) / 2 + 513, 445, 25, 25, 141, 50
    
    ' Draw the box for the selected item
    If i_Selected < 5 Then
        g_Bg.Blit (SCREEN_WIDTH - 532) / 2, 106 + (i_Selected * 57), 532, 57, 0, 91, , 256, 57
    Else
        If i_Selected = 5 Then
            g_Bg.Blit (SCREEN_WIDTH - 532) / 2, 391, 266, 57, 0, 91, , 256, 57
        Else
            g_Bg.Blit (SCREEN_WIDTH - 532) / 2 + 266, 391, 266, 57, 0, 91, , 256, 57
        End If
    End If
    
    ' Draw the options labels
    For i = 0 To 4
        g_Labels.Blit (SCREEN_WIDTH - 520) / 2, 117 + (i * 57), 100, 35, 0, (i * 35)
    Next i
    
    ' Draw the game mode setting
    If i_GameType = 1 Then
        g_Labels.Blit SCREEN_WIDTH / 2, 117, 100, 22, 150, 4
        g_Labels.Blit SCREEN_WIDTH / 2 + 120, 122, 140, 21, 140, 28, 128
    Else
        g_Labels.Blit SCREEN_WIDTH / 2, 122, 100, 22, 150, 4, 128
        g_Labels.Blit SCREEN_WIDTH / 2 + 120, 117, 140, 21, 140, 28
    End If
    
    ' Draw the speed level
    g_Labels.Blit SCREEN_WIDTH / 2 + 120, 179, 22, 25, 12, 197
    g_Labels.Blit SCREEN_WIDTH / 2 + 170, 179, 22, 25, 36, 197
    g_Text.DrawText i_Speed, SCREEN_WIDTH / 2 + 138, 179
    
    ' Draw the music setting
    If b_Music = True Then
        If b_MusicChange Then
            g_Music.StartMusic S_MUSIC_MENU_INTRO
            b_MusicChange = False
        End If
        g_Labels.Blit SCREEN_WIDTH / 2 + 120, 293, 38, 18, 218, 122
        g_Labels.Blit SCREEN_WIDTH / 2 + 170, 298, 55, 18, 208, 144, 128
    Else
        g_Music.StopMusic
        g_Labels.Blit SCREEN_WIDTH / 2 + 120, 298, 38, 18, 218, 122, 128
        g_Labels.Blit SCREEN_WIDTH / 2 + 170, 293, 55, 18, 208, 144
    End If
    
    ' Draw the difficulty level stuff
    Select Case i_GameLevel
    Case 1
        g_Labels.Blit SCREEN_WIDTH / 2 - 10, 236, 65, 21, 189, 51
        g_Labels.Blit SCREEN_WIDTH / 2 + 70, 241, 91, 21, 166, 75, 128
        g_Labels.Blit SCREEN_WIDTH / 2 + 170, 241, 71, 21, 192, 97, 128
    Case 2
        g_Labels.Blit SCREEN_WIDTH / 2 - 10, 241, 65, 21, 189, 51, 128
        g_Labels.Blit SCREEN_WIDTH / 2 + 70, 236, 91, 21, 166, 75
        g_Labels.Blit SCREEN_WIDTH / 2 + 170, 241, 71, 21, 192, 97, 128
    Case 3
        g_Labels.Blit SCREEN_WIDTH / 2 - 10, 241, 65, 21, 189, 51, 128
        g_Labels.Blit SCREEN_WIDTH / 2 + 70, 241, 91, 21, 166, 75, 128
        g_Labels.Blit SCREEN_WIDTH / 2 + 170, 236, 71, 21, 192, 97
    End Select
    
    ' Draw the video mode stuff
    If b_Fullscreen = True Then
        g_Labels.Blit SCREEN_WIDTH / 2 - 40, 345, 142, 21, 116, 189
        g_Labels.Blit SCREEN_WIDTH / 2 + 120, 355, 142, 21, 132, 164, 128
    Else
        g_Labels.Blit SCREEN_WIDTH / 2 - 40, 355, 142, 21, 116, 189, 128
        g_Labels.Blit SCREEN_WIDTH / 2 + 120, 345, 142, 21, 132, 164
    End If
    
    ' Draw the start and quit buttons
    g_Labels.Blit (SCREEN_WIDTH - 470) / 2, 425, 166, 8, 86, 212
    g_Labels.Blit (SCREEN_WIDTH - 470) / 2 + 270, 425, 166, 8, 86, 212
    g_Labels.Blit (SCREEN_WIDTH - 470) / 2 + 50, 405, 64, 29, 87, 226
    g_Labels.Blit (SCREEN_WIDTH - 470) / 2 + 335, 405, 48, 29, 166, 225

End Sub

Private Sub HandleKeyboard()

    ' If the key counter is at zero or a key was actually pressed this frame
    ' we can take directional requestes
    If i_KeyCounter = 0 Or b_KeyPressed = True Then
    
        ' Handle left key...
        If b_KeyLeft = True Then

            ' See what we're going to be changing
            Select Case i_Selected
            Case 0
                
                ' Change the game mode
                If i_GameType = 2 Then _
                    i_GameType = 1
                    
            Case 1
            
                ' Decrement the speed
                If i_Speed > 1 Then _
                    i_Speed = i_Speed - 1
            
            Case 3
            
                ' Change the music
                If b_Music = False Then
                    b_MusicChange = True
                    b_Music = True
                End If
            
            Case 4
            
                ' Change the windowed mode
                If b_Fullscreen = False Then
                    b_Fullscreen = True
                    g_Game.ChangeVideoMode
                End If
            
            Case 2
            
                ' Increment the game level
                If i_GameLevel > 1 Then i_GameLevel = i_GameLevel - 1
            
            Case 6
            
                ' Move over to start
                i_Selected = 5
            
            End Select

        End If
        
        ' ...right key...
        If b_KeyRight = True Then
            
            
            ' See what we're going to be changing
            Select Case i_Selected
            Case 0
                
                ' Change the game mode
                If i_GameType = 1 Then _
                    i_GameType = 2
            
            Case 1
            
                ' Increment the speed
                If i_Speed < 99 Then _
                    i_Speed = i_Speed + 1
            
            Case 3
            
                ' Change the music
                If b_Music Then
                    b_Music = False
                    b_MusicChange = True
                End If
            
            Case 2
            
                ' Decrement the game level
                If i_GameLevel < 3 Then i_GameLevel = i_GameLevel + 1
            
            Case 4
            
                ' Move to windowed mode
                If b_Fullscreen = True Then
                    b_Fullscreen = False
                    g_Game.ChangeVideoMode
                End If
            
            Case 5
            
                ' Move over to quit
                i_Selected = 6
            
            End Select
            
            
        End If
        
        ' ...up key...
        If b_KeyUp = True And i_Selected > 0 Then
            i_Selected = i_Selected - 1
            
        End If
        
        ' ...and the down key
        If b_KeyDown = True And i_Selected < 5 Then
            i_Selected = i_Selected + 1
            
        End If
        
        ' Set key pressed back to false
        b_KeyPressed = False
    
    End If

    If b_KeyEnter = True And i_Selected = 5 Then
        b_Finished = True
        g_Game.NewGame i_GameSpeed, i_GameType, i_GameLevel
        g_Game.b_Fade = True
        g_Game.i_FadeTo = GS_COUNTDOWN
        g_Game.i_FadeInc = 10
        g_Game.i_FadeLevel = 0
        g_Music.StopMusic
        i_Selected = 0
        
        ' Save the menu settings
        Open App.Path & "\settings.dat" For Binary As #1
        Put #1, , b_Fullscreen
        Put #1, , b_Music
        Put #1, , i_GameLevel
        Put #1, , i_Speed
        Put #1, , i_GameType
        Close #1
        
    End If
    
    If b_KeyEnter = True And i_Selected = 6 Then
        End
    End If

    ' If the key counter isn't zero, count it down
    If i_KeyCounter > 0 Then _
        i_KeyCounter = i_KeyCounter - 1

End Sub

Public Sub PauseMenu()

    Static i_BgX As Integer, i_BgY As Integer

    ' Update the background coordinates
    i_BgX = i_BgX + 1
    i_BgY = i_BgY + 1
    If i_BgX > 32 Then
        i_BgX = 0
        i_BgY = 0
    End If

    ' Draw the background
    For i = 0 To 8
        For j = 0 To 5
        
            If (i > 0 And i < 8) And (j > 0 And j < 5) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2) + i_BgY + ((j - 1) * 32), 32, 32, 0, 0, 225
            ElseIf i = 8 And (j > 0 And j < 5) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2) + i_BgY + ((j - 1) * 32), i_BgX, 32, 0, 0, 225
            ElseIf i = 0 And (j > 0 And j < 5) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2), ((SCREEN_HEIGHT - 160) / 2) + i_BgY + ((j - 1) * 32), 32 - i_BgX, 32, i_BgX, 0, 225
            ElseIf j = 5 And (i > 0 And i < 8) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2) + i_BgY + ((j - 1) * 32), 32, 32 - i_BgY, 0, 0, 225
            ElseIf j = 0 And (i > 0 And i < 8) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2), 32, i_BgY, 0, 32 - i_BgY, 225
            ElseIf j = 0 And i = 0 Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2), ((SCREEN_HEIGHT - 160) / 2), 32 - i_BgX, i_BgY, i_BgX, 32 - i_BgY, 225
            ElseIf j = 0 And i = 8 Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2), i_BgX, i_BgY, 0, 32 - i_BgY, 225
            ElseIf j = 5 And i = 0 Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2), ((SCREEN_HEIGHT - 160) / 2) + ((j - 1) * 32) + i_BgY, 32 - i_BgX, 32 - i_BgY, i_BgX, 0, 225
            Else
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2) + ((j - 1) * 32) + i_BgY, i_BgX, 32 - i_BgY, 0, 0, 225
            End If
        
        Next j
    Next i

    ' Draw some nice text
    g_Text.DrawText "Do you want", (SCREEN_WIDTH - 170) / 2, (SCREEN_HEIGHT - 150) / 2
    g_Text.DrawText "to quit?", (SCREEN_WIDTH - 112) / 2, (SCREEN_HEIGHT - 100) / 2
    
    ' Handle keyboard stuff
    If b_KeyLeft And i_Selected = 1 Then
        i_Selected = 0
    End If
    
    If b_KeyRight And i_Selected = 0 Then
        i_Selected = 1
    End If
    
    If b_KeyEnter Then
    
        If i_Selected = 0 Then
            g_Game.i_GameState = GS_NORMAL
            b_KeyEnter = False
        Else
            ' Go back to the menu
            g_Game.b_Fade = True
            g_Game.i_FadeTo = GS_MENU
            g_Game.i_FadeInc = 10
            g_Game.i_FadeLevel = 0
            b_MusicChange = True
            b_KeyEnter = False
            i_Selected = 0
            g_Music.StopMusic
            b_Finished = False
        End If
    
    End If
    
    ' Draw the options
    If i_Selected = 0 Then
        g_Labels.Blit (SCREEN_WIDTH - 150) / 2, SCREEN_HEIGHT / 2, 35, 18, 124, 124
        g_Labels.Blit (SCREEN_WIDTH / 2) + 25, (SCREEN_HEIGHT / 2) + 5, 52, 20, 10, 228, 128
    Else
        g_Labels.Blit (SCREEN_WIDTH - 150) / 2, (SCREEN_HEIGHT / 2) + 5, 35, 18, 124, 124, 128
        g_Labels.Blit (SCREEN_WIDTH / 2) + 25, (SCREEN_HEIGHT / 2), 52, 20, 10, 228
    End If

End Sub

Public Sub GameOver()

    Static i_BgX As Integer, i_BgY As Integer

    ' Update the background coordinates
    i_BgX = i_BgX + 1
    i_BgY = i_BgY + 1
    If i_BgX > 32 Then
        i_BgX = 0
        i_BgY = 0
    End If

    ' Draw the background
    For i = 0 To 8
        For j = 0 To 5
        
            If (i > 0 And i < 8) And (j > 0 And j < 5) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2) + i_BgY + ((j - 1) * 32), 32, 32, 0, 0, 225
            ElseIf i = 8 And (j > 0 And j < 5) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2) + i_BgY + ((j - 1) * 32), i_BgX, 32, 0, 0, 225
            ElseIf i = 0 And (j > 0 And j < 5) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2), ((SCREEN_HEIGHT - 160) / 2) + i_BgY + ((j - 1) * 32), 32 - i_BgX, 32, i_BgX, 0, 225
            ElseIf j = 5 And (i > 0 And i < 8) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2) + i_BgY + ((j - 1) * 32), 32, 32 - i_BgY, 0, 0, 225
            ElseIf j = 0 And (i > 0 And i < 8) Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2), 32, i_BgY, 0, 32 - i_BgY, 225
            ElseIf j = 0 And i = 0 Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2), ((SCREEN_HEIGHT - 160) / 2), 32 - i_BgX, i_BgY, i_BgX, 32 - i_BgY, 225
            ElseIf j = 0 And i = 8 Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2), i_BgX, i_BgY, 0, 32 - i_BgY, 225
            ElseIf j = 5 And i = 0 Then
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2), ((SCREEN_HEIGHT - 160) / 2) + ((j - 1) * 32) + i_BgY, 32 - i_BgX, 32 - i_BgY, i_BgX, 0, 225
            Else
                g_Bg.Blit ((SCREEN_WIDTH - 256) / 2) + (i * 32) - i_BgX, ((SCREEN_HEIGHT - 160) / 2) + ((j - 1) * 32) + i_BgY, i_BgX, 32 - i_BgY, 0, 0, 225
            End If
        
        Next j
    Next i

    ' Draw some nice text
    g_Text.DrawText "Want to play", (SCREEN_WIDTH - 170) / 2, (SCREEN_HEIGHT - 150) / 2
    g_Text.DrawText "again?", (SCREEN_WIDTH - 112) / 2, (SCREEN_HEIGHT - 100) / 2
    
    ' Handle keyboard stuff
    If b_KeyLeft And i_Selected = 1 Then
        i_Selected = 0
    End If
    
    If b_KeyRight And i_Selected = 0 Then
        i_Selected = 1
    End If
    
    If b_KeyEnter Then
    
        If i_Selected = 0 Then
            g_Game.NewGame i_GameSpeed, i_GameType, i_GameLevel
            g_Music.StopMusic
            b_KeyEnter = False
        Else
            ' Go back to the menu
            g_Game.b_Fade = True
            g_Game.i_FadeTo = GS_MENU
            g_Game.i_FadeInc = 10
            g_Game.i_FadeLevel = 0
            b_MusicChange = True
            b_KeyEnter = False
            i_Selected = 0
            g_Music.StopMusic
            b_Finished = False
            'g_Music.StopMusic
        End If
    
    End If
    
    ' Draw the options
    If i_Selected = 0 Then
        g_Labels.Blit (SCREEN_WIDTH - 150) / 2, SCREEN_HEIGHT / 2, 52, 20, 10, 228
        g_Labels.Blit (SCREEN_WIDTH / 2) + 25, (SCREEN_HEIGHT / 2) + 5, 35, 18, 124, 124, 128
    Else
        g_Labels.Blit (SCREEN_WIDTH - 150) / 2, (SCREEN_HEIGHT / 2) + 5, 52, 20, 10, 228, 128
        g_Labels.Blit (SCREEN_WIDTH / 2) + 25, (SCREEN_HEIGHT / 2), 35, 18, 124, 124
    End If

End Sub

