VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim b_Started As Boolean, b_Intro As Boolean
Dim i_Music As Integer

Public Sub InitMusic()

    ' Load the sounds
    g_Sound.LoadSound App.Path & "\sound\mus_normal_intro.wav", S_MUSIC_NORMAL_INTRO
    g_Sound.LoadSound App.Path & "\sound\mus_normal_loop.wav", S_MUSIC_NORMAL_LOOP
    g_Sound.LoadSound App.Path & "\sound\mus_menu_intro.wav", S_MUSIC_MENU_INTRO
    g_Sound.LoadSound App.Path & "\sound\mus_menu_loop.wav", S_MUSIC_MENU_LOOP
    g_Sound.LoadSound App.Path & "\sound\mus_panic_intro.wav", S_MUSIC_PANIC_INTRO
    g_Sound.LoadSound App.Path & "\sound\mus_panic_loop.wav", S_MUSIC_PANIC_LOOP

End Sub

Public Sub HandleMusic()

    ' If music has been turned off or the intro has already played jump out
    If b_Music = False Or b_Intro = True Then Exit Sub

    ' If the music hasn't been started, jump ship
    If b_Started = False Then Exit Sub

    ' Check to see if the current loop is still playing
    If g_Sound.GetPlayStatus(i_Music) <> DSBSTATUS_PLAYING Then
    
        ' The intro is done play the loop
        g_Sound.PlaySound i_Music + 1, DSBPLAY_LOOPING
        b_Intro = True
    
    End If

End Sub

Public Sub StartMusic(ByVal i_MusicNum As Integer)
    
    ' If the busic wasn't loaded or has already been started, get out of here
    If b_Music = False Or (b_Started = True And i_MusicNum = i_Music) Then Exit Sub
    
    ' Stop anything that's currently playing
    StopMusic
    
    ' Change the music number
    i_Music = i_MusicNum
    
    ' Start from the beginning and play
    g_Sound.PlaySound i_Music
    b_Started = True
    b_Intro = False
    
End Sub

Public Sub StopMusic()

    ' If there isn't any music playing bail out
    If b_Started = False Or i_Music = 0 Then Exit Sub

    ' Set the playing falg to false
    b_Intro = False
    b_Started = False
    g_Sound.StopSound i_Music
    g_Sound.StopSound i_Music + 1

End Sub
