Attribute VB_Name = "modMain"
Public g_Game As New clsEngine, b_Music As Boolean

Public Sub Main()

    ' Show the form
    frmMain.Show

    ' If there was a problem initializing the game close
    If g_Game.InitGame(frmMain.hWnd) = False Then End
    
    ' The Loop
    Do
    
        ' Let Windows handle stuff
        DoEvents
        
        ' Run the main game loop
        g_Game.GameLoop
    
        ' Handle the music stuff
        g_Music.HandleMusic
    
    Loop

End Sub
