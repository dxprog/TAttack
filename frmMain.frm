VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "TAttack"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        b_KeyLeft = False
        b_KeyUp = False
        b_KeyDown = False
        b_KeyRight = False
        i_KeyCounter = 20
        b_KeyPressed = True
    End If
    
    Select Case KeyCode
    Case vbKeyLeft
        b_KeyLeft = True
    Case vbKeyRight
        b_KeyRight = True
    Case vbKeyUp
        b_KeyUp = True
    Case vbKeyDown
        b_KeyDown = True
    Case vbKeyEscape
        End
    Case vbKeySpace
        b_KeySwap = True
    Case vbKeyShift
        b_KeyAdvance = True
    Case vbKeyReturn
        b_KeyEnter = True
    End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyLeft
        b_KeyLeft = False
    Case vbKeyRight
        b_KeyRight = False
    Case vbKeyUp
        b_KeyUp = False
    Case vbKeyDown
        b_KeyDown = False
    Case vbKeySpace
        b_KeySwap = False
    Case vbKeyL
        b_KeyL = True
    Case vbKeyShift
        b_KeyAdvance = False
    Case vbKeyReturn
        b_KeyEnter = False
    End Select

End Sub

Private Sub Form_Load()
    
    'Dim i_Counter As Long, i_FramesDrawn As Long, i_LastFrame As Long, i_FrameRate As Long
    'Dim g_Bg As DirectDrawSurface7
    'Dim i_FPS As Integer, i_LastTick As Long
    'i_FPS = 1000 / 60
    
    'Me.Show
    
    'Call thing.SetDisplayMode(dx, Me.hWnd, 640, 480, 16)
    'Set g_Cursor = thing.CreateSurface(App.Path & "\images\cursor.bmp", 64, 32, True)
    'Set g_Bg = thing.CreateSurface(App.Path & "\images\field.bmp", 640, 480, False)
    'g_Blocks.InitBlocks thing
    'g_Blocks.NewGame
    
    'i_CursorX = 3
    'i_CursorY = 10
    
    'Do
    
    '    If i_LastTick + FRAME_RATE < dx.TickCount Then
        
    '        DoEvents
            
    '        thing.ClearScreen
            
            'thing.Blit g_Bg, 0, 0, 640, 480, 0, 0
            
    '        If i_KeyCounter = 0 Then
            
    '            If b_KeyLeft And i_CursorX > 1 Then
    '                i_CursorX = i_CursorX - 1
    '                i_KeyCounter = 30
    '            End If
                
    '            If b_KeyRight And i_CursorX < FIELD_WIDTH - 1 Then
    '                i_CursorX = i_CursorX + 1
    '                i_KeyCounter = 30
    '            End If
            
    '            If b_KeyUp And i_CursorY > 1 Then
    '                i_CursorY = i_CursorY - 1
    '                i_KeyCounter = 30
    '            End If
                
    '            If b_KeyDown And i_CursorY < FIELD_HEIGHT Then
    '                i_CursorY = i_CursorY + 1
    '                i_KeyCounter = 30
    '            End If
            
    '        Else
    '            i_KeyCounter = i_KeyCounter - 1
    '        End If
            
    '        g_Blocks.HandleBlocks thing
    '        g_Blocks.DrawBlocks thing
            
    '        If b_NewLine Then _
    '            i_CursorY = i_CursorY - 1
    '        If i_CursorY = 0 Then i_CursorY = 1
            
    '        X = 224 + ((i_CursorX - 1) * 32)
    '        Y = 48 + ((i_CursorY - 1) * 32)
    '        thing.Blit g_Cursor, X, Y - i_StackHeight, 64, 32, 0, 0, True
    '        thing.DrawText i_FramesDrawn & " FPS", 0, 0
            'thing.DrawText "x" & i_Chain, 0, 20
    '        g_Blocks.BlockDebug i_CursorX, i_CursorY, thing
    '        thing.Flip DDFLIP_WAIT
    '        i_FramesDrawn = i_FramesDrawn + 1
            
    '        b_NewLine = False
        
    '    End If
        
    '    If dx.TickCount > i_Counter + 1000 Then
    '        i_FrameRate = i_FramesDrawn
    '        i_FramesDrawn = 0
    '        i_Counter = dx.TickCount
    '    End If
        
    'Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
