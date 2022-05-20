VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The DDraw variables
Private g_DDraw As DirectDraw7
Private g_Primary As DirectDrawSurface7
Private g_Backbuffer As DirectDrawSurface7

' Screen resolution
Private i_ScreenW As Integer, i_ScreenH As Integer

' The form handle
Private g_hWnd As Long

Public Function SetDisplayMode(ByRef g_Dx As DirectX7, ByVal h_Wnd As Long, ByVal i_Width As Integer, ByVal i_Height As Integer, ByVal i_BPP As Integer) As Boolean

    On Local Error GoTo errOut
    Dim g_Ddsd As DDSURFACEDESC2, g_Caps As DDSCAPS2

    ' Create the DDraw object
    Set g_DDraw = g_Dx.DirectDrawCreate("")
    
    ' Set the cooperative level and resolution
    g_DDraw.SetCooperativeLevel h_Wnd, DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
    g_DDraw.SetDisplayMode i_Width, i_Height, i_BPP, 0, DDSDM_DEFAULT
    
    ' Set up the primary surface
    g_Ddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    g_Ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    g_Ddsd.lBackBufferCount = 2 ' Tells DDraw we'll be doing triple buffering
    Set g_Primary = g_DDraw.CreateSurface(g_Ddsd)
    
    ' Set up the backbuffer
    g_Caps.lCaps = DDSCAPS_BACKBUFFER
    Set g_Backbuffer = g_Primary.GetAttachedSurface(g_Caps)
    g_Backbuffer.GetSurfaceDesc g_Ddsd
    
    ' Set the screen resolution and window handle for later use
    i_ScreenW = g_Ddsd.lWidth
    i_ScreenH = g_Ddsd.lHeight
    g_hWnd = h_Wnd
    
    ' Return true
    SetDisplayMode = True

errOut:

    ' If there was an error log it and return false
    If Err.Number <> 0 Then
    
        Open App.Path & "\error.log" For Append As #1
        Print #1, "---------------------------------"
        Print #1, "SetDisplayMode"
        Print #1, "Error " & Err.Number
        Print #1, Err.Description
        Close #1
        
        SetDisplayMode = False
    
    End If

End Function

Public Function CreateSurface(ByVal s_FileName As String, ByVal i_Width As Long, ByVal i_Height As Integer, Optional b_Colorkey As Boolean = False) As DirectDrawSurface7

    On Local Error GoTo errOut
    
    Dim g_Ddsd As DDSURFACEDESC2, g_ColKey As DDCOLORKEY, g_PixF As DDPIXELFORMAT
    Dim b_SurfCreated As Boolean
    
    ' Set to use a standard surface
    g_Ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    g_Ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    ' Load the surface
    Set CreateSurface = g_DDraw.CreateSurfaceFromFile(s_FileName, g_Ddsd)

    ' If a colorkey was requested set that up now
    If b_Colorkey Then
    
        ' Get the pixel format
        CreateSurface.GetPixelFormat g_PixF
        
        ' Set the color key to be magenta
        g_ColKey.high = g_PixF.lBBitMask + g_PixF.lRBitMask
        g_ColKey.low = g_ColKey.high
        
        ' Set the color key
        CreateSurface.SetColorKey DDCKEY_SRCBLT, g_ColKey
    
    End If
    
errOut:

    ' If there was an error log it and return false
    If Err.Number <> 0 Then
    
        Open App.Path & "\error.log" For Append As #1
        Print #1, "---------------------------------"
        Print #1, "CreateSurface (" & s_FileName & ")"
        Print #1, "Surface created: " & b_SurfCreated
        Print #1, "Error " & Err.Number
        Print #1, Err.Description
        Close #1
        
    End If

End Function

Public Sub Blit(ByRef g_Surface As DirectDrawSurface7, ByVal i_DestX As Integer, ByVal i_DestY As Integer, ByVal i_Width As Integer, ByVal i_Height As Integer, ByVal i_SrcX As Integer, ByVal i_SrcY As Integer, Optional b_Colorkey As Boolean = False)

On Local Error GoTo errOut

    Dim t_Rect As RECT, l_RetVal As Long
    
    ' Make sure the object is still on the screen before doing anything else
    If (i_DestX + i_Width < 0 Or i_DestX > i_ScreenW) Or (i_DestY + i_Height < 0 Or i_DestY > i_ScreenH) Then _
        Exit Sub
    
    ' Set up the rect parameters
    With t_Rect
    
        ' If the object is off the left side of the screen adjust the rect so it
        ' doesn't disappear
        If i_DestX < 0 Then
            t_Rect.Left = i_SrcX + Abs(i_DestX)
            i_DestX = 0
        Else
            t_Rect.Left = i_SrcX
        End If
        
        ' Do the same for the right side...
        If i_DestX + i_Width > i_ScreenW Then
            t_Rect.Right = i_SrcX + (i_Width - ((i_DestX + i_Width) - i_ScreenW))
        Else
            t_Rect.Right = i_SrcX + i_Width
        End If
        
        ' ...and the top...
        If i_DestY < 0 Then
            t_Rect.Top = i_SrcY + Abs(i_DestY)
            i_DestY = 0
        Else
            t_Rect.Top = i_SrcY
        End If
        
        ' ...and finally, the bottom.
        If i_DestY + l_Height > i_ScreenH Then
            t_Rect.Bottom = i_SrcY + (i_Height - ((i_DestY + i_Height) - i_ScreenH))
        Else
            t_Rect.Bottom = i_SrcY + i_Height
        End If
    
    End With

    ' Blit the stupid thing (use colorkey if necessary)
    If b_Colorkey Then
        l_RetVal = g_Backbuffer.BltFast(i_DestX, i_DestY, g_Surface, t_Rect, DDBLTFAST_SRCCOLORKEY)
    Else
        l_RetVal = g_Backbuffer.BltFast(i_X, i_Y, g_Surface, t_Rect, DDBLTFAST_WAIT)
    End If

    ' If ddraw is reporting that the surface was lost restore all surfaces and try again
    If l_RetVal = DDERR_SURFACELOST Then
        
        ' Restore the surfaces
        g_DDraw.RestoreAllSurfaces
        
        ' Re-blit
        If b_Colorkey Then
            l_RetVal = g_Backbuffer.BltFast(i_DestX, i_DestY, g_Surface, t_Rect, DDBLTFAST_SRCCOLORKEY)
        Else
            l_RetVal = g_Backbuffer.BltFast(i_X, i_Y, g_Surface, t_Rect, DDBLTFAST_WAIT)
        End If
    
    End If
    
errOut:

    If Err.Number <> 0 Then _
        Debug.Print Err.Description
    
End Sub

Public Sub Flip(ByVal e_Flip As CONST_DDFLIPFLAGS)

    g_Primary.Flip Nothing, e_Flip

End Sub

Private Sub Class_Terminate()

    ' Close everything down
    g_DDraw.RestoreDisplayMode
    g_DDraw.SetCooperativeLevel g_hWnd, DDSCL_NORMAL

End Sub

Public Sub DrawText(ByVal s_Text As String, ByVal X As Long, ByVal Y As Long)

    Dim i_OldColor As Long
    i_OldColor = g_Backbuffer.GetForeColor
    g_Backbuffer.SetForeColor vbWhite
    g_Backbuffer.DrawText X, Y, s_Text, False
    g_Backbuffer.SetForeColor i_OldColor

End Sub

Public Sub ClearScreen()

    ' Clear the screen for the next flip
    g_Backbuffer.SetFillColor 0
    g_Backbuffer.SetFillStyle 0
    g_Backbuffer.DrawBox 0, 0, i_ScreenW, i_ScreenH
    
End Sub

Public Function CreateSurfaceFromCompressed(ByVal s_FileName As String, ByVal i_Width As Long, ByVal i_Height As Integer, Optional b_Colorkey As Boolean = False) As DirectDrawSurface7

    On Local Error GoTo errOut
    
    Dim g_Ddsd As DDSURFACEDESC2, g_ColKey As DDCOLORKEY, g_PixF As DDPIXELFORMAT
    Dim b_SurfCreated As Boolean, g_SurfPic As Picture, b_JPEGLoaded As Boolean
    
    ' Load in the JEPG
    Set g_SurfPic = LoadPicture(s_FileName)
    SavePicture g_SurfPic, App.Path & "\temp.bmp"
    b_JPEGLoaded = True
    
    ' Set to use a standard surface
    g_Ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    g_Ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    ' Load the surface
    Set CreateSurfaceFromCompressed = g_DDraw.CreateSurfaceFromFile(App.Path & "\temp.bmp", g_Ddsd)

    ' Delete the temp picture and clean up the picture object
    Kill App.Path & "\temp.bmp"
    Set g_SurfPic = Nothing

    ' If a colorkey was requested set that up now
    If b_Colorkey Then
    
        ' Get the pixel format
        CreateSurfaceFromCompressed.GetPixelFormat g_PixF
        
        ' Set the color key to be magenta
        g_ColKey.high = g_PixF.lBBitMask + g_PixF.lRBitMask
        g_ColKey.low = g_ColKey.high
        
        ' Set the color key
        CreateSurfaceFromCompressed.SetColorKey DDCKEY_SRCBLT, g_ColKey
    
    End If
    
errOut:

    ' If there was an error log it and return false
    If Err.Number <> 0 Then
    
        Open App.Path & "\error.log" For Append As #1
        Print #1, "---------------------------------"
        Print #1, "CreateSurfaceFromJPEG (" & s_FileName & ")"
        Print #1, "JPEG Loaded: " & b_JPEGLoaded
        Print #1, "Surface created: " & b_SurfCreated
        Print #1, "Error " & Err.Number
        Print #1, Err.Description
        Close #1
        
    End If


End Function
