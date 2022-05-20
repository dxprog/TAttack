VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirect3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The Direct3D object and device
Private g_Direct3D As Direct3D8
Public g_Device As Direct3DDevice8
Public g_D3dx As New D3DX8

' Handle to the window
Private g_hWnd As Long

' The screen dimensions
Private i_ScreenWidth As Integer, i_ScreenHeight As Integer

' Flag telling us whether the scene has been started or not
Private b_StartScene As Boolean

' The starting display mode
Private g_Display As D3DDISPLAYMODE

Public Function Initialize(ByVal hWnd As Long, ByRef g_Dx As DirectX8, ByVal i_Width As Integer, ByVal i_Height As Integer, ByVal i_BPP As Integer, Optional b_Fullscreen As Boolean) As Boolean

    On Local Error GoTo errOut
    
    Dim l_RetVal As Long, g_DispMode As D3DPRESENT_PARAMETERS
    
    ' Save the window handle and screen dimensions
    g_hWnd = hWnd
    i_ScreenWidth = i_Width
    i_ScreenHeight = i_Height
    b_StartScene = False
    
    ' Create the Direct3D object
    Set g_Direct3D = g_Dx.Direct3DCreate()

    ' Get the current display mode
    g_Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, g_Display

    ' Ask the user if they want to run in fullscreen mode
    If b_Fullscreen = True Then

        ' Set up the display mode
        With g_DispMode
        
            ' Set the width, height, backbuffer count and window settings
            .BackBufferCount = 1
            .BackBufferWidth = i_Width
            .BackBufferHeight = i_Height
            .Windowed = False
        
            ' Set the bit depth
            Select Case i_BPP
            Case 16
                .BackBufferFormat = D3DFMT_R5G6B5
            Case 32
                .BackBufferFormat = D3DFMT_X8R8G8B8
            End Select
        
            i_DisplayFormat = g_DispMode.BackBufferFormat
        
            ' Set the window handle and swap effect
            .hDeviceWindow = g_hWnd
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        
        End With
    Else
        
        ' Set the properties
        g_DispMode.Windowed = 1
        g_DispMode.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        g_DispMode.BackBufferCount = 1
        g_DispMode.BackBufferFormat = g_Display.Format
        i_DisplayFormat = g_Display.Format
        
    End If

    ' Test the settings on the device with hardware acceleration
    If g_Direct3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, g_DispMode.BackBufferFormat, g_DispMode.BackBufferFormat, True) <> D3D_OK Then
    
        ' We'll have to use the software renderer
        Set g_Device = g_Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, g_hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, g_DispMode)
    
    Else
    
        ' Things check out so use the hardware accel features. We'll use software renderer since
        ' 2D isn't itensive enough to justify running hardware T&L checks.
        Set g_Device = g_Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, g_hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, g_DispMode)
    
    End If
    
    ' Turn off lighting and texture filtering
    g_Device.SetRenderState D3DRS_LIGHTING, False
    
    ' Return true
    Initialize = True

errOut:

    ' If there's been an error log it and bail
    If Err.Number <> 0 Then _
        Log "There was an error initializing Direct3D." & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Err Description: " & Err.Description

End Function

Private Sub Class_Terminate()

    ' Trash the Direct3D objects
    Set g_Device = Nothing
    Set g_Direct3D = Nothing

End Sub

Public Sub Clear()
    g_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, vbRed, 1#, 0
End Sub

Private Sub Log(ByVal s_Text As String)

    Open App.Path & "\debug.log" For Append As #1
    Print #1, vbCrLf & "-------------------------------------------"
    Print #1, Date & Time
    Print #1, s_Text

End Sub

Public Sub StartRender()
    
    ' Begin the scene
    g_Device.BeginScene
    b_StartScene = True
    
    ' Set the alpha blending stuff
    g_Device.SetRenderState D3DRS_ALPHABLENDENABLE, True
    g_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    g_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
End Sub

Public Sub Flip()

    ' If we haven't started a scene get ou before we hit an automation error
    If b_StartScene = False Then Exit Sub

    ' End the scene
    g_Device.EndScene
    b_StartScene = False

    ' Flip everything to the screen
    g_Device.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Sub RenderSurface(ByRef g_Surface As Direct3DSurface8)

On Local Error GoTo errOut

    Dim g_Backbuffer As Direct3DSurface8
    
    ' Get the backbuffer
    Set g_Backbuffer = g_Device.GetRenderTarget
    
    ' Blit the image
    g_Device.CopyRects g_Surface, ByVal 0, 0, g_Backbuffer, ByVal 0

errOut:
    If Err.Number <> 0 Then _
        Debug.Print Err.Number & ": " & Err.Description

End Sub

Public Sub ChangeVideoMode(ByVal i_Width As Integer, ByVal i_Height As Integer, ByVal i_BPP As Integer, ByVal b_Fullscreen As Boolean)

    Dim g_DispMode As D3DPRESENT_PARAMETERS

    ' Ask the user if they want to run in fullscreen mode
    If b_Fullscreen = True Then

        ' Set up the display mode
        With g_DispMode
        
            ' Set the width, height, backbuffer count and window settings
            .BackBufferCount = 1
            .BackBufferWidth = i_Width
            .BackBufferHeight = i_Height
            .Windowed = False
        
            ' Set the bit depth
            Select Case i_BPP
            Case 16
                .BackBufferFormat = D3DFMT_R5G6B5
            Case 32
                .BackBufferFormat = D3DFMT_X8R8G8B8
            End Select
        
            i_DisplayFormat = g_DispMode.BackBufferFormat
        
            ' Set the window handle and swap effect
            .hDeviceWindow = g_hWnd
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        
        End With
    Else
        
        ' Set the properties
        g_DispMode.Windowed = 1
        g_DispMode.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        g_DispMode.BackBufferCount = 1
        g_DispMode.BackBufferFormat = g_Display.Format
        i_DisplayFormat = g_Display.Format
        
    End If
    
    ' Reset the device
    g_Device.Reset g_DispMode
    b_StartScene = False

End Sub
