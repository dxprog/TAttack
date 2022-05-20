VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The surface
Private g_Surface As Direct3DTexture8

' The size of the surface image
Private i_SurfWidth As Integer, i_SurfHeight As Integer

' Geometry stuff
Private Const Pi = 3.14
Private Const Rad = Pi / 180

Private g_Sprite As D3DXSprite
Private g_D3dx As New D3DX8
Private s_File As String

Public Sub CreateSurface(ByRef g_Device As Direct3DDevice8, ByVal s_FileName As String, ByVal i_Width As Integer, ByVal i_Height As Integer)
    
    On Local Error GoTo errOut
    
    Dim g_D3dx As New D3DX8

    Set g_Sprite = g_D3dx.CreateSprite(g_Device)

    ' If the a color key was requested set that
    Set g_Surface = g_D3dx.CreateTextureFromFileEx(g_Device, s_FileName, i_Width, i_Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, ByVal 0, ByVal 0)
    
    ' Save the surface parameters
    i_SurfWidth = i_Width
    i_SurfHeight = i_Height

s_File = s_FileName
errOut:
    
    If Err.Number <> 0 Then _
        Debug.Print "Error loading " & s_FileName

End Sub

Public Sub Blit(ByVal i_DestX As Long, ByVal i_DestY As Long, ByVal i_DestWidth As Long, ByVal i_DestHeight As Long, ByVal i_SrcX As Long, ByVal i_SrcY As Long, Optional i_Alpha As Integer = 255, Optional i_SrcWidth As Integer, Optional i_SrcHeight As Integer, Optional i_Rotate As Integer, Optional i_RotateX As Integer = 0, Optional i_RotateY As Integer = 0)
 
    Dim t_Area As RECT, t_Translation As D3DVECTOR2, t_Rotate As D3DVECTOR2, t_Scaling As D3DVECTOR2
    
    ' If a source height and width weren't set use the destination dimensions
    If i_SrcHeight = 0 Or i_SrcWidth = 0 Then
        i_SrcHeight = i_DestHeight
        i_SrcWidth = i_DestWidth
    End If
    
    If i_DestWidth = 0 Or i_DestHeight = 0 Then Exit Sub
    
    ' Set the source coordinates
    With t_Area
        .Left = i_SrcX
        .Top = i_SrcY
        .Right = i_SrcX + i_SrcWidth
        .bottom = i_SrcY + i_SrcHeight
    End With
    
    ' Set the X and Y coordinates
    t_Translation.x = i_DestX
    t_Translation.y = i_DestY
    
    ' Set the scaling
    t_Scaling.x = i_DestWidth / i_SrcWidth
    t_Scaling.y = i_DestHeight / i_SrcHeight
        
    
    ' Set the rotation center
    t_Rotate.x = i_RotateX
    t_Rotate.y = i_RotateY
    
    ' Draw the sprite
    g_Sprite.Draw g_Surface, t_Area, t_Scaling, t_Rotate, i_Rotate * Rad, t_Translation, D3DColorRGBA(255, 255, 255, i_Alpha)

End Sub

Private Sub Class_Terminate()
    
    Set g_Surface = Nothing
    
End Sub
