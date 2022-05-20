VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  The surface
Private g_Surface As New clsSurface

Public Sub Initialize(ByRef g_Graphics As clsDirect3D)
    g_Surface.CreateSurface g_Graphics.g_Device, App.Path & "\images\font.png", 256, 256
End Sub

Public Sub DrawText(ByVal s_Text As String, ByVal i_X As Long, ByVal i_Y As Long)

    For i = 1 To Len(s_Text)
    
        s_Let = Mid(s_Text, i, 1)
        i_SrcY = Int((Asc(s_Let) - 33) / 15)
        i_SrcX = Int((Asc(s_Let) - 33) - (i_SrcY * 15))
        
        If s_Let <> " " Then _
            g_Surface.Blit i_X + ((i - 1) * 15), i_Y, 17, 30, i_SrcX * 17, i_SrcY * 30
        
    Next i

End Sub
