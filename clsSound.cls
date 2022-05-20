VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The DSound objects
Private g_DSound As DirectSound8

' The sound buffer
Private g_Sounds(MAX_SOUNDS) As DirectSoundSecondaryBuffer8

Public Sub Initialize(ByRef g_Dx As DirectX8, ByVal hWnd As Long)

    ' Create the DSound object
    Set g_DSound = g_Dx.DirectSoundCreate("")

    ' Set the cooperative level
    g_DSound.SetCooperativeLevel hWnd, DSSCL_PRIORITY

End Sub

Public Sub LoadSound(ByVal s_FileName As String, ByVal i_Buffer As Integer)

    Dim g_Desc As DSBUFFERDESC
    
    ' Set the settings
    g_Desc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME

    ' Load the sound
    Set g_Sounds(i_Buffer) = g_DSound.CreateSoundBufferFromFile(s_FileName, g_Desc)

End Sub

Public Sub PlaySound(ByVal i_Buffer As Integer, Optional i_PlayMethod As CONST_DSBPLAYFLAGS)

    ' Reset the sound position
    g_Sounds(i_Buffer).SetCurrentPosition 0
    
    ' Play the sound
    g_Sounds(i_Buffer).Play i_PlayMethod
    
End Sub

Public Function GetPlayStatus(ByVal i_Buffer As Integer) As Integer

    GetPlayStatus = g_Sounds(i_Buffer).GetStatus

End Function

Public Sub StopSound(ByVal i_Buffer As Integer)

    g_Sounds(i_Buffer).Stop

End Sub

Private Sub Class_Terminate()
    
    ' Clear all the sound buffers
    For i = 1 To 100
        Set g_Sounds(i) = Nothing
    Next i
    
    ' Clear up the dsound object
    Set g_DSound = Nothing
    
End Sub
