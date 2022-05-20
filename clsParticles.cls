VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The particle type
Private Type Particle
    i_VelX As Integer
    i_VelY As Integer
    i_X As Integer
    i_Y As Integer
    i_MaxLife As Integer
    i_Life As Integer
End Type

' The particle array
Private a_Particles(100) As Particle

Public Sub CreateParticles(ByVal i_NumParticles As Integer, ByVal i_Life As Integer, ByVal i_X As Integer, ByVal i_Y As Integer)

    Dim i_Index As Integer

    ' Loop through and create the particles
    For i = 1 To i_NumParticles
    
        ' Get a free spot in the array
        i_Index = GetFreeIndex
        
        ' Set the X and Y velocities
        Select Case i
        Case 1
            a_Particles(i_Index).i_VelX = -2
            a_Particles(i_Index).i_VelY = -2
        Case 2
            a_Particles(i_Index).i_VelX = 2
            a_Particles(i_Index).i_VelY = -2
        Case 3
            a_Particles(i_Index).i_VelX = -2
            a_Particles(i_Index).i_VelY = 2
        Case 4
            a_Particles(i_Index).i_VelX = 2
            a_Particles(i_Index).i_VelY = 2
        End Select
        
        ' Set the particle's life and X, Y coords
        a_Particles(i_Index).i_Life = i_Life
        a_Particles(i_Index).i_MaxLife = i_Life
        a_Particles(i_Index).i_X = i_X
        a_Particles(i_Index).i_Y = i_Y
    
    Next i

End Sub

Public Sub HandleParticles()

    For i = 1 To 100
    
        ' Before doing anything make sure this is a live particle
        If a_Particles(i).i_Life > 0 Then
        
            ' Decrement the particle's life
            a_Particles(i).i_Life = a_Particles(i).i_Life - 1
        
            ' Adjust the X and Y values
            a_Particles(i).i_X = a_Particles(i).i_X + a_Particles(i).i_VelX
            a_Particles(i).i_Y = a_Particles(i).i_Y + a_Particles(i).i_VelY
            
            ' Draw the particle
            g_MainImg.Blit a_Particles(i).i_X, a_Particles(i).i_Y, 30, 30, 186, 55, Int((a_Particles(i).i_Life / a_Particles(i).i_MaxLife) * 255), 30, 30, a_Particles(i).i_Life * (a_Particles(i).i_VelX * 4), 15, 15
            
        End If
    
    Next i

End Sub

Private Function GetFreeIndex() As Integer

    For i = 1 To 100
        If a_Particles(i).i_Life = 0 Then
            GetFreeIndex = i
            Exit Function
        End If
    Next i

End Function
