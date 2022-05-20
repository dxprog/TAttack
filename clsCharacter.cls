VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type CharRect
    t_Area As RECT
    i_X As Integer
    i_Y As Integer
    i_Angle As Integer
    i_RotateX As Integer
    i_RotateY As Integer
End Type

Private Type Command
    i_OpCode As Integer
    a_Parameters() As Integer
End Type

Private Type Frame
    i_Length As Integer
    i_NumInstructions As Integer
    a_Instructions() As Command
End Type

Private Type Section
    a_Frames() As Frame
    i_NumFrames As Integer
End Type

Private g_Surface As New clsSurface
Private a_Rects(10) As CharRect
Private a_Sections(6) As Section
Private i_State As CharacterState

Private i_Frame As Integer, i_Counter As Integer

Public Sub Initialize(ByRef g_Graphics As clsDirect3D, Optional b_ReadFile As Boolean = True)

    Dim i_Byte As Byte, i_CurrentSection As Integer, i_RectID As Integer, i_CurrentInstruction As Integer
    Dim i_Param As Integer

    ' Set the character's state to normal
    i_State = CS_NORMAL

    ' Load the surface
    g_Surface.CreateSurface g_Graphics.g_Device, App.Path & "\images\character.png", 256, 256
    
    If b_ReadFile = True Then
    
        ' Load the character
        Open App.Path & "\char.chr" For Binary As #1
        
        ' Loop through the file and start sorting things out
        Do While Not EOF(1)
        
            ' Get the command
            Get #1, , i_Byte
            
            ' See what we're dealing with
            Select Case i_Byte
            Case 1 ' Section label
                
                ' Get the section ID
                Get #1, , i_Byte
                i_CurrentSection = i_Byte
                
            Case 2 ' SetRect
        
                ' Get the rect ID
                Get #1, , i_Byte
                i_RectID = i_Byte
        
                ' Set the rect coordinates as they come in
                Get #1, , i_Byte ' SrcX
                a_Rects(i_RectID).t_Area.Left = i_Byte
                
                Get #1, , i_Byte ' SrcY
                a_Rects(i_RectID).t_Area.Top = i_Byte
        
                Get #1, , i_Byte ' Width
                a_Rects(i_RectID).t_Area.Right = i_Byte
                
                Get #1, , i_Byte ' Height
                a_Rects(i_RectID).t_Area.bottom = i_Byte
            
            Case 4
            
                ' Resize the frame array for this section
                a_Sections(i_CurrentSection).i_NumFrames = a_Sections(i_CurrentSection).i_NumFrames + 1
                ReDim Preserve a_Sections(i_CurrentSection).a_Frames(a_Sections(i_CurrentSection).i_NumFrames)
                
                ' Set the frame length and number of instructions
                With a_Sections(i_CurrentSection).a_Frames(a_Sections(i_CurrentSection).i_NumFrames)
                    Get #1, , i_Byte
                    .i_Length = i_Byte
                    Get #1, , i_Byte
                    .i_NumInstructions = i_Byte
                    
                    ' Resize the intructions array
                    ReDim .a_Instructions(.i_NumInstructions)
                    
                End With
                
                ' Reset the instruction counter
                i_CurrentInstruction = 0
            
            Case 5
            
                ' Increment the instruction counter
                i_CurrentInstruction = i_CurrentInstruction + 1
                
                ' Save the opcode
                With a_Sections(i_CurrentSection).a_Frames(a_Sections(i_CurrentSection).i_NumFrames)
                    
                    .a_Instructions(i_CurrentInstruction).i_OpCode = i_Byte
                    
                    ' Resize the parameter array
                    ReDim .a_Instructions(i_CurrentInstruction).a_Parameters(5)
                    
                    ' Get each parameter
                    For i = 1 To 5
                    
                        ' Get the byte
                        Get #1, , i_Byte
                        
                        ' If it's less than 255 then go ahead and save it
                        If i_Byte < 255 Then
                            .a_Instructions(i_CurrentInstruction).a_Parameters(i) = i_Byte
                        Else
                            
                            ' Start tallying up the final value
                            i_Param = 255
                            
                            ' Loop until we hit a number that's less than 255
                            Do Until i_Byte < 255
                                Get #1, , i_Byte
                                i_Param = i_Param + i_Byte
                            Loop
                            
                            ' Set the parameter
                            .a_Instructions(i_CurrentInstruction).a_Parameters(i) = i_Param
                            
                        End If
                    
                    Next i
                    
                End With
            Case 6
                ' Increment the instruction counter
                i_CurrentInstruction = i_CurrentInstruction + 1
                
                ' Save the opcode
                With a_Sections(i_CurrentSection).a_Frames(a_Sections(i_CurrentSection).i_NumFrames)
                    
                    .a_Instructions(i_CurrentInstruction).i_OpCode = i_Byte
                    
                    ' Resize the parameter array
                    ReDim .a_Instructions(i_CurrentInstruction).a_Parameters(5)
                    
                    ' Get each parameter
                    For i = 1 To 5
                    
                        ' Get the byte
                        Get #1, , i_Byte
                        
                        ' If it's less than 255 then go ahead and save it
                        If i_Byte < 255 Then
                            .a_Instructions(i_CurrentInstruction).a_Parameters(i) = i_Byte
                        Else
                            
                            ' Start tallying up the final value
                            i_Param = 255
                            
                            ' Loop until we hit a number that's less than 255
                            Do Until i_Byte < 255
                                Get #1, , i_Byte
                                i_Param = i_Param + i_Byte
                            Loop
                            
                            ' Set the parameter
                            .a_Instructions(i_CurrentInstruction).a_Parameters(i) = i_Param
                            
                        End If
                    
                    Next i
                    
                End With
            
            Case 8
            
                ' Increment the instruction counter
                i_CurrentInstruction = i_CurrentInstruction + 1
            
                ' Save the opcode
                With a_Sections(i_CurrentSection).a_Frames(a_Sections(i_CurrentSection).i_NumFrames)
                    
                    .a_Instructions(i_CurrentInstruction).i_OpCode = i_Byte
                    
                    ' Resize the parameter array
                    ReDim .a_Instructions(i_CurrentInstruction).a_Parameters(1)
                    
                    ' Get the rect number
                    Get #1, , i_Byte
                    .a_Instructions(i_CurrentInstruction).a_Parameters(1) = i_Byte
                    
                End With
            
            End Select
        
        Loop
    
        ' Close the file
        Close #1
        
    End If

End Sub

Public Sub TimeStep()

    Dim i_InterX As Single, i_InterY As Single, i_InterR As Single

    ' If a frame hasn't been set go to 1 by default
    If i_Frame = 0 Then i_Frame = 1

    ' Increment the frame counter
    i_Counter = i_Counter + 1
    
    ' See if the state needs changing
    If i_State = CS_NORMAL Then
        If i_NumCleared > 3 Then
            i_State = CS_COMBO
            i_Frame = 1
            i_Counter = 0
            g_Sound.PlaySound S_CHARACTER_CHEER
        End If
    
        If b_ChainFrame Then
            If i_Chain < 6 Then
                i_State = CS_CHAIN
            Else
                i_State = CS_BIGCHAIN
            End If
            i_Frame = 1
            i_Counter = 0
            g_Sound.PlaySound S_CHARACTER_CHEER
        End If
    End If
        
    
    ' If we've hit the length of this frame then move to the next one or
    ' start back at the beginning and change the state back to normal
    If i_Counter >= a_Sections(i_State).a_Frames(i_Frame).i_Length Then
        i_Counter = 0
        If a_Sections(i_State).i_NumFrames > i_Frame Then
            i_Frame = i_Frame + 1
        Else
            i_State = CS_NORMAL
            i_Frame = 1
        End If
    End If

    For i = 1 To a_Sections(i_State).a_Frames(i_Frame).i_NumInstructions
    
        With a_Sections(i_State).a_Frames(i_Frame)
                    
        
            Select Case .a_Instructions(i).i_OpCode
            Case 5
            
                ' Figure up the X and Y interpolation
                i_InterX = (.a_Instructions(i).a_Parameters(4) - .a_Instructions(i).a_Parameters(2)) / .i_Length
                i_InterY = (.a_Instructions(i).a_Parameters(5) - .a_Instructions(i).a_Parameters(3)) / .i_Length
                
                ' Set the X and Y coordinates
                a_Rects(.a_Instructions(i).a_Parameters(1)).i_X = .a_Instructions(i).a_Parameters(2) + (i_InterX * i_Counter)
                a_Rects(.a_Instructions(i).a_Parameters(1)).i_Y = .a_Instructions(i).a_Parameters(3) + (i_InterY * i_Counter)
            
            Case 6
            
                ' Figure up the angle interpolation
                i_InterR = (.a_Instructions(i).a_Parameters(3) - .a_Instructions(i).a_Parameters(2)) / .i_Length
                
                ' Set the rotation
                a_Rects(.a_Instructions(i).a_Parameters(1)).i_Angle = .a_Instructions(i).a_Parameters(2) + (i_InterR * i_Counter)
                a_Rects(.a_Instructions(i).a_Parameters(1)).i_RotateX = .a_Instructions(i).a_Parameters(4)
                a_Rects(.a_Instructions(i).a_Parameters(1)).i_RotateY = .a_Instructions(i).a_Parameters(5)
            
            Case 8
            
                ' Draw the rect in question
                With a_Rects(.a_Instructions(i).a_Parameters(1))
                    g_Surface.Blit .i_X, .i_Y, .t_Area.Right, .t_Area.bottom, .t_Area.Left, .t_Area.Top, , , , .i_Angle, .i_RotateX, .i_RotateY
                    
                    ' Now that the object has been drawn clear it's draw properties
                    .i_Angle = 0
                    .i_RotateX = 0
                    .i_RotateY = 0
                    .i_X = 0
                    .i_Y = 0
                    
                End With

            End Select
        
        End With
    
    Next i

End Sub

Public Sub ChangeState(ByVal i_NewState As CharacterState)

    If i_State = CS_NORMAL Then i_State = i_NewState

End Sub
