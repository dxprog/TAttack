VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The block type
Private Type Block
    
    ' Block type
    i_Type As Byte
    
    ' Block state
    i_State As BLOCK_STATE
    
    ' Counter. Used to keep track of block states
    i_Counter As Integer
    
    ' Used to keep track of which block is first in the swap chain
    b_First As Boolean
    
    ' Used to keep track of which block is last in the popping chain
    b_Last As Boolean
    
    ' The clear box
    t_ClearArea As RECT
    
    ' The block that started a clear
    i_ClearStart As Integer
    
    ' Set to true if the block is falling due to a block underneath clearing
    ' Used to figure up chains
    b_ClearFall As Boolean
    
    ' Used to link it with other blocks that cleared when it did
    i_ClearID As Long
    
    ' Set to true if the block has already been handled
    b_Handled As Boolean
    
    i_OldState As BLOCK_STATE
    
    ' Tells the main routine to check this block for a clear
    b_CheckClear As Boolean
    
    ' Where in the line of popping blocks this falls
    i_PopIndex As Integer
    
End Type

' The array of blocks (the size of the playing field plus an extra row for the next line)
Private a_Blocks(FIELD_WIDTH * FIELD_HEIGHT + FIELD_WIDTH) As Block

Public Sub NewGame()
    
    'On Error GoTo errOut
    
    Dim x As Integer, y As Integer, i_Index As Integer, a_UsedBlocks(NUM_BLOCKS) As Integer
    Dim i_Type As Integer, i_Last As Integer, t_Blank As Block
    
    ' We'll randomly generate 30 blocks and then apply gravity on them to achieve
    ' the effect in Tetris Attack
    
    ' Fire up the random number generator
    Randomize Timer
    
    ' Blank out the block array
    For i = 1 To (FIELD_WIDTH * FIELD_HEIGHT)
        a_Blocks(i) = t_Blank
    Next i
    
    ' First thing we'll do is set the block positions
    
    ' Now we go through and start adding blocks
    For i = 1 To 30
    
        ' Get ourselves some coordinates
        x = Int(Rnd * FIELD_WIDTH) + 1
        y = Int(Rnd * FIELD_HEIGHT) + 1
        i_Index = ((y - 1) * FIELD_WIDTH) + x
    
        ' If we hit a block that's already been used loop until we've got a blank
        ' spot
        Do While a_Blocks(i_Index).i_Type <> 0
            ' Get ourselves some coordinates
            x = Int(Rnd * FIELD_WIDTH) + 1
            y = Int(Rnd * FIELD_HEIGHT) + 1
            i_Index = ((y - 1) * FIELD_WIDTH) + x
        Loop
    
        ' Set the block
        a_Blocks(i_Index).i_Type = 1
    
    Next i

    ' Now apply instant gravity
    InstantGravity

    ' Now that we have a map of where the blocks will be we'll go in and set the
    ' types. We need to make sure that no block is repeated more than two times
    ' in a row and that there aren't two blocks in a row vertically
    For i = 1 To FIELD_HEIGHT
    
        For j = 1 To FIELD_WIDTH
        
            ' Get the array index
            i_Index = ((i - 1) * FIELD_WIDTH) + j
        
            ' If this block is set move on to the next phase
            If a_Blocks(i_Index).i_Type > 0 Then
            
                ' Generate a block type
                a_Blocks(i_Index).i_Type = Int(Rnd * NUM_BLOCKS) + 1
                
                ' Set the current block type
                a_Blocks(i_Index).i_State = BS_NORMAL
                    
                ' Loop until we have a block type that is different
                Do While a_Blocks(i_Index).i_Type = i_Last Or a_Blocks(i_Index).i_Type = a_Blocks(i_Index - FIELD_WIDTH).i_Type
                    a_Blocks(i_Index).i_Type = Int(Rnd * NUM_BLOCKS) + 1
                Loop
            
            End If
        
            ' Set the type that was used so it's not repeated next time 'round
            i_Last = a_Blocks(i_Index).i_Type
        
        Next j
    
    Next i

'Open "C:\level.puz" For Input As #1
'For i = 1 To 72
'    Line Input #1, s_Temp
'    If CInt(s_Temp) = 0 Then
'        a_Blocks(i).i_State = BS_EMPTY
'    Else
'        a_Blocks(i).i_State = BS_NORMAL
'        a_Blocks(i).i_Type = CInt(s_Temp)
'    End If
'Next i
'Close #1
    ' Generate a new line
    NewLine

errOut:

    If Err.Number <> 0 Then _
        Debug.Print Err.Description

End Sub

Public Sub DrawBlocks(Optional b_Jiggle As Boolean, Optional i_Alpha As Integer = 255)
    
    Dim x As Integer, y As Integer, i_Index As Integer, i_OffsetY As Integer
    
    ' Run through the blocks and display them
    For i = 1 To FIELD_HEIGHT + 1
    
        For j = 1 To FIELD_WIDTH
        
            ' Rig up our array index value
            i_Index = ((i - 1) * FIELD_WIDTH) + j
            
            ' If this block isn't set don't display it
            If a_Blocks(i_Index).i_Type <> 0 Then
                
                ' Figure up the x and y coordinates (center the blocks on the screen)
                x = ((SCREEN_WIDTH - (FIELD_WIDTH * BLOCK_WIDTH)) / 2) + ((j - 1) * BLOCK_WIDTH)
                y = ((SCREEN_HEIGHT - (FIELD_HEIGHT * BLOCK_HEIGHT)) / 2) + ((i - 1) * BLOCK_HEIGHT)
                
                ' If jiggle is on add a random amount to the X and Y values
                If b_Jiggle = True Then
                    x = x + (Int(Rnd * 4) - 2)
                    y = y + (Int(Rnd * 4) - 2)
                End If
                
                ' Check to see if this is comming from the new line of blocks
                If i <= FIELD_HEIGHT Then

                    ' Blit the block according to state
                    Select Case Int(a_Blocks(i_Index).i_State)
                    Case BS_SWAPPING
                        g_MainImg.Blit x + a_Blocks(i_Index).i_Counter, y - i_StackHeight, BLOCK_WIDTH, BLOCK_HEIGHT, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, 0, i_Alpha
                    Case BS_FALLING
                        g_MainImg.Blit x, y - i_StackHeight, BLOCK_WIDTH, BLOCK_HEIGHT, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, 0, i_Alpha
                    Case BS_PREFALL
                        g_MainImg.Blit x, y - i_StackHeight, BLOCK_WIDTH, BLOCK_HEIGHT, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, 0, i_Alpha
                    Case BS_NORMAL
                        If a_Blocks(i_Index).i_Counter > 2 Then
                            g_MainImg.Blit x, y - i_StackHeight + 2, BLOCK_WIDTH, BLOCK_HEIGHT - 2, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, 0, i_Alpha, BLOCK_WIDTH, BLOCK_HEIGHT
                        ElseIf a_Blocks(i_Index).i_Counter > 0 Then
                            g_MainImg.Blit x, y - i_StackHeight - 2, BLOCK_WIDTH, BLOCK_HEIGHT + 2, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, 0, i_Alpha, BLOCK_WIDTH, BLOCK_HEIGHT
                        Else
                            g_MainImg.Blit x, y - i_StackHeight, BLOCK_WIDTH, BLOCK_HEIGHT, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, 0, i_Alpha
                        End If
                    Case BS_FLASHING
                        i_OffsetY = a_Blocks(i_Index).i_Counter Mod 2
                        g_MainImg.Blit x, y - i_StackHeight, BLOCK_WIDTH, BLOCK_HEIGHT, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, BLOCK_HEIGHT, 255 - (128 * i_OffsetY)
                    Case BS_CLEARING
                        If a_Blocks(i_Index).i_Counter > 0 Then _
                            g_MainImg.Blit x, y - i_StackHeight, BLOCK_WIDTH, BLOCK_HEIGHT, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, BLOCK_HEIGHT, i_Alpha
                    End Select
                Else
                    ' Blit the block and shave some off the bottom so it doesn't blit
                    ' over the border
                    g_MainImg.Blit x, y - i_StackHeight, BLOCK_WIDTH, i_StackHeight, (a_Blocks(i_Index).i_Type - 1) * BLOCK_WIDTH, BLOCK_HEIGHT * 2
                End If
                
            End If
        
        Next j
    
    Next i

errOut:

End Sub

Private Sub InstantGravity()

    On Error GoTo errOut

    Dim i_Index As Integer, i_Next As Integer, y As Integer, t_Blank As Block
    
    ' We need to loop from the bottom up for this to work properly
    For i = FIELD_HEIGHT To 1 Step -1
    
        For j = 1 To FIELD_WIDTH
        
            ' Create an array index from the x, y coordinates
            i_Index = ((i - 1) * FIELD_WIDTH) + j
            
            ' Check to see if there is a block at this position
            If a_Blocks(i_Index).i_Type > 0 Then
                
                ' Set the index
                i_Next = i_Index
                
                ' Now loop until we either hit the bottom or another block
                Do While i_Next < ((FIELD_HEIGHT * FIELD_WIDTH) - FIELD_WIDTH) + 1
                    
                    ' Go to the block immediately under this one
                    i_Next = i_Next + FIELD_WIDTH
                    
                    ' If there's a block here decrement the y value and jump ship
                    If a_Blocks(i_Next).i_Type > 0 Then
                        i_Next = i_Next - FIELD_WIDTH
                        Exit Do
                    End If
                    
                Loop
                
                ' Move the block if necessary
                If i_Next <> i_Index Then
                    a_Blocks(i_Next) = a_Blocks(i_Index)
                    a_Blocks(i_Index) = t_Blank
                End If
                
            End If
        
        Next j
    
    Next i

errOut:
    If Err.Number <> 0 Then _
        Debug.Print Err.Description & ": " & i_Index & ", " & i_Next

End Sub

Public Function HandleBlocks() As Integer

    On Error GoTo errOut
    
    Dim a_NewLine As Variant, b_Clear As Boolean, i_Cleared As Integer
    Dim i_Temp As Integer, b_WaitForChain As Boolean, i_OffsetY As Integer, i_First As Integer
    Dim a_Fell(FIELD_WIDTH * FIELD_HEIGHT) As Boolean, i_NumChains As Integer, b_Pause As Boolean
    
    ' Set the first block to 72 that way we don't have weird problems if there are no blocks on the screen
    i_First = 72
    
    ' Now we go through the blocks (from the bottom up) and update them according to status
    For i = FIELD_HEIGHT * FIELD_WIDTH To 1 Step -1
    
        ' If this is the first block in the stack save it's value so we can figure
        ' up whether the player has lost later
        If a_Blocks(i).i_State <> BS_EMPTY Then _
            i_First = i
    
        ' If this block hasn't been handled yet handle it
        If a_Blocks(i).b_Handled = False Then
            ' See what the block is doing
            Select Case a_Blocks(i).i_State
            Case BS_NORMAL
                
                If a_Blocks(i).i_Counter > 0 Then _
                    a_Blocks(i).i_Counter = a_Blocks(i).i_Counter - 1
                    
            Case BS_PREFALL
                BsPrefall i
            Case BS_FALLING
                BsFalling i
            Case BS_SWAPPING
                BsSwapping i
            Case BS_FLASHING
                BsFlashing i
            Case BS_CLEARING
                BsClearing i
            End Select
        End If
        
        ' Set the block handling flag back to false
        a_Blocks(i).b_Handled = False
        
    Next i
    
    ' Now that the states of all the blocks have been updated test each block for falling
    ' and clearing
    For i = (FIELD_WIDTH * FIELD_HEIGHT) To 1 Step -1

        If a_Blocks(i).b_ClearFall = True Then b_WaitForChain = True

        If a_Blocks(i).i_State = BS_NORMAL Then
            
            If CheckFall(i) = False Then
                
                ' Only check this block if we need to
                If a_Blocks(i).b_CheckClear Then
                    i_Cleared = CheckClear(i)
                    
                    If i_Cleared < 3 Then
                        a_Blocks(i).b_ClearFall = False
                    Else
                        i_NumCleared = i_NumCleared + i_Cleared
                    End If
                    a_Blocks(i).b_CheckClear = False
                End If
                
            End If
        End If

        ' If the block's state is anything but empty or normal stop the stack advancement
        If a_Blocks(i).i_State <> BS_EMPTY And a_Blocks(i).i_State <> BS_NORMAL Then b_Pause = True

    Next i
    
    ' If there aren't any blocks falling because of clears clear the chain value
    If b_WaitForChain = False Then _
        i_Chain = 0
    
    ' If no blocks are clearing, falling, swapping, etc. move the stack up
    If b_Pause = False And i_StopTime = 0 Then
        ' If we're advancing, make the stack rise faster
        If b_Advance = True Then
            i_StackHeight = i_StackHeight + 2
        Else
            i_StackCounter = i_StackCounter + 1
            If i_StackCounter >= i_StackSpeed Then
                i_StackHeight = i_StackHeight + 1
                i_StackCounter = 0
            End If
        End If
    ElseIf b_Pause = True Or i_StopTime <> 0 Then
        b_Advance = False
    End If
    
    ' Get the actual Y value of the first block
    i_First = Int((i_First - 1) / FIELD_WIDTH)
    
    ' Handle the music stuffs
    If i_First <= 2 Then
    
        ' Play the panic music
        g_Music.StartMusic S_MUSIC_PANIC_INTRO
        
    Else
    
        ' Play the normal music
        g_Music.StartMusic S_MUSIC_NORMAL_INTRO
    
    End If
    
    ' If we've advanced one block, shift all the blocks up in the array and add the
    ' new line
    If i_StackHeight >= BLOCK_HEIGHT Then
    
        ' If we're over the line end the game
        If i_First = 1 Then
        
            ' Change the game state
            g_Game.i_GameState = GS_GAMEOVER
            Exit Function
        
        End If
    
        ' If the stack was advanced on purpose add one to the score
        If b_Advance Then
            g_Score.i_Score = g_Score.i_Score + 1
            b_Advance = False
        End If
    
        ' Reset the stack height and speed
        i_StackHeight = 0
        i_StackSpeed = 60 - (i_Speed * 1.5)
        i_StackCounter = 0
        b_Advance = False
    
        b_NewLine = True
    
        ' Bump everything up
        For i = 1 To FIELD_HEIGHT * FIELD_WIDTH
            
            a_Blocks(i) = a_Blocks(i + FIELD_WIDTH)
        
            ' If this is the last line have a clear check flagged for next frame
            If i >= (FIELD_HEIGHT * FIELD_WIDTH) - FIELD_WIDTH Then _
                a_Blocks(i).b_CheckClear = True
        
        Next i
    
        ' Get our new line
        NewLine
    
    End If
    
    ' Lastly, if we had a chain or a combo this round create cards for each
    If b_ChainFrame Or i_NumCleared > 3 Then
        
        ' Check to see if we're at critical level
        If i_First > 2 Then
            
            ' If there was a combo give the player 2 seconds of stop time
            If i_NumCleared > 3 Then
                i_StopTime = i_StopTime + 2000
            End If
            
            ' If there was a chain give the user 4 seconds
            If b_ChainFrame Then
                i_StopTime = i_StopTime + 4000
            End If
        
            ' Set the time that our stop counter started
            i_StopStart = i_LastTick
        
        Else
        
            ' Give the player ten seconds of stop time
            i_StopTime = i_StopTime + 10000
            i_StopStart = i_LastTick
        
        End If
        
        ' Clamp the stop time down to 20 seconds
        If i_StopTime > 20000 Then i_StopTime = 20000
        
        ' Loop through the stack and find the first panel with a clear ID for this frame
        For i = 1 To 72
            If a_Blocks(i).i_ClearID = i_LastTick Then
                If b_ChainFrame = True Then
                    g_Cards.CreateChainCard i_Chain, i
                    g_Score.ChainScore i_Chain
                    i_OffsetY = 1
                End If
                
                If i_NumCleared > 3 Then
                    g_Cards.CreateComboCard i_NumCleared, i, i_OffsetY
                    g_Score.ComboScore i_NumCleared
                End If
                
                Exit For
            End If
        Next i
    ElseIf i_NumCleared = 3 And i_Chain > 0 Then
        g_Score.ChainScore i_Chain
    End If
    
    ' Return the stack height
    HandleBlocks = i_First
    
errOut:
    If Err.Number <> 0 Then _
        Debug.Print i
End Function

Public Sub Swap(ByVal i_X As Long, ByVal i_Y As Long)

    Dim t_Block As Block, i_Index As Integer
    
    ' Get the index value of the first block
    i_Index = ((i_Y - 1) * FIELD_WIDTH) + i_X
    
    ' Make sure the two blocks being swapped are standing still
    If a_Blocks(i_Index).i_State = BS_NORMAL Or a_Blocks(i_Index).i_State = BS_EMPTY Or a_Blocks(i_Index).i_State = BS_FALLING Then
        If a_Blocks(i_Index + 1).i_State = BS_NORMAL Or a_Blocks(i_Index + 1).i_State = BS_EMPTY Or a_Blocks(i_Index + 1).i_State = BS_FALLING Then
    
            ' Check the blocks above to make sure there isn't anything trying to fall
            If i_Index - 6 > 0 Then _
                If a_Blocks(i_Index - FIELD_WIDTH).i_State = BS_PREFALL Or a_Blocks(i_Index - (FIELD_WIDTH - 1)).i_State = BS_PREFALL Or a_Blocks(i_Index - FIELD_WIDTH).i_State = BS_FALLING Or a_Blocks(i_Index - (FIELD_WIDTH - 1)).i_State = BS_FALLING Then Exit Sub
            
            ' If both blocks are empties don't do anything
            If a_Blocks(i_Index).i_State = BS_EMPTY And a_Blocks(i_Index + 1).i_State = BS_EMPTY Then Exit Sub
    
            ' Set the swap states to true and set which was first (farthest to the left)
            a_Blocks(i_Index).i_OldState = a_Blocks(i_Index).i_State
            a_Blocks(i_Index + 1).i_OldState = a_Blocks(i_Index + 1).i_State
            a_Blocks(i_Index).i_State = BS_SWAPPING
            a_Blocks(i_Index).b_First = True
            a_Blocks(i_Index + 1).i_State = BS_SWAPPING
        
            ' Play the swap sound
            g_Sound.PlaySound S_SWAP_BLOCK
            b_Swap = True
        
        End If
    End If
    

End Sub

Private Function CheckClear(ByVal i_Index As Integer, Optional b_CheckOnce As Boolean = False) As Integer

    On Local Error Resume Next

    ' Used for the clear box (explained below)
    Dim t_Clear As RECT, i_Width As Integer, i_Height As Integer, i_ExtraClear As Integer
    Dim i_TempX As Integer, i_TempY As Integer, i_First As Integer, i_Reserved As Integer
    Dim i_X As Integer, i_Y As Integer, i_Original As Integer, a_Temp(FIELD_WIDTH * FIELD_HEIGHT) As Integer
    
    ' Make sure there is a block here before continuing
    If a_Blocks(i_Index).i_State = BS_EMPTY Then Exit Function
    
    ' Get the X and Y coordinates of the block
    i_First = i_Index
    i_Y = Int((i_Index - 1) / FIELD_WIDTH)
    i_X = ((i_Index - 1) - (i_Y * FIELD_WIDTH)) + 1
    
    ' Save the type of the block
    i_Original = a_Blocks(i_Index).i_Type
    
    ' A little explination of how this is going to be done. First we need to get the
    ' height and width if the "clear box" as it will be called. Example:
    '   12345
    '   -----
    ' 1|00010
    ' 2|01111
    ' 3|00010
    '
    ' The box would start at coordinates 2, 1 (top-left) and go to 5, 3 (bottom-right)
    ' We then loop through starting at the top and give each block a countdown to
    ' pop. We'll label the last block so that when it pops the type for all these blocks
    ' are set to 0 and any other blocks resting on them can fall into place
    
    ' Set the initial search coordinates
    With t_Clear
        .Left = -1
        .Right = 1
        .Top = -1
        .bottom = 1
    End With
    
    ' Search left of the block
    Do While (i_X + t_Clear.Left) > 0
    
        ' If we've hit a block of a different type or the block isn't sitting still exit the loop
        If a_Blocks(i_Index + t_Clear.Left).i_Type <> i_Original Or a_Blocks(i_Index + t_Clear.Left).i_State <> BS_NORMAL Then Exit Do
        
        ' Set this block's state to reserved and dump it's index value into our array
        a_Blocks(i_Index + t_Clear.Left).i_State = BS_RESERVED
        i_Reserved = i_Reserved + 1
        a_Temp(i_Reserved) = i_Index + t_Clear.Left
        
        ' Decrement the left side
        t_Clear.Left = t_Clear.Left - 1
    
    Loop

    ' Because of the repeatition from here down I will not comment the next three
    ' loops. If you need comments see the loop above. It's pretty much the same
    ' stuff
    Do While (i_X + t_Clear.Right) < (FIELD_WIDTH + 1)
        If a_Blocks(i_Index + t_Clear.Right).i_Type <> i_Original Or a_Blocks(i_Index + t_Clear.Right).i_State <> BS_NORMAL Then Exit Do
        a_Blocks(i_Index + t_Clear.Right).i_State = BS_RESERVED
        i_Reserved = i_Reserved + 1
        a_Temp(i_Reserved) = i_Index + t_Clear.Right
        t_Clear.Right = t_Clear.Right + 1
    Loop

    Do While (i_Y + t_Clear.Top) > 0
        If a_Blocks(i_Index + (t_Clear.Top * FIELD_WIDTH)).i_Type <> i_Original Or a_Blocks(i_Index + (t_Clear.Top * FIELD_WIDTH)).i_State <> BS_NORMAL Then Exit Do
        a_Blocks(i_Index + (t_Clear.Top * FIELD_WIDTH)).i_State = BS_RESERVED
        i_Reserved = i_Reserved + 1
        a_Temp(i_Reserved) = i_Index + (t_Clear.Top * FIELD_WIDTH)
        t_Clear.Top = t_Clear.Top - 1
    Loop

    Do While (i_Y + t_Clear.bottom) < FIELD_HEIGHT
        If a_Blocks(i_Index + (t_Clear.bottom * FIELD_WIDTH)).i_Type <> i_Original Or a_Blocks(i_Index + (t_Clear.bottom * FIELD_WIDTH)).i_State <> BS_NORMAL Then Exit Do
        a_Blocks(i_Index + (t_Clear.bottom * FIELD_WIDTH)).i_State = BS_RESERVED
        i_Reserved = i_Reserved + 1
        a_Temp(i_Reserved) = i_Index + (t_Clear.bottom * FIELD_WIDTH)
        t_Clear.bottom = t_Clear.bottom + 1
    Loop

    ' We need to adjust our values just a smidge to avoid bogus numbers
    With t_Clear
        .bottom = .bottom - 1
        .Top = .Top + 1
        .Left = .Left + 1
        .Right = .Right - 1
    End With

    ' Now that we have our box we'll figure up the height and width
    i_Width = (Abs(t_Clear.Left) + t_Clear.Right) + 1
    i_Height = (Abs(t_Clear.Top) + t_Clear.bottom) + 1
    
    ' If the height or width is less than 3 clamp them down to 0 so weird things
    ' don't happen
    If i_Width < 3 Then
        i_Width = 1
        t_Clear.Left = 0
        t_Clear.Right = 0
    End If
    
    If i_Height < 3 Then
        i_Height = 1
        t_Clear.Top = 0
        t_Clear.bottom = 0
    End If
    
    ' Before we go clearing blocks make sure we had at least three in a row
    If ((i_Width + i_Height) - 1) >= 3 Then
    
        ' Now loop through the box (top to bottom, left to right) and set the clear
        ' times
        For i = t_Clear.bottom To t_Clear.Top Step -1
            
            For j = t_Clear.Left To t_Clear.Right
            
                ' Figure up the array index and coordinates
                i_TempX = i_X + j
                i_TempY = i_Y + i
                i_Index = (i_TempY * FIELD_WIDTH) + i_TempX
                
                ' Make sure we're in line before continuing
                If ((i_TempY = i_Y And i_Width >= 3) Or (i_TempX = i_X And i_Height >= 3)) And a_Blocks(i_Index).i_Type = i_Original Then
                    
                    ' Set the block's state to flashing and also set it's clear ID
                    a_Blocks(i_Index).i_State = BS_FLASHING
                    a_Blocks(i_Index).i_ClearID = i_LastTick
                    a_Blocks(i_Index).i_Counter = 0
                    a_Blocks(i_Index).b_Handled = True
                    
                    ' If this isn't the original block and we're supposed to be checking
                    ' for sub clears, check it for a clear
'                    If i_Index <> i_First And b_CheckOnce = False Then
'                        i_Cleared = CheckClear(i_Index, True)
'                        If i_Cleared > 2 Then _
'                            i_ExtraClear = i_ExtraClear + i_Cleared - 1
'                    End If
                    
                    ' If we haven't already incremented the chain counter this frame
                    ' and this block would cause a chain, increment the counter
                    If a_Blocks(i_Index).b_ClearFall = True And b_ChainFrame = False Then
                        i_Chain = i_Chain + 1
                        b_ChainFrame = True
                        i_ChainID = i_LastTick
                    End If
                    
                End If
            
            Next j
            
        Next i
        
    End If
    
    ' Loop through all the blocks we screwed with and set the reserved ones back to normal
    For i = 1 To i_Reserved
        If a_Blocks(a_Temp(i)).i_State = BS_RESERVED Then _
            a_Blocks(a_Temp(i)).i_State = BS_NORMAL
    Next i
    
    ' Return the amount of blocks cleared
    CheckClear = (i_Width + i_Height) - 1 + i_ExtraClear

End Function

Private Sub NewLine()

    Dim i_Last As Integer, i_Index As Integer

    For i = 1 To FIELD_WIDTH
    
        ' Generate a random block
        i_Index = (FIELD_WIDTH * FIELD_HEIGHT) + i
        a_Blocks(i_Index).i_Type = Int(Rnd * NUM_BLOCKS) + 1
    
        ' Loop until we have a block type that wasn't the last one
        Do While a_Blocks(i_Index).i_Type = i_Last Or a_Blocks(i_Index).i_Type = a_Blocks(i_Index - FIELD_WIDTH).i_Type
            a_Blocks(i_Index).i_Type = Int(Rnd * NUM_BLOCKS) + 1
        Loop
        
        a_Blocks(i_Index).i_State = BS_NORMAL
        i_Last = a_Blocks(i_Index).i_Type
    
    Next i

End Sub

Private Function CheckFall(ByVal i_Index As Integer) As Boolean

    ' Make sure that this block isn't empty
    If a_Blocks(i_Index).i_State = BS_EMPTY Then Exit Function
    
    ' Make sure we aren't going out of bounds by checking below
    If (i_Index + 6) > (FIELD_WIDTH * FIELD_HEIGHT) Then Exit Function
    
    ' Now that we're done with saftey checks take a peek below and see
    ' if there is a block there. If there isn't or the block is in prefall
    ' state set this block to prefall
    If a_Blocks(i_Index + FIELD_WIDTH).i_State = BS_EMPTY Or a_Blocks(i_Index + 6).i_State = BS_PREFALL Then
        
        ' Set the block's state to prefall
        a_Blocks(i_Index).i_State = BS_PREFALL
        
        ' If the block below is in prefall state set this block's clear/fall state
        ' to that of the one below
        If a_Blocks(i_Index + FIELD_WIDTH).i_State = BS_PREFALL And a_Blocks(i_Index).b_ClearFall = False Then _
            a_Blocks(i_Index).b_ClearFall = a_Blocks(i_Index + FIELD_WIDTH).b_ClearFall
        
        ' Return that the block is falling
        CheckFall = True
        
    End If

End Function

' Block state handlers
' ---------------------------------------------------------
Private Sub BsPrefall(ByVal i_Index As Integer)

    ' Increment the block's counter
    a_Blocks(i_Index).i_Counter = a_Blocks(i_Index).i_Counter + 1

    ' See if we've hit the prefall length
    If a_Blocks(i_Index).i_Counter >= i_PrefallLength Then
        ' Set the block to falling state and reset the counter
        a_Blocks(i_Index).i_State = BS_FALLING
        a_Blocks(i_Index).i_Counter = 0
    End If

End Sub

Private Sub BsFalling(ByVal i_Index As Integer)

    Dim t_Blank As Block

    ' Make sure we haven't hit the bottom
    If (i_Index + FIELD_WIDTH) <= (FIELD_WIDTH * FIELD_HEIGHT) Then

        ' If there's nothing underneath this block move it down
        If a_Blocks(i_Index + FIELD_WIDTH).i_State = BS_EMPTY Then
            ' Move the block down
            a_Blocks(i_Index + FIELD_WIDTH) = a_Blocks(i_Index)
            
            ' Blank out the previous spot
            a_Blocks(i_Index) = t_Blank
        ElseIf a_Blocks(i_Index + FIELD_WIDTH).i_State = BS_FALLING Or a_Blocks(i_Index + FIELD_WIDTH).i_State = BS_PREFALL Or a_Blocks(i_Index + FIELD_WIDTH).i_State = BS_SWAPPING Then
            ' If the block below is falling, in a prefall state or swapping just wait
            Exit Sub
        Else
                
            ' Set the block state back to normal
            a_Blocks(i_Index).i_State = BS_NORMAL
            a_Blocks(i_Index).i_Counter = 4
            a_Blocks(i_Index).b_CheckClear = True
            g_Sound.PlaySound S_BLOCK_FALL
            
            
        End If
    Else
        
        ' Set the block's state to normal
        a_Blocks(i_Index).i_State = BS_NORMAL
        a_Blocks(i_Index).i_Counter = 4
        a_Blocks(i_Index).b_CheckClear = True
        g_Sound.PlaySound S_BLOCK_FALL
        
    End If

End Sub

Private Sub BsSwapping(ByVal i_Index As Integer)

    Dim t_Temp As Block

    ' If this block is first in the swapp chain we add to its counter
    If a_Blocks(i_Index).b_First = True Then
        a_Blocks(i_Index).i_Counter = a_Blocks(i_Index).i_Counter + (BLOCK_WIDTH / 4)
    Else
        a_Blocks(i_Index).i_Counter = a_Blocks(i_Index).i_Counter - (BLOCK_WIDTH / 4)
    End If

    ' If the last block is finished swapping we'll swap the actual blocks
    If a_Blocks(i_Index).b_First = False And Abs(a_Blocks(i_Index).i_Counter) >= BLOCK_WIDTH Then
    
        ' Swap the blocks
        t_Temp = a_Blocks(i_Index)
        a_Blocks(i_Index) = a_Blocks(i_Index - 1)
        a_Blocks(i_Index - 1) = t_Temp
        
        ' Clear the first flag, set the states back to normal and reset the counters
        a_Blocks(i_Index).b_First = False
        a_Blocks(i_Index - 1).i_State = a_Blocks(i_Index - 1).i_OldState
        a_Blocks(i_Index).i_State = a_Blocks(i_Index).i_OldState
        a_Blocks(i_Index).i_Counter = 0
        a_Blocks(i_Index - 1).i_Counter = 0
        
        ' We need to check this for a clear
        a_Blocks(i_Index).b_CheckClear = True
        a_Blocks(i_Index - 1).b_CheckClear = True
        
    End If

End Sub

Private Sub BsFlashing(ByVal i_Index As Long)

    ' Increment the counter
    a_Blocks(i_Index).i_Counter = a_Blocks(i_Index).i_Counter + 1

    ' See if we're done flashing yet
    If a_Blocks(i_Index).i_Counter >= i_FlashLength Then
        ' Set the state to clearing
        a_Blocks(i_Index).i_State = BS_CLEARING
        
        ' Set the counter to -1 so that when the next frame comes around all the
        ' counters can be assign properly
        a_Blocks(i_Index).i_Counter = -1
    End If
        

End Sub

Private Sub BsClearing(ByVal i_Index)

    Dim i_Last As Integer, i_PopIndex As Integer, t_Blank As Block
    
    ' If this block just got off prefall we need to loop through the entire stack
    ' and set the counters for all blocks with the same clear ID
    If a_Blocks(i_Index).i_Counter = -1 Then
    
        For i = 1 To FIELD_WIDTH * FIELD_HEIGHT
        
            ' Check to see if this has the same clear ID
            If a_Blocks(i).i_ClearID = a_Blocks(i_Index).i_ClearID Then
                
                ' Increment the pop counter
                i_PopIndex = i_PopIndex + 1
                
                ' Set the new counter value and label this block as having been handled
                a_Blocks(i).i_Counter = (i_PopIndex * 10) + i_PopDelay
                a_Blocks(i).b_Handled = True
                a_Blocks(i).i_PopIndex = i_PopIndex
                
                ' Set this as the last edited block
                i_Last = i
                
            End If
        
        Next i
    
        ' Flag the last block to pop as the last block to pop
        a_Blocks(i_Last).b_Last = True
    
    Else
    
        ' If the counter isn't down to 0 decrement the value
        If a_Blocks(i_Index).i_Counter > 0 Then
            
            a_Blocks(i_Index).i_Counter = a_Blocks(i_Index).i_Counter - 1
            
            ' If the counter is now 0 add to the score
            If a_Blocks(i_Index).i_Counter = 0 Then
                g_Score.BlockPop
                
                ' Figure up the proper sound to play
                If a_Blocks(i_Index).i_PopIndex < 5 Then
                    If i_Chain < 4 Then
                        g_Sound.PlaySound (i_Chain * 4) + a_Blocks(i_Index).i_PopIndex
                    Else
                        g_Sound.PlaySound 12 + a_Blocks(i_Index).i_PopIndex
                    End If
                Else
                    If i_Chain < 4 Then
                        g_Sound.PlaySound (i_Chain * 4) + 4
                    Else
                        g_Sound.PlaySound 16
                    End If
                End If
                
                ' Create some particles
                i_Y = Int((i_Index - 1) / FIELD_WIDTH)
                i_X = i_Index - (i_Y * FIELD_WIDTH)
                g_Particles.CreateParticles 4, 15, ((i_X - 1) * BLOCK_WIDTH) + ((SCREEN_WIDTH - (BLOCK_WIDTH * FIELD_WIDTH)) / 2), (i_Y * BLOCK_HEIGHT) + ((SCREEN_HEIGHT - (FIELD_HEIGHT * BLOCK_HEIGHT)) / 2)
                
            End If
            
        Else
            
            ' If this was the last block to pop we need to go through the stack and
            ' destroy any blocks with this clear ID
            If a_Blocks(i_Index).b_Last = True Then
                For i = 1 To FIELD_WIDTH * FIELD_HEIGHT
                    If a_Blocks(i).i_ClearID = a_Blocks(i_Index).i_ClearID Then
                        a_Blocks(i) = t_Blank
                        
                        ' If the block above isn't empty label it as falling because of
                        ' a clear
                        If i - FIELD_WIDTH > 0 Then
                            If a_Blocks(i - FIELD_WIDTH).i_State = BS_NORMAL Then a_Blocks(i - FIELD_WIDTH).b_ClearFall = True
                        End If
                        
                        ' If this is the last block in the chain break out of the loop
                        If i = i_Index Then Exit For
                    End If
                Next i
            End If
        End If
        
    End If

End Sub
