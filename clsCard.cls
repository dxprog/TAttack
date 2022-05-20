VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The card type
Private Type Card
    i_X As Integer
    i_Y As Integer
    i_OffsetX As Integer
    i_OffsetY As Integer
    i_Counter As Integer
End Type

' The array of cards
Private a_Cards(10) As Card

Public Sub CreateChainCard(ByVal i_ChainAmount As Integer, ByVal i_Block As Integer)
    
    Dim i_Index As Integer
    
    ' Get an empty spot in the array
    i_Index = GetFreeSlot
    
    ' If the chain is more than 13 then we automatically clamp down to x?
    If i_ChainAmount > 12 Then
        a_Cards(i_Index).i_OffsetX = 160
        a_Cards(i_Index).i_OffsetY = 139
    Else
        If i_ChainAmount > 7 Then
            a_Cards(i_Index).i_OffsetY = 139
            a_Cards(i_Index).i_OffsetX = (i_ChainAmount - 8) * BLOCK_WIDTH
        Else
            a_Cards(i_Index).i_OffsetY = 107
            a_Cards(i_Index).i_OffsetX = (i_ChainAmount - 1) * BLOCK_WIDTH
        End If
    End If
    
    ' Set the X and Y values
    a_Cards(i_Index).i_Y = Int((i_Block - 1) / FIELD_WIDTH)
    a_Cards(i_Index).i_X = (i_Block - (a_Cards(i_Index).i_Y * FIELD_WIDTH))

End Sub

Public Function HandleCards()

    Dim t_Blank As Card, i_X As Long, i_Y As Long, i_Alpha As Integer
    
    ' Loop through the cards, adjust them, and draw them
    For i = 1 To 10
        
        ' If the card isn't blank handle it
        If a_Cards(i).i_X <> 0 Then
        
            ' Increment the Y counter
            a_Cards(i).i_Counter = a_Cards(i).i_Counter + 1
        
            ' Set the alpha
            i_Alpha = (((a_Cards(i).i_Counter - BLOCK_HEIGHT) * -1) / BLOCK_HEIGHT) * 255
        
            ' If this card has reached the height of a block, kill it
            If a_Cards(i).i_Counter >= BLOCK_HEIGHT Then
                a_Cards(i) = t_Blank
            Else
                ' Draw the card
                i_X = ((SCREEN_WIDTH - (FIELD_WIDTH * BLOCK_WIDTH)) / 2) + ((a_Cards(i).i_X - 1) * BLOCK_WIDTH)
                i_Y = ((SCREEN_HEIGHT - (FIELD_HEIGHT * BLOCK_HEIGHT)) / 2) + ((a_Cards(i).i_Y - 1) * BLOCK_HEIGHT)
                
                If a_Cards(i).i_Counter > 16 Then
                    g_MainImg.Blit i_X, i_Y - a_Cards(i).i_Counter, BLOCK_WIDTH, BLOCK_HEIGHT, a_Cards(i).i_OffsetX, a_Cards(i).i_OffsetY, (Abs(a_Cards(i).i_Y - 32) / 16) * 255
                Else
                    g_MainImg.Blit i_X, i_Y - a_Cards(i).i_Counter, BLOCK_WIDTH, BLOCK_HEIGHT, a_Cards(i).i_OffsetX, a_Cards(i).i_OffsetY, 255
                End If
                
            End If
        End If
        
    Next i

End Function

Private Function GetFreeSlot() As Integer

    ' Loop through the array until we hit an unused spot
    For i = 1 To 10
        If a_Cards(i).i_X = 0 Then
            GetFreeSlot = i
            Exit Function
        End If
    Next i

End Function

Public Sub CreateComboCard(ByVal i_Combo As Integer, ByVal i_Block As Integer, Optional i_OffsetY As Integer)

    Dim i_Index As Integer
    
    ' Get an empty spot in the array
    i_Index = GetFreeSlot
    
    ' If the chain is more than 13 then we automatically clamp down to x?
    If i_Combo > 15 Then
        a_Cards(i_Index).i_OffsetY = 203
        a_Cards(i_Index).i_OffsetX = 160
    Else
        If i_Combo > 10 Then
            a_Cards(i_Index).i_OffsetY = 203
            a_Cards(i_Index).i_OffsetX = (i_Combo - 11) * BLOCK_WIDTH
        Else
            a_Cards(i_Index).i_OffsetY = 171
            a_Cards(i_Index).i_OffsetX = (i_Combo - 4) * BLOCK_WIDTH
        End If
    End If
    
    ' Set the X and Y values
    a_Cards(i_Index).i_Y = Int((i_Block - 1) / FIELD_WIDTH)
    a_Cards(i_Index).i_X = (i_Block - (a_Cards(i_Index).i_Y * FIELD_WIDTH))
    a_Cards(i_Index).i_Y = a_Cards(i_Index).i_Y - i_OffsetY

End Sub
