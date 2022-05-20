VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Arrays to store the combo and chain values
Private a_Combo As Variant, a_Chain As Variant

' The score
Public i_Score As Long

Private Sub Class_Initialize()

    ' Set up the chain and combo scores
    a_Chain = Array(50, 80, 150, 300, 400, 500, 700, 900, 1100, 1300, 1500, 1800, 2000)
    a_Combo = Array(20, 30, 50, 60, 70, 80, 100, 140, 170, 200, 240, 300, 1000)

End Sub

Public Sub BlockPop()

    ' Increment the score
    i_Score = i_Score + 10

    ' Since we're already here, go ahead and do some figuring on the speed
    i_SpeedChange = i_SpeedChange - 1
    
    ' If we've hit 0 up the speed and reset the change counter
    If i_SpeedChange <= 0 Then
        i_Speed = i_Speed + 1
        i_SpeedChange = i_Speed * 1.5
        i_StackSpeed = 60 - (i_Speed * 2)
    End If

End Sub

Public Sub ChainScore(ByVal i_ChainAmount As Integer)
    
    If i_ChainAmount > 12 Then
        i_Score = i_Score + a_Chain(12)
    Else
        i_Score = i_Score + a_Chain(i_ChainAmount - 1)
    End If
    
End Sub

Public Sub ComboScore(ByVal i_Combo As Integer)

    If i_Combo - 4 > 12 Then
        i_Score = i_Score + a_Combo(12)
    Else
        i_Score = i_Score + a_Combo(i_Combo - 4)
    End If

End Sub
