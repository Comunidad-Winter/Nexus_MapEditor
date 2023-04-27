Attribute VB_Name = "mPooChar"
Option Explicit

Public Sub Char_Make(ByVal CharIndex As Integer, _
                     ByVal Body As Integer, _
                     ByVal Head As Integer, _
                     ByVal Heading As Byte, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal Arma As Integer, _
                     ByVal Escudo As Integer, _
                     ByVal Casco As Integer, _
                     ByVal AuraAnim As Long, _
                     ByVal AuraColor As Long)
 
    'Apuntamos al ultimo Char
 
    If CharIndex > LastChar Then
        LastChar = CharIndex

    End If
 
    NumChars = NumChars + 1

    If Arma = 0 Then Arma = 2
    If Escudo = 0 Then Escudo = 2
    If Casco = 0 Then Casco = 2
        
    With charlist(CharIndex)
       
        'If the char wasn't allready active (we are rewritting it) don't increase char count
                
        .iHead = Head
        .iBody = Body
                
        .Head = HeadData(Head)
        .Body = BodyData(Body)
                
        .Arma = WeaponAnimData(Arma)
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)

        Call InitGrh(.AuraAnim, AuraAnim)
        .AuraColor = AuraColor
        
        .Heading = Heading
         
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0

        'Update position
        .Pos.X = X
        .Pos.Y = Y
           
    End With
   
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
       
End Sub

Public Sub Char_Erase(ByVal CharIndex As Integer)
    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************
 
    If (CharIndex = 0) Then Exit Sub
    If (CharIndex > LastChar) Then Exit Sub
 
    With charlist(CharIndex)
                
        If Map_InBounds(.Pos.X, .Pos.Y) Then  '// Posicion valida
            MapData(.Pos.X, .Pos.Y).CharIndex = 0  '// Borramos el user
        End If
       
        'Update lastchar
        If CharIndex = LastChar Then
 
            Do Until charlist(LastChar).Heading > 0
               
                LastChar = LastChar - 1
 
                If LastChar = 0 Then
                                
                    NumChars = 0

                    Exit Sub

                End If
                       
            Loop
 
        End If
   
        Call Char_ResetInfo(CharIndex)
                
        'Update NumChars
        NumChars = NumChars - 1
 
        Exit Sub
 
    End With
 
End Sub

Sub Char_CleanAll()
    
    '// Borramos los obj y char que esten

    Dim X         As Long, Y As Long
    Dim CharIndex As Integer, obj As Long
    
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
          
            'Erase NPCs
            CharIndex = Char_MapPosExits(CByte(X), CByte(Y))
 
            If (CharIndex > 0) Then
                Call Char_Erase(CharIndex)
            End If
                        
            'Erase OBJs
            obj = Map_PosExitsObject(CByte(X), CByte(Y))

            If (obj > 0) Then
                Call Map_DestroyObject(CByte(X), CByte(Y))
            End If

        Next Y
    Next X

End Sub

Public Function Char_MapPosExits(ByVal X As Byte, ByVal Y As Byte) As Integer
 
    '*****************************************************************
    'Checks to see if a tile position has a char_index and return it
    '*****************************************************************
   
    If (Map_InBounds(X, Y)) Then
        Char_MapPosExits = MapData(X, Y).CharIndex
    Else
        Char_MapPosExits = 0
    End If
  
End Function

Private Sub Char_ResetInfo(ByVal CharIndex As Integer)

    '*****************************************************************
    'Author: Ao 13.0
    'Last Modify Date: 13/12/2013
    'Reset Info User
    '*****************************************************************

    With charlist(CharIndex)
        'Remove particles
        'Call Char_Particle_Group_Remove_All(CharIndex)
            
        .active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
            
        .Moving = 0
        .muerto = False
        .Nombre = vbNullString
        .Clan = vbNullString
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
            
    End With
 
End Sub

Function NextOpenChar() As Integer
    '*****************************************************************
    'Finds next open char slot in CharList
    '*****************************************************************
    
    On Error GoTo NextOpenChar_Err
    
    Dim loopc As Integer

    loopc = 1

    Do While charlist(loopc).active
        loopc = loopc + 1
    Loop

    NextOpenChar = loopc

    
    Exit Function

NextOpenChar_Err:
    Call LogError(Err.Number, Err.Description, "mPooChar.NextOpenChar", Erl)
    Resume Next
    
End Function
