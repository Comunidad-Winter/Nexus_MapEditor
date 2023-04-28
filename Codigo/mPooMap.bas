Attribute VB_Name = "mPooMap"
Option Explicit

Function Map_InBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Checks to see if a tile position is in the maps bounds
      '*****************************************************************

      If (X < XMinMapSize) Or (X > XMaxMapSize) Or (Y < YMinMapSize) Or (Y > YMaxMapSize) Then
            Map_InBounds = False

            Exit Function

      End If
    
      Map_InBounds = True
End Function

Public Function Map_PosExitsObject(ByVal X As Byte, ByVal Y As Byte) As Long
 
      '*****************************************************************
      'Checks to see if a tile position has a char_index and return it
      '*****************************************************************

      If (Map_InBounds(X, Y)) Then
        If MapData(X, Y).ObjGrh.GrhIndex > 0 Then
            Map_PosExitsObject = MapData(X, Y).ObjGrh.GrhIndex
        ElseIf MapData(X, Y).Particle_Group_Index > 0 Then
            Map_PosExitsObject = MapData(X, Y).Particle_Group_Index
        End If
      Else
            Map_PosExitsObject = 0
      End If
 
End Function

Public Sub Map_DestroyObject(ByVal X As Byte, ByVal Y As Byte)

      If (Map_InBounds(X, Y)) Then

            With MapData(X, Y)
                  '.objgrh.GrhIndex = 0
                  .OBJInfo.OBJIndex = 0
                  .OBJInfo.Amount = 0

                  Call GrhUninitialize(.ObjGrh)
        
            End With

      End If

End Sub

Function Map_LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Author: ZaMa
      'Last Modification: 06/04/2020
      'Checks to see if a tile position is legal, including if there is a casper in the tile
      '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
      '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
      '12/01/2020: Recox - Now we manage monturas.
      '06/04/2020: FrankoH298 - Si estamos montados, no nos deja ingresar a las casas.
      '*****************************************************************

      Dim CharIndex As Integer
    
      'Limites del mapa

      If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then

            Exit Function

      End If
    
      'Tile Bloqueado?

      If (Map_GetBlocked(X, Y)) Then
         
            Exit Function

      End If
    
      CharIndex = (Char_MapPosExits(CByte(X), CByte(Y)))
        
      'Hay un personaje?

      If (CharIndex > 0) Then
    
            If (Map_GetBlocked(UserPos.X, UserPos.Y)) Then
                
                  Exit Function

            End If
        
            With charlist(CharIndex)
                  ' Si no es casper, no puede pasar

                  If WalkMode Then
                              
                        Exit Function

                  Else
                        ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)

                        If (Map_CheckWater(UserPos.X, UserPos.Y)) Then
                              If Not (Map_CheckWater(X, Y)) Then
                                            
                                    Exit Function

                              End If

                        Else
                              ' No puedo intercambiar con un casper que este en la orilla (Lado agua)

                              If (Map_CheckWater(X, Y)) Then
                                             
                                    Exit Function

                              End If
                                        
                        End If

                  End If

            End With

      End If
   
      
    
      Map_LegalPos = True
End Function

Public Function Map_GetBlocked(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
      'Last Modify Date: 10/07/2002
      'Checks to see if a tile position is blocked
      '*****************************************************************

      If (Map_InBounds(X, Y)) Then
            Map_GetBlocked = (MapData(X, Y).Blocked)
      End If

End Function

Function Map_CheckWater(ByVal X As Integer, ByVal Y As Integer) As Boolean

      If Map_InBounds(X, Y) Then

            With MapData(X, Y)

                  If ((.Graphic(1).GrhIndex >= 1505 And .Graphic(1).GrhIndex <= 1520) Or (.Graphic(1).GrhIndex >= 5665 And .Graphic(1).GrhIndex <= 5680) Or (.Graphic(1).GrhIndex >= 13547 And .Graphic(1).GrhIndex <= 13562)) And .Graphic(2).GrhIndex = 0 Then
                        Map_CheckWater = True
                  Else
                        Map_CheckWater = False
                  End If

            End With

      End If
                  
End Function

