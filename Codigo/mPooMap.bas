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
