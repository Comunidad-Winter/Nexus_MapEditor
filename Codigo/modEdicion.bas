Attribute VB_Name = "modEdicion"
Option Explicit

Public maskBloqueo As Byte
Public TriggerBox As Byte

''
' Acciona la operacion al hacer doble click en una posicion del mapa
'
' @param tX Especifica la posicion X en el mapa
' @param tY Espeficica la posicion Y en el mapa

Sub DobleClick(tX As Integer, tY As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    ' Selecciones
    
    On Error GoTo DobleClick_Err
    
    ' Translados
    Dim tTrans As WorldPos
    tTrans = MapData(tX, tY).TileExit

    If tTrans.Map > 0 Then

        If MapInfo.Changed = 1 Then
            If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then _
                modMapIO.GuardarMapa frmMain.Dialog.FileName

        End If
    
        If LenB(frmMain.Dialog.FileName) <> 0 Then
            If FileExist(PATH_Save & NameMap_Save & tTrans.Map & ".csm", vbArchive) = True Then
                Call modMapIO.NuevoMapa
                frmMain.Dialog.FileName = PATH_Save & NameMap_Save & tTrans.Map & ".csm"
                modMapIO.AbrirMapa frmMain.Dialog.FileName
                UserPos.X = tTrans.X
                UserPos.y = tTrans.y

                If WalkMode = True Then
                    Call Char_MovebyPos(UserCharIndex, UserPos.X, UserPos.y - 7)
                    charlist(UserCharIndex).Heading = SOUTH

                End If

                frmMain.mnuReAbrirMapa.Enabled = True

            End If

        End If

    End If

    
    Exit Sub

DobleClick_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.DobleClick", Erl)
    Resume Next
    
End Sub

Public Sub Superficie_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 07/12/2018
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then _
        Exit Sub
    
    'TODO
    'modEdicion.Deshacer_Add "Superficie en area" ' Hago deshacer

    For y = y1 To y2
        For X = x1 To x2
            If Poner = True Then
                If frmConfigSup.MOSAICO.value = vbChecked Then
                    Dim aux As Integer
                    aux = Val(frmSuperficies.cGrh.Text) + _
                    ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
                     MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                    'Setup GRH
                    InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
                Else
                    'Else Place graphic
                    MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                    'Setup GRH
                    InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
                End If
            Else
                MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = 0
            End If
        Next X
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Coloca la superficie seleccionada al azar en el mapa
'

Public Sub Superficie_Azar()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    On Error Resume Next
    Dim y As Integer
    Dim X As Integer
    Dim Cuantos As Integer
    Dim k As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    Cuantos = InputBox("Cuantos Grh se deben poner en este mapa?", "Poner Grh Al Azar", 0)
    
    If Cuantos > 0 Then
        'TODO
        'modEdicion.Deshacer_Add "Insertar Superficie al Azar" ' Hago deshacer
        
        For k = 1 To Cuantos
            X = RandomNumber(10, 90)
            y = RandomNumber(10, 90)
            
            If frmConfigSup.MOSAICO.value = vbChecked Then
              Dim aux As Integer
              Dim dy As Integer
              Dim dX As Integer
              If frmConfigSup.DespMosaic.value = vbChecked Then
                dy = Val(frmConfigSup.DMLargo)
                dX = Val(frmConfigSup.DMAncho.Text)
              Else
                dy = 0
                dX = 0
              End If
                    
              If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                    aux = Val(frmSuperficies.cGrh.Text) + _
                    (((y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
                    If frmBloqueos.cInsertarBloqueo.value = True Then
                        MapData(X, y).Blocked = 1
                    Else
                        MapData(X, y).Blocked = 0
                    End If
                    MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                    InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
              Else
                    Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                    tXX = X
                    tYY = y
                    desptile = 0
                    For i = 1 To frmConfigSup.mLargo.Text
                        For j = 1 To frmConfigSup.mAncho.Text
                            aux = Val(frmSuperficies.cGrh.Text) + desptile
                             
                            If frmBloqueos.cInsertarBloqueo.value = True Then
                                MapData(tXX, tYY).Blocked = 1
                            Else
                                MapData(tXX, tYY).Blocked = 0
                            End If
    
                             MapData(tXX, tYY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                             
                             InitGrh MapData(tXX, tYY).Graphic(Val(frmSuperficies.cCapas.Text)), aux
                             tXX = tXX + 1
                             desptile = desptile + 1
                        Next
                        tXX = X
                        tYY = tYY + 1
                    Next
                    tYY = y
              End If
            End If
        Next
    End If
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Coloca la superficie seleccionada en todos los bordes
'

Public Sub Superficie_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    Dim y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    'TODO
    'modEdicion.Deshacer_Add "Insertar Superficie en todos los bordes" ' Hago deshacer
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    
              If frmConfigSup.MOSAICO.value = vbChecked Then
                Dim aux As Integer
                aux = Val(frmSuperficies.cGrh.Text) + _
                ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
                If frmBloqueos.cInsertarBloqueo.value = True Then
                    MapData(X, y).Blocked = 1
                Else
                    MapData(X, y).Blocked = 0
                End If
                MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                'Setup GRH
                InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
              Else
                'Else Place graphic
                If frmBloqueos.cInsertarBloqueo.value = True Then
                    MapData(X, y).Blocked = 1
                Else
                    MapData(X, y).Blocked = 0
                End If
                
                MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                
                'Setup GRH
        
                InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
            End If
                 'Erase NPCs
                If MapData(X, y).NPCIndex > 0 Then
                    Call Char_Erase(MapData(X, y).CharIndex)
                    MapData(X, y).NPCIndex = 0
                End If
    
                'Erase Objs
                MapData(X, y).OBJInfo.objindex = 0
                MapData(X, y).OBJInfo.Amount = 0
                MapData(X, y).ObjGrh.GrhIndex = 0
    
                'Clear exits
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0
    
            End If
    
        Next X
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Coloca la misma superficie seleccionada en todo el mapa
'

Public Sub Superficie_Todo()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    'TODO
    'modEdicion.Deshacer_Add "Insertar Superficie en todo el mapa" ' Hago deshacer
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If frmConfigSup.MOSAICO.value = vbChecked Then
                Dim aux As Integer
                aux = Val(frmSuperficies.cGrh.Text) + _
                ((y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
                 MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                'Setup GRH
                InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), aux
            Else
                'Else Place graphic
                MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                'Setup GRH
                InitGrh MapData(X, y).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
            End If
    
        Next X
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Manda una advertencia de Edicion Critica
'
' @return   Nos devuelve si acepta o no el cambio

Private Function EditWarning() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If MsgBox(MSGDang, vbExclamation + vbYesNo) = vbNo Then
        EditWarning = True
    Else
        EditWarning = False
    End If
End Function

''
' Realiza una operacion de edicion aislada sobre el mapa
'
' @param Button Indica el estado del Click del mouse
' @param tX Especifica la posicion X en el mapa
' @param tY Especifica la posicion Y en el mapa

Sub ClickEdit(Button As Integer, _
              tX As Byte, _
              tY As Byte, _
              Optional ByVal Deshacer As Boolean = True)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo ClickEdit_Err
    
    Dim loopc    As Integer

    Dim NPCIndex As Integer

    Dim objindex As Integer

    Dim Head     As Integer

    Dim Body     As Integer

    Dim Heading  As Byte
    
    If tY < YMinMapSize Or tY > YMaxMapSize Then Exit Sub
    If tX < XMinMapSize Or tX > XMaxMapSize Then Exit Sub
    
    If Button = 0 Then
        ' Pasando sobre :P
        SobreY = tY
        SobreX = tX
        
        Exit Sub

    End If
    
    'Right
    
    If Button = vbRightButton Then
        ' Posicion
        frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & ENDL & ENDL & "Posición " & tX & "," & tY
        
        ' Bloqueos
        If MapData(tX, tY).Blocked > 0 Then frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (BLOQ)"
        
        ' Translados
        If MapData(tX, tY).TileExit.Map <> 0 Then
            If frmMain.mnuAutoCapturarTranslados.Checked = True Then
                frmTraslados.tTMapa.Text = MapData(tX, tY).TileExit.Map
                frmTraslados.tTX.Text = MapData(tX, tY).TileExit.X
                frmTraslados.tTY = MapData(tX, tY).TileExit.y

            End If

            frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (Trans.: " & MapData(tX, tY).TileExit.Map & "," & MapData(tX, tY).TileExit.X & "," & MapData(tX, tY).TileExit.y & ")"

        End If
        
        ' NPCs
        If MapData(tX, tY).NPCIndex > 0 Then
            If MapData(tX, tY).NPCIndex > 499 Then
                frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (NPC-Hostil: " & MapData(tX, tY).NPCIndex & " - " & NpcData(MapData(tX, tY).NPCIndex).Name & ")"
            Else
                frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (NPC: " & MapData(tX, tY).NPCIndex & " - " & NpcData(MapData(tX, tY).NPCIndex).Name & ")"

            End If

        End If
        
        ' OBJs
        If MapData(tX, tY).OBJInfo.objindex > 0 Then
            frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (Obj: " & MapData(tX, tY).OBJInfo.objindex & " - " & ObjData(MapData(tX, tY).OBJInfo.objindex).Name & " - Cant.:" & MapData(tX, tY).OBJInfo.Amount & ")"

        End If
        
        ' Capas
        frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & ENDL & "Capa1: " & MapData(tX, tY).Graphic(1).GrhIndex & " - Capa2: " & MapData(tX, tY).Graphic(2).GrhIndex & " - Capa3: " & MapData(tX, tY).Graphic(3).GrhIndex & " - Capa4: " & MapData(tX, tY).Graphic(4).GrhIndex

        If frmMain.mnuAutoCapturarSuperficie.Checked = True And frmSuperficies.cSeleccionarSuperficie.value = False Then
            If MapData(tX, tY).Graphic(4).GrhIndex <> 0 Then
                frmSuperficies.cCapas.Text = 4
                frmSuperficies.cGrh.Text = MapData(tX, tY).Graphic(4).GrhIndex
            ElseIf MapData(tX, tY).Graphic(3).GrhIndex <> 0 Then
                frmSuperficies.cCapas.Text = 3
                frmSuperficies.cGrh.Text = MapData(tX, tY).Graphic(3).GrhIndex
            ElseIf MapData(tX, tY).Graphic(2).GrhIndex <> 0 Then
                frmSuperficies.cCapas.Text = 2
                frmSuperficies.cGrh.Text = MapData(tX, tY).Graphic(2).GrhIndex
            ElseIf MapData(tX, tY).Graphic(1).GrhIndex <> 0 Then
                frmSuperficies.cCapas.Text = 1
                frmSuperficies.cGrh.Text = MapData(tX, tY).Graphic(1).GrhIndex

            End If

            frmRemplazo.GrhReplaceFrom.Text = frmSuperficies.cGrh.Text

        End If
        
        ' Limpieza
        If Len(frmConsola.StatTxt.Text) > 4000 Then
            frmConsola.StatTxt.Text = Right(frmConsola.StatTxt.Text, 3000)

        End If

        frmConsola.StatTxt.SelStart = Len(frmConsola.StatTxt.Text)
        
        Exit Sub

    End If
    
    'Left click
    If Button = vbLeftButton Then
            
        'Erase 2-3
        If frmSuperficies.cQuitarEnTodasLasCapas.value = True Then
            'If Deshacer Then modEdicion.Deshacer_Add "Quitar Todas las Capas (2/3)" ' Hago deshacer
            MapInfo.Changed = 1 'Set changed flag

            For loopc = 2 To 3
                MapData(tX, tY).Graphic(loopc).GrhIndex = 0
            Next loopc

            Call DibujarMinimapa
            Exit Sub

        End If
    
        'Borrar "esta" Capa
        If frmSuperficies.cQuitarEnEstaCapa.value = True Then
            If Val(frmSuperficies.cCapas.Text) = 1 Then
                If MapData(tX, tY).Graphic(1).GrhIndex <> 1 Then
                    'If Deshacer Then modEdicion.Deshacer_Add "Quitar Capa 1" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Graphic(1).GrhIndex = 1
                    Exit Sub

                End If

            ElseIf MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> 0 Then
                'If Deshacer Then modEdicion.Deshacer_Add "Quitar Capa " & frmSuperficies.cCapas.Text  ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = 0
                Call DibujarMinimapa
                Exit Sub

            End If

        End If
    
        '************** Place grh
        If frmSuperficies.cSeleccionarSuperficie.value = True Then
            
            If frmConfigSup.MOSAICO.value = vbChecked Then

                Dim aux As Long

                Dim dy  As Integer

                Dim dX  As Integer

                If frmConfigSup.DespMosaic.value = vbChecked Then
                    dy = Val(frmConfigSup.DMLargo)
                    dX = Val(frmConfigSup.DMAncho.Text)
                Else
                    dy = 0
                    dX = 0

                End If
                    
                If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                    'If Deshacer Then modEdicion.Deshacer_Add "Insertar Superficie' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    aux = Val(frmSuperficies.cGrh.Text) + (((tY + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((tX + dX) Mod frmConfigSup.mAncho.Text)

                    If MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> aux Or MapData(tX, tY).Blocked <> frmMain.SelectPanel(2).value Then
                        MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                        InitGrh MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)), aux

                    End If

                Else
                    'If Deshacer Then modEdicion.Deshacer_Add "Insertar Auto-Completar Superficie' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag

                    Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer

                    tXX = tX
                    tYY = tY
                    desptile = 0

                    For i = 1 To frmConfigSup.mLargo.Text
                        For j = 1 To frmConfigSup.mAncho.Text
                            aux = Val(frmSuperficies.cGrh.Text) + desptile

                            If tYY > 100 Then Exit Sub
                            If tXX > 100 Then Exit Sub
                            MapData(tXX, tYY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                            InitGrh MapData(tXX, tYY).Graphic(Val(frmSuperficies.cCapas.Text)), aux
                            tXX = tXX + 1
                            desptile = desptile + 1
                        Next
                        tXX = tX
                        tYY = tYY + 1
                    Next
                    tYY = tY
                    
                End If
              
            Else

                'Else Place graphic
                If MapData(tX, tY).Blocked <> frmMain.SelectPanel(2).value Or MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> Val(frmSuperficies.cGrh.Text) Then
                    'If Deshacer Then modEdicion.Deshacer_Add "Quitar Superficie en Capa " & frmSuperficies.cCapas.Text ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = Val(frmSuperficies.cGrh.Text)
                    'Setup GRH
                    InitGrh MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)), Val(frmSuperficies.cGrh.Text)
                    
                End If

            End If

            Call DibujarMinimapa
            
        End If

        '************** Place blocked tile
        If frmBloqueos.cInsertarBloqueo.value = True Then
            If MapData(tX, tY).Blocked <> maskBloqueo Then
                'If Deshacer Then modEdicion.Deshacer_Add "Insertar Bloqueo" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Blocked = maskBloqueo
                
            End If

        ElseIf frmBloqueos.cQuitarBloqueo.value = True Then

            If MapData(tX, tY).Blocked <> 0 Then
                'If Deshacer Then modEdicion.Deshacer_Add "Quitar Bloqueo" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Blocked = 0

            End If

        End If
    
        '************** Place exit
        If frmTraslados.cInsertarTrans.value = True Then
            If Cfg_TrOBJ > 0 And Cfg_TrOBJ <= NumOBJs And frmTraslados.cInsertarTransOBJ.value = True Then
                If ObjData(Cfg_TrOBJ).ObjType = 19 Then
                    'If Deshacer Then modEdicion.Deshacer_Add "Insertar Objeto de Translado" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    InitGrh MapData(tX, tY).ObjGrh, ObjData(Cfg_TrOBJ).GrhIndex
                    MapData(tX, tY).OBJInfo.objindex = Cfg_TrOBJ
                    MapData(tX, tY).OBJInfo.Amount = 1

                End If

            End If

            If Val(frmTraslados.tTMapa.Text) < -1 Or Val(frmTraslados.tTMapa.Text) > 9000 Then
                MsgBox "Valor de Mapa invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf Val(frmTraslados.tTX.Text) < 0 Or Val(frmTraslados.tTX.Text) > 100 Then
                MsgBox "Valor de X invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf Val(frmTraslados.tTY.Text) < 0 Or Val(frmTraslados.tTY.Text) > 100 Then
                MsgBox "Valor de Y invalido", vbCritical + vbOKOnly
                Exit Sub

            End If

            If frmTraslados.cUnionManual.value = True Then
                'If Deshacer Then modEdicion.Deshacer_Add "Insertar Translado de Union Manual' Hago deshacer"
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).TileExit.Map = Val(frmTraslados.tTMapa.Text)

                If tX >= 90 Then ' 21 ' derecha
                    MapData(tX, tY).TileExit.X = 12
                    MapData(tX, tY).TileExit.y = tY
                ElseIf tX <= 11 Then ' 9 ' izquierda
                    MapData(tX, tY).TileExit.X = 91
                    MapData(tX, tY).TileExit.y = tY

                End If

                If tY >= 91 Then ' 94 '''' hacia abajo
                    MapData(tX, tY).TileExit.y = 11
                    MapData(tX, tY).TileExit.X = tX
                ElseIf tY <= 10 Then ''' hacia arriba
                    MapData(tX, tY).TileExit.y = 90
                    MapData(tX, tY).TileExit.X = tX

                End If

            Else
                'If Deshacer Then modEdicion.Deshacer_Add "Insertar Translado" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).TileExit.Map = Val(frmTraslados.tTMapa.Text)
                MapData(tX, tY).TileExit.X = Val(frmTraslados.tTX.Text)
                MapData(tX, tY).TileExit.y = Val(frmTraslados.tTY.Text)

            End If

        ElseIf frmTraslados.cQuitarTrans.value = True Then
            'If Deshacer Then modEdicion.Deshacer_Add "Quitar Translado" ' Hago deshacer
            MapInfo.Changed = 1 'Set changed flag
            MapData(tX, tY).TileExit.Map = 0
            MapData(tX, tY).TileExit.X = 0
            MapData(tX, tY).TileExit.y = 0

        End If
    
        '************** Place NPC
        If frmNpcs.cInsertarFunc.value = True Then
            If frmNpcs.cNPC.Text > 0 Then
                NPCIndex = frmNpcs.cNPC.Text

                If NPCIndex <> MapData(tX, tY).NPCIndex Then
                    'If Deshacer Then modEdicion.Deshacer_Add "Insertar NPC" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call Char_Make(NextOpenChar(), Body, Head, Heading, tX, tY, 2, 2, 2, 0, 0)
                    MapData(tX, tY).NPCIndex = NPCIndex

                End If

            End If

        ElseIf frmNpcs.cInsertarFunc.value = True Then

            If frmNpcs.cNPC.Text > 0 Then
                NPCIndex = frmNpcs.cNPC.Text

                If NPCIndex <> (MapData(tX, tY).NPCIndex) Then
                    'If Deshacer Then modEdicion.Deshacer_Add "Insertar NPC Hostil' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call Char_Make(NextOpenChar(), Body, Head, Heading, tX, tY, 2, 2, 2, 0, 0)
                    MapData(tX, tY).NPCIndex = NPCIndex

                End If

            End If

        ElseIf frmNpcs.cQuitarNpc.value = True Then

            If MapData(tX, tY).NPCIndex > 0 Then
                'If Deshacer Then modEdicion.Deshacer_Add "Quitar NPC" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).NPCIndex = 0
                Call Char_Erase(MapData(tX, tY).CharIndex)
                
            End If

        End If
    
        ' ***************** Control de Funcion de Objetos *****************
        If frmObjetos.cInsertarObj.value = True Then ' Insertar Objeto
            If frmObjetos.cOBJ.Text > 0 Then
                objindex = frmObjetos.cOBJ.Text

                If MapData(tX, tY).OBJInfo.objindex <> objindex Or MapData(tX, tY).OBJInfo.Amount <> Val(frmObjetos.cCantidad.Text) Then
                    'If Deshacer Then modEdicion.Deshacer_Add "Insertar Objeto" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    InitGrh MapData(tX, tY).ObjGrh, ObjData(objindex).GrhIndex
                    MapData(tX, tY).OBJInfo.objindex = objindex
                    MapData(tX, tY).OBJInfo.Amount = Val(frmObjetos.cCantidad.Text)

                    Select Case ObjData(objindex).ObjType

                        Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                            MapData(tX, tY).Graphic(3) = MapData(tX, tY).ObjGrh
                            MapData(tX, tY).Blocked = &HF

                    End Select

                End If

            End If

        ElseIf frmObjetos.cQuitarObj.value = True Then ' Quitar Objeto

            If MapData(tX, tY).OBJInfo.objindex <> 0 Or MapData(tX, tY).OBJInfo.Amount <> 0 Then
                'If Deshacer Then modEdicion.Deshacer_Add "Quitar Objeto" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag

                If MapData(tX, tY).Graphic(3).GrhIndex = MapData(tX, tY).ObjGrh.GrhIndex Then MapData(tX, tY).Graphic(3).GrhIndex = 0
                MapData(tX, tY).ObjGrh.GrhIndex = 0
                MapData(tX, tY).OBJInfo.objindex = 0
                MapData(tX, tY).OBJInfo.Amount = 0
                MapData(tX, tY).Blocked = 0

            End If

        End If
        
        ' ***************** Control de Funcion de Triggers *****************
        If frmTriggers.cInsertarTrigger.value = True Then ' Insertar Trigger
            If TriggerBox < 10 Then
                TriggerBox = frmTriggers.lListado.ListIndex

            End If

            If MapData(tX, tY).Trigger <> TriggerBox Then
                'If Deshacer Then modEdicion.Deshacer_Add "Insertar Trigger" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Trigger = TriggerBox

            End If

        ElseIf frmTriggers.cQuitarTrigger.value = True Then ' Quitar Trigger

            If MapData(tX, tY).Trigger <> 0 Then
                'If Deshacer Then modEdicion.Deshacer_Add "Quitar Trigger" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Trigger = 0

            End If
            
        End If
    End If
    
        Exit Sub

ClickEdit_Err:
        Call LogError(Err.Number, Err.Description, "modEdicion.ClickEdit", Erl)

        Resume Next
    
    End Sub
