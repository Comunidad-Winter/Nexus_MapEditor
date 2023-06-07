Attribute VB_Name = "modEdicion"
Option Explicit

Public TriggerBox As Byte

' Deshacer
Public Const maxDeshacer As Integer = 10
Public MapData_Deshacer() As MapBlock

Type tDeshacerInfo

    Libre As Boolean
    Desc As String

End Type

Public MapData_Deshacer_Info(1 To maxDeshacer) As tDeshacerInfo

Public Sub InitDeshacer()
    '*************************************************
    'Author: Lorwik
    'Last modified: 22/03/2021
    '*************************************************
    
    ReDim MapData_Deshacer(1 To maxDeshacer, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
End Sub

''
' Vacia el Deshacer
'
Public Sub Deshacer_Clear()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo Deshacer_Clear_Err
    
    Dim i As Integer

    ' Vacio todos los campos afectados
    For i = 1 To maxDeshacer
        MapData_Deshacer_Info(i).Libre = True
    Next
    ' no ahi que deshacer
    frmMain.mnuDeshacer.Enabled = False

    Exit Sub

Deshacer_Clear_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Deshacer_Clear", Erl)
    Resume Next
    
End Sub

''
' Deshacer un paso del Deshacer
'
Public Sub Deshacer_Recover()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo Deshacer_Recover_Err
    
    Dim i       As Integer
    Dim F       As Integer
    Dim j       As Integer
    Dim Body    As Integer
    Dim Head    As Integer
    Dim Heading As Byte

    If MapData_Deshacer_Info(1).Libre = False Then

        ' Aplico deshacer
        For F = XMinMapSize To XMaxMapSize
            For j = YMinMapSize To YMaxMapSize

                If (MapData(F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex <> MapData_Deshacer(1, F, j).NPCIndex) Or (MapData(F, j).NPCIndex <> 0 And MapData_Deshacer(1, F, j).NPCIndex = 0) Then
                    ' Si ahi un NPC, y en el deshacer es otro lo borramos
                    ' (o) Si aun no NPC y en el deshacer no esta
                    MapData(F, j).NPCIndex = 0
                    Call Char_Erase(MapData(F, j).CharIndex)

                End If

                If MapData_Deshacer(1, F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex = 0 Then
                    ' Si ahi un NPC en el deshacer y en el no esta lo hacemos
                    Body = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Body
                    Head = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Head
                    Heading = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Heading
                    Call Char_Make(NextOpenChar(), Body, Head, Heading, F, j, 0, 0, 0, 0, 0)
                Else
                    MapData(F, j) = MapData_Deshacer(1, F, j)

                End If

            Next
        Next
        MapData_Deshacer_Info(1).Libre = True

        ' Desplazo todos los deshacer uno hacia adelante
        For i = 1 To maxDeshacer - 1
            For F = XMinMapSize To XMaxMapSize
                For j = YMinMapSize To YMaxMapSize
                    MapData_Deshacer(i, F, j) = MapData_Deshacer(i + 1, F, j)
                Next
            Next
            MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i + 1)
        Next
        ' borro el ultimo
        MapData_Deshacer_Info(maxDeshacer).Libre = True

        ' ahi para deshacer?
        If MapData_Deshacer_Info(1).Libre = True Then
            frmMain.mnuDeshacer.Caption = "&Deshacer (no ahi nada que deshacer)"
            frmMain.mnuDeshacer.Enabled = False
        Else
            frmMain.mnuDeshacer.Caption = "&Deshacer (Ultimo: " & MapData_Deshacer_Info(1).Desc & ")"
            frmMain.mnuDeshacer.Enabled = True

        End If

    Else
        MsgBox "No ahi acciones para deshacer", vbInformation

    End If

    Call DibujarMinimapa

    
    Exit Sub


Deshacer_Recover_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Deshacer_Recover", Erl)
    Resume Next
    
End Sub

''
' Agrega un Deshacer
'
Public Sub Deshacer_Add(ByVal Desc As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo Deshacer_Add_Err
    
    If frmMain.mnuUtilizarDeshacer.Checked = False Then Exit Sub

    Dim i As Integer
    Dim F As Integer
    Dim j As Integer

    ' Desplazo todos los deshacer uno hacia atras
    For i = maxDeshacer To 2 Step -1
        For F = XMinMapSize To XMaxMapSize
            For j = YMinMapSize To YMaxMapSize
                MapData_Deshacer(i, F, j) = MapData_Deshacer(i - 1, F, j)
            Next
        Next
        MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i - 1)
    Next

    ' Guardo los valores
    For F = XMinMapSize To XMaxMapSize
        For j = YMinMapSize To YMaxMapSize
            MapData_Deshacer(1, F, j) = MapData(F, j)
        Next
    Next
    MapData_Deshacer_Info(1).Desc = Desc
    MapData_Deshacer_Info(1).Libre = False
    frmMain.mnuDeshacer.Caption = "&Deshacer (Ultimo: " & MapData_Deshacer_Info(1).Desc & ")"
    frmMain.mnuDeshacer.Enabled = True

    
    Exit Sub

Deshacer_Add_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Deshacer_Add", Erl)
    Resume Next
    
End Sub

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
    
    Seleccionando = False ' GS
    SeleccionIX = 0
    SeleccionIY = 0
    SeleccionFX = 0
    SeleccionFY = 0
    
    ' Translados
    Dim tTrans As WorldPos
    tTrans = MapData(tX, tY).TileExit

    If tTrans.Map > 0 Then

        If MapInfo.Changed = 1 Then
            If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then _
                modMapIO.GuardarMapa frmMain.Dialog.FileName

        End If
    
        If LenB(frmMain.Dialog.FileName) <> 0 Then
            If FileExist(PATH_Save & NameMap_Save & tTrans.Map & MapFormat, vbArchive) = True Then
                Call modMapIO.NuevoMapa
                frmMain.Dialog.FileName = PATH_Save & NameMap_Save & tTrans.Map & MapFormat
                modMapIO.AbrirMapa frmMain.Dialog.FileName, MapFormat
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
    
    modEdicion.Deshacer_Add "Superficie en area" ' Hago deshacer

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

Public Sub Bloqueos_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Inserta As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 07/12/2018
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Bloqueo en Area" ' Hago deshacer

    For y = y1 To y2
        For X = x1 To x2
    
            If Inserta = True Then
                MapData(X, y).Blocked = 1
            Else
                MapData(X, y).Blocked = 0
            End If
    
        Next X
    Next y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Triggers_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 25/03/2021
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Triggers en Area" ' Hago deshacer

    For y = y1 To y2
        For X = x1 To x2
            If Poner = True Then
                If frmConfigSup.MOSAICO.value = vbChecked Then
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(X, y).Trigger = frmTriggers.lListado.ListIndex
                Else
                    MapInfo.Changed = 1
                    'Else Place trigger
                    MapData(X, y).Trigger = 0

                End If
            Else
                MapInfo.Changed = 1
                MapData(X, y).Trigger = 0
                
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

        modEdicion.Deshacer_Add "Insertar Superficie al Azar" ' Hago deshacer
        
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
    
    modEdicion.Deshacer_Add "Insertar Superficie en todos los bordes" ' Hago deshacer
    
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
    
    modEdicion.Deshacer_Add "Insertar Superficie en todo el mapa" ' Hago deshacer
    
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
              tX As Integer, _
              tY As Integer, _
              Optional ByVal Deshacer As Boolean = True)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo ClickEdit_Err
    
    Dim LoopC    As Integer
    Dim NPCIndex As Integer
    Dim objindex As Integer
    Dim Head     As Integer
    Dim Body     As Integer
    Dim Heading  As Byte
    
    If tY < YMinMapSize Or tY > YMaxMapSize Then Exit Sub
    If tX < XMinMapSize Or tX > XMaxMapSize Then Exit Sub
    
    If Button = 0 Then
        'Pasando sobre :P
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
                frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (NPC-Hostil: " & MapData(tX, tY).NPCIndex & " - " & NpcData(MapData(tX, tY).NPCIndex).name & ")"
            Else
                frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (NPC: " & MapData(tX, tY).NPCIndex & " - " & NpcData(MapData(tX, tY).NPCIndex).name & ")"

            End If

        End If
        
        ' OBJs
        If MapData(tX, tY).OBJInfo.objindex > 0 Then
            frmConsola.StatTxt.Text = frmConsola.StatTxt.Text & " (Obj: " & MapData(tX, tY).OBJInfo.objindex & " - " & ObjData(MapData(tX, tY).OBJInfo.objindex).name & " - Cant.:" & MapData(tX, tY).OBJInfo.Amount & ")"

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
            If Deshacer Then modEdicion.Deshacer_Add "Quitar Todas las Capas (2/3)" ' Hago deshacer
            MapInfo.Changed = 1 'Set changed flag

            For LoopC = 2 To 3
                MapData(tX, tY).Graphic(LoopC).GrhIndex = 0
            Next LoopC

            Call DibujarMinimapa
            Exit Sub

        End If
    
        'Borrar "esta" Capa
        If frmSuperficies.cQuitarEnEstaCapa.value = True Then
            If Val(frmSuperficies.cCapas.Text) = 1 Then
                If MapData(tX, tY).Graphic(1).GrhIndex <> 1 Then
                    If Deshacer Then modEdicion.Deshacer_Add "Quitar Capa 1" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Graphic(1).GrhIndex = 1
                    Exit Sub

                End If

            ElseIf MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> 0 Then

                If Deshacer Then modEdicion.Deshacer_Add "Quitar Capa " & frmSuperficies.cCapas.Text  ' Hago deshacer
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
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar Superficie' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    aux = Val(frmSuperficies.cGrh.Text) + (((tY + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((tX + dX) Mod frmConfigSup.mAncho.Text)

                    If MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex <> aux Or MapData(tX, tY).Blocked <> frmMain.SelectPanel(2).value Then
                        MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)).GrhIndex = aux
                        InitGrh MapData(tX, tY).Graphic(Val(frmSuperficies.cCapas.Text)), aux

                    End If

                Else

                    If Deshacer Then modEdicion.Deshacer_Add "Insertar Auto-Completar Superficie' Hago deshacer"
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
                    If Deshacer Then modEdicion.Deshacer_Add "Quitar Superficie en Capa " & frmSuperficies.cCapas.Text ' Hago deshacer
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
            If MapData(tX, tY).Blocked <> 1 Then
                If Deshacer Then modEdicion.Deshacer_Add "Insertar Bloqueo" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Blocked = 1
                
            End If

        ElseIf frmBloqueos.cQuitarBloqueo.value = True Then

            If MapData(tX, tY).Blocked <> 0 Then
                If Deshacer Then modEdicion.Deshacer_Add "Quitar Bloqueo" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Blocked = 0

            End If

        End If
    
        '************** Place exit
        If frmTraslados.cInsertarTrans.value = True Then
            If Cfg_TrOBJ > 0 And Cfg_TrOBJ <= NumOBJs And frmTraslados.cInsertarTransOBJ.value = True Then
                If ObjData(Cfg_TrOBJ).ObjType = 19 Then
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar Objeto de Translado" ' Hago deshacer
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
                If Deshacer Then modEdicion.Deshacer_Add "Insertar Translado de Union Manual' Hago deshacer"
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

                If Deshacer Then modEdicion.Deshacer_Add "Insertar Translado" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).TileExit.Map = Val(frmTraslados.tTMapa.Text)
                MapData(tX, tY).TileExit.X = Val(frmTraslados.tTX.Text)
                MapData(tX, tY).TileExit.y = Val(frmTraslados.tTY.Text)

            End If

        ElseIf frmTraslados.cQuitarTrans.value = True Then

            If Deshacer Then modEdicion.Deshacer_Add "Quitar Translado" ' Hago deshacer
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
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar NPC" ' Hago deshacer
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
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar NPC Hostil' Hago deshacer"
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
                If Deshacer Then modEdicion.Deshacer_Add "Quitar NPC" ' Hago deshacer
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
                    If Deshacer Then modEdicion.Deshacer_Add "Insertar Objeto" ' Hago deshacer
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
                If Deshacer Then modEdicion.Deshacer_Add "Quitar Objeto" ' Hago deshacer
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
                If Deshacer Then modEdicion.Deshacer_Add "Insertar Trigger" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Trigger = TriggerBox

            End If

        ElseIf frmTriggers.cQuitarTrigger.value = True Then ' Quitar Trigger

            If MapData(tX, tY).Trigger <> 0 Then
                If Deshacer Then modEdicion.Deshacer_Add "Quitar Trigger" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).Trigger = 0

            End If
            
        End If
        
        ' ***************** Control de Funcion de Particulas *****************
        If frmParticulas.cSeleccionarParticula.value = True Then ' Insertar Particle
            If Deshacer Then modEdicion.Deshacer_Add "Insertar Particula" ' Hago deshacer
            MapInfo.Changed = 1
            MapData(tX, tY).Particle_Group_Index = General_Particle_Create(CLng(frmParticulas.cParticula.Text), tX, tY)
            MapData(tX, tY).Particle_Index = Val(frmParticulas.cParticula.Text)
                
        ElseIf frmParticulas.cQuitarParticula.value = True Then ' Quitar Particle

            If Deshacer Then modEdicion.Deshacer_Add "Quitar Particula" ' Hago deshacer
            MapInfo.Changed = 1 'Set changed flag
            MapData(tX, tY).Particle_Group_Index = 0
                
        End If
        
        ' ***************** Control de Funcion de Luces *****************
        If frmLuces.cInsertarLuz.value Then
            If Val(frmLuces.cRango = 0) Then Exit Sub
            Call mDx8_Luces.Create_Light_To_Map(tX, tY, frmLuces.cRango, Val(frmLuces.R), Val(frmLuces.G), Val(frmLuces.B))
            Call mDx8_Luces.LightRenderAll
                
            With MapData(tX, tY).Light
                .active = True
                .range = frmLuces.cRango
                .RGBCOLOR.a = 255
                .RGBCOLOR.R = Val(frmLuces.R)
                .RGBCOLOR.G = Val(frmLuces.G)
                .RGBCOLOR.B = Val(frmLuces.B)
                    
            End With
                
            If Deshacer Then modEdicion.Deshacer_Add "Insertar Luz" ' Hago deshacer
            MapInfo.Changed = 1 'Set changed flag
                
        ElseIf frmLuces.cQuitarLuz.value Then
            
            With MapData(tX, tY).Light
                .range = 0
                .RGBCOLOR.a = 255
                .RGBCOLOR.R = Val(frmLuces.R)
                .RGBCOLOR.G = Val(frmLuces.G)
                .RGBCOLOR.B = Val(frmLuces.B)
                    
            End With
    
            mDx8_Luces.Delete_Light_To_Map tX, tY
    
            If Deshacer Then modEdicion.Deshacer_Add "Quitar Luz" ' Hago deshacer
            MapInfo.Changed = 1 'Set changed flag
                
        End If

    End If
    
    Exit Sub

ClickEdit_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.ClickEdit", Erl)
    Resume Next
    
End Sub

''
' Bloquea los Bordes del Mapa
'

Public Sub Bloquear_Bordes(Optional ByVal ac As Byte = 1)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Bloquear_Bordes_Err
    
    Dim y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then Exit Sub
        
    modEdicion.Deshacer_Add "Bloquear los bordes" ' Hago deshacer
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
                MapData(X, y).Blocked = ac
            End If
        Next X
    Next y
    
    MapData(MinXBorder, MinYBorder).Blocked = ac
    MapData(MaxXBorder, MinYBorder).Blocked = ac
    MapData(MinXBorder, MaxYBorder).Blocked = ac
    MapData(MaxXBorder, MaxYBorder).Blocked = ac
    
    'Set changed flag
    MapInfo.Changed = 1

    Call DibujarMinimapa

    Exit Sub

Bloquear_Bordes_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Bloquear_Bordes", Erl)
    Resume Next
    
End Sub

''
' Modifica los bloqueos de todo mapa
'
' @param Valor Especifica el estado de Bloqueo que se asignara

Public Sub Bloqueo_Todo(ByVal Valor As Byte)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Bloqueo_Todo_Err

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    modEdicion.Deshacer_Add "Bloquear todo el mapa" ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            MapData(X, y).Blocked = Valor
        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    Exit Sub

Bloqueo_Todo_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Bloqueo_Todo", Erl)
    Resume Next
    
End Sub

''
' Elimita todos los translados del mapa
'

Public Sub Quitar_Translados()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo Quitar_Translados_Err

    If EditWarning Then Exit Sub

    modEdicion.Deshacer_Add "Quitar todos los Translados" ' Hago deshacer

    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).TileExit.Map <> 0 Then
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    
    Exit Sub

Quitar_Translados_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Quitar_Translados", Erl)
    Resume Next
    
End Sub

''
' Elimina todos los Triggers del mapa
'

Public Sub Quitar_Triggers()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_Triggers_Err
    
    If EditWarning Then Exit Sub

    modEdicion.Deshacer_Add "Quitar todos los Triggers" ' Hago deshacer

    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).Trigger > 0 Then
                MapData(X, y).Trigger = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    Exit Sub

Quitar_Triggers_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Quitar_Triggers", Erl)
    Resume Next
    
End Sub

''
' Elimita los NPCs del mapa
'
' @param Hostiles Indica si elimita solo hostiles o solo npcs no hostiles

Public Sub Quitar_NPCs(ByVal Hostiles As Boolean)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_NPCs_Err
    
    modEdicion.Deshacer_Add "Quitar todos los NPCs" & IIf(Hostiles = True, " Hostiles", "") ' Hago deshacer

    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).NPCIndex > 0 Then
                Call Char_Erase(MapData(X, y).CharIndex)
                MapData(X, y).NPCIndex = 0

            End If
        
        Next X
    Next y

    bRefreshRadar = True ' Radar

    'Set changed flag
    MapInfo.Changed = 1

    Exit Sub

Quitar_NPCs_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Quitar_NPCs", Erl)
    Resume Next
    
End Sub

''
' Elimita todos los Objetos del mapa
'

Public Sub Quitar_Objetos()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_Objetos_Err
    
    If EditWarning Then Exit Sub
    
    modEdicion.Deshacer_Add "Quitar todos los Objetos" ' Hago deshacer

    Dim y As Integer
    Dim X As Integer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, y).OBJInfo.objindex > 0 Then
                If MapData(X, y).Graphic(3).GrhIndex = MapData(X, y).ObjGrh.GrhIndex Then MapData(X, y).Graphic(3).GrhIndex = 0
                MapData(X, y).OBJInfo.objindex = 0
                MapData(X, y).OBJInfo.Amount = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    Exit Sub

Quitar_Objetos_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Quitar_Objetos", Erl)
    Resume Next
    
End Sub

''
' Elimita todo lo que se encuentre en los bordes del mapa
'

Public Sub Quitar_Bordes()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_Bordes_Err

    If EditWarning Then Exit Sub

    '*****************************************************************
    'Clears a border in a room with current GRH
    '*****************************************************************

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then Exit Sub

    modEdicion.Deshacer_Add "Quitar todos los Bordes" ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        
                MapData(X, y).Graphic(1).GrhIndex = 1
                InitGrh MapData(X, y).Graphic(1), 1
                MapData(X, y).Blocked = 0
            
                'Erase NPCs
                If MapData(X, y).NPCIndex > 0 Then
                    Char_Erase MapData(X, y).CharIndex
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
            
                ' Triggers
                MapData(X, y).Trigger = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    Exit Sub

Quitar_Bordes_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Quitar_Bordes", Erl)
    Resume Next
    
End Sub

''
' Elimita una capa completa del mapa
'
' @param Capa Especifica la capa

Public Sub Quitar_Capa(ByVal Capa As Byte)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Quitar_Capa_Err

    If EditWarning Then Exit Sub

    '*****************************************************************
    'Clears one layer
    '*****************************************************************

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then Exit Sub

    modEdicion.Deshacer_Add "Quitar Capa " & Capa ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If Capa = 1 Then
                MapData(X, y).Graphic(Capa).GrhIndex = 1
            Else
                MapData(X, y).Graphic(Capa).GrhIndex = 0

            End If

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    Exit Sub

Quitar_Capa_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Quitar_Capa", Erl)
    Resume Next
    
End Sub

''
' Borra todo el Mapa menos los Triggers
'

Public Sub Borrar_Mapa()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Borrar_Mapa_Err
    
    If EditWarning Then Exit Sub

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then Exit Sub

    modEdicion.Deshacer_Add "Borrar todo el mapa" ' Hago deshacer
    
    'Borramos las particulas activas en el mapa.
    Call Particle_Group_Remove_All
        
    'Borramos todas las luces
    Call LightRemoveAll

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            MapData(X, y).Graphic(1).GrhIndex = 1
            'Change blockes status
            MapData(X, y).Blocked = 0

            'Erase layer 2 and 3
            MapData(X, y).Graphic(2).GrhIndex = 0
            MapData(X, y).Graphic(3).GrhIndex = 0
            MapData(X, y).Graphic(4).GrhIndex = 0

            'Erase NPCs
            If MapData(X, y).NPCIndex > 0 Then
                Char_Erase MapData(X, y).CharIndex
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
        
            InitGrh MapData(X, y).Graphic(1), 1
            
            MapData(X, y).Light.active = False
            MapData(X, y).Particle_Group_Index = 0
            
            MapData(X, y).Trigger = 0

        Next X
    Next y

    'Set changed flag
    MapInfo.Changed = 1

    Exit Sub

Borrar_Mapa_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.Borrar_Mapa", Erl)
    Resume Next
    
End Sub

Public Sub CopiarSeleccion(Optional ByVal Borde As Boolean = False)
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    'podria usar copy mem , pero por las dudas no XD
    
    On Error GoTo CopiarSeleccion_Err
    
    Dim X As Integer
    Dim y As Integer

    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock

    If Not Borde Then
        For X = 0 To SeleccionAncho - 1
            For y = 0 To SeleccionAlto - 1
                SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
            Next
        Next
        
    Else
        For X = 0 To SeleccionAncho - 1
            For y = 0 To SeleccionAlto - 1
                SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
            Next
        Next
    End If
    MapInfo.Changed = 1

    Exit Sub

CopiarSeleccion_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.CopiarSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub CortarSeleccion()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    
    On Error GoTo CortarSeleccion_Err
    
    CopiarSeleccion
    Dim X     As Integer
    Dim y     As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            MapData(X + SeleccionIX, y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
    Call DibujarMinimapa

    Exit Sub

CortarSeleccion_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.CortarSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub BlockearSeleccion()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    
    On Error GoTo BlockearSeleccion_Err
    
    Dim X     As Integer
    Dim y     As Integer
    Dim Vacio As MapBlock
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1

            If MapData(X + SeleccionIX, y + SeleccionIY).Blocked > 0 Then
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 0
            Else
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = &HF

            End If

        Next
    Next
    Seleccionando = False

    Exit Sub

BlockearSeleccion_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.BlockearSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub AccionSeleccion()

    Working = True

    '*************************************************
    'Author: Loopzera
    'Last modified: 21/11/07
    '*************************************************
    On Error Resume Next

    Dim X As Integer
    Dim y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + y, False
        Next y
    Next X
    
    Seleccionando = False
    
    If frmConsola.Visible Then _
        Call AddtoRichTextBox(frmConsola.StatTxt, "Tarea finalizada.", 255, 0, 0, False, True, False)
        
    Working = False
    Call DibujarMinimapa

End Sub

Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    'podria usar copy mem , pero por las dudas no XD
    
    On Error GoTo PegarSeleccion_Err
    
    Static UltimoX As Integer
    Static UltimoY As Integer
    'If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    
    Debug.Print SobreX
    Debug.Print SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1

            If y + SobreY > 100 Then Exit For
            If X + SobreX > 100 Then Exit For
            'NO copia tile exit - LADDER
  
            DeSeleccionMap(X, y).TileExit.Map = MapData(X + SobreX, y + SobreY).TileExit.Map
            DeSeleccionMap(X, y).TileExit.X = MapData(X + SobreX, y + SobreY).TileExit.X
            DeSeleccionMap(X, y).TileExit.y = MapData(X + SobreX, y + SobreY).TileExit.y
            DeSeleccionMap(X, y) = MapData(X + SobreX, y + SobreY)

            MapData(X + SobreX, y + SobreY).NPCIndex = 0 'NO copia NPC
        Next
    Next

    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1

            If y + SobreY > 100 Then Exit For
            If X + SobreX > 100 Then Exit For
            'NO copia tile exit - LADDER
            SeleccionMap(X, y).TileExit.Map = MapData(X + SobreX, y + SobreY).TileExit.Map
            SeleccionMap(X, y).TileExit.X = MapData(X + SobreX, y + SobreY).TileExit.X
            SeleccionMap(X, y).TileExit.y = MapData(X + SobreX, y + SobreY).TileExit.y
        
            MapData(X + SobreX, y + SobreY) = SeleccionMap(X, y)
            MapData(X + SobreX, y + SobreY).NPCIndex = 0 'NO copia NPC

        Next
    Next
    Seleccionando = False
    Call DibujarMinimapa
    
    Exit Sub

PegarSeleccion_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.PegarSeleccion", Erl)
    Resume Next
    
End Sub

Public Sub DePegar()
    '*************************************************
    'Author: Loopzer
    'Last modified: 21/11/07
    '*************************************************
    
    On Error GoTo DePegar_Err
    
    Dim X As Integer
    Dim y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
            MapData(X + DeSeleccionOX, y + DeSeleccionOY) = DeSeleccionMap(X, y)
        Next
    Next
    
    Exit Sub

DePegar_Err:
    Call LogError(Err.Number, Err.Description, "modEdicion.DePegar", Erl)
    Resume Next
    
End Sub
