Attribute VB_Name = "modMapIO"
Option Explicit

'********************************
'Load Map with .CSM format
'********************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    r As Integer
    g As Integer
    b As Integer
    range As Byte
    X As Integer
    y As Integer
End Type

Private Type tDatosParticulas
    X As Integer
    y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    y As Integer
    objindex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    ambient As String
    lvlMinimo As String
    LuzBase As Long
    version As Long
    NoTirarItems As Boolean
End Type

Public MapSize As tMapSize
Private MapDat As tMapDat
'********************************
'END - Load Map with .CSM format
'********************************


''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByVal Path As String)
    
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2023
    '*************************************************
    
    On Error GoTo AbrirMapa_Err

    Call MapaCSM_Cargar(Path)
    
    Exit Sub

AbrirMapa_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.AbrirMapa", Erl)
    Resume Next
    
End Sub

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************

    frmMain.Dialog.CancelError = True
On Error GoTo ErrHandler
    
    If LenB(Path) = 0 Then
        frmMain.ObtenerNombreArchivo True
        Path = frmMain.Dialog.FileName
        If LenB(Path) = 0 Then Exit Sub
    End If
    
    If frmMain.Dialog.FilterIndex = 1 Then _
        Call GuardarMapa_CSM(Path)

ErrHandler:
End Sub

''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 21/05/06
    '*************************************************

    On Error Resume Next

    Dim loopc As Integer
    Dim y     As Integer
    Dim X     As Integer

    frmMain.mnuReAbrirMapa.Enabled = False

    MapaCargado = False

    For loopc = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(loopc).Enabled = False
    Next

    frmMain.MousePointer = 11

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            ' Capa 1
            MapData(X, y).Graphic(1).GrhIndex = 1
        
            ' Bloqueos
            MapData(X, y).Blocked = 0

            ' Capas 2, 3 y 4
            MapData(X, y).Graphic(2).GrhIndex = 0
            MapData(X, y).Graphic(3).GrhIndex = 0
            MapData(X, y).Graphic(4).GrhIndex = 0

            ' NPCs
            If MapData(X, y).CharIndex > 0 Then
                Call Char_Erase(MapData(X, y).CharIndex)
                MapData(X, y).NPCIndex = 0

            End If

            ' OBJs
            MapData(X, y).OBJInfo.objindex = 0
            MapData(X, y).OBJInfo.Amount = 0
            MapData(X, y).ObjGrh.GrhIndex = 0

            ' Translados
            MapData(X, y).TileExit.Map = 0
            MapData(X, y).TileExit.X = 0
            MapData(X, y).TileExit.y = 0
        
            ' Triggers
            MapData(X, y).Trigger = 0
        
            InitGrh MapData(X, y).Graphic(1), 1
        Next X
    Next y

    MapInfo.MapVersion = 0
    MapInfo.Name = "Nuevo Mapa"
    MapInfo.Music = 0
    MapInfo.Pk = True
    MapInfo.MagiaSinEfecto = 0
    MapInfo.InviSinEfecto = 0
    MapInfo.ResuSinEfecto = 0
    MapInfo.Terreno = "BOSQUE"
    MapInfo.Zona = "CAMPO"
    MapInfo.Restringir = 0

    'Call MapInfo_Actualizar
    Call DibujarMinimapa

    bRefreshRadar = True ' Radar

    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0

    MapaCargado = True
    EngineRun = True

    'FrmMain.SetFocus

End Sub

Private Sub MapaCSM_Cargar(ByVal RutaMapa As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 27/04/2023
    '***************************************************
    
    On Error GoTo MapaCSM_Cargar_Err
    
    Dim fh              As Integer
    Dim MH              As tMapHeader
    Dim Blqs()          As tDatosBloqueados
    Dim L1()            As Long
    Dim L2()            As tDatosGrh
    Dim L3()            As tDatosGrh
    Dim L4()            As tDatosGrh
    Dim Triggers()      As tDatosTrigger
    Dim Luces()         As tDatosLuces
    Dim Particulas()    As tDatosParticulas
    Dim Objetos()       As tDatosObjs
    Dim NPCs()          As tDatosNPC
    Dim TEs()           As tDatosTE
    Dim MapSize         As tMapSize
    Dim MapDat          As tMapDat
    Dim npcfile         As String
    Dim i               As Long
    Dim j               As Long
    Dim LaCabecera      As tCabecera
    
    Dim Body         As Integer
    Dim Head         As Integer
    Dim Heading      As Byte
    
    fh = FreeFile
    
    Open RutaMapa For Binary Access Read As fh
    
        Get #fh, , LaCabecera
    
        Get #fh, , MH
        Get #fh, , MapSize
        Get #fh, , MapDat
        
        ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
        
        Get #fh, , L1
        
        With MH
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
                For i = 1 To .NumeroBloqueados
                    MapData(Blqs(i).X, Blqs(i).y).Blocked = 1
                Next i
            End If
            
            If .NumeroLayers(2) > 0 Then
                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2
                For i = 1 To .NumeroLayers(2)
                    Call InitGrh(MapData(L2(i).X, L2(i).y).Graphic(2), L2(i).GrhIndex)
                Next i
            End If
            
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3
                For i = 1 To .NumeroLayers(3)
                    Call InitGrh(MapData(L3(i).X, L3(i).y).Graphic(3), L3(i).GrhIndex)
                Next i
            End If
            
            If .NumeroLayers(4) > 0 Then
                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4
                For i = 1 To .NumeroLayers(4)
                    Call InitGrh(MapData(L4(i).X, L4(i).y).Graphic(4), L4(i).GrhIndex)
                Next i
            End If
            
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
                For i = 1 To .NumeroTriggers
                    MapData(Triggers(i).X, Triggers(i).y).Trigger = Triggers(i).Trigger
                Next i
            End If
            
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas
                
                For i = 1 To .NumeroParticulas
    
                    With Particulas(i)
                    
                        'MapData(.X, .Y).Particle_Group_Index = General_Particle_Create(.Particula, .X, .Y)
    
                    End With
    
                Next i
            End If
            
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Get #fh, , Luces
                
                For i = 1 To .NumeroLuces
    
                    With Luces(i)
    
                        Call Create_Light_To_Map(.X, .y, .range, .r, .g, .b)
    
                    End With
    
                Next i
            
                Call LightRenderAll
            End If
            
            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos
                For i = 1 To .NumeroOBJs
                    MapData(Objetos(i).X, Objetos(i).y).OBJInfo.objindex = Objetos(i).objindex
                    MapData(Objetos(i).X, Objetos(i).y).OBJInfo.Amount = Objetos(i).ObjAmmount
                Next i
            End If
                
            If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs

                For i = 1 To .NumeroNPCs
                
                    MapData(NPCs(i).X, NPCs(i).y).NPCIndex = NPCs(i).NPCIndex
    
                    If MapData(NPCs(i).X, NPCs(i).y).NPCIndex > 0 Then
                        Body = NpcData(MapData(NPCs(i).X, NPCs(i).y).NPCIndex).Body
                        Head = NpcData(MapData(NPCs(i).X, NPCs(i).y).NPCIndex).Head
                        Heading = NpcData(MapData(NPCs(i).X, NPCs(i).y).NPCIndex).Heading
                        Call Char_Make(NextOpenChar(), Body, Head, Heading, NPCs(i).X, NPCs(i).y, 2, 2, 2, 0, 0)
                    End If

                Next i

            End If
                
            If .NumeroTE > 0 Then
                ReDim TEs(1 To .NumeroTE)
                Get #fh, , TEs
                For i = 1 To .NumeroTE
                    MapData(TEs(i).X, TEs(i).y).TileExit.Map = TEs(i).DestM
                    MapData(TEs(i).X, TEs(i).y).TileExit.X = TEs(i).DestX
                    MapData(TEs(i).X, TEs(i).y).TileExit.y = TEs(i).DestY
                Next i
            End If
            
        End With
    
    Close fh
    
        
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                Call InitGrh(MapData(i, j).Graphic(1), L1(i, j))
            End If
        Next i
    Next j
    
    Call Actualizar_Estado(Estado_Actual_Date)
    Call DibujarMinimapa
    
    bRefreshRadar = True ' Radar
    
    'Set changed flag
    MapInfo.Changed = 0
    
    Call Pestanias(RutaMapa)
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True
    
    Exit Sub

MapaCSM_Cargar_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.MapaCSM_Cargar", Erl)
    Resume Next
End Sub

Private Sub GuardarMapa_CSM(ByVal RutaMapa As String)

End Sub

''
' Calcula la orden de Pestañas
'
' @param Map Especifica path del mapa

Public Sub Pestanias(ByVal Map As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo Pestañas_Err
    
    Dim loopc As Integer

    For loopc = Len(Map) To 1 Step -1

        If mid(Map, loopc, 1) = "\" Then
            PATH_Save = Left(Map, loopc)
            Exit For

        End If

    Next
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))

    For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

        If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
            NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
            NameMap_Save = Left(Map, loopc)
            Exit For

        End If

    Next

    For loopc = (NumMap_Save - 4) To (NumMap_Save + 6)

        If FileExist(PATH_Save & NameMap_Save & loopc & ".csm", vbArchive) = True Then
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
        Else
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False

        End If

    Next
    
    Exit Sub

Pestañas_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.Pestanias", Erl)

    Resume Next
    
End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo DeseaGuardarMapa_Err
    

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            GuardarMapa Path

        End If

    End If

    
    Exit Sub

DeseaGuardarMapa_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.DeseaGuardarMapa", Erl)
    Resume Next
    
End Sub
