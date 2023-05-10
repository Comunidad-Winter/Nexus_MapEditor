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
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    R As Integer
    G As Integer
    B As Integer
    range As Byte
    X As Integer
    Y As Integer
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    objindex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
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

Public MapFormat As String
Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByVal Path As String, Optional ByVal MapaTipo As Byte)
    
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2023
    '*************************************************
    
    On Error GoTo AbrirMapa_Err

    Select Case MapaTipo
    
        Case 0
            Call MapaCSM_Cargar(Path)
            
        Case 1
            Call MapaAO_Cargar(Path)
            
        Case 2
            Call MapaAO_Cargar(Path, False)
            
        Case 3
            Call MapaMCL_Cargar(Path)
    
    End Select
    
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
        Call MapaCSM_Guardar(Path)

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

    Dim LoopC As Integer

    Dim Y     As Integer

    Dim X     As Integer

    bAutoGuardarMapaCount = 0

    frmMain.mnuReAbrirMapa.Enabled = False
    frmMain.TimAutoGuardarMapa.Enabled = False

    MapaCargado = False

    For LoopC = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(LoopC).Enabled = False
    Next

    frmMain.MousePointer = 11
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            With MapData(X, Y)
                ' Capa 1
                .Graphic(1).GrhIndex = 1
            
                ' Bloqueos
                .Blocked = 0
    
                ' Capas 2, 3 y 4
                .Graphic(2).GrhIndex = 0
                .Graphic(3).GrhIndex = 0
                .Graphic(4).GrhIndex = 0
    
                ' NPCs
                If .CharIndex > 0 Then
                    Call Char_Erase(.CharIndex)
                    .NPCIndex = 0
    
                End If
    
                ' OBJs
                .OBJInfo.objindex = 0
                .OBJInfo.Amount = 0
                Call GrhUninitialize(.ObjGrh)
    
                ' Translados
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            
                ' Triggers
                .Trigger = 0
            
                InitGrh .Graphic(1), 1
                
                .Light.active = False
                .Light.range = 0
                .Light.map_x = 0
                .Light.map_y = 0
                .Light.RGBCOLOR.a = 0
                .Light.RGBCOLOR.R = 0
                .Light.RGBCOLOR.G = 0
                .Light.RGBCOLOR.B = 0
                
            End With
            
        Next X
    Next Y
    
    'Limpieza adicional del mapa. PARCHE: Solucion a bug de clones.
    Call Char_CleanAll
    
    'Borramos las particulas activas en el mapa.
    Call Particle_Group_Remove_All
        
    'Borramos todas las luces
    Call LightRemoveAll

    With MapInfo
    
        .MapVersion = 0
        .name = "Nuevo Mapa"
        .Music = 0
        .ambient = 0
        .Pk = True
        .MagiaSinEfecto = 0
        .InviSinEfecto = 0
        .ResuSinEfecto = 0
        .Terreno = "BOSQUE"
        .Zona = "CAMPO"
        .Restringir = 0
        .LuzBase = 0
    
    End With
    
    MapFormat = ".csm"

    Call MapInfo_Actualizar
    Call DibujarMinimapa
    bRefreshRadar = True ' Radar
    
    ' Vacio deshacer
    Call modEdicion.Deshacer_Clear
    
    Estado_Actual = Estados(e_estados.MedioDia)
    Call Actualizar_Estado

    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0

    MapaCargado = True
    EngineRun = True

    'FrmMain.SetFocus

End Sub

Public Sub MapaCSM_Cargar(ByVal RutaMapa As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 27/04/2023
    '***************************************************
    
    On Error GoTo MapaCSM_Cargar_Err
    
    Dim fh           As Integer
    Dim MH           As tMapHeader
    Dim Blqs()       As tDatosBloqueados
    Dim L1()         As Long
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh
    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE
    Dim MapSize      As tMapSize
    Dim npcfile      As String
    Dim i            As Long
    Dim j            As Long
    Dim LaCabecera   As tCabecera
    Dim Body         As Integer
    Dim Head         As Integer
    Dim Heading      As Byte
    
    fh = FreeFile
    
    MapFormat = ".csm"
    
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
                MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i

        End If
            
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                Call InitGrh(MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex)
            Next i

        End If
            
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                Call InitGrh(MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex)
            Next i

        End If
            
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                Call InitGrh(MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex)
            Next i

        End If
            
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i

        End If
            
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
                
            For i = 1 To .NumeroParticulas
    
                With Particulas(i)
                    
                    MapData(.X, .Y).Particle_Group_Index = General_Particle_Create(.Particula, .X, .Y)
    
                End With
    
            Next i

        End If
            
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
                
            For i = 1 To .NumeroLuces
    
                With MapData(Luces(i).X, Luces(i).Y)
                    .Light.range = Luces(i).range
                    .Light.RGBCOLOR.a = 255
                    .Light.RGBCOLOR.R = Luces(i).R
                    .Light.RGBCOLOR.G = Luces(i).G
                    .Light.RGBCOLOR.B = Luces(i).B
                    .Light.active = True

                End With
    
                With Luces(i)
                    Call Create_Light_To_Map(.X, .Y, .range, .R, .G, .B)

                End With
    
            Next i
            
            Call LightRenderAll

        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos

            For i = 1 To .NumeroOBJs
                MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex = Objetos(i).objindex
                MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount

                If MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex > NumOBJs Then
                    InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, 20299
                Else
                    InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, ObjData(MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex).GrhIndex

                End If

            Next i

        End If
                
        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs

            For i = 1 To .NumeroNPCs
                
                MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
    
                If MapData(NPCs(i).X, NPCs(i).Y).NPCIndex > 0 Then
                    Body = NpcData(MapData(NPCs(i).X, NPCs(i).Y).NPCIndex).Body
                    Head = NpcData(MapData(NPCs(i).X, NPCs(i).Y).NPCIndex).Head
                    Heading = NpcData(MapData(NPCs(i).X, NPCs(i).Y).NPCIndex).Heading
                    Call Char_Make(NextOpenChar(), Body, Head, Heading, NPCs(i).X, NPCs(i).Y, 2, 2, 2, 0, 0)

                End If

            Next i

        End If
                
        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs

            For i = 1 To .NumeroTE
                MapData(TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                MapData(TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                MapData(TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
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
    
    Call CSMInfoCargar
    Call Actualizar_Estado
    Call DibujarMinimapa
    
    bRefreshRadar = True ' Radar
    
    ' Vacio deshacer
    modEdicion.Deshacer_Clear
    
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

Public Function MapaCSM_Guardar(ByVal RutaMapa As String) As Boolean
    '***************************************************
    'Author: Lorwik
    'Last Modification: 14/03/2021
    '***************************************************
    
On Error GoTo ErrorHandler

    Dim fh As Integer
    Dim MH As tMapHeader
    Dim Blqs() As tDatosBloqueados
    Dim L1() As Long
    Dim L2() As tDatosGrh
    Dim L3() As tDatosGrh
    Dim L4() As tDatosGrh
    Dim Triggers() As tDatosTrigger
    Dim Luces() As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos() As tDatosObjs
    Dim NPCs() As tDatosNPC
    Dim TEs() As tDatosTE
    
    Dim i As Integer
    Dim j As Integer
    
    If NoSobreescribir = False Then
        If FileExist(RutaMapa, vbNormal) = True Then
            If MsgBox("¿Desea sobrescribir " & RutaMapa & "?", vbCritical + vbYesNo) = vbNo Then
                Exit Function
            Else
                'Kill MapRoute
            End If
        End If
    End If
    
    frmMain.MousePointer = 11
    MapSize.XMax = XMaxMapSize
    MapSize.YMax = YMaxMapSize
    MapSize.XMin = XMinMapSize
    MapSize.YMin = YMinMapSize
    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax)
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            With MapData(i, j)
                If .Blocked Then
                    MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                    ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                    Blqs(MH.NumeroBloqueados).X = i
                    Blqs(MH.NumeroBloqueados).Y = j
                End If
                
                L1(i, j) = .Graphic(1).GrhIndex
                
                If .Graphic(2).GrhIndex > 0 Then
                    MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                    ReDim Preserve L2(1 To MH.NumeroLayers(2))
                    L2(MH.NumeroLayers(2)).X = i
                    L2(MH.NumeroLayers(2)).Y = j
                    L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2).GrhIndex
                End If
                
                If .Graphic(3).GrhIndex > 0 Then
                    MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                    ReDim Preserve L3(1 To MH.NumeroLayers(3))
                    L3(MH.NumeroLayers(3)).X = i
                    L3(MH.NumeroLayers(3)).Y = j
                    L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3).GrhIndex
                End If
                
                If .Graphic(4).GrhIndex > 0 Then
                    MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                    ReDim Preserve L4(1 To MH.NumeroLayers(4))
                    L4(MH.NumeroLayers(4)).X = i
                    L4(MH.NumeroLayers(4)).Y = j
                    L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4).GrhIndex
                End If
                
                If .Trigger > 0 Then
                    MH.NumeroTriggers = MH.NumeroTriggers + 1
                    ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                    Triggers(MH.NumeroTriggers).X = i
                    Triggers(MH.NumeroTriggers).Y = j
                    Triggers(MH.NumeroTriggers).Trigger = .Trigger
                End If
                
                If .Particle_Group_Index > 0 Then
                    MH.NumeroParticulas = MH.NumeroParticulas + 1
                    ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                    Particulas(MH.NumeroParticulas).X = i
                    Particulas(MH.NumeroParticulas).Y = j
                    Particulas(MH.NumeroParticulas).Particula = .Particle_Group_Index
    
                End If
               
               '¿Hay luz activa en este punto?
                If .Light.active Then
                    MH.NumeroLuces = MH.NumeroLuces + 1
                    ReDim Preserve Luces(1 To MH.NumeroLuces)

                    Luces(MH.NumeroLuces).R = .Light.RGBCOLOR.R
                    Luces(MH.NumeroLuces).G = .Light.RGBCOLOR.G
                    Luces(MH.NumeroLuces).B = .Light.RGBCOLOR.B
                    Luces(MH.NumeroLuces).range = .Light.range
                    Luces(MH.NumeroLuces).X = .Light.map_x
                    Luces(MH.NumeroLuces).Y = .Light.map_y
                End If
                
                If .OBJInfo.objindex > 0 Then
                    MH.NumeroOBJs = MH.NumeroOBJs + 1
                    ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                    Objetos(MH.NumeroOBJs).objindex = .OBJInfo.objindex
                    Objetos(MH.NumeroOBJs).ObjAmmount = .OBJInfo.Amount
                    Objetos(MH.NumeroOBJs).X = i
                    Objetos(MH.NumeroOBJs).Y = j
                End If
                
                If .NPCIndex > 0 Then
                    MH.NumeroNPCs = MH.NumeroNPCs + 1
                    ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                    NPCs(MH.NumeroNPCs).NPCIndex = .NPCIndex
                    NPCs(MH.NumeroNPCs).X = i
                    NPCs(MH.NumeroNPCs).Y = j
                End If
                
                If .TileExit.Map > 0 Then
                    MH.NumeroTE = MH.NumeroTE + 1
                    ReDim Preserve TEs(1 To MH.NumeroTE)
                    TEs(MH.NumeroTE).DestM = .TileExit.Map
                    TEs(MH.NumeroTE).DestX = .TileExit.X
                    TEs(MH.NumeroTE).DestY = .TileExit.Y
                    TEs(MH.NumeroTE).X = i
                    TEs(MH.NumeroTE).Y = j
                End If
            End With
        Next i
    Next j
    
    Call CSMInfoSave
              
    fh = FreeFile
    Open RutaMapa For Binary As fh
        
        Put #fh, , MiCabecera
        
        Put #fh, , MH
        Put #fh, , MapSize
        Put #fh, , MapDat
        Put #fh, , L1
    
        With MH
            If .NumeroBloqueados > 0 Then _
                Put #fh, , Blqs
            If .NumeroLayers(2) > 0 Then _
                Put #fh, , L2
            If .NumeroLayers(3) > 0 Then _
                Put #fh, , L3
            If .NumeroLayers(4) > 0 Then _
                Put #fh, , L4
            If .NumeroTriggers > 0 Then _
                Put #fh, , Triggers
            If .NumeroParticulas > 0 Then _
                Put #fh, , Particulas
            If .NumeroLuces > 0 Then _
                Put #fh, , Luces
            If .NumeroOBJs > 0 Then _
                Put #fh, , Objetos
            If .NumeroNPCs > 0 Then _
                Put #fh, , NPCs
            If .NumeroTE > 0 Then _
                Put #fh, , TEs
        End With
    
    Close fh
    
    Call Pestanias(RutaMapa)
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    NoSobreescribir = False
    
    MapaCSM_Guardar = True
    
    If frmConsola.Visible Then _
        Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa " & RutaMapa & " guardado...", 0, 255, 0)
    Exit Function

ErrorHandler:
    If fh <> 0 Then Close fh
End Function

''
' Abrir Mapa con el formato de AO
'
' @param Map Especifica el Path del mapa

Public Sub MapaAO_Cargar(ByVal Map As String, Optional ByVal EsInteger As Boolean = True)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim LoopC       As Integer

    Dim tempint     As Integer

    Dim Body        As Integer

    Dim Head        As Integer

    Dim Heading     As Byte

    Dim Y           As Integer

    Dim X           As Integer

    Dim i           As Byte

    Dim ByFlags     As Byte

    Dim FreeFileMap As Long

    Dim FreeFileInf As Long

    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
    
    MapFormat = ".map"
    
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = Left$(Map, Len(Map) - 4)
    Map = Map & ".inf"
    
    FreeFileInf = FreeFile
    Open Map For Binary As FreeFileInf
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , tempint
    Get FreeFileMap, , tempint
    Get FreeFileMap, , tempint
    Get FreeFileMap, , tempint
    
    'Cabecera inf
    Get FreeFileInf, , tempint
    Get FreeFileInf, , tempint
    Get FreeFileInf, , tempint
    Get FreeFileInf, , tempint
    Get FreeFileInf, , tempint

    'Load arrays
    For Y = YMinMapSize To 100
        For X = XMinMapSize To 100
            
            With MapData(X, Y)
            
                Get FreeFileMap, , ByFlags
                .Blocked = (ByFlags And 1)
            
               'Layer 1
                If EsInteger Then
                    Get FreeFileMap, , .Graphic(1).GrhIndexInt
                    Call InitGrh(.Graphic(1), .Graphic(1).GrhIndexInt)
                Else
                    Get FreeFileMap, , .Graphic(1).GrhIndex
                    Call InitGrh(.Graphic(1), .Graphic(1).GrhIndex)
                End If
            
                'Layer 2 used?
                If ByFlags And 2 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(2).GrhIndexInt
                        Call InitGrh(.Graphic(2), .Graphic(2).GrhIndexInt)
                    Else
                        Get FreeFileMap, , .Graphic(2).GrhIndex
                        Call InitGrh(.Graphic(2), .Graphic(2).GrhIndex)
                    End If
 
                Else
                
                    .Graphic(2).GrhIndex = 0
                    
                End If
                
                'Layer 3 used?
                If ByFlags And 4 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(3).GrhIndexInt
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndexInt)
                    Else
                        Get FreeFileMap, , .Graphic(3).GrhIndex
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndex)
                    End If

                Else
                
                    .Graphic(3).GrhIndex = 0
                    
                End If
                
                'Layer 4 used?
                If ByFlags And 8 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(4).GrhIndexInt
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndexInt)
                    Else
                        Get FreeFileMap, , .Graphic(4).GrhIndex
                        Call InitGrh(.Graphic(4), .Graphic(4).GrhIndex)
                    End If

                Else
                    
                    .Graphic(4).GrhIndex = 0

                End If

                'Trigger used?
                If ByFlags And 16 Then
                    Get FreeFileMap, , .Trigger
                Else
                    .Trigger = 0

                End If

                'Cargamos el archivo ".INF"
                Get FreeFileInf, , ByFlags

                If ByFlags And 1 Then

                    With .TileExit

                        Get FreeFileInf, , .Map
                        Get FreeFileInf, , .X
                        Get FreeFileInf, , .Y

                    End With

                End If

                If ByFlags And 2 Then

                    'Get and make NPC
                    Get FreeFileInf, , .NPCIndex

                    If .NPCIndex < 0 Then
                        .NPCIndex = 0
                    Else
                        Body = NpcData(.NPCIndex).Body
                        Head = NpcData(.NPCIndex).Head
                        Heading = NpcData(.NPCIndex).Heading
                        Call Char_Make(NextOpenChar(), Body, Head, Heading, X, Y, 0, 0, 0, 0, 0)

                    End If

                End If

                If ByFlags And 4 Then

                    'Get and make Object
                    Get FreeFileInf, , .OBJInfo.objindex
                    Get FreeFileInf, , .OBJInfo.Amount

                    If .OBJInfo.objindex > 0 Then
                        Call InitGrh(.ObjGrh, ObjData(.OBJInfo.objindex).GrhIndex)

                    End If

                End If
            
            End With
    
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
    
    Call Pestanias(Map)
    
    Map = Left$(Map, Len(Map) - 4) & ".dat"
    
    Call MapInfoAO_Cargar(Map)
    
    With frmMain
    
        frmMapInfo.txtMapVersion.Text = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        Call modEdicion.Deshacer_Clear
        
        'Change mouse icon
        .MousePointer = 0
    
    End With
    
    MapaCargado = True

End Sub

''
' Guardar Mapa con el formato V2
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaAO_Guardar(ByVal SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim LoopC       As Long
    Dim tempint     As Integer
    Dim Y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte

    If FileExist(SaveAs, vbNormal) = True Then
        
        If NoSobreescribir = False Then
            If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            Else
                Call Kill(SaveAs)
            End If
        
        Else
            Call Kill(SaveAs)
            
        End If
        
    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Call Kill(Left$(SaveAs, Len(SaveAs) - 4) & ".inf")
    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"

    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMapInfo.txtMapVersion.Text < 32767 Then
        frmMapInfo.txtMapVersion.Text = frmMapInfo.txtMapVersion + 1
    End If

    Put FreeFileMap, , CInt(frmMapInfo.txtMapVersion.Text)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , tempint
    Put FreeFileMap, , tempint
    Put FreeFileMap, , tempint
    Put FreeFileMap, , tempint
    
    'inf Header
    Put FreeFileInf, , tempint
    Put FreeFileInf, , tempint
    Put FreeFileInf, , tempint
    Put FreeFileInf, , tempint
    Put FreeFileInf, , tempint
    
    'Write .map file
    For Y = YMinMapSize To 100
        For X = XMinMapSize To 100
            
            With MapData(X, Y)
            
                ByFlags = 0
                
                If .Blocked = 1 Then ByFlags = ByFlags Or 1
                
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8

                If .Trigger Then ByFlags = ByFlags Or 16
                    
                Put FreeFileMap, , ByFlags
                    
                Put FreeFileMap, , .Graphic(1).GrhIndex
                
                For LoopC = 2 To 4
                    
                    If .Graphic(LoopC).GrhIndex Then Put FreeFileMap, , .Graphic(LoopC).GrhIndex

                Next LoopC
                    
                If .Trigger Then Put FreeFileMap, , .Trigger
                
                'Escribimos el archivo ".INF"
                ByFlags = 0
                    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                If .NPCIndex Then ByFlags = ByFlags Or 2
                
                If .OBJInfo.objindex Then ByFlags = ByFlags Or 4
                    
                Put FreeFileInf, , ByFlags
                    
                If .TileExit.Map Then
                    Put FreeFileInf, , .TileExit.Map
                    Put FreeFileInf, , .TileExit.X
                    Put FreeFileInf, , .TileExit.Y
                End If
                    
                If .NPCIndex Then
                    Put FreeFileInf, , CInt(.NPCIndex)
                End If
                    
                If .OBJInfo.objindex Then
                    Put FreeFileInf, , .OBJInfo.objindex
                    Put FreeFileInf, , .OBJInfo.Amount
                End If
            
            End With
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestanias(SaveAs)

    'write .dat file
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    Call MapInfoAO_Guardar(SaveAs)

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    NoSobreescribir = False

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description

End Sub

Public Sub MapaMCL_Cargar(ByVal RutaMapa As String)

    On Error Resume Next

    Dim LoopC    As Integer

    Dim Y        As Integer

    Dim X        As Integer

    Dim tempint  As Integer

    Dim InfoTile As Byte

    Dim i        As Integer

    Open RutaMapa For Binary As #1
    Seek #1, 1

    Get #1, , MapInfo.MapVersion

    For Y = 1 To 100
        For X = 1 To 100

            Get #1, , InfoTile
        
            MapData(X, Y).Blocked = (InfoTile And 1)
        
            Get #1, , MapData(X, Y).Graphic(1).GrhIndexInt
        
            For i = 2 To 4

                If InfoTile And (2 ^ (i - 1)) Then
                    Get #1, , MapData(X, Y).Graphic(i).GrhIndexInt
                    Call InitGrh(MapData(X, Y).Graphic(i), MapData(X, Y).Graphic(i).GrhIndexInt)
                    Else: MapData(X, Y).Graphic(i).GrhIndex = 0

                End If

            Next
        
            MapData(X, Y).Trigger = 0
        
            For i = 4 To 6

                If (InfoTile And 2 ^ i) Then MapData(X, Y).Trigger = MapData(X, Y).Trigger Or 2 ^ (i - 4)
            Next
        
            Call InitGrh(MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndexInt)
    
            If MapData(X, Y).CharIndex > 0 Then Call Char_Erase(MapData(X, Y).CharIndex)
            MapData(X, Y).ObjGrh.GrhIndex = 0
        Next X
    Next Y

    Close #1

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
    
    Dim LoopC As Integer

    For LoopC = Len(Map) To 1 Step -1

        If mid(Map, LoopC, 1) = "\" Then
            PATH_Save = Left(Map, LoopC)
            Exit For

        End If

    Next
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))

    For LoopC = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

        If IsNumeric(mid(Left(Map, Len(Map) - 4), LoopC, 1)) = False Then
            NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - LoopC)
            NameMap_Save = Left(Map, LoopC)
            Exit For

        End If

    Next

    For LoopC = (NumMap_Save - 4) To (NumMap_Save + 6)

        If FileExist(PATH_Save & NameMap_Save & LoopC & MapFormat, vbArchive) = True Then
            frmMain.MapPest(LoopC - NumMap_Save + 4).Visible = True
            frmMain.MapPest(LoopC - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(LoopC - NumMap_Save + 4).Caption = NameMap_Save & LoopC
        Else
            frmMain.MapPest(LoopC - NumMap_Save + 4).Visible = False

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
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then _
            GuardarMapa Path

    End If

    Exit Sub

DeseaGuardarMapa_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.DeseaGuardarMapa", Erl)
    Resume Next
    
End Sub

Public Sub CSMInfoSave()
'**********************************
'Autor: Lorwik
'Fecha: 14/03/2021
'**********************************

    On Error GoTo CSMInfoSave_Err

    MapDat.map_name = MapInfo.name
    MapDat.music_number = MapInfo.Music
    MapDat.ambient = MapInfo.ambient
    MapDat.lvlMinimo = MapInfo.lvlMinimo
    
    If frmMapInfo.chkLuzClimatica = Checked Then
        MapDat.LuzBase = MapInfo.LuzBase
        
    Else
        MapDat.LuzBase = 0
        
    End If
    
    MapDat.version = MapInfo.MapVersion
    
    If MapInfo.Pk = True Then
        MapDat.battle_mode = True
    Else
        MapDat.battle_mode = False
    End If
    
    MapDat.terrain = MapInfo.Terreno
    MapDat.zone = MapInfo.Zona
    MapDat.restrict_mode = MapInfo.Restringir
    MapDat.backup_mode = MapInfo.BackUp
    
    Exit Sub
CSMInfoSave_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.CSMInfoSave", Erl)
    Resume Next
    
End Sub

Public Sub CSMInfoCargar()
'**********************************
'Autor: Lorwik
'Fecha: 14/03/2021
'**********************************

    On Error GoTo CSMInfoCargar_Err:

    Dim tR As Byte
    Dim tG As Byte
    Dim tB As Byte
    
    MapInfo.name = MapDat.map_name
    MapInfo.Music = MapDat.music_number
    MapInfo.ambient = MapDat.ambient
    MapInfo.lvlMinimo = Val(MapDat.lvlMinimo)
    MapInfo.LuzBase = MapDat.LuzBase
    
    If MapDat.LuzBase <> 0 Then
        frmMapInfo.chkLuzClimatica = Checked
        Call ConvertLongToRGB(MapDat.LuzBase, tR, tG, tB)
        
        frmMapInfo.r1.Text = tR
        frmMapInfo.G1.Text = tG
        frmMapInfo.b1.Text = tB
        
        Estado_Custom.a = 255
        Estado_Custom.R = tR
        Estado_Custom.G = tG
        Estado_Custom.B = tB
        
        Call Actualizar_Estado
    Else
        frmMapInfo.chkLuzClimatica = Unchecked
    End If
    
    MapInfo.MapVersion = MapDat.version
    
    If MapDat.battle_mode = True Then
        MapInfo.Pk = True
    Else
        MapInfo.Pk = False
    End If
    
    MapInfo.Terreno = MapDat.terrain
    MapInfo.Zona = MapDat.zone
    MapInfo.Restringir = MapDat.restrict_mode
    MapInfo.BackUp = MapDat.backup_mode

    Call MapInfo_Actualizar
    
    Exit Sub
    
CSMInfoCargar_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.CSMInfoCargar", Erl)
    Resume Next
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error GoTo MapInfo_Actualizar_Err

    With frmMapInfo
        .txtMapNombre.Text = MapInfo.name
        .txtMapMusica.Text = MapInfo.Music
        .txtMapTerreno.Text = MapInfo.Terreno
        .txtMapZona.Text = MapInfo.Zona
        .txtMapRestringir.Text = MapInfo.Restringir
        '   .chkMapBackup.value = MapInfo.BackUp
        .chkMapPK.value = IIf(MapInfo.Pk = True, 1, 0)
        .TxtAmbient.Text = MapInfo.ambient
        .TxtlvlMinimo = MapInfo.lvlMinimo
        .chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
        .chkMapInviSinEfecto.value = IIf(MapInfo.InviSinEfecto, vbChecked, vbUnchecked)
        .chkMapResuSinEfecto.value = IIf(MapInfo.ResuSinEfecto, vbChecked, vbUnchecked)
        .txtMapVersion = MapInfo.MapVersion

    End With

    Exit Sub
    
MapInfo_Actualizar_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.MapInfo_Actualizar", Erl)

    Resume Next
    
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfoAO_Cargar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next

    Dim Leer  As New clsIniManager

    Dim LoopC As Integer

    Dim Path  As String

    MapTitulo = Empty
    Leer.Initialize Archivo

    For LoopC = Len(Archivo) To 1 Step -1

        If mid(Archivo, LoopC, 1) = "\" Then
            Path = Left(Archivo, LoopC)
            Exit For

        End If

    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(Left(Archivo, Len(Archivo) - 4))

    MapInfo.name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
    MapInfo.ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
    
    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.Pk = True
    Else
        MapInfo.Pk = False

    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    Call MapInfo_Actualizar
    
End Sub

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfoAO_Guardar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save

    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", str(MapInfo.BackUp))

    If MapInfo.Pk Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")

    End If

End Sub
