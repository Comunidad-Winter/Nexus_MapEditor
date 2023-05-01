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
                .ObjGrh.GrhIndex = 0
    
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
        .Pk = True
        .MagiaSinEfecto = 0
        .InviSinEfecto = 0
        .ResuSinEfecto = 0
        .Terreno = "BOSQUE"
        .Zona = "CAMPO"
        .Restringir = 0
    
    End With

    'Call MapInfo_Actualizar
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

Private Sub MapaCSM_Cargar(ByVal RutaMapa As String)
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

    Dim MapDat       As tMapDat

    Dim npcfile      As String

    Dim i            As Long

    Dim j            As Long

    Dim LaCabecera   As tCabecera
    
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
    
    Call Actualizar_Estado
    Call DibujarMinimapa
    Call CSMInfoCargar
    
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

Private Function MapaCSM_Guardar(ByVal RutaMapa As String) As Boolean
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
    
    Call AddtoRichTextBox(frmConsola.StatTxt, "Mapa " & RutaMapa & " guardado...", 0, 255, 0)
    Exit Function

ErrorHandler:
    If fh <> 0 Then Close fh
End Function

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

        If FileExist(PATH_Save & NameMap_Save & LoopC & ".csm", vbArchive) = True Then
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
    
    MapDat.ambient = MapInfo.ambient
    MapDat.terrain = MapInfo.Terreno
    MapDat.zone = MapInfo.Zona
    MapDat.restrict_mode = MapInfo.Restringir
    MapDat.backup_mode = MapInfo.BackUp
    
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
    
    MapInfo.lvlMinimo = Val(MapDat.lvlMinimo)
    MapInfo.LuzBase = MapDat.LuzBase
    
    If MapDat.LuzBase <> 0 Then
        frmMapInfo.chkLuzClimatica = Checked
        Call ConvertLongToRGB(MapDat.LuzBase, tR, tG, tB)
        
        frmMapInfo.r1.Text = tR
        frmMapInfo.G1.Text = tG
        frmMapInfo.b1.Text = tB
    Else
        frmMapInfo.chkLuzClimatica = Unchecked
    End If
    
    MapInfo.MapVersion = MapDat.version
    
    If MapDat.battle_mode = True Then
        MapInfo.Pk = True
    Else
        MapInfo.Pk = False
    End If
    
    MapInfo.ambient = MapDat.ambient
    
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
