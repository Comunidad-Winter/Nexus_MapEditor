Attribute VB_Name = "mod_TileEngine"
Option Explicit

Private NotFirstRender As Boolean

Dim temp_verts(3) As TLVERTEX

Public OffsetCounterX As Single
Public OffsetCounterY As Single
    
Public WeatherFogX1 As Single
Public WeatherFogY1 As Single
Public WeatherFogX2 As Single
Public WeatherFogY2 As Single
Public WeatherFogCount As Byte

Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer
Public LastOffsetY As Integer

'Map sizes in tiles
Public XMaxMapSize As Byte
Public XMinMapSize As Byte
Public YMaxMapSize As Byte
Public YMinMapSize As Byte

Private Const GrhFogata As Long = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

'Posicion en un mapa
Public Type Position
    X As Long
    y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamano y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    speed As Single
    
    mini_map_color As Long
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Long
    GrhIndexInt As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

Public CurrentGrh  As Grh

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Apariencia del personaje
Public Type Char
    Movement As Boolean
    active As Byte
    Heading As E_Heading
    Pos As Position
    moved As Boolean
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    AuraAnim As Grh
    AuraColor As Long
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Atacable As Byte
    WorldBoss As Boolean
    
    Nombre As String
    Clan As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
    ParticleIndex As Integer
    Particle_Count As Long
    Particle_Group() As Long
End Type

'Info de un objeto
Public Type obj
    objindex As Integer
    Amount As Integer
End Type

Private Type tLight
    RGBCOLOR As D3DCOLORVALUE
    active As Boolean
    map_x As Integer
    map_y As Integer
    range As Byte
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    
    Engine_Light(0 To 3) As Long
    Particle_Index As Integer
    Light As tLight
    
    Particle_Group_Index As Long 'Particle Engine
    
    fX As Grh
    FxIndex As Integer
End Type

'Info del mapa
Type MapInfo

    NumUsers As Integer
    Music As String
    ambient As String
    name As String
    StartPos As WorldPos
    OnDeathGoTo As WorldPos
    
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    
    ' Anti Magias/Habilidades
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    OcultarSinEfecto As Byte
    InvocarSinEfecto As Byte
    
    RoboNpcsPermitido As Byte
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    
    lvlMinimo As Byte
    
    NoTirarItems As Byte
    
    LuzBase As Long
    
    Changed As Byte
End Type

Public SeleccionIX           As Integer
Public SeleccionFX           As Integer
Public SeleccionIY           As Integer
Public SeleccionFY           As Integer
Public SeleccionAncho        As Integer
Public SeleccionAlto         As Integer
Public Seleccionando         As Boolean
Public SeleccionMap()        As MapBlock

Public DeSeleccionOX         As Integer
Public DeSeleccionOY         As Integer
Public DeSeleccionIX         As Integer
Public DeSeleccionFX         As Integer
Public DeSeleccionIY         As Integer
Public DeSeleccionFY         As Integer
Public DeSeleccionAncho      As Integer
Public DeSeleccionAlto       As Integer
Public DeSeleccionando       As Boolean
Public DeSeleccionMap()      As MapBlock

'Bordes del mapa
Public MinXBorder As Integer
Public MaxXBorder As Integer
Public MinYBorder As Integer
Public MaxYBorder As Integer

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Public FPSLastCheck As Long

'Tamano del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Tamano de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer

Private MouseTileX As Integer
Private MouseTileY As Integer

'?????????Graficos???????????
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'?????????????????????????

'?????????Mapa????????????
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'?????????????????????????

Public Normal_RGBList(3) As Long
Public NoUsa_RGBList(3) As Long
Public Color_Paralisis As Long
Public Color_Invisibilidad As Long
Public Color_Montura As Long

'   Control de Lluvia
Public bTecho       As Boolean 'hay techo?
Public bFogata       As Boolean

Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'?????????????????????????

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    On Error GoTo ConverCPtoTP_Err
    
    tX = UserPos.X + viewPortX \ 32 - frmMain.MainViewPic.ScaleWidth \ 64
    tY = UserPos.y + viewPortY \ 32 - frmMain.MainViewPic.ScaleHeight \ 64
    
    Exit Sub
    
ConverCPtoTP_Err:
    Call LogError(Err.Number, Err.Description, "mod_TileEngine.ConvertCPtoTP(", Erl)
    Resume Next
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Grh.GrhIndex < 1 Then Exit Sub
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.GrhIndex).speed
End Sub

Public Sub InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer)
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
'Configures the engine to start running.
'***************************************************

On Error GoTo ErrorHandler:

    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    
    Call ChangeView

    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    Call modEdicion.InitDeshacer
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

On Error GoTo 0
    
    'Cargamos indice de graficos.
    'TODO: No usar variable de compilacion y acceder a esto desde el config.ini
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarMinimapa
    Call LoadGraphics
    Call CargarParticulas

    Exit Sub
    
ErrorHandler:

    Call LogError(Err.Number, Err.Description, "Mod_TileEngine.InitTileEngine")
    
    Call CloseClient
    
End Sub

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.byMemory)
End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, _
                  ByVal DisplayFormLeft As Integer, _
                  ByVal MouseViewX As Integer, _
                  ByVal MouseViewY As Integer)

On Error GoTo ErrorHandler:

    If EngineRun Then
        
        Call Engine_BeginScene
            
        Call DesvanecimientoTechos
        Call DesvanecimientoMsg
            
        If UserMoving Then
            
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
    
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
    
                End If
                    
            End If
                
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
    
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.y = 0
                    UserMoving = False
                        
                End If
    
            End If
    
        End If
            
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
            
        '****** Update screen ******
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)

        ' Calculamos los FPS y los mostramos
        Call Engine_Update_FPS
        Call DrawText(5, 5, "FPS: " & mod_TileEngine.FPS, -1, False)
        Call DrawText(5, 20, "X: " & UserPos.X & " Y: " & UserPos.y, -1, False)
        Call DrawText(5, 35, "Mouse X: " & frmMain.tX & " Y: " & frmMain.tY, -1, False)
    
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
            
        Call Engine_EndScene(MainScreenRect, 0)
    
    End If
    
ErrorHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        
        Call mDx8_Engine.Engine_DirectX8_Init
        
        Call LoadGraphics
    
    End If
  
End Sub

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************
    
    On Error GoTo RenderScreen_Err
    
    Dim y                As Long     'Keeps track of where on map we are

    Dim X                As Long     'Keeps track of where on map we are
    
    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen
    
    Dim minY             As Integer  'Start Y pos on current map

    Dim maxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim maxX             As Integer  'End X pos on current map
    
    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen
    
    Dim minXOffset       As Integer

    Dim minYOffset       As Integer
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    Dim ElapsedTime      As Single
    
    Dim HeadingIt        As Byte
    
    Dim i                As Integer
    
    Dim Grh              As Grh
    
    Dim Sobre            As Long
    
    Dim bCapa            As Byte
    
    ElapsedTime = Engine_ElapsedTime()
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize * 2 ' WyroX: Parche para que no desaparezcan techos y arboles
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then
        screenmaxY = screenmaxY + 1
    Else
        screenmaxY = YMaxMapSize
    End If
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then
        screenmaxX = screenmaxX + 1
    Else
        screenmaxX = XMaxMapSize
    End If
    
    If Val(frmSuperficies.cCapas.Text) >= 1 And (frmSuperficies.cCapas.Text) <= 4 Then
        bCapa = Val(frmSuperficies.cCapas.Text)
    Else
        bCapa = 1

    End If
    
    Call GenerarVista
    
    If screenmaxX > XMaxMapSize Then screenmaxX = XMaxMapSize

    If screenminX < XMinMapSize Then
        ScreenX = XMinMapSize - screenminX
        screenminX = XMinMapSize

    End If

    If screenminY < YMinMapSize Then
        ScreenY = YMinMapSize - screenminY
        screenminY = YMinMapSize

    End If
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
            
            If SobreX = X And SobreY = y Then

                ' Pone Grh !
                Sobre = -1

                If frmSuperficies.cSeleccionarSuperficie.value = True Then
                    Sobre = MapData(X, y).Graphic(bCapa).GrhIndex

                    If frmConfigSup.MOSAICO.value = vbChecked Then

                        Dim aux As Long
                        Dim dy  As Integer
                        Dim dX  As Integer

                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo.Text)
                            dX = Val(frmConfigSup.DMAncho.Text)
                            
                        Else
                            dy = 0
                            dX = 0

                        End If

                        If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                            aux = Val(frmSuperficies.cGrh.Text) + (((y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)

                            If MapData(X, y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, y).Graphic(bCapa), aux

                            End If

                        Else
                            aux = Val(frmSuperficies.cGrh.Text) + (((y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)

                            If MapData(X, y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, y).Graphic(bCapa), aux

                            End If

                        End If

                    Else

                        If MapData(X, y).Graphic(bCapa).GrhIndex <> Val(frmSuperficies.cGrh.Text) Then
                            MapData(X, y).Graphic(bCapa).GrhIndex = Val(frmSuperficies.cGrh.Text)
                            InitGrh MapData(X, y).Graphic(bCapa), Val(frmSuperficies.cGrh.Text)

                        End If

                    End If

                End If

            Else
            
                Sobre = -1
            
            End If
            
            'Layer 1 **********************************
            If MapData(X, y).Graphic(1).GrhIndex <> 0 And VerCapa1 Then Call Draw_Grh(MapData(X, y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, y).Engine_Light(), 1)
            '******************************************

            'Layer 2 **********************************
            If MapData(X, y).Graphic(2).GrhIndex <> 0 And VerCapa2 Then Call Draw_Grh(MapData(X, y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, y).Engine_Light(), 1)
            '******************************************
            
            If Sobre >= 0 Then
                If MapData(X, y).Graphic(bCapa).GrhIndex <> Sobre Then
                    MapData(X, y).Graphic(bCapa).GrhIndex = Sobre
                    InitGrh MapData(X, y).Graphic(bCapa), Sobre

                End If

            End If
            
            ScreenX = ScreenX + 1
            
            If X > XMaxMapSize Then Exit For
        Next X
    
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
        
        If y > YMaxMapSize Then Exit For
    Next y
                 
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - TileBufferSize

    For y = minY To maxY
        
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX

            If Map_InBounds(X, y) Then
            
                PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
                
                With MapData(X, y)
                
                    'Object Layer **********************************
                    If .ObjGrh.GrhIndex <> 0 And VerObjetos Then Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                    '***********************************************

                    'Char layer********************************
                    If .CharIndex <> 0 And VerNpcs Then Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                    '*************************************************

                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 And VerCapa3 Then Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                    '************************************************
                    
                    'Particulas *****************************************
                    If .Particle_Group_Index And VerParticulas Then _
                        Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                    '************************************************
                    
                    If Not .FxIndex = 0 Then
                        Call Draw_Grh(.fX, PixelOffsetXTemp + FxData(MapData(X, y).FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxIndex).OffsetY, 1, .Engine_Light(), 1, True)

                        If .fX.Started = 0 Then .FxIndex = 0

                    End If
                    
                End With
                
            End If
            
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next y
    
    '<----- Layer 4 ----->
    ScreenY = minYOffset - TileBufferSize

    For y = minY To maxY

        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
            
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            'Layer 4
            If MapData(X, y).Graphic(4).GrhIndex And VerCapa4 Then
            
                If bTecho Then
                    Call Draw_Grh(MapData(X, y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, temp_rgb(), 1)
                Else
                
                    If ColorTecho = 250 Then
                        Call Draw_Grh(MapData(X, y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, y).Engine_Light(), 1)
                    Else
                        Call Draw_Grh(MapData(X, y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, temp_rgb(), 1)

                    End If

                End If

            End If
                           
            If MapData(X, y).TileExit.Map <> 0 And VerTranslados Then
                Grh.GrhIndex = 3
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 1)
                        
            End If
                
            'Show blocked tiles
            If VerBlockeados And MapData(X, y).Blocked = 1 Then
                Grh.GrhIndex = 4
                Grh.FrameCounter = 1
                Grh.Started = 0
                    
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 1)
                        
            End If
                
            If VerGrilla Then
                Grh.GrhIndex = 2
                Grh.FrameCounter = 1
                Grh.Started = 0
                    
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 0)
                        
            End If

            If VerTriggers Then If MapData(X, y).Trigger > 0 Then Call DrawText(PixelOffsetXTemp + 5, PixelOffsetYTemp - 13, MapData(X, y).Trigger, -1, False, 2)

            If Seleccionando Then
          
                If X >= SeleccionIX And y >= SeleccionIY Then
                    If X <= SeleccionFX And y <= SeleccionFY Then
                   
                        Engine_Draw_Box PixelOffsetXTemp, PixelOffsetYTemp, 30, 30, D3DColorARGB(75, 0, 200, 100)
    
                    End If

                End If

            End If
            
            ScreenX = ScreenX + 1
            
        Next X

        ScreenY = ScreenY + 1
    Next y
    
    '   Set Offsets
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    
RenderScreen_Err:

    If Err.Number Then
        Call LogError(Err.Number, Err.Description, "Mod_TileEngine.RenderScreen")

    End If

End Sub

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Draw char's to screen without offcentering them
    '***************************************************
    On Error GoTo Char_Render_Err
    
    Dim moved As Boolean
    Dim AuraColorFinal(0 To 3) As Long
    Dim ColorFinal(0 To 3) As Long
    
    With charlist(CharIndex)

        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If

        If .Heading = 0 Then Exit Sub
        
        'If done moving stop animation
        If Not moved Then
        
            'Evito runtime
            If Not .Heading <> 0 Then .Heading = EAST
        
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False

        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
    
        ColorFinal(0) = MapData(.Pos.X, .Pos.y).Engine_Light()(0)
        ColorFinal(1) = MapData(.Pos.X, .Pos.y).Engine_Light()(1)
        ColorFinal(2) = MapData(.Pos.X, .Pos.y).Engine_Light()(2)
        ColorFinal(3) = MapData(.Pos.X, .Pos.y).Engine_Light()(3)
    
        'Draw Body
        If .Body.Walk(.Heading).GrhIndex Then _
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal, 1)

        'Draw Head
        If .Head.Head(.Heading).GrhIndex Then _
            Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 1, PixelOffsetY + .Body.HeadOffset.y, 1, ColorFinal(), 0)
                
        'Draw Helmet
        If .Casco.Head(.Heading).GrhIndex Then _
            Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 1, PixelOffsetY + .Body.HeadOffset.y + OFFSET_HEAD, 1, ColorFinal(), 0)
                
        'Draw Weapon
        If .Arma.WeaponWalk(.Heading).GrhIndex Then _
            Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
                
        'Draw Shield
        If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
            Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
        
    End With

    Exit Sub

Char_Render_Err:
    Call LogError(Err.Number, Err.Description, "mod_TileEngine.CharRender", Erl)
    Resume Next

End Sub

Sub Draw_GrhIndex(ByVal GrhIndex As Long, ByVal X As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal angle As Single = 0, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        'Draw
        Call Device_Textured_Render(X, y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha)
    End With
    
End Sub

Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, ByVal Animate As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
On Error GoTo Error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.speed) * Movement_Speed

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth * ScaleX - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Call Device_Textured_Render(X, y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
        
    End With
    
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        #If Desarrollo = 0 Then
            Call LogError(Err.Number, "Error in Draw_Grh, " & Err.Description, "Draw_Grh", Erl)
            MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
            Call CloseClient
        
        #Else
            Debug.Print "Error en Draw_Grh en el grh" & CurrentGrhIndex & ", " & Err.Description & ", (" & Err.Number & ")"
        #End If
    End If
End Sub

Public Sub Device_Textured_Render(ByVal X As Single, ByVal y As Single, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef color() As Long, _
                                  Optional ByVal Alpha As Boolean = False, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal ScaleX As Single = 1!, _
                                  Optional ByVal ScaleY As Single = 1!)

        Dim Texture As Direct3DTexture8
        
        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
        
        With SpriteBatch

                Call .SetTexture(Texture)
                    
                Call .SetAlpha(Alpha)
                
                If TextureWidth <> 0 And TextureHeight <> 0 Then
                    Call .Draw(X, y, Width * ScaleX, Height * ScaleY, color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
                Else
                    Call .Draw(X, y, TextureWidth * ScaleX, TextureHeight * ScaleY, color, , , , , angle)
                End If
                
        End With
        
End Sub

Function InMapBounds(ByVal X As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Public Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub GrhUninitialize(Grh As Grh)
        '*****************************************************************
        'Author: Aaron Perkins
        'Last Modify Date: 1/04/2003
        'Resets a Grh
        '*****************************************************************

        With Grh
        
                'Copy of parameters
                .GrhIndex = 0
                .Started = False
                .Loops = 0
        
                'Set frame counters
                .FrameCounter = 0
                .speed = 0
                
        End With

End Sub

Private Sub DesvanecimientoTechos()
 
    If bTecho Then
        If Not Val(ColorTecho) = 50 Then ColorTecho = ColorTecho - 1
    Else
        If Not Val(ColorTecho) = 250 Then ColorTecho = ColorTecho + 1
    End If
    
    If Not Val(ColorTecho) = 250 Then
        Call Engine_Long_To_RGB_List(temp_rgb(), D3DColorARGB(ColorTecho, ColorTecho, ColorTecho, ColorTecho))
    End If
    
End Sub

Public Sub DesvanecimientoMsg()
'*****************************************************************
'Author: FrankoH
'Last Modify Date: 04/09/2019
'DESVANECIMIENTO DE LOS TEXTOS DEL RENDER
'*****************************************************************
    Static lastmovement As Long
    
    If GetTickCount - lastmovement > 1 Then
        lastmovement = GetTickCount
    Else
        Exit Sub
    End If

    If LenB(renderText) Then
        If Not Val(colorRender) = 0 Then colorRender = colorRender - 1
    ElseIf LenB(renderText) = 0 Then
        Exit Sub
    Else
        If Not Val(colorRender) = 240 Then colorRender = colorRender + 1
    End If
    
    If Not Val(colorRender) = 240 Then
        Call Engine_Long_To_RGB_List(render_msg(), ARGB(255, 255, 255, colorRender))
    End If
    
    If colorRender = 0 Then renderMsgReset
    
End Sub

Public Sub renderMsgReset()

    renderFont = 1
    renderText = vbNullString

End Sub

Public Sub DibujarMinimapa(Optional ByVal Refrescar As Boolean = False)

    Dim map_x As Byte

    Dim map_y As Byte

    Dim Capas As Byte
    
    If Working Then Exit Sub
    
    If Not Refrescar Then
    
        'Primero limpiamos el minimapa anterior
        frmMiniMapa.Render.Cls
    
        For map_y = YMinMapSize To YMaxMapSize
            For map_x = XMinMapSize To XMaxMapSize
            
                For Capas = 1 To 2

                    If MapData(map_x, map_y).Graphic(Capas).GrhIndex > 0 Then
                        SetPixel frmMiniMapa.Render.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(Capas).GrhIndex).mini_map_color

                    End If
                    
                Next Capas
            Next map_x
        Next map_y
   
        'Refrescamos
        frmMiniMapa.Render.Refresh

    End If
    
    frmMiniMapa.ApuntadorRadar.Left = (UserPos.X) - 9
    frmMiniMapa.ApuntadorRadar.Top = (UserPos.y) - 8

End Sub

Public Sub ChangeView()
    WindowTileHeight = Round(frmMain.MainViewPic.Height / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.Width / 32, 0)

    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2
End Sub

Public Sub MapCapture(ByRef Format As Boolean, ByVal ConBordes As Boolean)
    
    On Error GoTo MapCapture_Err
    

    '*************************************************
    'Author: Torres Patricio(Pato)
    'Last modified:12/03/11
    '*************************************************
    
    Dim D3DWindow        As D3DPRESENT_PARAMETERS
    Dim y                As Long     'Keeps track of where on map we are
    Dim X                As Long     'Keeps track of where on map we are
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim Grh              As Grh      'Temp Grh for show tile and blocked
    Static srcRect       As RECT
    Static destRect      As RECT

    Call Engine_BeginScene

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            'Layer 1 **********************************
            If MapData(X, y).Graphic(1).GrhIndex <> 0 Then _
                Call Draw_Grh(MapData(X, y).Graphic(1), (X - 1) * 32 + TilePixelWidth - 35, (y - 1) * 32 + TilePixelHeight - 35, 0, Normal_RGBList(), 1)
            '******************************************
        
        Next X
    Next y
    
    'Draw floor layer 2
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            'Layer 2 **********************************
            If (MapData(X, y).Graphic(2).GrhIndex <> 0) And VerCapa2 Then _
                Call Draw_Grh(MapData(X, y).Graphic(2), (X - 1) * 32 + TilePixelWidth - 35, (y - 1) * 32 + TilePixelHeight - 35, 1, Normal_RGBList(), 1)
            '******************************************
        
        Next X
    Next y

    'Draw Transparent Layers
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
                
            PixelOffsetXTemp = (X - 1) * 32 + TilePixelWidth
            PixelOffsetYTemp = (y - 1) * 32 + TilePixelHeight
        
            With MapData(X, y)
            
                'Object Layer **********************************
                If (.ObjGrh.GrhIndex <> 0) And VerObjetos Then _
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, Normal_RGBList(), 1)
                '***********************************************
            
                'Layer 3 *****************************************
                If (.Graphic(3).GrhIndex <> 0) And VerCapa3 Then _
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, Normal_RGBList(), 1)
                '************************************************
            
            End With

        Next X
    Next y
    
    'Draw layer 4
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(X, y)
            
                'Layer 4 **********************************
                If (.Graphic(4).GrhIndex <> 0) And VerCapa4 Then _
                    Call Draw_Grh(.Graphic(4), (X - 1) * 32 + TilePixelWidth - 35, (y - 1) * 32 + TilePixelHeight - 35, 1, Normal_RGBList(), 1)

                '**********************************
            
            End With

        Next X
    Next y

    'Draw trans, bloqs, triggers and select tiles
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(X, y)
                PixelOffsetXTemp = (X - 1) * 32 + TilePixelWidth
                PixelOffsetYTemp = (y - 1) * 32 + TilePixelHeight
            
                '**********************************
                Grh.FrameCounter = 1
                Grh.Started = 0

                If (.TileExit.Map <> 0) And VerTranslados Then
                    Grh.GrhIndex = 3
                
                    Call Draw_Grh(Grh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, Normal_RGBList(), 0)

                End If
            
                'Show blocked tiles
                If (.Blocked = 1) And VerBlockeados Then
                    Grh.GrhIndex = 4
                    Call Draw_Grh(Grh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, Normal_RGBList(), 0)

                End If

                '******************************************
            
            End With

        Next X
    Next y

    ToWorldMap2 = True
    
    With srcRect
        If ConBordes Then
            .Left = 0
            .Top = 0
            .Right = 3200
            .Bottom = 3200
        Else
            .Left = 13 * 32
            .Top = 10 * 32
            .Right = (100 - 13) * 32
            .Bottom = (100 - 10) * 32
        End If
    End With
    
    With destRect
        .Left = 0
        .Top = 0
        
        If ConBordes Then
            .Right = 3200
            .Bottom = 3200
        Else
            .Right = (100 - 26) * 32
            .Bottom = (100 - 20) * 32
        End If
    End With
    
    If ConBordes Then
        FrmRender.picMap.Width = 800
        FrmRender.picMap.Height = 800
    Else
        FrmRender.picMap.Width = (100 - 26) * 8
        FrmRender.picMap.Height = (100 - 20) * 8
    End If
    
    Call Engine_EndScene(srcRect, FrmRender.picMap.hWnd)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''Guardo la imagen''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call FrmRender.Capturar_Imagen(FrmRender.picMap, FrmRender.picMap)
    
    'Si no existe la carpeta de MiniMapas, la hacemos.
    If Not FileExist(App.Path & "\Render\", vbDirectory) Then
        Call MkDir(App.Path & "\Render\")

    End If
    
    'SavePicture FrmRender.picMap, "C:\imagen.bmp"
    SavePicture FrmRender.picMap.Picture, App.Path & "\Render\" & frmMain.MapPest(4).Caption & ".bmp"
    
    'Call SavePicture(FrmRender.picMap, App.Path & "\Render\" & NumMap_Save & ".bmp")
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    
    Exit Sub

MapCapture_Err:
    Call LogError(Err.Number, Err.Description, "mod_TileEngine.MapCapture", Erl)
    Resume Next
    
End Sub
