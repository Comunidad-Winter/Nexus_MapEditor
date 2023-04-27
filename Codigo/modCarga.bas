Attribute VB_Name = "modCarga"
Option Explicit

Public Type tSetupMods

    ' VIDEO
    byMemory    As Integer
    OverrideVertexProcess As Byte
    LimiteFPS As Boolean
    FPSShow As Boolean
    
End Type

Public ClientSetup As tSetupMods

Public Type tCabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Private Lector As clsIniManager
Public Const CLIENT_FILE As String = "Config.ini"

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
    r As Integer
    g As Integer
    b As Integer
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
    NpcIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    OBJIndex As Integer
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

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Long
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer
End Type

Public Type tIndiceArmas
    weapon(1 To 4) As Long
End Type

Public Type tIndiceEscudos
    shield(1 To 4) As Long
End Type

Public NumHeads As Integer
Public NumCascos As Integer
Public NumEscudosAnims As Integer
Private grhCount As Long

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
        .CRC = Rnd * 245
        .MagicWord = Rnd * 92
    End With
    
End Sub

Public Sub CargarConfiguracion()
    On Local Error GoTo fileErr:
    
    Call IniciarCabecera

    Set Lector = New clsIniManager
    Call Lector.Initialize(App.Path & "\Config.ini")
    
    DirRecursos = Lector.GetValue("RUTAS", "Recursos")
    
    With ClientSetup
        ' VIDEO
        .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
        .OverrideVertexProcess = CByte(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
    End With
    
    Exit Sub
  
fileErr:

    If Err.Number <> 0 Then
       MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.Number & " : " & Err.Description)
       End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
End Sub

''
' Loads grh data using the new file format.
'

Public Sub LoadGrhData()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Graficos
'*************************************
On Error GoTo ErrorHandler:

    Dim Grh         As Long
    Dim Frame       As Long
    Dim fileVersion As Long
    Dim LaCabecera  As tCabecera
    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte
    
    InfoHead = File_Find(DirRecursos & "\Scripts" & Formato, LCase$("Graficos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Graficos.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        fileVersion = fileBuff.getLong
        
        grhCount = fileBuff.getLong
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Grh <> grhCount
            Grh = fileBuff.getLong

            With GrhData(Grh)
            
                '.active = True
                .NumFrames = fileBuff.getInteger
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        .Frames(Frame) = fileBuff.getLong
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    .speed = fileBuff.getSingle
                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                    
                Else
                    
                    .FileNum = fileBuff.getLong
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = fileBuff.getInteger
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = fileBuff.getInteger
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .sX = fileBuff.getInteger
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    .sY = fileBuff.getInteger
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
Exit Sub

ErrorHandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarCabezas()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cabezas
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim i           As Integer
    Dim j           As Integer
    Dim LaCabecera  As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "\Scripts" & Formato, LCase$("Head.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Head.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
        
        NumHeads = fileBuff.getInteger()  'cantidad de cabezas
    
        ReDim HeadData(0 To NumHeads) As HeadData
        ReDim Miscabezas(0 To NumHeads) As tIndiceCabeza
                      
        For i = 1 To NumHeads
        
            Miscabezas(i).Head(1) = fileBuff.getLong()
            Miscabezas(i).Head(2) = fileBuff.getLong()
            Miscabezas(i).Head(3) = fileBuff.getLong()
            Miscabezas(i).Head(4) = fileBuff.getLong()
                
            If Miscabezas(i).Head(1) Then
                Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
                Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
                Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
                Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
            End If
        Next i
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Head.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarCascos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cascos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Integer
    Dim j           As Integer
    Dim LaCabecera  As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "\Scripts" & Formato, LCase$("Helmet.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Helmet.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        NumCascos = fileBuff.getInteger()   'cantidad de cascos
             
        ReDim CascoAnimData(0 To NumCascos) As HeadData
        ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
             
        For i = 1 To NumCascos
        
            Miscabezas(i).Head(1) = fileBuff.getLong()
            Miscabezas(i).Head(2) = fileBuff.getLong()
            Miscabezas(i).Head(3) = fileBuff.getLong()
            Miscabezas(i).Head(4) = fileBuff.getLong()
            
            If Miscabezas(i).Head(1) Then
                Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
                Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
                Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
                Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
            End If
        Next i
         
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Helmet.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarCuerpos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cuerpos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    Dim LaCabecera As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "\Scripts" & Formato, LCase$("Personajes.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Personajes.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de cabezas
        NumCuerpos = fileBuff.getInteger()
    
        'Resize array
        ReDim BodyData(0 To NumCuerpos) As BodyData
        ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
        For i = 1 To NumCuerpos
            MisCuerpos(i).Body(1) = fileBuff.getLong()
            MisCuerpos(i).Body(2) = fileBuff.getLong()
            MisCuerpos(i).Body(3) = fileBuff.getLong()
            MisCuerpos(i).Body(4) = fileBuff.getLong()
            MisCuerpos(i).HeadOffsetX = fileBuff.getInteger()
            MisCuerpos(i).HeadOffsetY = fileBuff.getInteger()
            
            If MisCuerpos(i).Body(1) Then
                Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
                Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
                Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
                Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
                
                BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
                BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarFxs()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Fxs
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Long
    Dim NumFxs      As Integer
    Dim LaCabecera  As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "\Scripts" & Formato, LCase$("FXs.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("FXs.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de Fxs
        NumFxs = fileBuff.getInteger()
        
        'Resize array
        ReDim FxData(1 To NumFxs) As tIndiceFx
        
        For i = 1 To NumFxs
            FxData(i).Animacion = fileBuff.getLong()
            FxData(i).OffsetX = fileBuff.getInteger()
            FxData(i).OffsetY = fileBuff.getInteger()
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing

errhandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Fxs.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If

End Sub

Public Sub CargarMinimapa()

    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte
    Dim i           As Long
    
    InfoHead = File_Find(DirRecursos & "\Scripts" & Formato, LCase$("minimap.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("minimap.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        For i = 1 To grhCount
            If Grh_Check(i) Then
                GrhData(i).mini_map_color = fileBuff.getLong
            End If
        Next i
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
End Sub

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grhCount Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If
End Function
