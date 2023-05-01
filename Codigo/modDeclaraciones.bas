Attribute VB_Name = "modDeclaraciones"
Option Explicit

Public DirRecursos   As String

Public DirDats       As String

Public DirInit       As String

Public Form_Caption  As String

Public PATH_Save     As String

Public NumMap_Save   As Integer

Public NameMap_Save  As String

Public MapaActual    As Integer

' Client Config
Public ClienteHeight As Integer

Public ClienteWidth  As Integer

'Control
Public prgRun        As Boolean 'When true the program ends

Public pausa         As Boolean

'Direcciones
Public Enum E_Heading

    nada = 0
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

' COLORES RENDER
Public ColorTecho            As Byte

Public temp_rgb(3)           As Long

Public renderText            As String

Public renderFont            As Integer

Public colorRender           As Byte

Public render_msg(3)         As Long

'Caminata fluida
Public Movement_Speed        As Single

Public MapaCargado           As Boolean

Public bRefreshRadar         As Boolean

Public bAutoGuardarMapa      As Byte

Public bAutoGuardarMapaCount As Byte

Public HotKeysAllow          As Boolean  ' Control Automatico de HotKeys

Public vMostrando            As Byte

'Map editor variables
Public WalkMode              As Boolean

Public dLastWalk             As Double

Public VerBlockeados         As Boolean

Public VerTriggers           As Boolean

Public VerMarco              As Boolean ' Marco

Public VerGrilla             As Boolean ' grilla

Public VerCapa1              As Boolean

Public VerCapa2              As Boolean

Public VerCapa3              As Boolean

Public VerCapa4              As Boolean

Public VerTranslados         As Boolean

Public VerObjetos            As Boolean

Public VerNpcs               As Boolean

Public VerParticulas         As Boolean

Public VerLuces              As Boolean

Public Cfg_TrOBJ             As Integer

Public UltimoClickX          As Integer

Public UltimoClickY          As Integer

Public SobreX                As Byte   ' Posicion X bajo el Cursor

Public SobreY                As Byte   ' Posicion Y bajo el Cursor

Public NoSobreescribir       As Boolean

Public Working               As Boolean

Public Const OFFSET_HEAD     As Integer = 0

Public Const MSGMod          As String = "Este mapa há sido modificado." & vbCrLf & "Si no lo guardas perderas todos los cambios ¿Deseas guardarlo?"

Public Const MSGDang         As String = "CUIDADO! Este comando puede arruinar el mapa." & vbCrLf & "¿Estas seguro que desea continuar?"

Public Const ENDL            As String * 2 = vbCrLf

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpFileName As String) As Long

Public Declare Function getprivateprofilestring _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nSize As Long, _
                                                 ByVal lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'For KeyInput
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
