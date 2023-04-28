Attribute VB_Name = "modDeclaraciones"
Option Explicit

Public DirRecursos As String
Public DirDats As String
Public DirInit As String
Public Form_Caption As String

'Control
Public prgRun As Boolean 'When true the program ends
Public pausa As Boolean

'Direcciones
Public Enum E_Heading
    nada = 0
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

' COLORES RENDER
Public ColorTecho As Byte
Public temp_rgb(3) As Long
Public renderText As String
Public renderFont As Integer
Public colorRender As Byte
Public render_msg(3) As Long

'Caminata fluida
Public Movement_Speed As Single

Public MapaCargado As Boolean
Public bRefreshRadar As Boolean

Public HotKeysAllow As Boolean  ' Control Automatico de HotKeys
Public vMostrando As Byte

'Map editor variables
Public WalkMode As Boolean
Public dLastWalk As Double

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'For KeyInput
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
