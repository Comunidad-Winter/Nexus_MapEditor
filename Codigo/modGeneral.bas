Attribute VB_Name = "modGeneral"
Option Explicit

Private lFrameTimer As Long

Sub Main()
    Static lastFlush As Long
    
    frmCargando.Show

    Call CargarConfiguracion
    Call GenerateContra
    
    ChDrive App.Path
    ChDir App.Path
    Windows_Temp_Dir = General_Get_Temp_Dir
    Form_Caption = "Nexus MapEditor v" & App.Major & "." & App.Minor & "." & App.Revision
    
    '##############
    ' MOTOR GRAFICO
    
    'Iniciamos el Engine de DirectX 8
    frmCargando.lblCargando.Caption = "Iniciando Motor Grafico..."
    Call mDx8_Engine.Engine_DirectX8_Init
          
    'Tile Engine
    frmCargando.lblCargando.Caption = "Cargando Tile Engine..."
    Call InitTileEngine(frmMain.hwnd, 32, 32, 8, 8)
    
    '##############
    ' Dats
    
    frmCargando.lblCargando.Caption = "Cargando Superficies..."
    Call CargarIndicesSuperficie
    
    frmCargando.lblCargando.Caption = "Cargando NPC's..."
    Call CargarIndicesNPC
    
    'Inicializacion de variables globales
    prgRun = True
    pausa = False
    
    Unload frmCargando
    modMapIO.NuevoMapa
    frmMain.Show
    
    Do While prgRun
    
        'Solo dibujamos si la ventana no esta minimizada
        If frmMain.WindowState <> vbMinimized And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
        End If
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            
            lFrameTimer = GetTickCount
        End If
        
        If timeGetTime >= lastFlush Then
            ' If there is anything to be sent, we send it
            lastFlush = timeGetTime + 10
        End If
        DoEvents
    Loop

End Sub

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    
    EngineRun = False
    
    'Stop tile engine
    Call Engine_DirectX8_End
    
    Set SurfaceDB = Nothing
    
    Call UnloadAllForms
    
End Sub

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
    
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    
    On Error GoTo FileExist_Err
    

    '*************************************************
    'Author: Unkwown
    'Last modified: 26/05/06
    '*************************************************
    If LenB(Dir(File, FileType)) = 0 Then
        FileExist = False
    Else
        FileExist = True

    End If

    
    Exit Function

FileExist_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.FileExist", Erl)
    Resume Next
    
End Function

Public Sub LogError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
'**********************************************************
'Author: Jopi
'Guarda una descripcion detallada del error en Errores.log
'**********************************************************
    Dim File As Integer
        File = FreeFile
        
    Open App.Path & "\logs\Errores.log" For Append As #File
    
        Print #File, "Error: " & Numero
        Print #File, "Descripcion: " & Descripcion
        
        If LenB(Linea) <> 0 Then
            Print #File, "Linea: " & Linea
        End If
        
        Print #File, "Componente: " & Componente
        Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        Print #File, vbNullString
        
    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
                
End Sub

Public Sub DibujarMinimapa()

    Dim map_x, map_y, Capas As Byte
    
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
End Sub

