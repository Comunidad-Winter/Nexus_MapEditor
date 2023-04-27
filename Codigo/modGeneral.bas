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
    Call mDx8_Engine.Engine_DirectX8_Init
          
    'Tile Engine
    Call InitTileEngine(frmMain.hWnd, 32, 32, 8, 8)
    
    Call mDx8_Engine.Engine_DirectX8_Aditional_Init
    
    'Inicializacion de variables globales
    prgRun = True
    pausa = False
    
    Unload frmCargando
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
