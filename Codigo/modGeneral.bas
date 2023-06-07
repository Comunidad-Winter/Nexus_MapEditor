Attribute VB_Name = "modGeneral"
Option Explicit

Private lFrameTimer As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub CheckKeys()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    'On Error GoTo CheckKeys_Err

    If HotKeysAllow = False Then Exit Sub
    '[Loopzer]
    'If GetKeyState(vbKeyControl) < 0 Then
    '    If Seleccionando Then
    '        If GetKeyState(vbKeyC) < 0 Then CopiarSeleccion
    '        If GetKeyState(vbKeyX) < 0 Then CortarSeleccion
    '        If GetKeyState(vbKeyB) < 0 Then BlockearSeleccion
    '        If GetKeyState(vbKeyD) < 0 Then AccionSeleccion
    ''    Else
    '        If GetKeyState(vbKeyS) < 0 Then DePegar ' GS
    '        If GetKeyState(vbKeyV) < 0 Then PegarSeleccion
    '    End If
    'End If
    '[/Loopzer]
    
    If GetKeyState(vbKeyUp) < 0 Then
        If UserPos.y < YMinMapSize Then Exit Sub ' 10
        
        If Map_LegalPos(UserPos.X, UserPos.y - 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.y = UserPos.y - 1
            Call Char_MovebyPos(UserCharIndex, UserPos.X, UserPos.y)
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.y = UserPos.y - 1

        End If
        
        bRefreshRadar = True ' Radar
        Call DibujarMinimapa(True)
        frmMain.SetFocus
        Exit Sub

    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If UserPos.X > XMaxMapSize Then Exit Sub ' 89
        
        If Map_LegalPos(UserPos.X + 1, UserPos.y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X + 1
            Call Char_MovebyPos(UserCharIndex, UserPos.X, UserPos.y)
            dLastWalk = GetTickCount
            
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X + 1
            
        End If
        
        bRefreshRadar = True ' Radar
        Call DibujarMinimapa(True)
        frmMain.SetFocus
        Exit Sub

    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If UserPos.y > YMaxMapSize Then Exit Sub ' 92
        
        If Map_LegalPos(UserPos.X, UserPos.y + 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.y = UserPos.y + 1
            Call Char_MovebyPos(UserCharIndex, UserPos.X, UserPos.y)
            dLastWalk = GetTickCount
            
        ElseIf WalkMode = False Then
            UserPos.y = UserPos.y + 1
            
        End If
        
        bRefreshRadar = True ' Radar
        Call DibujarMinimapa(True)
        frmMain.SetFocus
        Exit Sub
        
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If UserPos.X < XMinMapSize Then Exit Sub ' 12
        
        If Map_LegalPos(UserPos.X - 1, UserPos.y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X - 1
            Call Char_MovebyPos(UserCharIndex, UserPos.X, UserPos.y)
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X - 1

        End If

        bRefreshRadar = True ' Radar
        Call DibujarMinimapa(True)
        frmMain.SetFocus
        Exit Sub

    End If
    
'CheckKeys_Err:
'    Call LogError(Err.Number, Err.Description, "modGeneral.CheckKeys", Erl)
'    Resume Next
    
End Sub

Sub Main()
    
    frmCargando.Show

    Call CargarConfiguracion
    Call GenerateContra
    
    ChDrive App.Path
    ChDir App.Path
    Windows_Temp_Dir = General_Get_Temp_Dir
    
    '##############
    ' MOTOR GRAFICO
    
    'Iniciamos el Engine de DirectX 8
    frmCargando.lblCargando.Caption = "Iniciando Motor Grafico..."
    Call mDx8_Engine.Engine_DirectX8_Init
          
    'Tile Engine
    frmCargando.lblCargando.Caption = "Cargando Tile Engine..."
    Call InitTileEngine(frmMain.hWnd, 32, 32, 8, 8)
    
    '##############
    ' Dats
    
    frmCargando.lblCargando.Caption = "Cargando Superficies..."
    Call CargarIndicesSuperficie
    
    frmCargando.lblCargando.Caption = "Cargando Triggers..."
    Call CargarIndicesTriggers
    
    frmCargando.lblCargando.Caption = "Cargando NPC's..."
    Call CargarIndicesNPC
    
    frmCargando.lblCargando.Caption = "Cargando OBJ's..."
    Call CargarIndicesOBJ
    
    frmMiniMapa.Render.Width = XMaxMapSize
    frmMiniMapa.Render.Height = YMaxMapSize
    
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
            Call CheckKeys
            
        Else
            Sleep 10&
            
        End If
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            
            lFrameTimer = GetTickCount
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
    
    Erase MapData
    Erase SeleccionMap
    
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

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo ReadField_Err
    
    Dim i         As Integer
    Dim lastPos   As Integer
    Dim CurChar   As String * 1
    Dim FieldNum  As Integer
    Dim Seperator As String

    Seperator = Chr(SepASCII)
    lastPos = 0
    FieldNum = 0

    For i = 1 To Len(Text)
        CurChar = mid(Text, i, 1)

        If CurChar = Seperator Then
            FieldNum = FieldNum + 1

            If FieldNum = Pos Then
                ReadField = mid(Text, lastPos + 1, (InStr(lastPos + 1, Text, Seperator, vbTextCompare) - 1) - (lastPos))
                Exit Function

            End If

            lastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1

    If FieldNum = Pos Then
        ReadField = mid(Text, lastPos + 1)

    End If

    
    Exit Function

ReadField_Err:
    Call LogError(Err.Number, Err.Description, "modGeneral.ReadField", Erl)
    Resume Next
    
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    '******************************************r
    On Error GoTo AddtoRichTextBox_Err
    
    With RichTextBox

        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
    End With

    
    Exit Sub

AddtoRichTextBox_Err:
    Call LogError(Err.Number, Err.Description, "ModGeneral.AddtoRichTextBox", Erl)
    Resume Next
    
End Sub

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

Public Sub GenerarVista()
    '*************************************************
    'Author: Unknown
    'Last modified: ????
    '*************************************************
    
    On Error GoTo GenerarVista_Err
    
    VerBlockeados = frmMain.mnuVerBloqueos.Checked
    VerTriggers = frmMain.mnuVerTriggers.Checked
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    VerParticulas = frmMain.mnuVerParticulas.Checked
    VerLuces = frmMain.mnuVerParticulas.Checked
    
    
    Exit Sub

GenerarVista_Err:
    Call LogError(Err.Number, Err.Description, "modGeneral.GenerarVista", Erl)
    Resume Next
    
End Sub

Public Sub ToggleWalkMode()

    '*************************************************
    'Author: Unkwown
    'Last modified: 28/05/06 - GS
    '*************************************************
    On Error GoTo fin:

    If WalkMode = False Then
        WalkMode = True
    Else
        frmMain.mnuModoCaminata.Checked = False
        WalkMode = False

    End If

    If WalkMode = False Then
        'Erase character
        Call Char_Erase(UserCharIndex)
        MapData(UserPos.X, UserPos.y).CharIndex = 0
    Else

        'MakeCharacter
        If Map_LegalPos(UserPos.X, UserPos.y) Then
            Call Char_Make(NextOpenChar(), 1, 4, SOUTH, UserPos.X, UserPos.y, 0, 0, 0, 0, 0)
            UserCharIndex = MapData(UserPos.X, UserPos.y).CharIndex
            frmMain.mnuModoCaminata.Checked = True
        Else
            MsgBox "ERROR: Ubicacion ilegal."
            WalkMode = False

        End If

    End If

fin:

End Sub

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
    
    On Error GoTo SetTopMostWindow_Err
    

    If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopMostWindow = False

    End If

    
    Exit Function

SetTopMostWindow_Err:
    Call LogError(Err.Number, Err.Description, "ModGeneral.SetTopMostWindow", Erl)
    Resume Next
    
End Function
