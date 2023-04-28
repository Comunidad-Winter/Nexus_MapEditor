VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000006&
   Caption         =   "Nexus MapEditor"
   ClientHeight    =   11595
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   19200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   768
   ScaleMode       =   0  'User
   ScaleWidth      =   1280
   StartUpPosition =   2  'CenterScreen
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "Super."
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":0000
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":0028
      cback           =   -2147483633
   End
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10965
      Left            =   30
      ScaleHeight     =   731
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1275
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   19125
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   90
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   1
      Left            =   1065
      TabIndex        =   15
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "Trasl."
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":0C7A
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":0CA2
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   2
      Left            =   2010
      TabIndex        =   16
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "Bloq."
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":18F4
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":191C
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   3
      Left            =   2940
      TabIndex        =   17
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "NPC's"
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":256E
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":2596
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   4
      Left            =   3870
      TabIndex        =   18
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "Obj."
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":31E8
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":3210
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   19
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "Trigg."
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":3E62
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":3E8A
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   6
      Left            =   5730
      TabIndex        =   20
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "Partic."
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":4ADC
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":4B04
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   7
      Left            =   6660
      TabIndex        =   21
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "Luces"
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":5186
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":51AE
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   8
      Left            =   7590
      TabIndex        =   22
      Top             =   90
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      caption         =   "Bordes"
      capalign        =   2
      backstyle       =   2
      font            =   "frmMain.frx":5650
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      image           =   "frmMain.frx":5678
      cback           =   -2147483633
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   597
      X2              =   597
      Y1              =   5.961
      Y2              =   33.78
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   600
      X2              =   600
      Y1              =   5.961
      Y2              =   33.78
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   16800
      TabIndex        =   12
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9150
      TabIndex        =   11
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   16035
      TabIndex        =   10
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   15270
      TabIndex        =   9
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   14505
      TabIndex        =   8
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   13740
      TabIndex        =   7
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   12990
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   12210
      TabIndex        =   5
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   11445
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   10680
      TabIndex        =   3
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9915
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   17565
      TabIndex        =   1
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   18330
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine0 
         Caption         =   "-"
      End
      Begin VB.Menu menuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuMinimapa 
         Caption         =   "&Minimapa"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuVerParticulas 
         Caption         =   "...Parti&culas"
      End
      Begin VB.Menu mnuLinMostrar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tX                  As Byte
Public tY                  As Byte
Public MouseX              As Long
Public MouseY              As Long
Public MouseBoton          As Long
Public MouseShift          As Long
Private clicX              As Long
Private clicY              As Long

Public UltPos As Integer

Private Sub Form_Load()

    Me.Caption = Form_Caption
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    
    
    Exit Sub

Form_KeyDown_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.Form_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub menuSalir_Click()
    Call CloseClient
End Sub

Private Sub mnuAbrirMapa_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2023
    '*************************************************
    Dialog.CancelError = True

    On Error GoTo errhandler

    'TODO
    'DeseaGuardarMapa Dialog.filename

    ObtenerNombreArchivo False

    If Len(Dialog.filename) < 3 Then Exit Sub

    'TODO
    'If WalkMode = True Then
    '    Call modGeneral.ToggleWalkMode

    'End If
    
    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapa Dialog.filename
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
errhandler:

End Sub

Private Sub cmdQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdQuitarFunciones_Click_Err
    
    'TODO
    'Call mnuQuitarFunciones_Click
    
    Exit Sub

cmdQuitarFunciones_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.cmdQuitarFunciones_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAutoQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuAutoQuitarFunciones_Click_Err
    
    mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

    
    Exit Sub

mnuAutoQuitarFunciones_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuAutoQuitarFunciones_Click", Erl)
    Resume Next
    
End Sub

Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    With Dialog

        .Filter = "Mapas de NexusAO (*.csm)|*.csm"

        If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .filename = vbNullString
            .flags = cdlOFNPathMustExist
            .ShowSave
        Else
            .DialogTitle = "Cargar"
            .filename = vbNullString
            .flags = cdlOFNFileMustExist
            .ShowOpen

        End If

    End With

End Sub

Private Sub mnuMinimapa_Click()
    frmMiniMapa.Show , Me
End Sub

Public Sub SelectPanel_Click(Index As Integer)
    
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo SelectPanel_Click_Err
    
    Call VerFuncion(Index)

    'TODO
    'If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click

    Exit Sub

SelectPanel_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.SelectPanel_Click", Erl)
    Resume Next
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo MainViewPic_MouseMove_Err

    MouseX = X
    MouseY = Y

    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    Exit Sub

MainViewPic_MouseMove_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.MainViewPic_MouseMove", Erl)
    Resume Next
    
End Sub
