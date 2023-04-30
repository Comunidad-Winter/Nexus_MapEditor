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
      Index           =   9
      Left            =   8430
      TabIndex        =   21
      Top             =   90
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":0000
      cBack           =   -2147483643
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Super."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":30052
      cBack           =   -2147483633
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
      TabIndex        =   11
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
      Left            =   975
      TabIndex        =   13
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Trasl."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":30CA4
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   2
      Left            =   1890
      TabIndex        =   14
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Bloq."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":318F6
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   3
      Left            =   2820
      TabIndex        =   15
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "NPC's"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":32548
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   4
      Left            =   3750
      TabIndex        =   16
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Obj."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":3319A
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   5
      Left            =   4680
      TabIndex        =   17
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Trigg."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":33DEC
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   6
      Left            =   5610
      TabIndex        =   18
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Partic."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":34A3E
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   7
      Left            =   6540
      TabIndex        =   19
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Luces"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":350C0
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H SelectPanel 
      Height          =   375
      Index           =   8
      Left            =   7470
      TabIndex        =   20
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Bordes"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":35562
      cBack           =   -2147483633
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   697
      X2              =   697
      Y1              =   5.961
      Y2              =   33.78
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   700
      X2              =   700
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
      Left            =   18330
      TabIndex        =   10
      Top             =   180
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
      Left            =   10680
      TabIndex        =   9
      Top             =   180
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
      Left            =   17565
      TabIndex        =   8
      Top             =   180
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
      Left            =   16800
      TabIndex        =   7
      Top             =   180
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
      Left            =   16035
      TabIndex        =   6
      Top             =   180
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
      Left            =   15270
      TabIndex        =   5
      Top             =   180
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
      Left            =   14520
      TabIndex        =   4
      Top             =   180
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
      Left            =   13740
      TabIndex        =   3
      Top             =   180
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
      Left            =   12975
      TabIndex        =   2
      Top             =   180
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
      Left            =   12210
      TabIndex        =   1
      Top             =   180
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
      Left            =   11445
      TabIndex        =   0
      Top             =   180
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
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuArchivoLine0 
         Caption         =   "-"
      End
      Begin VB.Menu menuSalir 
         Caption         =   "&Salir"
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
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
      End
      Begin VB.Menu mnuLineEdicion2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto completar superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuMinimapa 
         Caption         =   "&Minimapa"
      End
      Begin VB.Menu mnuConsola 
         Caption         =   "&Consola"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
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
   Begin VB.Menu mnuopciones 
      Caption         =   "Opciones"
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad Caminata"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "Actualizar indices"
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    clicX = X
    clicY = y
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    
    
    Exit Sub

Form_KeyDown_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.Form_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    
    On Error GoTo Form_QueryUnload_Err
    
    Dim ConfigFile As String
    
    ConfigFile = DirInit & "Config.ini"

    WriteVar ConfigFile, "PATH", "UltimoMapa", Dialog.FileName
    WriteVar ConfigFile, "MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "Translados", IIf(frmMain.mnuVerTranslados.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "Grilla", IIf(frmMain.mnuVerGrilla.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked = True, "1", "0")
    WriteVar ConfigFile, "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.y

    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

    Exit Sub

Form_QueryUnload_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.Form_QueryUnload", Erl)
    Resume Next
    
End Sub

Private Sub MapPest_Click(Index As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2023
    '*************************************************
    
    On Error GoTo MapPest_Click_Err

    Dim Mapa As Integer
    Mapa = Index + NumMap_Save - 4

    MapaActual = Mapa

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then _
            modMapIO.GuardarMapa Dialog.FileName

    End If

    Dialog.FileName = PATH_Save & NameMap_Save & Mapa & ".csm"

    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    Call modMapIO.NuevoMapa
    DoEvents
    modMapIO.AbrirMapa Dialog.FileName
    EngineRun = True
    Exit Sub

    Exit Sub

ErrHandler:
    MsgBox Err.Description

    
    Exit Sub

MapPest_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.MapPest_Click", Erl)
    Resume Next

End Sub

Public Sub NextMap()
    Call MapPest_Click(5)
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

    On Error GoTo ErrHandler

    DeseaGuardarMapa Dialog.FileName

    ObtenerNombreArchivo False

    If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then _
        Call modGeneral.ToggleWalkMode
    
    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapa Dialog.FileName
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuActualizarIndices_Click()
    
    On Error GoTo mnuActualizarIndices_Click_Err
    
    frmActualizarIndices.Show , Me

    Exit Sub

mnuActualizarIndices_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuActualizarIndices_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuConfigAvanzada_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuConfigAvanzada_Click_Err
    
    frmConfigSup.Show

    Exit Sub

mnuConfigAvanzada_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuConfigAvanzada_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuConsola_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2023
    '*************************************************

    If Not frmConsola.Visible Then
        frmConsola.Show , frmMain
    Else
        frmConsola.Visible = False
    End If

End Sub

Public Sub mnuGuardarMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    On Error GoTo mnuGuardarMapa_Click_Err

    modMapIO.GuardarMapa Dialog.FileName
    
    Exit Sub

mnuGuardarMapa_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarMapaComo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    
    On Error GoTo mnuGuardarMapaComo_Click_Err
    
    modMapIO.GuardarMapa
    
    Exit Sub

mnuGuardarMapaComo_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapaComo_Click", Erl)
    Resume Next
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
            .FileName = vbNullString
            .flags = cdlOFNPathMustExist
            .ShowSave
        Else
            .DialogTitle = "Cargar"
            .FileName = vbNullString
            .flags = cdlOFNFileMustExist
            .ShowOpen

        End If

    End With

End Sub

Private Sub mnuMinimapa_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2023
    '*************************************************
    
    If Not frmMiniMapa.Visible Then
        frmMiniMapa.Show , Me
    Else
        frmMiniMapa.Visible = False
    End If
End Sub

Private Sub mnuModoCaminata_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo mnuModoCaminata_Click_Err
    
    ToggleWalkMode
    
    Exit Sub

mnuModoCaminata_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuModoCaminata_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuNuevoMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuNuevoMapa_Click_Err
    

    Dim loopc As Integer

    DeseaGuardarMapa Dialog.FileName

    For loopc = 0 To frmMain.MapPest.Count
        frmMain.MapPest(loopc).Visible = False
    Next

    frmMain.Dialog.FileName = Empty

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If

    Call modMapIO.NuevoMapa

    Call SelectPanel_Click(9)

    Exit Sub

mnuNuevoMapa_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuNuevoMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuPreview_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2023
    '*************************************************
    
    If Not frmPreview.Visible Then
        frmPreview.Show , Me
    Else
        frmPreview.Visible = False
    End If
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

Private Sub MainViewPic_Click()
    
    On Error GoTo MainViewPic_Click_Err

    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    UltimoClickX = tX
    UltimoClickY = tY

    Exit Sub

MainViewPic_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.MainViewPic_Click", Erl)
    Resume Next
    
End Sub

Private Sub MainViewPic_DblClick()
    
    On Error GoTo MainViewPic_DblClick_Err
    
    Dim tX As Integer
    Dim tY As Integer

    If Not MapaCargado Then Exit Sub

    If SobreX > 0 And SobreY > 0 Then _
        DobleClick Val(SobreX), Val(SobreY)
    
    Exit Sub

MainViewPic_DblClick_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.MainViewPic_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo MainViewPic_MouseMove_Err

    MouseX = X
    MouseY = y

    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    Exit Sub

MainViewPic_MouseMove_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.MainViewPic_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    On Error GoTo MainViewPic_MouseDown_Err
    

    If Not MapaCargado Then Exit Sub

    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

    'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
    If Shift = 1 And Button = 1 Then
        'TODO
        'Seleccionando = True
        'SeleccionIX = tX '+ UserPos.X
        'SeleccionIY = tY '+ UserPos.Y
    Else
        ClickEdit Button, tX, tY

    End If

    
    Exit Sub

MainViewPic_MouseDown_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.MainViewPic_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerAutomatico_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerAutomatico_Click_Err
    
    mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)
    
    Exit Sub

mnuVerAutomatico_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerAutomatico_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerBloqueos_Click_Err
    
    frmBloqueos.cVerBloqueos.value = (frmBloqueos.cVerBloqueos.value = False)
    mnuVerBloqueos.Checked = frmBloqueos.cVerBloqueos.value

    Exit Sub

mnuVerBloqueos_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerBloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa1_Click()
    
    On Error GoTo mnuVerCapa1_Click_Err
    
    mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)

    Exit Sub

mnuVerCapa1_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerCapa1_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa2_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa2_Click_Err
    
    mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)

    Exit Sub

mnuVerCapa2_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerCapa2_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa3_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa3_Click_Err
    
    mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)

    Exit Sub

mnuVerCapa3_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerCapa3_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerCapa4_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerCapa4_Click_Err
    
    mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)

    Exit Sub

mnuVerCapa4_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerCapa4_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerGrilla_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    
    On Error GoTo mnuVerGrilla_Click_Err
    
    VerGrilla = (VerGrilla = False)
    mnuVerGrilla.Checked = VerGrilla

    Exit Sub

mnuVerGrilla_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerGrilla_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerLuces_Click()
    
    On Error GoTo mnuVerLuces_Click_Err
    
    'mnuVerLuces.Checked = (mnuVerLuces.Checked = False)
    
    Exit Sub

mnuVerLuces_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerLuces_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerNPCs_Click_Err
    
    mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

    Exit Sub

mnuVerNPCs_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerNPCs_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerObjetos_Click_Err
    
    mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)
    
    Exit Sub

mnuVerObjetos_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerObjetos_Click", Erl)
    Resume Next
    
End Sub

Public Sub mnuVerParticulas_Click()
    '************************************************
    'Author: Lorwik
    'Last modified: ???
    '************************************************
    
    On Error GoTo mnuVerParticulas_Click_Err
    
    mnuVerParticulas.Checked = (mnuVerParticulas.Checked = False)
    
    Exit Sub

mnuVerParticulas_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerParticulas_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo mnuVerTranslados_Click_Err
    
    mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)
    
    Exit Sub

mnuVerTranslados_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerTranslados_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuVerTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuVerTriggers_Click_Err
    
    frmTriggers.cVerTriggers.value = (frmTriggers.cVerTriggers.value = False)
    mnuVerTriggers.Checked = frmTriggers.cVerTriggers.value

    
    Exit Sub

mnuVerTriggers_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.mnuVerTriggers_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuInsertarSuperficieAlAzar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Superficie_Azar
End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Superficie_Bordes
End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Superficie_Todo
End Sub
