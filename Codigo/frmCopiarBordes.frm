VERSION 5.00
Begin VB.Form frmCopiarBordes 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copiar Translados manual"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5310
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPegarEste 
      Caption         =   "Pegar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1980
      TabIndex        =   8
      Top             =   1710
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdPegarOeste 
      Caption         =   "Pegar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1980
      TabIndex        =   7
      Top             =   1710
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdPegarSur 
      Caption         =   "Pegar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2010
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdPegarNorte 
      Caption         =   "Pegar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1980
      TabIndex        =   5
      Top             =   1740
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSur 
      Caption         =   "Copiar al Sur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1740
      TabIndex        =   4
      Top             =   3060
      Width           =   1935
   End
   Begin VB.CommandButton cmdNorte 
      Caption         =   "Copiar al Norte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1740
      TabIndex        =   3
      Top             =   420
      Width           =   1935
   End
   Begin VB.CommandButton cmdEste 
      Caption         =   "Copiar al Este"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4230
      TabIndex        =   2
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdOeste 
      Caption         =   "Copiar al Oeste"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblMapaOeste 
      BackStyle       =   0  'Transparent
      Caption         =   "Oeste"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1380
      TabIndex        =   12
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label lblMapaEste 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Este"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2580
      TabIndex        =   11
      Top             =   1860
      Width           =   1455
   End
   Begin VB.Label lblMapaSur 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1380
      TabIndex        =   10
      Top             =   2700
      Width           =   2535
   End
   Begin VB.Label lblMapaNorte 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Norte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1740
      TabIndex        =   9
      Top             =   1260
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COPIAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1380
      TabIndex        =   0
      Top             =   60
      Width           =   2535
   End
End
Attribute VB_Name = "frmCopiarBordes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdNorte_Click()
    'Superior
    
    On Error GoTo Command1_Click_Err
    
    'Call VerMapaTraslado

    SeleccionIX = 1
    SeleccionFX = 100
    SeleccionIY = 11
    SeleccionFY = 22
    cmdNorte.Visible = False
    cmdOeste.Visible = False
    cmdEste.Visible = False
    cmdSur.Visible = False
    cmdPegarNorte.Visible = True
    cmdPegarSur.Visible = False
    cmdPegarOeste.Visible = False
    cmdPegarEste.Visible = False
    Call CopiarSeleccion
    MapInfo.Changed = 1
    frmMain.mnuGuardarMapa_Click

    
    Exit Sub

Command1_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.cmdNorte_Click", Erl)
    Resume Next
    
End Sub

Public Sub cmdOeste_Click()
    'copiar izquierdo
    
    On Error GoTo Command2_Click_Err
    
    SeleccionIX = 12
    SeleccionFX = 24
    SeleccionIY = 1
    SeleccionFY = 100
    cmdNorte.Visible = False
    cmdOeste.Visible = False
    cmdEste.Visible = False
    cmdSur.Visible = False
    cmdPegarNorte.Visible = False
    cmdPegarSur.Visible = False
    cmdPegarOeste.Visible = True
    cmdPegarEste.Visible = False
    Call CopiarSeleccion
    MapInfo.Changed = 1
    frmMain.mnuGuardarMapa_Click

    
    Exit Sub

Command2_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.cmdOeste_Click", Erl)
    Resume Next
    
End Sub

Public Sub cmdEste_Click()
    'copiar derecho
    
    On Error GoTo Command3_Click_Err
    
    SeleccionIX = 76
    SeleccionFX = 87
    SeleccionIY = 1
    SeleccionFY = 100
    cmdNorte.Visible = False
    cmdOeste.Visible = False
    cmdEste.Visible = False
    cmdSur.Visible = False
    cmdPegarNorte.Visible = False
    cmdPegarSur.Visible = False
    cmdPegarOeste.Visible = False
    cmdPegarEste.Visible = True
    Call CopiarSeleccion
    MapInfo.Changed = 1
    frmMain.mnuGuardarMapa_Click

    
    Exit Sub

Command3_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.cmdEste_Click", Erl)
    Resume Next
    
End Sub

Public Sub cmdSur_Click()
    'Copiar inferior ok!
    
    On Error GoTo Command4_Click_Err
    
    SeleccionIX = 1
    SeleccionFX = 100
    SeleccionIY = 81
    SeleccionFY = 90
    cmdNorte.Visible = False
    cmdOeste.Visible = False
    cmdEste.Visible = False
    cmdSur.Visible = False
    cmdPegarNorte.Visible = False
    cmdPegarSur.Visible = True
    cmdPegarOeste.Visible = False
    cmdPegarEste.Visible = False
    Call CopiarSeleccion
    MapInfo.Changed = 1
    frmMain.mnuGuardarMapa_Click

    
    Exit Sub

Command4_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.cmdSur_Click", Erl)
    Resume Next
    
End Sub

Public Sub cmdPegarNorte_Click()
    
    On Error GoTo Command5_Click_Err
    

    'Pegar Inferior OK!
    If lblMapaNorte.Caption <> "Norte" Then
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaNorte.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaNorte.Caption) & ".csm"

            modMapIO.AbrirMapa frmMain.Dialog.FileName, "CSM"
    
            frmMain.mnuReAbrirMapa.Enabled = True

        End If

        SobreX = 1
        SobreY = 91
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        MapInfo.Changed = 1
        UserPos.y = 87
        Unload Me

    End If

    
    Exit Sub

Command5_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.cmdPegarNorte_Click", Erl)
    Resume Next
    
End Sub

Public Sub cmdPegarSur_Click()
    'pegar superior
    
    On Error GoTo Command6_Click_Err
    

    If lblMapaSur.Caption <> "Sur" Then
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaSur.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaSur.Caption) & ".csm"
            modMapIO.AbrirMapa frmMain.Dialog.FileName, "CSM"
    
            frmMain.mnuReAbrirMapa.Enabled = True

        End If

        SobreX = 1
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        MapInfo.Changed = 1
        UserPos.y = 14
        Unload Me

    End If

    
    Exit Sub

Command6_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.cmdPegarSur_Click", Erl)
    Resume Next
    
End Sub

Public Sub cmdPegarOeste_Click()
    
    On Error GoTo Command7_Click_Err
    

    'pegar derecho OK!
    If lblMapaOeste.Caption <> "Oeste" Then
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaOeste.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaOeste.Caption) & ".csm"
            modMapIO.AbrirMapa frmMain.Dialog.FileName, "CSM"
    
            frmMain.mnuReAbrirMapa.Enabled = True

        End If

        SobreX = 89
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        MapInfo.Changed = 1
        UserPos.X = 83
        Unload Me

    End If

    
    Exit Sub

Command7_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.cmdPegarOeste_Click", Erl)
    Resume Next
    
End Sub

Public Sub cmdPegarEste_Click()
    
    On Error GoTo Command8_Click_Err
    

    'pegar izquierdo OK!
    If lblMapaEste.Caption <> "Este" Then
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaEste.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaEste.Caption) & ".csm"
            modMapIO.AbrirMapa frmMain.Dialog.FileName, "CSM"
    
            frmMain.mnuReAbrirMapa.Enabled = True

        End If

        SobreX = 1
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        MapInfo.Changed = 1
        UserPos.X = 19
        Unload Me

    End If

    
    Exit Sub

Command8_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.cmdPegarEste_Click", Erl)
    Resume Next
    
End Sub

Private Sub VerMapaTraslado()
    
    On Error GoTo VerMapaTraslado_Err
    
    Dim X As Integer
    Dim y As Integer

    'Izquierda
    X = 12

    For y = (11) To (90)

        If MapData(X, y).TileExit.Map <> 0 Then
            lblMapaOeste.Caption = MapData(X, y).TileExit.Map
            Exit For
        End If

    Next
    
    'arriba
    y = 10

    For X = (14) To (87)

        If MapData(X, y).TileExit.Map <> 0 Then
            lblMapaNorte.Caption = MapData(X, y).TileExit.Map
            Exit For
        End If

    Next
    
    'Derecha
    X = 89

    For y = (11) To (90)

        If MapData(X, y).TileExit.Map <> 0 Then
            lblMapaEste.Caption = MapData(X, y).TileExit.Map
            Exit For
        End If

    Next
    
    'Abajo
    y = 91

    For X = (14) To (87)

        If MapData(X, y).TileExit.Map <> 0 Then
            lblMapaSur.Caption = MapData(X, y).TileExit.Map
            Exit For
        End If

    Next

    
    Exit Sub

VerMapaTraslado_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.VerMapaTraslado", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call VerMapaTraslado

    If lblMapaSur.Caption = "Sur" Then frmCopiarBordes.cmdSur.Visible = False
    If lblMapaEste.Caption = "Este" Then frmCopiarBordes.cmdEste.Visible = False
    If lblMapaOeste.Caption = "Oeste" Then frmCopiarBordes.cmdOeste.Visible = False
    If lblMapaOeste.Caption = "Norte" Then frmCopiarBordes.cmdNorte.Visible = False

    
    Exit Sub

Form_Load_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.Form_Load", Erl)
    Resume Next
    
End Sub
