VERSION 5.00
Begin VB.Form frmCopiarBordes 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copiar Bordes"
   ClientHeight    =   3345
   ClientLeft      =   10395
   ClientTop       =   13350
   ClientWidth     =   6120
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   223
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraCopiarAutomatico 
      BackColor       =   &H80000006&
      Caption         =   "Copiar Manual"
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
      Height          =   3165
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   1995
      Begin Nexus_MapEditor.lvButtons_H LvBCopiar 
         Height          =   465
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   820
         Caption         =   "Copiar Norte"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBCopiar 
         Height          =   1485
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   2619
         Caption         =   "Copiar Oeste"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBCopiar 
         Height          =   1485
         Index           =   2
         Left            =   1050
         TabIndex        =   9
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   2619
         Caption         =   "Copiar Este"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBCopia 
         Height          =   465
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   2550
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   820
         Caption         =   "Copiar Sur"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame FraCopiarAuto 
      BackColor       =   &H80000006&
      Caption         =   "Copiar Automatico"
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
      Height          =   3165
      Left            =   2100
      TabIndex        =   0
      Top             =   120
      Width           =   3945
      Begin Nexus_MapEditor.lvButtons_H LvBCopiarManual 
         Height          =   465
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   420
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   820
         Caption         =   "Copiar Norte"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBCopiarManual 
         Height          =   1575
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   930
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   2778
         Caption         =   "Copiar Oeste"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBCopiarManual 
         Height          =   1545
         Index           =   2
         Left            =   2970
         TabIndex        =   3
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   2725
         Caption         =   "Copiar Este"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBCopiarManual 
         Height          =   465
         Index           =   3
         Left            =   960
         TabIndex        =   4
         Top             =   2460
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   820
         Caption         =   "Copiar Sur"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBPegar 
         Height          =   795
         Index           =   0
         Left            =   1530
         TabIndex        =   5
         Top             =   1260
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         Caption         =   "Pegar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBPegar 
         Height          =   795
         Index           =   1
         Left            =   1530
         TabIndex        =   15
         Top             =   1260
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         Caption         =   "Pegar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBPegar 
         Height          =   795
         Index           =   2
         Left            =   1530
         TabIndex        =   16
         Top             =   1260
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         Caption         =   "Pegar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Nexus_MapEditor.lvButtons_H LvBPegar 
         Height          =   795
         Index           =   3
         Left            =   1530
         TabIndex        =   17
         Top             =   1260
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         Caption         =   "Pegar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblMapaEste 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Este"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   1350
         TabIndex        =   14
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Label lblMapaSur 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sur"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   690
         TabIndex        =   13
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label lblMapaNorte 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Norte"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblMapaOeste 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oeste"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   990
         TabIndex        =   11
         Top             =   1530
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCopiarBordes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo Form_Load_Err
    
    Call VerMapaTraslado

    If lblMapaSur.Caption = "Sur" Then LvBCopiarManual(3).Visible = False
    If lblMapaEste.Caption = "Este" Then LvBCopiarManual(2).Visible = False
    If lblMapaOeste.Caption = "Oeste" Then LvBCopiarManual(1).Visible = False
    If lblMapaOeste.Caption = "Norte" Then LvBCopiarManual(0).Visible = False
    
    Exit Sub

Form_Load_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.Form_Load", Erl)
    Resume Next
End Sub

Private Sub LvBCopiar_Click(index As Integer)
    
    Select Case index
    
        Case 0 'Norte
            Call Copiar(0)
            Call LvBPegar_Click(0)
            Call Copiar(3)
            Call LvBPegar_Click(3)
            
        Case 1 'Oeste
            Call Copiar(1)
            Call LvBPegar_Click(1)
            Call Copiar(2)
            Call LvBPegar_Click(2)
            
        Case 2 'Este
            Call Copiar(2)
            Call LvBPegar_Click(2)
            Call Copiar(1)
            Call LvBPegar_Click(1)
            
        Case 3 'Sur
            Call Copiar(3)
            Call LvBPegar_Click(3)
            Call Copiar(0)
            Call LvBPegar_Click(0)
    End Select
    
End Sub

Private Sub LvBCopiarManual_Click(index As Integer)

    On Error GoTo LvBCopiarManual_Click_Err
        
    Call Copiar(index)
    
    Exit Sub

LvBCopiarManual_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.LvBCopiarManual_Click", Erl)
    Resume Next
End Sub

Private Sub Copiar(index As Integer)
    On Error GoTo Copiar_Err

    Select Case index
    
        Case 0 'Norte
            SeleccionIX = 1
            SeleccionFX = 100
            SeleccionIY = 11
            SeleccionFY = 22
            LvBPegar(0).Visible = True
            LvBPegar(1).Visible = False
            LvBPegar(2).Visible = False
            LvBPegar(3).Visible = False

            
        Case 1 'Oeste
            SeleccionIX = 14
            SeleccionFX = 27
            SeleccionIY = 1
            SeleccionFY = 100
            LvBPegar(0).Visible = False
            LvBPegar(1).Visible = True
            LvBPegar(2).Visible = False
            LvBPegar(3).Visible = False
            
        Case 2 'Este
            SeleccionIX = 75
            SeleccionFX = 87
            SeleccionIY = 1
            SeleccionFY = 100
            LvBPegar(0).Visible = False
            LvBPegar(1).Visible = False
            LvBPegar(2).Visible = True
            LvBPegar(3).Visible = False
            
        Case 3 'Sur
            SeleccionIX = 1
            SeleccionFX = 100
            SeleccionIY = 81
            SeleccionFY = 90
            LvBPegar(0).Visible = False
            LvBPegar(1).Visible = False
            LvBPegar(2).Visible = False
            LvBPegar(3).Visible = True
    
    End Select
    
    LvBCopiarManual(0).Visible = False
    LvBCopiarManual(1).Visible = False
    LvBCopiarManual(2).Visible = False
    LvBCopiarManual(3).Visible = False

    Call CopiarSeleccion
    MapInfo.Changed = 1
    frmMain.mnuGuardarMapa_Click
    
    Exit Sub

Copiar_Err:
    Call LogError(Err.Number, Err.Description, "frmCopiarBordes.Copiar", Erl)
    Resume Next
End Sub

Private Sub LvBPegar_Click(index As Integer)

    On Error GoTo LvBPegar_Click_Err

    Call Pegar(index)
    
Exit Sub

LvBPegar_Click_Err:
Call LogError(Err.Number, Err.Description, "frmCopiarBordes.LvBPegar_Click", Erl)

Resume Next

End Sub

Private Sub Pegar(index As Integer)
    On Error GoTo Pegar_Err

    Select Case index
    
        Case 0

            If lblMapaNorte.Caption <> "Norte" Then
                If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaNorte.Caption) & ".csm", vbArchive) = True Then
                    Call modMapIO.NuevoMapa
                    frmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaNorte.Caption) & ".csm"
                    modMapIO.AbrirMapa frmMain.Dialog.FileName
    
                    frmMain.mnuReAbrirMapa.Enabled = True

                End If
                
            Else
                Exit Sub
                
            End If
        
        Case 1
        
            If lblMapaOeste.Caption <> "Oeste" Then
                If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaOeste.Caption) & ".csm", vbArchive) = True Then
                    Call modMapIO.NuevoMapa
                    frmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaOeste.Caption) & ".csm"
                    modMapIO.AbrirMapa frmMain.Dialog.FileName
    
                    frmMain.mnuReAbrirMapa.Enabled = True

                End If
                
            Else
                Exit Sub
    
            End If

        Case 2

            If lblMapaEste.Caption <> "Este" Then
                If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaEste.Caption) & ".csm", vbArchive) = True Then
                    Call modMapIO.NuevoMapa
                    frmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaEste.Caption) & ".csm"
                    modMapIO.AbrirMapa frmMain.Dialog.FileName
    
                    frmMain.mnuReAbrirMapa.Enabled = True

                End If
                
            Else
                
                Exit Sub

            End If
            
        Case 3

            If lblMapaSur.Caption <> "Sur" Then
                If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaSur.Caption) & ".csm", vbArchive) = True Then
                    Call modMapIO.NuevoMapa
                    frmMain.Dialog.FileName = PATH_Save & NameMap_Save & CLng(lblMapaSur.Caption) & ".csm"
                    modMapIO.AbrirMapa frmMain.Dialog.FileName
    
                    frmMain.mnuReAbrirMapa.Enabled = True

                End If
                
            Else
                Exit Sub

            End If

    End Select

    SobreX = 1
    SobreY = 1
    Call PegarSeleccion
    Call modEdicion.Bloquear_Bordes
    MapInfo.Changed = 1
    UserPos.X = 19
    Unload Me
    
Exit Sub

Pegar_Err:
Call LogError(Err.Number, Err.Description, "frmCopiarBordes.Pegar", Erl)

Resume Next
End Sub

Private Sub VerMapaTraslado()
    
    On Error GoTo VerMapaTraslado_Err
    
    Dim X As Integer
    Dim y As Integer

    'Izquierda
    X = 13

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
    X = 88

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
