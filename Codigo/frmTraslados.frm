VERSION 5.00
Begin VB.Form frmTraslados 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Traslados"
   ClientHeight    =   4005
   ClientLeft      =   23415
   ClientTop       =   13395
   ClientWidth     =   3630
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
   ScaleHeight     =   4005
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tTMapa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Text            =   "1"
      Top             =   210
      Width           =   2265
   End
   Begin VB.TextBox tTX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Text            =   "1"
      Top             =   570
      Width           =   2265
   End
   Begin VB.TextBox tTY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Text            =   "1"
      Top             =   930
      Width           =   2265
   End
   Begin Nexus_MapEditor.lvButtons_H cInsertarTrans 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   1410
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      Caption         =   "Insertar Traslado"
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
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cInsertarTransOBJ 
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   1830
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      Caption         =   "Colocar automaticamente Objeto"
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
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cUnionManual 
      Height          =   375
      Left            =   180
      TabIndex        =   5
      Top             =   2310
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      Caption         =   "Union con Mapa Adyacente (manual)"
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
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cUnionAuto 
      Height          =   375
      Left            =   180
      TabIndex        =   9
      Top             =   2700
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      Caption         =   "Union con Mapas Adyacentes (auto)"
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
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cQuitarTrans 
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   3540
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      Caption         =   "Quitar Traslados"
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
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H LvBDesplazarTraslados 
      Height          =   375
      Left            =   180
      TabIndex        =   11
      Top             =   3090
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      Caption         =   "Desplazar Traslados"
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
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label lMapN 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   210
      TabIndex        =   3
      Top             =   270
      Width           =   435
   End
   Begin VB.Label lXhor 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "X horizontal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   210
      TabIndex        =   2
      Top             =   630
      Width           =   900
   End
   Begin VB.Label lYver 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Y vertical:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   210
      TabIndex        =   1
      Top             =   990
      Width           =   735
   End
End
Attribute VB_Name = "frmTraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cInsertarTrans_Click()
    
    On Error GoTo cInsertarTrans_Click_Err
    
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    If cInsertarTrans.value = True Then
        cQuitarTrans.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        cQuitarTrans.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If
    
    Exit Sub

cInsertarTrans_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.cInsertarTrans_Click", Erl)
    Resume Next
    
End Sub

Private Sub cUnionManual_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cUnionManual_Click_Err
    
    cInsertarTrans.value = (cUnionManual.value = True)
    Call cInsertarTrans_Click

    Exit Sub

cUnionManual_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.cUnionManual_Click", Erl)
    Resume Next
    
End Sub

Private Sub cUnionAuto_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    'Call MapPest_Click(4)
    
    On Error GoTo cUnionAuto_Click_Err
    
    frmUnionAdyacente.Show

    Exit Sub

cUnionAuto_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.cUnionAuto_Click", Erl)
    Resume Next
    
End Sub

Private Sub LvBDesplazarTraslados_Click()
    '********************************
    'Author: Lorwik
    '
    '********************************
    
    frmDesplazarTranslados.Show , frmMain
    
End Sub
