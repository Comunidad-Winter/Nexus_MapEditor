VERSION 5.00
Begin VB.Form frmDesplazarMapa 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desplazar Mapa"
   ClientHeight    =   2985
   ClientLeft      =   17850
   ClientTop       =   11070
   ClientWidth     =   4380
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
   ScaleHeight     =   2985
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraAutomatizar 
      BackColor       =   &H80000006&
      Caption         =   "Automatizar"
      ForeColor       =   &H8000000B&
      Height          =   1695
      Left            =   2370
      TabIndex        =   9
      Top             =   180
      Width           =   1875
      Begin VB.CheckBox chkAutomatizar 
         BackColor       =   &H80000006&
         Caption         =   "Automatizar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   270
         TabIndex        =   14
         Top             =   1230
         Width           =   1335
      End
      Begin VB.TextBox txtMapMax 
         Height          =   315
         Left            =   660
         TabIndex        =   13
         Text            =   "168"
         Top             =   690
         Width           =   1035
      End
      Begin VB.TextBox txtMapMin 
         Height          =   315
         Left            =   660
         TabIndex        =   11
         Text            =   "1"
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         ForeColor       =   &H8000000B&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         ForeColor       =   &H8000000B&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   510
      End
   End
   Begin VB.TextBox txtAlto 
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Text            =   "100"
      Top             =   1530
      Width           =   645
   End
   Begin VB.TextBox txtAncho 
      Height          =   315
      Left            =   1650
      TabIndex        =   6
      Text            =   "100"
      Top             =   1080
      Width           =   615
   End
   Begin Nexus_MapEditor.lvButtons_H LvBDesplazar 
      Height          =   405
      Left            =   540
      TabIndex        =   4
      Top             =   2460
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   714
      Caption         =   "Desplazar"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox Ydest 
      Height          =   315
      Left            =   1230
      TabIndex        =   3
      Text            =   "15"
      Top             =   660
      Width           =   1035
   End
   Begin VB.TextBox Xdest 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "15"
      Top             =   210
      Width           =   1065
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esperando..."
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   2070
      Width           =   3855
   End
   Begin VB.Label lblAltoA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alto a desplazar:"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   210
      TabIndex        =   7
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label lblAnchoA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho a desplazar:"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   1110
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y de destino:"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   690
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X de destino:"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmDesplazarMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LvBDesplazar_Click()
'****************************************
'Author: Lorwik
'Fecha: 01/05/2023
'****************************************

    Dim i As Integer

    If chkAutomatizar.value = Unchecked Then
        Call DesplazarMapa
        
    Else
    
        For i = Val(txtMapMin.Text) To Val(txtMapMax.Text)
        
            If FileExist(DirMapas & "mapa" & i & ".csm", vbNormal) = True Then
            
                Call modMapIO.NuevoMapa
                Call modMapIO.MapaCSM_Cargar(DirMapas & "mapa" & i & ".csm")
                DoEvents
                Call DesplazarMapa
                DoEvents
                NoSobreescribir = True
                Call modMapIO.MapaCSM_Guardar(DirMapas & "mapa" & i & ".csm")
            
                lblInfo.Caption = "Mapa" & i & " convertido correctamente!"
                
            Else
                lblInfo.Caption = "Mapa" & i & ".csm no existe!"
                
            End If
        
        Next i
    
    End If
    
End Sub

Private Sub DesplazarMapa()
'****************************************
'Author: Lorwik
'Fecha: 01/05/2023
'****************************************

    Dim X As Integer
    Dim Y As Integer
    Dim i As Byte
    Dim Vacio As MapBlock
    
    Working = True
    
    ReDim SeleccionMap(Val(txtAncho.Text), Val(txtAlto.Text)) As MapBlock
    
    'Copiamos el trozo de mapa
    For X = 1 To Val(txtAncho.Text)
        For Y = 1 To Val(txtAlto.Text)
        
            SeleccionMap(X, Y) = MapData(X, Y)
        
        Next Y
    Next X
    
    'Borramos el trozo de mapa
    For X = 1 To Val(txtAncho.Text)
        For Y = 1 To Val(txtAlto.Text)
        
            MapData(X, Y) = Vacio
            
        Next Y
    Next X
    
    'Pegamos el trozo de mapa en la nueva ubicación
    For X = 1 To Val(txtAncho.Text)
        For Y = 1 To Val(txtAlto.Text)
            MapData(X + Val(Xdest.Text), Y + Val(Ydest.Text)) = SeleccionMap(X, Y)
        Next Y
    Next X
    
    Working = False
    
End Sub
