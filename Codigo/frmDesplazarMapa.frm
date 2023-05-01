VERSION 5.00
Begin VB.Form frmDesplazarMapa 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desplazar Mapa"
   ClientHeight    =   2565
   ClientLeft      =   17850
   ClientTop       =   11070
   ClientWidth     =   2340
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
   ScaleHeight     =   2565
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAlto 
      Height          =   315
      Left            =   1590
      TabIndex        =   8
      Text            =   "100"
      Top             =   1530
      Width           =   645
   End
   Begin VB.TextBox txtAncho 
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Text            =   "100"
      Top             =   1080
      Width           =   615
   End
   Begin Nexus_MapEditor.lvButtons_H LvBDesplazar 
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2085
      _ExtentX        =   3678
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
      Left            =   1200
      TabIndex        =   3
      Text            =   "10"
      Top             =   660
      Width           =   1035
   End
   Begin VB.TextBox Xdest 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Text            =   "10"
      Top             =   210
      Width           =   1035
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

    Dim x As Integer
    Dim y As Integer
    Dim i As Byte
    Dim Vacio As MapBlock
    
    ReDim SeleccionMap(Val(txtAncho.Text), Val(txtAlto.Text)) As MapBlock
    
    'Copiamos el trozo de mapa
    For x = 1 To Val(txtAncho.Text)
        For y = 1 To Val(txtAlto.Text)
        
            SeleccionMap(x, y) = MapData(x, y)
        
        Next y
    Next x
    
    'Borramos el trozo de mapa
    For x = 1 To Val(txtAncho.Text)
        For y = 1 To Val(txtAlto.Text)
        
            MapData(x, y) = Vacio
            
        Next y
    Next x
    
    'Pegamos el trozo de mapa en la nueva ubicación
    For x = 1 To Val(txtAncho.Text)
        For y = 1 To Val(txtAlto.Text)
            MapData(x + Val(Xdest.Text), y + Val(Ydest.Text)) = SeleccionMap(x, y)
        Next y
    Next x
    
End Sub
