VERSION 5.00
Begin VB.Form frmSuperficies 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Superficies"
   ClientHeight    =   4935
   ClientLeft      =   9525
   ClientTop       =   9210
   ClientWidth     =   4080
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
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   ShowInTaskbar   =   0   'False
   Begin Nexus_MapEditor.lvButtons_H cQuitarEnTodasLasCapas 
      Height          =   345
      Left            =   90
      TabIndex        =   8
      Top             =   4530
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      Caption         =   "Quitar en Capas 2 y 3"
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
   Begin Nexus_MapEditor.lvButtons_H cQuitarEnEstaCapa 
      Height          =   375
      Left            =   90
      TabIndex        =   7
      Top             =   4110
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Quitar en esta Capa"
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
   Begin Nexus_MapEditor.lvButtons_H cSeleccionarSuperficie 
      Height          =   765
      Left            =   2100
      TabIndex        =   6
      Top             =   4110
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1349
      Caption         =   "Insertar Superficie"
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
   Begin VB.ComboBox cCapas 
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
      Height          =   330
      ItemData        =   "frmSuperficies.frx":0000
      Left            =   1020
      List            =   "frmSuperficies.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3660
      Width           =   855
   End
   Begin VB.ComboBox cFiltro 
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
      Height          =   330
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3300
      Width           =   3285
   End
   Begin VB.ComboBox cGrh 
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
      Height          =   330
      Left            =   2850
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3660
      Width           =   1125
   End
   Begin VB.ListBox lListado 
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
      Height          =   3180
      ItemData        =   "frmSuperficies.frx":001A
      Left            =   -30
      List            =   "frmSuperficies.frx":001C
      TabIndex        =   0
      Tag             =   "-1"
      Top             =   0
      Width           =   4125
   End
   Begin VB.Label lblFiltrar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   3330
      Width           =   480
   End
   Begin VB.Label lbGrh 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Sup Actual:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   3720
      Width           =   825
   End
   Begin VB.Label lbCapas 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Capa Actual:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   3735
      Width           =   930
   End
End
Attribute VB_Name = "frmSuperficies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cSeleccionarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cSeleccionarSuperficie.value = True Then
        cQuitarEnTodasLasCapas.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        cQuitarEnTodasLasCapas.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If
End Sub

Private Sub cQuitarEnEstaCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cQuitarEnEstaCapa.value = True Then
        lListado.Enabled = False
        cFiltro.Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnTodasLasCapas.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        lListado.Enabled = True
        cFiltro.Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnTodasLasCapas.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If
End Sub

Private Sub cQuitarEnTodasLasCapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cQuitarEnTodasLasCapas.value = True Then
        cCapas.Enabled = False
        lListado.Enabled = False
        cFiltro.Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        cCapas.Enabled = True
        lListado.Enabled = True
        cFiltro.Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If
End Sub
