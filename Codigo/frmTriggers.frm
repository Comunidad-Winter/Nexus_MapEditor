VERSION 5.00
Begin VB.Form frmTriggers 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Triggers"
   ClientHeight    =   4155
   ClientLeft      =   23415
   ClientTop       =   6900
   ClientWidth     =   4110
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
   ScaleHeight     =   4155
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
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
      ItemData        =   "frmTriggers.frx":0000
      Left            =   0
      List            =   "frmTriggers.frx":0002
      TabIndex        =   1
      Tag             =   "-1"
      Top             =   0
      Width           =   4125
   End
   Begin Nexus_MapEditor.lvButtons_H cQuitarTrigger 
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   3720
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   661
      Caption         =   "Quitar Trigger's"
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
   Begin Nexus_MapEditor.lvButtons_H cVerTriggers 
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   3330
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   661
      Caption         =   "Mostrar Trigger's"
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
   Begin Nexus_MapEditor.lvButtons_H cInsertarTrigger 
      Height          =   765
      Left            =   2160
      TabIndex        =   3
      Top             =   3330
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1349
      Caption         =   "Insertar Trigger's"
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
End
Attribute VB_Name = "frmTriggers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cverTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cverTriggers_Click_Err
    
    frmMain.mnuVerTriggers.Checked = cVerTriggers.value
    
    Exit Sub

cverTriggers_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmTriggers.cverTriggers_Click", Erl)
    Resume Next
    
End Sub
