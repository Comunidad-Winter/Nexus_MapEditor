VERSION 5.00
Begin VB.Form frmParticulas 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Particulas"
   ClientHeight    =   4455
   ClientLeft      =   19710
   ClientTop       =   5970
   ClientWidth     =   3390
   ControlBox      =   0   'False
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
   ScaleHeight     =   4455
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cParticula 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   1350
      TabIndex        =   3
      Top             =   3330
      Width           =   1935
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
      ItemData        =   "frmParticulas.frx":0000
      Left            =   30
      List            =   "frmParticulas.frx":0002
      TabIndex        =   0
      Tag             =   "-1"
      Top             =   0
      Width           =   3375
   End
   Begin Nexus_MapEditor.lvButtons_H cSeleccionarParticula 
      Height          =   525
      Left            =   1680
      TabIndex        =   1
      Top             =   3780
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   926
      Caption         =   "Insertar Particula"
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
   Begin Nexus_MapEditor.lvButtons_H cQuitarParticula 
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   3780
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   926
      Caption         =   "Quitar Particula"
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
   Begin VB.Label lbGrh 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Particula Actual:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3390
      Width           =   1170
   End
End
Attribute VB_Name = "frmParticulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cQuitarParticula_Click()
'**********************************
'Author: Lorwik
'Ultima Modificación: 01/05/2023
'**********************************

    On Error GoTo cQuitarParticula_Click_Err

    If cQuitarParticula.value = True Then
        cSeleccionarParticula.Enabled = False
    Else
        cSeleccionarParticula.Enabled = True

    End If
    
    Exit Sub

cQuitarParticula_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmParticulas.cQuitarParticula_Click", Erl)
    Resume Next
End Sub

Private Sub cSeleccionarParticula_Click()
'**********************************
'Author: Lorwik
'Ultima Modificación: 01/05/2023
'**********************************

    On Error GoTo cSeleccionarParticula_Click_Err
    
    modEdicion.Deshacer_Add "Insertar Particula" ' Hago deshacer
    
    If cSeleccionarParticula.value = True Then
        cQuitarParticula.Enabled = False
    Else
        cQuitarParticula.Enabled = True

    End If
    
    Exit Sub

cSeleccionarParticula_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmParticulas.cSeleccionarParticula_Click", Erl)
    Resume Next
    
End Sub

Private Sub lListado_Click()
'**********************************
'Author: Lorwik
'Ultima Modificación: 01/05/2023
'**********************************
    On Error GoTo lListado_Click_Err

    If HotKeysAllow = False Then _
        lListado.Tag = lListado.ListIndex

    cParticula.Text = lListado.Text
    
    Call fPreviewGrh(cParticula.Text)
    Call modPaneles.RenderPreview
    
    Exit Sub
    
lListado_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmParticulas.lListado_Click", Erl)
    Resume Next
End Sub
