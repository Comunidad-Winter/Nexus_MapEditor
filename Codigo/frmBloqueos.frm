VERSION 5.00
Begin VB.Form frmBloqueos 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bloqueos"
   ClientHeight    =   2025
   ClientLeft      =   24330
   ClientTop       =   12000
   ClientWidth     =   3000
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
   ScaleHeight     =   2025
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin Nexus_MapEditor.lvButtons_H cInsertarBloqueo 
      Height          =   615
      Left            =   90
      TabIndex        =   1
      Top             =   690
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      Caption         =   "Insertar Bloqueos"
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
   Begin Nexus_MapEditor.lvButtons_H cVerBloqueos 
      Height          =   435
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   767
      Caption         =   "Mostrar Bloqueos"
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
   Begin Nexus_MapEditor.lvButtons_H cQuitarBloqueo 
      Height          =   615
      Left            =   90
      TabIndex        =   2
      Top             =   1350
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      Caption         =   "Quitar Bloqueos"
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
Attribute VB_Name = "frmBloqueos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cverBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cverBloqueos_Click_Err
    
    frmMain.mnuVerBloqueos.Checked = cVerBloqueos.value

    
    Exit Sub

cverBloqueos_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmBloqueos.cverBloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    cInsertarBloqueo.Tag = vbNullString
    If cInsertarBloqueo.value = True Then
        cQuitarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
        
    Else
        cQuitarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)
        
    End If
End Sub

Private Sub cQuitarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    cInsertarBloqueo.Tag = vbNullString
    If cQuitarBloqueo.value = True Then
        cInsertarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
        
    Else
        cInsertarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)
        
    End If
End Sub

Public Sub InsertarBloqueo()
'*************************************************
'Author: Lorwik
'Last modified: 29/04/2023
'*************************************************

    Call cInsertarBloqueo_Click
    
End Sub
