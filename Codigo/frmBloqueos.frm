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
      _extentx        =   4895
      _extenty        =   1085
      caption         =   "Insertar Bloqueos"
      capalign        =   2
      backstyle       =   2
      font            =   "frmBloqueos.frx":0000
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cVerBloqueos 
      Height          =   435
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   2775
      _extentx        =   4895
      _extenty        =   767
      caption         =   "Mostrar Bloqueos"
      capalign        =   2
      backstyle       =   2
      font            =   "frmBloqueos.frx":0028
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cQuitarBloqueo 
      Height          =   615
      Left            =   90
      TabIndex        =   2
      Top             =   1350
      Width           =   2775
      _extentx        =   4895
      _extenty        =   1085
      caption         =   "Quitar Bloqueos"
      capalign        =   2
      backstyle       =   2
      font            =   "frmBloqueos.frx":0050
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      cback           =   -2147483633
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
