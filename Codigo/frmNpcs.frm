VERSION 5.00
Begin VB.Form frmNpcs 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Npc's"
   ClientHeight    =   4965
   ClientLeft      =   9525
   ClientTop       =   9210
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
   ScaleHeight     =   4965
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
      ItemData        =   "frmNpcs.frx":0000
      Left            =   0
      List            =   "frmNpcs.frx":0002
      TabIndex        =   5
      Tag             =   "-1"
      Top             =   0
      Width           =   4125
   End
   Begin VB.ComboBox cNPC 
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
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3690
      Width           =   2595
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
      ItemData        =   "frmNpcs.frx":0004
      Left            =   690
      List            =   "frmNpcs.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3300
      Width           =   3285
   End
   Begin Nexus_MapEditor.lvButtons_H cQuitarFunc 
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   4530
      Width           =   1935
      _extentx        =   3413
      _extenty        =   609
      caption         =   "Quitar NPC's"
      capalign        =   2
      backstyle       =   2
      font            =   "frmNpcs.frx":0008
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cAgregarFuncalAzar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4110
      Width           =   1935
      _extentx        =   3413
      _extenty        =   661
      caption         =   "Insertar NPC's al azar"
      capalign        =   2
      backstyle       =   2
      font            =   "frmNpcs.frx":0030
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cInsertarFunc 
      Height          =   765
      Left            =   2130
      TabIndex        =   2
      Top             =   4110
      Width           =   1845
      _extentx        =   3254
      _extenty        =   1349
      caption         =   "Insertar NPC"
      capalign        =   2
      backstyle       =   2
      font            =   "frmNpcs.frx":0058
      cgradient       =   0
      mode            =   1
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin VB.Label lbFiltrar 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lbnNPC 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero del NPC:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   3780
      Width           =   1215
   End
End
Attribute VB_Name = "frmNpcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cFiltro_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cFiltro_LostFocus_Err
    
    HotKeysAllow = True

    
    Exit Sub

cFiltro_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "FrmMain.cFiltro_LostFocus", Erl)
    Resume Next
    
End Sub
