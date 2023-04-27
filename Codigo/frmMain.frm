VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000006&
   Caption         =   "Nexus MapEditor"
   ClientHeight    =   11595
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   19200
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
   ScaleHeight     =   768
   ScaleMode       =   0  'User
   ScaleWidth      =   1280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   11115
      Left            =   0
      ScaleHeight     =   741
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1277
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   19155
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
   End
   Begin VB.Menu mnuNuevoMapa 
      Caption         =   "&Nuevo Mapa"
   End
   Begin VB.Menu mnuAbrirMapa 
      Caption         =   "&Abrir Mapa"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tX                  As Byte
Public tY                  As Byte
Public MouseX              As Long
Public MouseY              As Long
Public MouseBoton          As Long
Public MouseShift          As Long
Private clicX              As Long
Private clicY              As Long

Public UltPos As Integer

Private Sub Form_Load()

    EngineRun = True
    Me.Caption = Form_Caption
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub
