VERSION 5.00
Begin VB.Form frmMiniMapa 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MiniMapa"
   ClientHeight    =   1485
   ClientLeft      =   9525
   ClientTop       =   5505
   ClientWidth     =   1440
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
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   96
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Render 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   0
      MouseIcon       =   "frmMinimapa.frx":0000
      ScaleHeight     =   97
      ScaleMode       =   0  'User
      ScaleWidth      =   97
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   750
         Top             =   773
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmMiniMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

