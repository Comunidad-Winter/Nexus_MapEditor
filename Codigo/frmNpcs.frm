VERSION 5.00
Begin VB.Form frmNpcs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Npc's"
   ClientHeight    =   3165
   ClientLeft      =   9525
   ClientTop       =   9210
   ClientWidth     =   4050
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
   ScaleHeight     =   3165
   ScaleWidth      =   4050
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
      TabIndex        =   0
      Tag             =   "-1"
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frmNpcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

