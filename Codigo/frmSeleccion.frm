VERSION 5.00
Begin VB.Form frmSeleccion 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Herramienta de Selección"
   ClientHeight    =   1485
   ClientLeft      =   12555
   ClientTop       =   12270
   ClientWidth     =   3990
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
   ScaleWidth      =   266
   ShowInTaskbar   =   0   'False
   Begin Nexus_MapEditor.lvButtons_H LvBAreas 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   990
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      Caption         =   "Quitar Bloqueos"
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
   Begin VB.TextBox DY2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3330
      TabIndex        =   3
      Text            =   "5"
      Top             =   90
      Width           =   495
   End
   Begin VB.TextBox DY1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2370
      TabIndex        =   2
      Text            =   "1"
      Top             =   90
      Width           =   495
   End
   Begin VB.TextBox DX2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1410
      TabIndex        =   1
      Text            =   "5"
      Top             =   90
      Width           =   495
   End
   Begin VB.TextBox DX1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   450
      TabIndex        =   0
      Text            =   "1"
      Top             =   90
      Width           =   495
   End
   Begin Nexus_MapEditor.lvButtons_H LvBAreas 
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   540
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      Caption         =   "Insertar Bloqueos"
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
   Begin Nexus_MapEditor.lvButtons_H LvBAreas 
      Height          =   405
      Index           =   3
      Left            =   1575
      TabIndex        =   10
      Top             =   990
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Quit. Superficie"
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
   Begin Nexus_MapEditor.lvButtons_H LvBAreas 
      Height          =   405
      Index           =   2
      Left            =   1575
      TabIndex        =   11
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Ins. Superficie"
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
   Begin Nexus_MapEditor.lvButtons_H LvBAreas 
      Height          =   405
      Index           =   5
      Left            =   2850
      TabIndex        =   12
      Top             =   990
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   714
      Caption         =   "Quit. Trigger"
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
   Begin Nexus_MapEditor.lvButtons_H LvBAreas 
      Height          =   405
      Index           =   4
      Left            =   2850
      TabIndex        =   13
      Top             =   540
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   714
      Caption         =   "Ins. Trigger"
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
   Begin VB.Label lblX2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   105
      Width           =   255
   End
   Begin VB.Label lblX2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1125
      TabIndex        =   6
      Top             =   105
      Width           =   255
   End
   Begin VB.Label lblY1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2010
      TabIndex        =   5
      Top             =   105
      Width           =   255
   End
   Begin VB.Label lblY2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3045
      TabIndex        =   4
      Top             =   105
      Width           =   255
   End
End
Attribute VB_Name = "frmSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LvBAreas_Click(Index As Integer)
    If IsNumeric(DX1.Text) = False Or _
       IsNumeric(DX2.Text) = False Or _
       IsNumeric(DY1.Text) = False Or _
       IsNumeric(DY2.Text) = False Then
    
        Call MsgBox("Debes introducir valores nï¿½mericos. Estos pueden tener un mï¿½nimo de 1 y un mï¿½ximo de " & (YMinMapSize + XMinMapSize) / 2 & ".")
    
       Exit Sub
    End If
    
    Select Case Index
        Case 0
            Call Bloqueos_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 1
            Call Bloqueos_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 2
            Call Superficie_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 3
            Call Superficie_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 4
            Call Triggers_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 5
            Call Triggers_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
    End Select
End Sub

