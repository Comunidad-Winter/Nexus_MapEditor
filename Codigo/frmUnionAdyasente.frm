VERSION 5.00
Begin VB.Form frmUnionAdyacente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Union con Mapas Adyasentes"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmUnionAdyasente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Nexus_MapEditor.lvButtons_H cmdCancelar 
      Height          =   405
      Left            =   4710
      TabIndex        =   40
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      Caption         =   "Cancelar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   37
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox AutoMapeo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   35
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   26
      Text            =   "87"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   24
      Text            =   "14"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   5520
      TabIndex        =   22
      Text            =   "11"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   20
      Text            =   "90"
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   19
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox Aplicar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   16
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Text            =   "0"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   13
      Text            =   "13"
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Text            =   "89"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Text            =   "10"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Text            =   "91"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox AutoMapeo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto-Mapeo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   36
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Nexus_MapEditor.lvButtons_H cmdAplicar 
      Height          =   405
      Left            =   3390
      TabIndex        =   41
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      Caption         =   "Aplicar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin Nexus_MapEditor.lvButtons_H cmdDefault 
      Height          =   405
      Left            =   90
      TabIndex        =   42
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      Caption         =   "Default"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Label lblMapaAct 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   39
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblMapaActual 
      Caption         =   "Mapa Actual"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   38
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Leyenda sobre posiciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   4950
      Y2              =   5080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion Y del mapa actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3960
      TabIndex        =   32
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   4695
      Y2              =   4845
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion Y del mapa destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   3960
      TabIndex        =   31
      Top             =   4680
      Width           =   2025
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   4950
      Y2              =   5080
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion X del mapa actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   1680
      TabIndex        =   30
      Top             =   4920
      Width           =   1920
   End
   Begin VB.Line Line15 
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   4695
      Y2              =   4845
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion X del mapa destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   210
      Left            =   1680
      TabIndex        =   29
      Top             =   4680
      Width           =   2010
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00008000&
      X1              =   120
      X2              =   6000
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label13 
      Caption         =   "NOTA: Mapa 0, borra el translado de mapa."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1320
      TabIndex        =   28
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Line Line13 
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      X1              =   840
      X2              =   840
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Label Label12 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   255
   End
   Begin VB.Line Line12 
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   720
      Y2              =   3240
   End
   Begin VB.Label Label11 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   3120
      Width           =   255
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   1080
      X2              =   5280
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label10 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   3480
      Width           =   255
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   840
      X2              =   5040
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label9 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00404040&
      X1              =   960
      X2              =   5160
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      X1              =   5160
      X2              =   5160
      Y1              =   3480
      Y2              =   600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404040&
      X1              =   5160
      X2              =   960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   960
      X2              =   960
      Y1              =   3480
      Y2              =   600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008000&
      X1              =   120
      X2              =   6000
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label8 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   720
      X2              =   4920
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   1200
      X2              =   5400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   1080
      X2              =   1080
      Y1              =   840
      Y2              =   3600
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   960
      Top             =   600
      Width           =   4215
   End
   Begin VB.Menu mnuDefault 
      Caption         =   "mnuDefault"
      Visible         =   0   'False
      Begin VB.Menu mnuLegal 
         Caption         =   "Borde Legal Automatico"
      End
      Begin VB.Menu mnuBasica 
         Caption         =   "11,10 y 90,91 - Basica"
      End
      Begin VB.Menu mnuUlla 
         Caption         =   "9,7 y 92,94 - Ullathorpe"
      End
   End
End
Attribute VB_Name = "frmUnionAdyacente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit

Private Sub Aplicar_Click(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Aplicar_Click_Err
    
    Dim i As Byte
    cmdAplicar.Enabled = False

    For i = 0 To 3

        If Aplicar(i).value = 1 Then cmdAplicar.Enabled = True
    Next

    
    Exit Sub

Aplicar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.Aplicar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdAplicar_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If
    
    modEdicion.Deshacer_Add "Insertar Translados a mapas Adyasentes" ' Hago deshacer

    ' ARRIBA
    If Mapa(0).Text > -1 And Aplicar(0).value = 1 Then
        y = PosLim(1).Text

        For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)

            If (MapData(X, y).Blocked And &HF) <> &HF Then
                MapData(X, y).TileExit.Map = Mapa(0).Text

                If Mapa(0).Text = 0 Then
                    MapData(X, y).TileExit.X = 0
                    MapData(X, y).TileExit.y = 0
                Else
                    MapData(X, y).TileExit.X = X
                    MapData(X, y).TileExit.y = PosLim(4).Text
                    MapInfo.Changed = 1
                End If

            End If

        Next

    End If

    ' DERECHA
    If Mapa(1).Text > -1 And Aplicar(1).value = 1 Then
        X = PosLim(2).Text

        For y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)

            If (MapData(X, y).Blocked And &HF) <> &HF Then
                MapData(X, y).TileExit.Map = Mapa(1).Text

                If Mapa(1).Text = 0 Then
                    MapData(X, y).TileExit.X = 0
                    MapData(X, y).TileExit.y = 0
                Else
                    MapData(X, y).TileExit.X = PosLim(6).Text
                    MapData(X, y).TileExit.y = y
                    MapInfo.Changed = 1

                End If

            End If

        Next

    End If

    ' ABAJO
    If Mapa(2).Text > -1 And Aplicar(2).value = 1 Then
        y = PosLim(0).Text

        For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)

            If (MapData(X, y).Blocked And &HF) <> &HF Then
                MapData(X, y).TileExit.Map = Mapa(2).Text

                If Mapa(2).Text = 0 Then
                    MapData(X, y).TileExit.X = 0
                    MapData(X, y).TileExit.y = 0
                Else
                    MapData(X, y).TileExit.X = X
                    MapData(X, y).TileExit.y = PosLim(5).Text
                    MapInfo.Changed = 1
                End If

            End If

        Next

    End If

    ' IZQUIERDA
    If Mapa(3).Text > -1 And Aplicar(3).value = 1 Then
        X = PosLim(3).Text

        For y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)

            If (MapData(X, y).Blocked And &HF) <> &HF Then
                MapData(X, y).TileExit.Map = Mapa(3).Text

                If Mapa(3).Text = 0 Then
                    MapData(X, y).TileExit.X = 0
                    MapData(X, y).TileExit.y = 0
                Else
                    MapData(X, y).TileExit.X = PosLim(7).Text
                    MapData(X, y).TileExit.y = y
                    MapInfo.Changed = 1
                End If

            End If

        Next

    End If

    'Set changed flag
    'ver ReyarB bloqueo bordes
    'Call modEdicion.Bloquear_Bordes
    
    DoEvents
    Unload Me

End Sub

Private Sub cmdCancelar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdCancelar_Click_Err
    
    Unload Me
    
    Exit Sub

cmdCancelar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.cmdCancelar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdDefault_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdDefault_Click_Err
    
    Me.PopupMenu mnuDefault

    
    Exit Sub

cmdDefault_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.cmdDefault_Click", Erl)
    Resume Next
    
End Sub

''
'   Lee los Translados existentes en lugares claves en el Mapa
'

Private Sub LeerMapaExit()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    Dim X As Integer
    Dim y As Integer

    ' ARRIBA
    Mapa(0).Text = 0
    y = PosLim(1).Text

    For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)

        If MapData(X, y).TileExit.Map <> 0 Then
            Mapa(0).Text = MapData(X, y).TileExit.Map
            Exit For

        End If

    Next
    Aplicar(0).value = 0

    ' DERECHA
    Mapa(1).Text = 0
    X = PosLim(2).Text

    For y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)

        If MapData(X, y).TileExit.Map <> 0 Then
            Mapa(1).Text = MapData(X, y).TileExit.Map
            Exit For

        End If

    Next
    Aplicar(1).value = 0

    ' ABAJO
    Mapa(2).Text = 0
    y = PosLim(0).Text

    For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)

        If MapData(X, y).TileExit.Map <> 0 Then
            Mapa(2).Text = MapData(X, y).TileExit.Map
            Exit For

        End If

    Next
    Aplicar(2).value = 0

    ' IZQUIERDA
    Mapa(3).Text = 0
    X = PosLim(3).Text

    For y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)

        If MapData(X, y).TileExit.Map <> 0 Then
            Mapa(3).Text = MapData(X, y).TileExit.Map
            Exit For

        End If

    Next
    Aplicar(3).value = 0

End Sub

Private Sub Form_Load()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Form_Load_Err
    
    Call mnuLegal_Click
    frmUnionAdyacente.lblMapaAct.Caption = MapaActual
    
    Exit Sub

Form_Load_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Mapa_Change(Index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Mapa_Change_Err
    
    Aplicar(Index).value = 1
    
    Exit Sub

Mapa_Change_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.Mapa_Change", Erl)
    Resume Next
    
End Sub

Private Sub Mapa_KeyPress(Index As Integer, KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Mapa_KeyPress_Err
    
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0
        Exit Sub

    End If
    
    Exit Sub

Mapa_KeyPress_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.Mapa_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Mapa_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo Mapa_KeyUp_Err
    
    If LenB(Mapa(Index).Text) = 0 Then Mapa(Index).Text = 0
    If Mapa(Index).Text > 1024 Then Mapa(Index).Text = 1024

    Exit Sub

Mapa_KeyUp_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.Mapa_KeyUp", Erl)
    Resume Next
    
End Sub

Private Sub mnuBasica_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuBasica_Click_Err
    
    Call LeerMapaExit

    Exit Sub

mnuBasica_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.mnuBasica_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuLegal_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/10/06
    '*************************************************
    
    On Error GoTo mnuLegal_Click_Err
    
    PosLim(0).Text = MaxYBorder
    PosLim(1).Text = MinYBorder
    PosLim(2).Text = MaxXBorder
    PosLim(3).Text = MinXBorder
    PosLim(4).Text = MaxYBorder - 1
    PosLim(5).Text = MinYBorder + 1
    PosLim(6).Text = MinXBorder + 1
    PosLim(7).Text = MaxXBorder - 1
    Call LeerMapaExit

    Exit Sub

mnuLegal_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.mnuLegal_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuUlla_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo mnuUlla_Click_Err
    
    PosLim(0).Text = 94
    PosLim(1).Text = 7
    PosLim(2).Text = 92
    PosLim(3).Text = 9
    PosLim(4).Text = 93
    PosLim(5).Text = 8
    PosLim(6).Text = 10
    PosLim(7).Text = 91
    Call LeerMapaExit
    
    Exit Sub

mnuUlla_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.mnuUlla_Click", Erl)
    Resume Next
    
End Sub

Private Sub PosLim_KeyPress(Index As Integer, KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo PosLim_KeyPress_Err
    
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0
        Exit Sub

    End If
    
    Exit Sub

PosLim_KeyPress_Err:
    Call LogError(Err.Number, Err.Description, "frmUnionAdyacente.PosLim_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub PosLim_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    On Error Resume Next

    If LenB(PosLim(Index).Text) = 0 Then PosLim(Index).Text = 1
    If PosLim(Index).Text > 99 Then PosLim(Index) = 99
    If PosLim(Index).Text < 1 Then PosLim(Index) = 1

    Dim y As Integer
    Dim X As Integer

    ' ARRIBA
    y = PosLim(1).Text

    For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)

        If MapData(X, y).TileExit.Map <> 0 Then
            Mapa(0).Text = MapData(X, y).TileExit.Map
            Aplicar(0).value = 0
            Exit For

        End If

    Next

    ' DERECHA
    X = PosLim(2).Text

    For y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)

        If MapData(X, y).TileExit.Map <> 0 Then
            Mapa(1).Text = MapData(X, y).TileExit.Map
            Aplicar(1).value = 0
            Exit For

        End If

    Next

    ' ABAJO
    y = PosLim(0).Text

    For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)

        If MapData(X, y).TileExit.Map <> 0 Then
            Mapa(2).Text = MapData(X, y).TileExit.Map
            Aplicar(2).value = 0
            Exit For

        End If

    Next

    ' IZQUIERDA
    X = PosLim(3).Text

    For y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)

        If MapData(X, y).TileExit.Map <> 0 Then
            Mapa(3).Text = MapData(X, y).TileExit.Map
            Aplicar(3).value = 0
            Exit For

        End If

    Next

End Sub
