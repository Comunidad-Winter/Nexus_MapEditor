VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conversor de Mapas"
   ClientHeight    =   4335
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   6270
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
   ScaleHeight     =   4335
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraConversorDe 
      Caption         =   "Conversor de Formatos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin Nexus_MapEditor.lvButtons_H LvBConversion 
         Height          =   435
         Left            =   1500
         TabIndex        =   14
         Top             =   3450
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
         Caption         =   "Convertir"
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
      Begin VB.TextBox txtMin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "Automatizar proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   1850
         Width           =   1815
      End
      Begin VB.TextBox txtMax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox ComOpracion 
         Height          =   315
         ItemData        =   "frmConvert.frx":0000
         Left            =   600
         List            =   "frmConvert.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Info 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Esperando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   5535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblOperacion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Integer, Long, CSM, ImpC, IAO1.3, IAO1.4"
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   3840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carpetas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrucciones:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Metes el mapa en su carpeta de origen, en la conversion, aparecera en su carpeta de destino."
         Height          =   435
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   4680
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Automatico As Boolean

Private Sub chkAuto_Click()

    If chkAuto.value = False Then
        Label6.Visible = False
        Label7.Visible = False
        txtMax.Visible = False
        Automatico = False
        
    Else
    
        Label6.Visible = True
        Label7.Visible = True
        txtMax.Visible = True
        Automatico = True
        
    End If
    
End Sub

Private Sub Form_Load()
    ComOpracion.ListIndex = 0
End Sub

Private Sub LvBCerrar_Click()
    Unload Me
    
End Sub

Private Sub ConvertirInteger()

    Dim i As Integer
    
    If Automatico = False Then
        If FileExist(App.Path & "\Conversor\Integer\Mapa" & txtMin.Text & ".map", vbNormal) = True Then
            Call modMapIO.NuevoMapa
            Call modMapIO.MapaAO_Cargar(App.Path & "\Conversor\Integer\Mapa" & txtMin.Text & ".map", True)
            DoEvents
            Call modMapIO.MapaAO_Guardar(App.Path & "\Conversor\Long\Mapa" & txtMin.Text & ".map")
            
            Info.Caption = "Conversion realizada correctamente!"
                    
        Else
            Info.Caption = "Mapa" & txtMin.Text & ".map no existe!"
        End If
    Else
        For i = txtMin.Text To txtMax.Text
            If FileExist(App.Path & "\Conversor\Integer\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call modMapIO.MapaAO_Cargar(App.Path & "\Conversor\Integer\Mapa" & i & ".map", True)
                DoEvents
                Call modMapIO.MapaAO_Guardar(App.Path & "\Conversor\Long\Mapa" & i & ".map")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
                
            Else
                Info.Caption = "Mapa" & i & ".map no existe!"
                
            End If
            
        Next i
    End If
    
End Sub

Private Sub ConvertirLong()

    Dim i As Integer
    
    If Automatico = False Then
        If FileExist(App.Path & "\Conversor\Long\Mapa" & txtMin.Text & ".map", vbNormal) = True Then
            Call modMapIO.NuevoMapa
            Call modMapIO.MapaAO_Cargar(App.Path & "\Conversor\Long\Mapa" & txtMin.Text & ".map", False)
            DoEvents
            Call modMapIO.MapaCSM_Guardar(App.Path & "\Conversor\CSM\Mapa" & txtMin.Text & ".csm")
            
            Info.Caption = "Conversion realizada correctamente!"
            
        Else
            Info.Caption = "Mapa" & txtMin.Text & ".map no existe!"
            
        End If
        
    Else
        For i = txtMin.Text To txtMax.Text
            
            If FileExist(App.Path & "\Conversor\Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call modMapIO.MapaAO_Cargar(App.Path & "\Conversor\Long\Mapa" & i & ".map", False)
                DoEvents
                Call modMapIO.MapaCSM_Guardar(App.Path & "\Conversor\CSM\Mapa" & i & ".csm")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
                
            Else
                Info.Caption = "Mapa" & i & ".map no existe!"
                
            End If
        Next i
    End If
    
End Sub

Private Sub LvBConversion_Click()

    Select Case ComOpracion.ListIndex
    
        Case 0 'Int > Long
            Call ConvertirInteger
            
        Case 1 'Long > CSM
            Call ConvertirLong
            
    End Select
End Sub
