VERSION 5.00
Begin VB.Form frmActualizarIndices 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizar índices"
   ClientHeight    =   2775
   ClientLeft      =   16470
   ClientTop       =   10605
   ClientWidth     =   2535
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Botones 
      Caption         =   "Actualizar NPCs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Botones 
      Caption         =   "Actualizar objetos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Botones 
      Caption         =   "Actualizar triggers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Botones 
      Caption         =   "Actualizar cabezas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Botones 
      Caption         =   "Actualizar cuerpos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Botones 
      Caption         =   "ActualizarIndices.ini"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Botones 
      Caption         =   "Actualizar gráficos.ind"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmActualizarIndices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Botones_Click(Index As Integer)
    
    On Error GoTo Botones_Click_Err
    

    Select Case Index

        Case 0
            Call modCarga.LoadGrhData

        Case 1
            Call modCarga.CargarIndicesSuperficie

        Case 2
            Call modCarga.CargarCuerpos

        Case 3
            Call modCarga.CargarCabezas

        Case 4
            Call modCarga.CargarIndicesTriggers

        Case 5
            Call modCarga.CargarIndicesOBJ

        Case 6
            Call modCarga.CargarIndicesNPC

    End Select

    
    Exit Sub

Botones_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmActualizarIndices.Botones_Click", Erl)
    Resume Next
    
End Sub
