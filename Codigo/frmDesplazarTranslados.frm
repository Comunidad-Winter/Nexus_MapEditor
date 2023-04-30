VERSION 5.00
Begin VB.Form frmDesplazarTranslados 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desplazar Translados"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox DestinoDerecha 
      Height          =   375
      Left            =   4410
      TabIndex        =   20
      Text            =   "13"
      Top             =   1830
      Width           =   495
   End
   Begin VB.TextBox ActualXDerecha 
      Height          =   285
      Left            =   3450
      TabIndex        =   18
      Text            =   "92"
      Top             =   1710
      Width           =   495
   End
   Begin VB.TextBox DesplazadaXDerecha 
      Height          =   375
      Left            =   3450
      TabIndex        =   17
      Text            =   "88"
      Top             =   2070
      Width           =   495
   End
   Begin VB.TextBox DestinoIzquierda 
      Height          =   375
      Left            =   90
      TabIndex        =   16
      Text            =   "87"
      Top             =   1950
      Width           =   495
   End
   Begin VB.TextBox DesplazadaXIzquierda 
      Height          =   375
      Left            =   930
      TabIndex        =   15
      Text            =   "12"
      Top             =   2100
      Width           =   495
   End
   Begin VB.TextBox ActualXIzquierda 
      Height          =   285
      Left            =   930
      TabIndex        =   14
      Text            =   "9"
      Top             =   1740
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   1710
      TabIndex        =   12
      Top             =   1770
      Width           =   1575
   End
   Begin VB.TextBox DestinoYInferior 
      Height          =   285
      Left            =   2610
      TabIndex        =   11
      Text            =   "11"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox DestinoSuperior 
      Height          =   285
      Left            =   2370
      TabIndex        =   9
      Text            =   "90"
      Top             =   135
      Width           =   495
   End
   Begin VB.TextBox DesplazadaYInferior 
      Height          =   285
      Left            =   3570
      TabIndex        =   7
      Text            =   "91"
      Top             =   3150
      Width           =   495
   End
   Begin VB.TextBox DesplazadaYSuperior 
      Height          =   285
      Left            =   3570
      TabIndex        =   5
      Text            =   "10"
      Top             =   705
      Width           =   495
   End
   Begin VB.TextBox ActualYInferior 
      Height          =   285
      Left            =   1740
      TabIndex        =   3
      Text            =   "94"
      Top             =   3150
      Width           =   495
   End
   Begin VB.TextBox ActualYSuperior 
      Height          =   285
      Left            =   1770
      TabIndex        =   1
      Text            =   "7"
      Top             =   705
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual X"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3330
      TabIndex        =   19
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual X"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   930
      TabIndex        =   13
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino Y="
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1770
      TabIndex        =   10
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino Y="
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1530
      TabIndex        =   8
      Top             =   150
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000B&
      Height          =   3135
      Left            =   690
      Top             =   510
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Desplazar a Y="
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2370
      TabIndex        =   6
      Top             =   3195
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Desplazar a Y="
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2370
      TabIndex        =   4
      Top             =   750
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Y="
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Y="
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   930
      TabIndex        =   0
      Top             =   750
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      Height          =   2895
      Left            =   810
      Top             =   630
      Width           =   3375
   End
End
Attribute VB_Name = "frmDesplazarTranslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Dim X As Byte
    Dim y As Byte

    For X = 13 To 87
        For y = ActualYSuperior To ActualYSuperior

            If MapData(X, y).TileExit.Map <> 0 Then
                MapData(X, DesplazadaYSuperior).TileExit.Map = MapData(X, y).TileExit.Map
                MapData(X, DesplazadaYSuperior).TileExit.X = MapData(X, y).TileExit.X
                MapData(X, DesplazadaYSuperior).TileExit.y = DestinoSuperior
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0

            End If

        Next y
    Next X

    For X = 13 To 87
        For y = ActualYInferior To ActualYInferior

            If MapData(X, y).TileExit.Map <> 0 Then
                MapData(X, DesplazadaYInferior).TileExit.Map = MapData(X, y).TileExit.Map
                MapData(X, DesplazadaYInferior).TileExit.X = MapData(X, y).TileExit.X
                MapData(X, DesplazadaYInferior).TileExit.y = DestinoYInferior
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0

            End If

        Next y
    Next X

    For X = ActualXIzquierda To ActualXIzquierda
        For y = 11 To 90

            If MapData(X, y).TileExit.Map <> 0 Then
                MapData(DesplazadaXIzquierda, y).TileExit.Map = MapData(X, y).TileExit.Map
                MapData(DesplazadaXIzquierda, y).TileExit.X = DestinoIzquierda
                MapData(DesplazadaXIzquierda, y).TileExit.y = MapData(X, y).TileExit.y
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0

            End If

        Next y
    Next X

    For X = ActualXDerecha To ActualXDerecha
        For y = 11 To 90

            If MapData(X, y).TileExit.Map <> 0 Then
                MapData(DesplazadaXDerecha, y).TileExit.Map = MapData(X, y).TileExit.Map
                MapData(DesplazadaXDerecha, y).TileExit.X = DestinoDerecha
                MapData(DesplazadaXDerecha, y).TileExit.y = MapData(X, y).TileExit.y
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0

            End If

        Next y
    Next X
        
    
    Exit Sub

Command1_Click_Err:
    Call LogError(Err.Number, Err.Description, "DesplazarTranslados.Command1_Click", Erl)
    Resume Next
    
End Sub
