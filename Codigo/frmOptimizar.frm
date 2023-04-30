VERSION 5.00
Begin VB.Form frmOptimizar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Optimizar Mapa"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   Icon            =   "frmOptimizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Nexus_MapEditor.lvButtons_H cCancelar 
      Height          =   495
      Left            =   1890
      TabIndex        =   6
      Top             =   2700
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
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
   Begin VB.CheckBox chkBloquearArbolesEtc 
      Caption         =   "Bloquear Arboles, Carteles, Foros y Yacimientos"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkMapearArbolesEtc 
      Caption         =   "Mapear Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTodoBordes 
      Caption         =   "Quitar NPCs, Objetos y Translados en los Bordes Exteriores"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigTrans 
      Caption         =   "Quitar Trigger's en Translados"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigBloq 
      Caption         =   "Quitar Trigger's Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrans 
      Caption         =   "Quitar Translados Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin Nexus_MapEditor.lvButtons_H cOptimizar 
      Height          =   495
      Left            =   210
      TabIndex        =   7
      Top             =   2700
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      Caption         =   "Optimizar"
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
      cBack           =   12648384
   End
End
Attribute VB_Name = "frmOptimizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Optimizar()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    
    On Error GoTo Optimizar_Err
    
    Dim y As Integer
    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    ' Quita Translados Bloqueados
    ' Quita Trigger's Bloqueados
    ' Quita Trigger's en Translados
    ' Quita NPCs, Objetos y Translados en los Bordes Exteriores
    ' Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa

    modEdicion.Deshacer_Add "Aplicar Optimizacion del Mapa" ' Hago deshacer

    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            ' ** Quitar NPCs, Objetos y Translados en los Bordes Exteriores
            If (X < 12 Or X > 88 Or y < 10 Or y > 91) And chkQuitarTodoBordes.value = 1 Then

                'Quitar NPCs
                If MapData(X, y).NPCIndex > 0 Then
                    Char_Erase MapData(X, y).CharIndex
                    MapData(X, y).NPCIndex = 0

                End If

                ' Quitar Objetos
                MapData(X, y).OBJInfo.objindex = 0
                MapData(X, y).OBJInfo.Amount = 0
                MapData(X, y).ObjGrh.GrhIndex = 0
                ' Quitar Translados
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.X = 0
                MapData(X, y).TileExit.y = 0
                ' Quitar Triggers
                MapData(X, y).Trigger = 0

            End If

            ' ** Quitar Translados y Triggers en Bloqueo
            If MapData(X, y).Blocked = &HF Then
                If MapData(X, y).TileExit.Map <> 0 And chkQuitarTrans.value = 1 Then ' Quita Translado Bloqueado
                    MapData(X, y).TileExit.Map = 0
                    MapData(X, y).TileExit.y = 0
                    MapData(X, y).TileExit.X = 0
                ElseIf MapData(X, y).Trigger > 0 And chkQuitarTrigBloq.value = 1 Then ' Quita Trigger Bloqueado
                    MapData(X, y).Trigger = 0

                End If

            End If

            ' ** Quitar Triggers en Translado
            If MapData(X, y).TileExit.Map <> 0 And chkQuitarTrigTrans.value = 1 Then
                If MapData(X, y).Trigger > 0 Then ' Quita Trigger en Translado
                    MapData(X, y).Trigger = 0

                End If

            End If

            ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
            If MapData(X, y).OBJInfo.objindex > 0 And (chkMapearArbolesEtc.value = 1 Or chkBloquearArbolesEtc.value = 1) Then

                Select Case ObjData(MapData(X, y).OBJInfo.objindex).ObjType

                    Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos

                        If MapData(X, y).Graphic(3).GrhIndex <> MapData(X, y).ObjGrh.GrhIndex And chkMapearArbolesEtc.value = 1 Then _
                            MapData(X, y).Graphic(3) = MapData(X, y).ObjGrh

                        If chkBloquearArbolesEtc.value = 1 And MapData(X, y).Blocked = 0 Then _
                            MapData(X, y).Blocked = 1

                End Select

            End If

            ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
        Next X
    Next y
    
    Exit Sub

Optimizar_Err:
    Call LogError(Err.Number, Err.Description, "frmOptimizar.Optimizar", Erl)
    Resume Next
    
End Sub

Private Sub cCancelar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/09/06
    '*************************************************
    
    On Error GoTo cCancelar_Click_Err
    
    Unload Me

    Exit Sub

cCancelar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmOptimizar.cCancelar_Click", Erl)
    Resume Next
    
End Sub

Public Sub cOptimizar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/09/06
    '*************************************************
    
    On Error GoTo cOptimizar_Click_Err
    
    Call Optimizar
    MapInfo.Changed = 1

    Exit Sub

cOptimizar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmOptimizar.cOptimizar_Click", Erl)
    Resume Next
    
End Sub

