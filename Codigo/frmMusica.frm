VERSION 5.00
Begin VB.Form frmMusica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Musica"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmMusica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Nexus_MapEditor.lvButtons_H cmdCerrar 
      Height          =   465
      Left            =   2910
      TabIndex        =   4
      Top             =   1350
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   820
      Caption         =   "Cerrar"
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
   Begin Nexus_MapEditor.lvButtons_H cmdAplicarYCerrar 
      Height          =   465
      Left            =   2910
      TabIndex        =   3
      Top             =   750
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   820
      Caption         =   "Aplicar y cerrar"
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
   Begin Nexus_MapEditor.lvButtons_H cmdDetener 
      Height          =   465
      Left            =   4110
      TabIndex        =   2
      Top             =   120
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   820
      Caption         =   "Detener"
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
   Begin Nexus_MapEditor.lvButtons_H cmdEscuchar 
      Height          =   465
      Left            =   2910
      TabIndex        =   1
      Top             =   120
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      Caption         =   "Escuchar"
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
   Begin VB.FileListBox fleMusicas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   120
      Pattern         =   "*.mid"
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMusica"
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

Private MidiActual As String

''
' Aplica la Musica seleccionada y oculta la ventana
'

Private Sub cmdAplicarYCerrar_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If Len(MidiActual) >= 5 Then
        MapInfo.Music = Left(MidiActual, Len(MidiActual) - 4)
        frmMapInfo.txtMapMusica.Text = MapInfo.Music
        ' FrmMain.lblMapMusica = MapInfo.Music
        MidiActual = Empty

    End If

    Me.Hide

End Sub

''
' Oculta la ventana
'

Private Sub cmdCerrar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdCerrar_Click_Err
    
    Me.Hide

    Exit Sub

cmdCerrar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmMusica.cmdCerrar_Click", Erl)
    Resume Next
    
End Sub

''
' Detiene la Musica que se encuentra Reproduciendo
'

Private Sub cmdDetener_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdDetener_Click_Err

    cmdEscuchar.Enabled = True
    cmdDetener.Enabled = False
    'Play = False
    'TODO
    Exit Sub

cmdDetener_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmMusica.cmdDetener_Click", Erl)
    Resume Next
    
End Sub

''
' Inicia la reproduccion de la Musica Seleccionada
'

Private Sub cmdEscuchar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdEscuchar_Click_Err

    cmdDetener.Enabled = True
    cmdEscuchar.Enabled = False
    'Play = True
    'TODO
    Exit Sub

cmdEscuchar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmMusica.cmdEscuchar_Click", Erl)
    Resume Next
    
End Sub

''
' Selecciona una nueva Musica del listado
'

Private Sub fleMusicas_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo fleMusicas_Click_Err
    
    MidiActual = fleMusicas.List(fleMusicas.ListIndex)

    cmdAplicarYCerrar.Enabled = True

    'If Play = False Then cmdEscuchar.Enabled = True
    'TODO
    Exit Sub

fleMusicas_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmMusica.fleMusicas_Click", Erl)
    Resume Next
    
End Sub

