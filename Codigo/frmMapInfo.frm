VERSION 5.00
Begin VB.Form frmMapInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información del Mapa"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "frmMapInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtlvlMinimo 
      Height          =   285
      Left            =   1650
      TabIndex        =   33
      Text            =   "0"
      Top             =   2700
      Width           =   2655
   End
   Begin VB.TextBox TxtAmbient 
      Height          =   285
      Left            =   1680
      TabIndex        =   31
      Text            =   "0"
      Top             =   1230
      Width           =   2655
   End
   Begin VB.ComboBox txtMapRestringir 
      Height          =   315
      ItemData        =   "frmMapInfo.frx":000C
      Left            =   1650
      List            =   "frmMapInfo.frx":0022
      TabIndex        =   30
      Top             =   2310
      Width           =   2655
   End
   Begin Nexus_MapEditor.lvButtons_H cPrevia 
      Height          =   375
      Left            =   150
      TabIndex        =   29
      Top             =   5640
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      Caption         =   "Vista Previa"
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
   Begin Nexus_MapEditor.lvButtons_H cmdCerrar 
      Height          =   375
      Left            =   2730
      TabIndex        =   28
      Top             =   5670
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
   Begin Nexus_MapEditor.lvButtons_H cmdMusica 
      Height          =   315
      Left            =   3540
      TabIndex        =   27
      Top             =   840
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      Caption         =   "Más"
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
   Begin VB.CheckBox chkLuzClimatica 
      Caption         =   "Luz climatica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   26
      Top             =   4170
      Width           =   1335
   End
   Begin VB.TextBox b1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3390
      TabIndex        =   25
      Text            =   "255"
      Top             =   5130
      Width           =   495
   End
   Begin VB.TextBox G1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2550
      TabIndex        =   23
      Text            =   "255"
      Top             =   5130
      Width           =   495
   End
   Begin VB.TextBox r1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1710
      TabIndex        =   21
      Text            =   "255"
      Top             =   5130
      Width           =   495
   End
   Begin VB.TextBox txtColor 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1470
      TabIndex        =   18
      Text            =   "&HFFFFFF"
      Top             =   4650
      Width           =   2415
   End
   Begin VB.PictureBox PicColorMap 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   150
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   4410
      Width           =   1095
   End
   Begin VB.CheckBox chkMapResuSinEfecto 
      Caption         =   "ResuSinEfecto"
      Height          =   255
      Left            =   2430
      TabIndex        =   15
      Top             =   3090
      Width           =   1815
   End
   Begin VB.CheckBox chkMapInviSinEfecto 
      Caption         =   "InviSinEfecto"
      Height          =   255
      Left            =   150
      TabIndex        =   14
      Top             =   3090
      Width           =   2055
   End
   Begin VB.TextBox txtMapVersion 
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
      Left            =   1680
      TabIndex        =   12
      Text            =   "0"
      Top             =   480
      Width           =   2655
   End
   Begin VB.CheckBox chkMapPK 
      Caption         =   "PK (inseguro)"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
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
      Left            =   150
      TabIndex        =   10
      Top             =   3570
      Width           =   1575
   End
   Begin VB.ComboBox txtMapTerreno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":0054
      Left            =   1650
      List            =   "frmMapInfo.frx":0061
      TabIndex        =   9
      Top             =   1950
      Width           =   2655
   End
   Begin VB.ComboBox txtMapZona 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMapInfo.frx":007E
      Left            =   1650
      List            =   "frmMapInfo.frx":008B
      TabIndex        =   8
      Top             =   1590
      Width           =   2655
   End
   Begin VB.TextBox txtMapMusica 
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
      Left            =   1680
      TabIndex        =   7
      Text            =   "0"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtMapNombre 
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
      Left            =   1680
      TabIndex        =   6
      Text            =   "Nuevo Mapa"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox chkMapBackup 
      Caption         =   "Backup"
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
      Left            =   2430
      TabIndex        =   3
      Top             =   3330
      Width           =   1575
   End
   Begin VB.CheckBox chkMapMagiaSinEfecto 
      Caption         =   "Magia Sin Efecto"
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
      Left            =   150
      TabIndex        =   2
      Top             =   3330
      Width           =   1575
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel Minimo:"
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
      Left            =   90
      TabIndex        =   34
      Top             =   2700
      Width           =   930
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido Ambiental:"
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
      Left            =   120
      TabIndex        =   32
      Top             =   1230
      Width           =   1290
   End
   Begin VB.Label Label11 
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3150
      TabIndex        =   24
      Top             =   5175
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "G:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2310
      TabIndex        =   22
      Top             =   5175
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "R:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   20
      Top             =   5175
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Codigo de color:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1470
      TabIndex        =   19
      Top             =   4410
      Width           =   1170
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Luz Base"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   17
      Top             =   3930
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Versión del Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   165
      X2              =   4345
      Y1              =   3930
      Y2              =   3930
   End
   Begin VB.Label Label5 
      Caption         =   "Restringir:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   2310
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Terreno:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Zona:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Musica:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   4345
      Y1              =   3930
      Y2              =   3930
   End
End
Attribute VB_Name = "frmMapInfo"
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

Private Sub chkMapBackup_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo chkMapBackup_LostFocus_Err
    
    MapInfo.BackUp = chkMapBackup.value
    MapInfo.Changed = 1

    Exit Sub

chkMapBackup_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.chkMapBackup_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub chkMapMagiaSinEfecto_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo chkMapMagiaSinEfecto_LostFocus_Err
    
    MapInfo.MagiaSinEfecto = chkMapMagiaSinEfecto.value
    MapInfo.Changed = 1
    
    Exit Sub

chkMapMagiaSinEfecto_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.chkMapMagiaSinEfecto_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub chkMapInviSinEfecto_LostFocus()
    '*************************************************
    'Author:
    'Last modified:
    '*************************************************
    
    On Error GoTo chkMapInviSinEfecto_LostFocus_Err
    
    MapInfo.InviSinEfecto = chkMapInviSinEfecto.value
    MapInfo.Changed = 1

    Exit Sub

chkMapInviSinEfecto_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.chkMapInviSinEfecto_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub chkMapResuSinEfecto_LostFocus()
    '*************************************************
    'Author:
    'Last modified:
    '*************************************************
    
    On Error GoTo chkMapResuSinEfecto_LostFocus_Err
    
    MapInfo.ResuSinEfecto = chkMapResuSinEfecto.value
    MapInfo.Changed = 1
    
    Exit Sub

chkMapResuSinEfecto_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.chkMapResuSinEfecto_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub chkMapPK_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo chkMapPK_LostFocus_Err
    
    MapInfo.Pk = chkMapPK.value
    MapInfo.Changed = 1

    Exit Sub

chkMapPK_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.chkMapPK_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_Click()
    
    On Error GoTo cmdCerrar_Click_Err
    
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If txtColor = "" Then
        Me.Hide
        Exit Sub

    End If

    MapInfo.LuzBase = RGB(r1.Text, G1.Text, b1.Text)
    Call Actualizar_Estado
    Me.Hide
    MapInfo.Changed = 1

    Exit Sub

cmdCerrar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.cmdCerrar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdMusica_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdMusica_Click_Err
    
    frmMusica.Show

    Exit Sub

cmdMusica_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.cmdMusica_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo Form_QueryUnload_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide

    End If
    
    Exit Sub

Form_QueryUnload_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.Form_QueryUnload", Erl)
    Resume Next
    
End Sub

Private Sub cPrevia_Click()
    
    On Error GoTo lvButtons_H1_Click_Err

    MapInfo.LuzBase = RGB(r1.Text, G1.Text, b1.Text)
    Call Actualizar_Estado
    
    Exit Sub

lvButtons_H1_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.cPrevia_Click", Erl)
    Resume Next
    
End Sub

Public Function Selected_Color()
    
    On Error GoTo Selected_Color_Err

    Dim c   As Long
  
    Dim R   As Integer ' Red component value   (0 to 255)
    Dim G   As Integer ' Green component value (0 to 255)
    Dim B   As Integer ' Blue component value  (0 to 255)
  
    Dim Out As String  ' Function output string
    
    ' Setup the color selection palette dialog.
    With frmMain.Dialog
  
        ' Set initial flags to open the full palette and allow an
        ' initial default color selection.
        .flags = cdlCCFullOpen + cdlCCRGBInit
      
        .color = RGB(255, 255, 255)
      
        ' Display the full color palette
        .ShowColor
        c = .color
                      
    End With

    R = c And 255              ' Get lowest 8 bits  - Red
    G = Int(c / 256) And 255   ' Get middle 8 bits  - Green
    B = Int(c / 65536) And 255 ' Get highest 8 bits - Blue
  
    ' If H mode is selected, replace default with hex RGB values.
    Out = "&H" & Format(Hex(R), "0#") & Format(Hex(G), "0#") & Format(Hex(B), "0#")
    'frmMain.Picture3.BackColor = RGB(r, g, b)

    Selected_Color = Out

    Exit Function

Selected_Color_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.Selected_Color", Erl)
    Resume Next
    
End Function

Private Sub PicColorMap_Click()
    
    On Error GoTo PicColorMap_Click_Err
    
    If chkLuzClimatica.value = False Then Exit Sub
    
    frmColorPicker.Show

    Exit Sub

PicColorMap_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.PicColorMap_Click", Erl)
    Resume Next
    
End Sub

Private Sub txtMapMusica_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo txtMapMusica_LostFocus_Err
    
    MapInfo.Music = txtMapMusica.Text
    'FrmMain.lblMapMusica.Caption = MapInfo.Music
    MapInfo.Changed = 1

    Exit Sub

txtMapMusica_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapMusica_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub txtMapVersion_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    
    On Error GoTo txtMapVersion_LostFocus_Err
    
    MapInfo.MapVersion = txtMapVersion.Text
    'FrmMain.lblMapVersion.Caption = MapInfo.MapVersion
    MapInfo.Changed = 1

    Exit Sub

txtMapVersion_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapVersion_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub txtMapNombre_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo txtMapNombre_LostFocus_Err
    
    MapInfo.name = txtMapNombre.Text
    'FrmMain.lblMapNombre.Caption = MapInfo.name
    MapInfo.Changed = 1
    
    Exit Sub

txtMapNombre_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapNombre_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub txtMapRestringir_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo txtMapRestringir_KeyPress_Err
    
    KeyAscii = 0
    
    Exit Sub

txtMapRestringir_KeyPress_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapRestringir_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub txtMapRestringir_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    'MapInfo.Restringir = txtMapRestringir.Text
    
    On Error GoTo txtMapRestringir_LostFocus_Err
    
    MapInfo.Changed = 1
    
    Exit Sub

txtMapRestringir_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapRestringir_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub txtMapTerreno_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo txtMapTerreno_KeyPress_Err
    
    KeyAscii = 0

    Exit Sub

txtMapTerreno_KeyPress_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapTerreno_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub txtMapTerreno_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    'MapInfo.Terreno = txtMapTerreno.Text
    
    On Error GoTo txtMapTerreno_LostFocus_Err
    
    MapInfo.Changed = 1
    
    Exit Sub

txtMapTerreno_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapTerreno_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub txtMapZona_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo txtMapZona_KeyPress_Err
    
    KeyAscii = 0
    
    Exit Sub

txtMapZona_KeyPress_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapZona_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub txtMapZona_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo txtMapZona_LostFocus_Err
    
    MapInfo.Zona = txtMapZona.Text
    MapInfo.Changed = 1

    Exit Sub

txtMapZona_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "frmMapInfo.txtMapZona_LostFocus", Erl)
    Resume Next
    
End Sub

Public Sub CambiarColorMap()
    '*************************************************
    'Author: Lorwik
    'Last modified: ?????
    '*************************************************
    
    On Error GoTo PicColorMap_Err
    
    PicColorMap.BackColor = MapInfo.LuzBase
    frmMapInfo.PicColorMap.BackColor = PicColorMap.BackColor
    MapInfo.Changed = 1
    
    Exit Sub

PicColorMap_Err:
    Call LogError(Err.Number, Err.Description, " FrmMain.CambiarColorMap", Erl)
    Resume Next
    
End Sub
