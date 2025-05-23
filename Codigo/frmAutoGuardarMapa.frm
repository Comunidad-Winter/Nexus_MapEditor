VERSION 5.00
Begin VB.Form frmAutoGuardarMapa 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Guardar Mapa Automaticamente"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmAutoGuardarMapa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Nexus_MapEditor.lvButtons_H cmdCerrar 
      Height          =   405
      Left            =   3480
      TabIndex        =   3
      Top             =   1530
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
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
   Begin VB.ComboBox cmbMinutos 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmAutoGuardarMapa.frx":000C
      Left            =   1440
      List            =   "frmAutoGuardarMapa.frx":0025
      TabIndex        =   1
      Text            =   "10"
      Top             =   840
      Width           =   1215
   End
   Begin Nexus_MapEditor.lvButtons_H cmdDetener 
      Height          =   435
      Left            =   1830
      TabIndex        =   4
      Top             =   1530
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   767
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
   Begin Nexus_MapEditor.lvButtons_H cmdAceptar 
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   1530
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   767
      Caption         =   "Aceptar"
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
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   135
      X2              =   4915
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "minutos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   885
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Indique cada cuantos Minutos desea que se Guarde Automaticamente el Mapa con el que trabaja:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4915
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmAutoGuardarMapa"
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

Private Sub cmdAceptar_Click()
    
    On Error GoTo cmdAceptar_Click_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If IsNumeric(cmbMinutos.Text) = False Then
        MsgBox "Los minutos deben ingresarse de forma n�merica.", vbCritical + vbOKOnly
        Exit Sub
    ElseIf Val(cmbMinutos.Text) < 5 Or Val(cmbMinutos.Text) > 120 Then
        MsgBox "Los minutos ingresados son invalidos." & vbCrLf & "Solo estan permitidos los valores de entre 5 y 120 minutos inclusive.", vbCritical + vbOKOnly
        Exit Sub

    End If

    bAutoGuardarMapa = Val(cmbMinutos.Text)
    bAutoGuardarMapaCount = 0
    frmMain.TimAutoGuardarMapa.Enabled = True
    frmMain.mnuAutoGuardarMapas.Checked = True
    Unload Me

    Exit Sub

cmdAceptar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmAutoGuardarMapa.cmdAceptar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdCerrar_Click_Err
    
    Unload Me

    
    Exit Sub

cmdCerrar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmAutoGuardarMapa.cmdCerrar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdDetener_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cmdDetener_Click_Err
    
    frmMain.TimAutoGuardarMapa.Enabled = False
    frmMain.mnuAutoGuardarMapas.Checked = False
    Unload Me

    
    Exit Sub

cmdDetener_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmAutoGuardarMapa.cmdDetener_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo Form_Load_Err
    
    cmbMinutos.Text = bAutoGuardarMapa

    
    Exit Sub

Form_Load_Err:
    Call LogError(Err.Number, Err.Description, "frmAutoGuardarMapa.Form_Load", Erl)
    Resume Next
    
End Sub

