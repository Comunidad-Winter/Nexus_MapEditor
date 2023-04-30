VERSION 5.00
Begin VB.Form frmGRHaBMP 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GRH => BMP"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3765
   Icon            =   "frmGRHaBMP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Nexus_MapEditor.lvButtons_H cmdCerrar 
      Height          =   345
      Left            =   330
      TabIndex        =   4
      Top             =   1590
      Width           =   3075
      _extentx        =   5424
      _extenty        =   609
      caption         =   "&Cerrar"
      capalign        =   2
      backstyle       =   2
      font            =   "frmGRHaBMP.frx":000C
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin VB.TextBox txtGRH 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblBMP 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de BMP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de GRH:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   3600
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmGRHaBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo cmdCerrar_Click_Err
    
    Unload Me

    Exit Sub

cmdCerrar_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmGRHaBMP.cmdCerrar_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo Form_Load_Err
    
    Me.Icon = frmMain.Icon
    
    Exit Sub

Form_Load_Err:
    Call LogError(Err.Number, Err.Description, "frmGRHaBMP.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub txtGRH_Change()
    
    On Error GoTo txtGRH_Change_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    If txtGRH.Text <> "" And IsNumeric(txtGRH.Text) = True Then
        If txtGRH.Text > grhCount Then Exit Sub
        If txtGRH.Text < 1 Then Exit Sub
        lblBMP.Caption = GrhData(txtGRH.Text).FileNum

    End If

    Exit Sub

txtGRH_Change_Err:
    Call LogError(Err.Number, Err.Description, "frmGRHaBMP.txtGRH_Change", Erl)
    Resume Next
    
End Sub

Private Sub txtGRH_KeyPress(KeyAscii As Integer)
    
    On Error GoTo txtGRH_KeyPress_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0

    End If

    Exit Sub

txtGRH_KeyPress_Err:
    Call LogError(Err.Number, Err.Description, "frmGRHaBMP.txtGRH_KeyPress", Erl)
    Resume Next
    
End Sub
