VERSION 5.00
Begin VB.Form frmGraficos 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de Graficos"
   ClientHeight    =   7035
   ClientLeft      =   15585
   ClientTop       =   7305
   ClientWidth     =   3840
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
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox GraficosView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   263
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   3840
   End
   Begin VB.ListBox ListGraficosind 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   4050
      Width           =   3825
   End
End
Attribute VB_Name = "frmGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Dim lR As Long
    lR = SetTopMostWindow(frmGraficos.hWnd, True)
         
    Dim i As Long

    For i = 1 To grhCount
        ListGraficosind.AddItem i
    Next i

    Exit Sub

Form_Load_Err:
    Call LogError(Err.Number, Err.Description, "frmGraficos.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub ListGraficosind_Click()
    
    On Error GoTo ListGraficosind_Click_Err
    
    Dim DR As RECT
    
    With DR
        .Right = GraficosView.Width
        .Bottom = GraficosView.Height
    End With
    
    Call DrawGrhtoHdc(frmGraficos.GraficosView, str$(ListGraficosind.ListIndex And &HFFFF&) + 1, DR)
    frmConfigSup.MOSAICO.value = vbUnchecked
    frmConfigSup.mAncho.Text = "0"
    frmConfigSup.mLargo.Text = "0"
    HotKeysAllow = False

    Exit Sub

ListGraficosind_Click_Err:
    Call LogError(Err.Number, Err.Description, "frmGraficos.ListGraficosind_Click", Erl)
    Resume Next
    
End Sub

Private Sub ListGraficosind_DblClick()
    
    On Error GoTo ListGraficosind_DblClick_Err
    
    frmSuperficies.cGrh.Text = str$(ListGraficosind.ListIndex And &HFFFF&) + 1
    
    Exit Sub

ListGraficosind_DblClick_Err:
    Call LogError(Err.Number, Err.Description, "frmGraficos.ListGraficosind_DblClick", Erl)
    Resume Next
    
End Sub
