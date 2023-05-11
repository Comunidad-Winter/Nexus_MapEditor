VERSION 5.00
Begin VB.Form frmNpcs 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Npc's"
   ClientHeight    =   4965
   ClientLeft      =   9525
   ClientTop       =   9210
   ClientWidth     =   4110
   ControlBox      =   0   'False
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
   ScaleHeight     =   4965
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lListado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   3180
      ItemData        =   "frmNpcs.frx":0000
      Left            =   0
      List            =   "frmNpcs.frx":0002
      TabIndex        =   5
      Tag             =   "-1"
      Top             =   0
      Width           =   4125
   End
   Begin VB.ComboBox cNPC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   3690
      Width           =   2595
   End
   Begin VB.ComboBox cFiltro 
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   315
      ItemData        =   "frmNpcs.frx":0004
      Left            =   690
      List            =   "frmNpcs.frx":0006
      TabIndex        =   3
      Top             =   3300
      Width           =   3285
   End
   Begin Nexus_MapEditor.lvButtons_H cQuitarNpc 
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   4530
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      Caption         =   "Quitar NPC's"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cAgregarFuncalAzar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4110
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Insertar NPC's al azar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cInsertarFunc 
      Height          =   765
      Left            =   2130
      TabIndex        =   2
      Top             =   4110
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1349
      Caption         =   "Insertar NPC"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label lbFiltrar 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lbnNPC 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero del NPC:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   3780
      Width           =   1215
   End
End
Attribute VB_Name = "frmNpcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cFiltro_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cFiltro_LostFocus_Err
    
    HotKeysAllow = True

    
    Exit Sub

cFiltro_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "FrmNpcs.cFiltro_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub cFiltro_KeyPress(KeyAscii As Integer)

    '*************************************************
    'Author: Lorwik
    'Last modified: 04/05/2023
    '*************************************************
    If KeyAscii = 13 Then
        
        Dim vMaximo As Integer
        Dim vDatos  As String
        Dim NumI    As Integer
        Dim i       As Integer
        Dim j       As Integer
    
        If cFiltro.ListCount > 5 Then cFiltro.RemoveItem 0
        
        cFiltro.AddItem cFiltro.Text
        lListado.Clear
    
        vMaximo = NumNPCs - 1
        
        For i = 0 To vMaximo
        
            vDatos = NpcData(i + 1).name
            NumI = i + 1
            
            For j = 1 To Len(vDatos)

                If UCase$(mid$(vDatos & str(i), j, Len(cFiltro.Text))) = UCase$(cFiltro.Text) Or LenB(cFiltro.Text) = 0 Then
                    lListado.AddItem vDatos & " - #" & NumI
                    Exit For

                End If

            Next j
        Next i
    
    End If

End Sub

Private Sub cInsertarFunc_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cInsertarFunc_Click_Err
    
    If cInsertarFunc.value = True Then
        cQuitarNpc.Enabled = False
        cAgregarFuncalAzar.Enabled = False

        Call modPaneles.EstSelectPanel(3, True)
    Else
        cQuitarNpc.Enabled = True
        cAgregarFuncalAzar.Enabled = True

        Call modPaneles.EstSelectPanel(3, False)

    End If

    Exit Sub

cInsertarFunc_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmNpcs.cInsertarFunc_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarNpc_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cQuitarNpc_Click_Err
    
    If cQuitarNpc.value = True Then
        cInsertarFunc.Enabled = False
        cAgregarFuncalAzar.Enabled = False
        cNPC.Enabled = False
        cFiltro.Enabled = False
        lListado.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cInsertarFunc.Enabled = True
        cAgregarFuncalAzar.Enabled = True
        cNPC.Enabled = True
        cFiltro.Enabled = True
        lListado.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

    
    Exit Sub

cQuitarNpc_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmObjetos.cQuitarNpc_Click", Erl)
    Resume Next
    
End Sub

Private Sub cAgregarFuncalAzar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cAgregarFuncalAzar_Click_Err

    cAgregarFuncalAzar.Enabled = False
    Call PonerAlAzar(1, 2)
    cAgregarFuncalAzar.Enabled = True
    
    Exit Sub

cAgregarFuncalAzar_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmNpcs.cAgregarFuncalAzar_Click", Erl)
    Resume Next
    
End Sub

Private Sub lListado_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    
    On Error GoTo lListado_MouseMove_Err

    HotKeysAllow = False

    Exit Sub

lListado_MouseMove_Err:
    Call LogError(Err.Number, Err.Description, "FrmNPCs.lListado_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub lListado_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: ?????
    '*************************************************
    
    If HotKeysAllow = False Then _
        lListado.Tag = lListado.ListIndex
        
    cNPC.Text = ReadField(2, lListado.Text, Asc("#"))
End Sub
