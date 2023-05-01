VERSION 5.00
Begin VB.Form frmObjetos 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Objetos"
   ClientHeight    =   4905
   ClientLeft      =   9525
   ClientTop       =   11535
   ClientWidth     =   4125
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
   ScaleHeight     =   4905
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cCantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   3120
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
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
      ItemData        =   "frmObjetos.frx":0000
      Left            =   0
      List            =   "frmObjetos.frx":0002
      TabIndex        =   5
      Tag             =   "-1"
      Top             =   0
      Width           =   4125
   End
   Begin VB.ComboBox cOBJ 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.ComboBox cFiltro 
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000013&
      Height          =   315
      ItemData        =   "frmObjetos.frx":0004
      Left            =   690
      List            =   "frmObjetos.frx":0006
      TabIndex        =   3
      Top             =   3300
      Width           =   3285
   End
   Begin Nexus_MapEditor.lvButtons_H cQuitarObj 
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   4530
      Width           =   1935
      _extentx        =   3413
      _extenty        =   609
      caption         =   "Quitar OBJ"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmObjetos.frx":0008
      mode            =   1
      value           =   0
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cAgregarObjAzar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4110
      Width           =   1935
      _extentx        =   3413
      _extenty        =   661
      caption         =   "Insertar OBJ al azar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmObjetos.frx":0030
      mode            =   1
      value           =   0
      cback           =   -2147483633
   End
   Begin Nexus_MapEditor.lvButtons_H cInsertarObj 
      Height          =   765
      Left            =   2130
      TabIndex        =   2
      Top             =   4110
      Width           =   1845
      _extentx        =   3254
      _extenty        =   1349
      caption         =   "Insertar OBJ"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmObjetos.frx":0058
      mode            =   1
      value           =   0
      cback           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2370
      TabIndex        =   8
      Top             =   3750
      Width           =   705
   End
   Begin VB.Label lblFiltrar 
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lbnobj 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero del OBJ:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   3780
      Width           =   1200
   End
End
Attribute VB_Name = "frmObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cAgregarObjAzar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cAgregarObjAzar_Click_Err

    If IsNumeric(cCantidad.Text) = False Or cCantidad.Text > 200 Then
        MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
        Exit Sub

    End If

    cAgregarObjAzar.Enabled = False
    Call PonerAlAzar(cCantidad.Text, 0)
    cAgregarObjAzar.Enabled = True

    
    Exit Sub

cAgregarObjAzar_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmObjetos.cAgregarObjAzar_Click", Erl)
    Resume Next
End Sub

Private Sub cFiltro_LostFocus()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cFiltro_LostFocus_Err
    
    HotKeysAllow = True
    
    Exit Sub

cFiltro_LostFocus_Err:
    Call LogError(Err.Number, Err.Description, "FrmObjetos.cFiltro_LostFocus", Erl)
    Resume Next
    
End Sub

Private Sub cInsertarObj_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cInsertarFunc_Click_Err
    
    If cInsertarObj.value = True Then
        cQuitarObj.Enabled = False
        cAgregarObjAzar.Enabled = False

        Call modPaneles.EstSelectPanel(0, True)
    Else
        cQuitarObj.Enabled = True
        cAgregarObjAzar.Enabled = True

        Call modPaneles.EstSelectPanel(0, False)

    End If

    
    Exit Sub

cInsertarFunc_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmObjetos.cInsertarObj_Click", Erl)
    Resume Next
    
End Sub

Private Sub cQuitarObj_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo cQuitarFunc_Click_Err
    
    If cQuitarObj.value = True Then
        cInsertarObj.Enabled = False
        cAgregarObjAzar.Enabled = False
        cCantidad.Enabled = False
        cOBJ.Enabled = False
        cFiltro.Enabled = False
        lListado.Enabled = False
        Call modPaneles.EstSelectPanel(3, True)
    Else
        cInsertarObj.Enabled = True
        cAgregarObjAzar.Enabled = True
        cCantidad.Enabled = True
        cOBJ.Enabled = True
        cFiltro.Enabled = True
        lListado.Enabled = True
        Call modPaneles.EstSelectPanel(3, False)

    End If

    
    Exit Sub

cQuitarFunc_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmObjetos.cQuitarFunc_Click", Erl)
    Resume Next
End Sub

Private Sub lListado_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2023
    '*************************************************
    
    cOBJ.Text = ReadField(2, lListado.Text, Asc("#"))
    
    Call fPreviewGrh(ObjData(cOBJ.Text).GrhIndex)
    Call modPaneles.RenderPreview(True)
    
End Sub
