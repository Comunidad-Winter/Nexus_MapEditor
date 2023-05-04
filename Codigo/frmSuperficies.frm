VERSION 5.00
Begin VB.Form frmSuperficies 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Superficies"
   ClientHeight    =   4935
   ClientLeft      =   9525
   ClientTop       =   9210
   ClientWidth     =   4080
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
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   ShowInTaskbar   =   0   'False
   Begin Nexus_MapEditor.lvButtons_H cQuitarEnTodasLasCapas 
      Height          =   345
      Left            =   90
      TabIndex        =   8
      Top             =   4530
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      Caption         =   "Quitar en Capas 2 y 3"
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
   Begin Nexus_MapEditor.lvButtons_H cQuitarEnEstaCapa 
      Height          =   375
      Left            =   90
      TabIndex        =   7
      Top             =   4110
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Quitar en esta Capa"
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
   Begin Nexus_MapEditor.lvButtons_H cSeleccionarSuperficie 
      Height          =   765
      Left            =   2130
      TabIndex        =   6
      Top             =   4110
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1349
      Caption         =   "Insertar Superficie"
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
   Begin VB.ComboBox cCapas 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   315
      ItemData        =   "frmSuperficies.frx":0000
      Left            =   1020
      List            =   "frmSuperficies.frx":0002
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   3660
      Width           =   855
   End
   Begin VB.ComboBox cFiltro 
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   690
      TabIndex        =   2
      Top             =   3300
      Width           =   3285
   End
   Begin VB.ComboBox cGrh 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   2850
      TabIndex        =   1
      Top             =   3660
      Width           =   1125
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
      ItemData        =   "frmSuperficies.frx":0004
      Left            =   -30
      List            =   "frmSuperficies.frx":0006
      TabIndex        =   0
      Tag             =   "-1"
      Top             =   0
      Width           =   4125
   End
   Begin VB.Label lblFiltrar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   3330
      Width           =   480
   End
   Begin VB.Label lbGrh 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Sup Actual:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   3720
      Width           =   825
   End
   Begin VB.Label lbCapas 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Capa Actual:"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   3735
      Width           =   930
   End
End
Attribute VB_Name = "frmSuperficies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cSeleccionarSuperficie_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cSeleccionarSuperficie.value = True Then
        cQuitarEnTodasLasCapas.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        cQuitarEnTodasLasCapas.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If

End Sub

Private Sub cQuitarEnEstaCapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnEstaCapa.value = True Then
        lListado.Enabled = False
        cFiltro.Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnTodasLasCapas.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        lListado.Enabled = True
        cFiltro.Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnTodasLasCapas.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If

End Sub

Private Sub cQuitarEnTodasLasCapas_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnTodasLasCapas.value = True Then
        cCapas.Enabled = False
        lListado.Enabled = False
        cFiltro.Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        cCapas.Enabled = True
        lListado.Enabled = True
        cFiltro.Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If

End Sub

Private Sub cGrh_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo Fallo

    If KeyAscii = 13 Then
        Call fPreviewGrh(cGrh.Text)

        If cGrh.ListCount > 5 Then
            cGrh.RemoveItem 0

        End If

        cGrh.AddItem cGrh.Text

    End If

    Exit Sub
Fallo:
    cGrh.Text = 1

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
    Call LogError(Err.Number, Err.Description, "FrmMain.cFiltro_LostFocus", Erl)

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
    
        vMaximo = MaxSup
        
        For i = 0 To vMaximo
        
            vDatos = SupData(i).name
            NumI = i
            
            For j = 1 To Len(vDatos)

                If UCase$(mid$(vDatos & str(i), j, Len(cFiltro.Text))) = UCase$(cFiltro.Text) Or LenB(cFiltro.Text) = 0 Then
                    lListado.AddItem vDatos & " - #" & NumI
                    Exit For

                End If

            Next j
        Next i
    
    End If

End Sub

Private Sub Form_Load()

    Dim i As Byte
    
    For i = 1 To 4
        cCapas.AddItem i
    Next i
    
    cCapas.ListIndex = 0

End Sub

Private Sub lListado_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: ?????
    '*************************************************
    
    cGrh.Text = DameGrhIndex(ReadField(2, lListado.Text, Asc("#")))
                
    If SupData(ReadField(2, lListado.Text, Asc("#"))).Capa <> 0 Then
        If LenB(ReadField(2, lListado.Text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.Text
        cCapas.Text = SupData(ReadField(2, lListado.Text, Asc("#"))).Capa
    Else

        If LenB(cCapas.Tag) <> 0 Then
            cCapas.Text = cCapas.Tag
            cCapas.Tag = vbNullString

        End If

    End If
                
    If SupData(ReadField(2, lListado.Text, Asc("#"))).Block = True Then
        If LenB(frmBloqueos.cInsertarBloqueo.Tag) = 0 Then frmBloqueos.cInsertarBloqueo.Tag = IIf(frmBloqueos.cInsertarBloqueo.value = True, 1, 0)
        frmBloqueos.cInsertarBloqueo.value = True
        Call frmBloqueos.InsertarBloqueo
    Else

        If LenB(frmBloqueos.cInsertarBloqueo.Tag) <> 0 Then
            frmBloqueos.cInsertarBloqueo.value = IIf(Val(frmBloqueos.cInsertarBloqueo.Tag) = 1, True, False)
            frmBloqueos.cInsertarBloqueo.Tag = vbNullString
            Call frmBloqueos.InsertarBloqueo

        End If

    End If

    Call fPreviewGrh(cGrh.Text)
    Call modPaneles.RenderPreview

End Sub
