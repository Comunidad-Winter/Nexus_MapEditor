VERSION 5.00
Begin VB.Form FrmRender 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   12810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12495
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   854
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   833
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBuscarErrores 
      Caption         =   "Buscar Errores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8760
      TabIndex        =   5
      Top             =   120
      Width           =   2130
   End
   Begin VB.Timer SaveAllErrores 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11280
      Top             =   480
   End
   Begin VB.Timer SaveAll 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11880
      Top             =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Renderizar todos los mapas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Renderizar sin bordes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12000
      Left            =   240
      ScaleHeight     =   800
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   2
      Top             =   600
      Width           =   12000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Renderizar con bordes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11040
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Smallpic 
      Height          =   5535
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "FrmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'*************************************************************
' Capturar la imagen de controles
       
'  1 - Colocar un picturebox llamado picture1, un Command1 y un Command2 _
   2 - Agragar algunos controles _
   3 - Indicar en la Sub " Capturar_Imagen " .. el control a capturar
'*************************************************************
      
' Declaraciones del Api
      
'*************************************************************
' Funci�n BitBlt para copiar la imagen del control en un picturebox
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
      
' Recupera la imagen del �rea del control
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Dim handle As Integer

Private Sub cmdAceptar_Click()
    
    On Error GoTo cmdAceptar_Click_Err
    
    Call MapCapture(False, True)
    
    Exit Sub

cmdAceptar_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmRender.cmdAceptar_Click", Erl)
    Resume Next
    
End Sub

'*************************************************************
' Sub que copia la imagen del control en un picturebox
'*************************************************************
Public Sub Capturar_Imagen(Control As Control, Destino As Object)
          
    Dim hdc             As Long
    Dim Escala_Anterior As Integer
    Dim Ancho           As Long
    Dim Alto            As Long
          
    ' Para que se mantenga la imagen por si se repinta la ventana
    Destino.AutoRedraw = True
          
    On Error Resume Next

    ' Si da error es por que el control est� dentro de un Frame _
      ya que  los Frame no tiene  dicha propiedad
    Escala_Anterior = Control.Container.ScaleMode
          
    If Err.Number = 438 Then
        ' Si el control est� en un Frame, convierte la escala
        Ancho = ScaleX(Control.Width, vbTwips, vbPixels)
        Alto = ScaleY(Control.Height, vbTwips, vbPixels)
    Else
        ' Si no cambia la escala del  contenedor a pixeles
        Control.Container.ScaleMode = vbPixels
        Ancho = Control.Width
        Alto = Control.Height

    End If
          
    ' limpia el error
    On Error GoTo 0

    ' Captura el �rea de pantalla correspondiente al control
    hdc = GetWindowDC(Control.hWnd)
    
    ' Copia esa �rea al picturebox
    If ToWorldMap2 Then
        'Call BitBlt(Destino.hdc, 0 - 50, 0 - 50, Ancho - 50, Alto - 50, hdc, 0, 0, vbSrcCopy) '
        Call BitBlt(Destino.hdc, 0, 0, Ancho, Alto, hdc, 0, 0, vbSrcCopy)
    Else
        Call BitBlt(Destino.hdc, 0, 0, 3000, 3000, hdc, 0, 0, vbSrcCopy)
        

    End If
    
    ' Convierte la imagen anterior en un Mapa de bits
    Destino.Picture = Destino.Image
    
    ' Borra la imagen ya que ahora usa el Picture
    Call Destino.Cls
          
    On Error Resume Next

    If Err.Number = 0 Then
        ' Si el control no est� en un  Frame, restaura la escala del contenedor
        Control.Container.ScaleMode = Escala_Anterior

    End If
          
End Sub

Private Sub cmdBuscarErrores_Click()

    Dim FileName As String
    FileName = PATH_Save & NameMap_Save & "1.csm"

    If FileExist(FileName, vbArchive) = False Then
        Unload Me
        MsgBox "Primero abre alg�n mapa de la carpeta a convertir.", vbOKOnly, "Error"
        Exit Sub
    End If
    
    Call AbrirMapa(FileName)

    SaveAllErrores.Enabled = True
    
    handle = FreeFile

    If Dir(App.Path & "\errores.txt", vbArchive) <> "" Then
        Kill (App.Path & "\errores.txt")
    End If

    Open App.Path & "\errores.txt" For Append As #handle
    
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Unload Me

    
    Exit Sub

Command1_Click_Err:
    Call LogError(Err.Number, Err.Description, "FrmRender.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    Call MapCapture(False, False)
End Sub

Private Sub Command3_Click()

    Dim FileName As String
    
    frmBloqueos.cVerBloqueos.value = (frmBloqueos.cVerBloqueos.value = False)
    frmMain.mnuVerBloqueos.Checked = frmBloqueos.cVerBloqueos.value
    
    frmMain.mnuVerTranslados.Checked = (frmMain.mnuVerTranslados.Checked = False)
    
    frmTriggers.cVerTriggers.value = (frmTriggers.cVerTriggers.value = False)
    frmMain.mnuVerTriggers.Checked = frmTriggers.cVerTriggers.value
            
    FileName = PATH_Save & NameMap_Save & "1.csm"


    If FileExist(FileName, vbArchive) = False Then
        Unload Me
        MsgBox "Primero abre alg�n mapa de la carpeta a convertir.", vbOKOnly, "Error"
        Exit Sub
    End If
    
    Call AbrirMapa(FileName)
    
    SaveAll.Enabled = True

End Sub

Private Function IsBlock(ByVal X As Integer, ByVal y As Integer) As Boolean

    If X - 1 < XMinMapSize Or X + 1 > XMaxMapSize Then
        IsBlock = True
        Exit Function
    End If
    
    If y - 1 < YMinMapSize Or y + 1 > YMaxMapSize Then
        IsBlock = True
        Exit Function
    End If
    
    IsBlock = (MapData(X, y).Blocked And &HF) = &HF
    
End Function

Private Sub SaveAll_Timer()

    Call MapCapture(False, False)
        
    If Not frmMain.MapPest(5).Visible Then
        SaveAll.Enabled = False
        Exit Sub
    End If

    Call frmMain.NextMap
End Sub

Private Sub SaveAllErrores_Timer()
     
    Dim X As Integer, y As Integer, BordeX As Integer, BordeY As Integer
    
    BordeX = 11
    BordeY = 8
    
    If frmMapInfo.txtMapNombre.Text = "" Then
        Print #handle, frmMain.MapPest(4).Caption & " :::: " & "Este mapa no tiene Nombre."
    End If
               
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
                
            If Not IsBlock(X, y) Then
                If IsBlock(X - 1, y) And IsBlock(X + 1, y) And IsBlock(X, y + 1) And IsBlock(X, y - 1) Then
                    Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: Falta Bloqueo."
                End If
            End If

            If MapData(X, y).NPCIndex Then
                If NpcData(MapData(X, y).NPCIndex).Body = 0 Then
                    Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: NPC BODY 0 "; MapData(X, y).NPCIndex
                Else

                    If BodyData(NpcData(MapData(X, y).NPCIndex).Body).Walk(1).GrhIndex = 0 Then
                        Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: NPC BODY SIN GRH "; MapData(X, y).NPCIndex
                    End If
                End If
            End If
            
            If MapData(X, y).OBJInfo.objindex Then
                If ObjData(MapData(X, y).OBJInfo.objindex).ObjType = 6 And ObjData(MapData(X, y).OBJInfo.objindex).Subtipo = 0 Then
                
                    If X > (XMinMapSize + BordeX) And X < (XMaxMapSize - BordeX) And y > (YMinMapSize + BordeY) And y < (YMaxMapSize - BordeY) Then
                        If Not IsBlock(X + 1, y) And Not ((MapData(X + 1, y).Blocked And 1) <> 0 And (MapData(X + 1, y + 1).Blocked And 4) <> 0) Then
                            Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y & " :::: FALTA BLOQUEO TOTAL"
                        End If
                    
                        If ObjData(MapData(X, y).OBJInfo.objindex).Cerrada = 1 Then
                            If (MapData(X, y).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y + 1).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y + 1).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                        Else

                            If (MapData(X, y).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y + 1).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y + 1 & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y + 1).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y + 1 & " :::: HAY BLOQUEO PARCIAL"
                            End If
                        End If
                    End If
                End If
                
                If ObjData(MapData(X, y).OBJInfo.objindex).ObjType = 6 And ObjData(MapData(X, y).OBJInfo.objindex).Subtipo = 2 Then
                
                    If X > (XMinMapSize + BordeX) And X < (XMaxMapSize - BordeX) And y > (YMinMapSize + BordeY) And y < (YMaxMapSize - BordeY) Then
                    
                        If Not IsBlock(X + 2, y - 1) And Not ((MapData(X + 2, y - 1).Blocked And 1) <> 0 And (MapData(X + 2, y + 2).Blocked And 4) <> 0) Then
                            Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 2 & ", " & y - 1 & " :::: FALTA BLOQUEO TOTAL"
                        End If
                        
                        If Not IsBlock(X - 2, y - 1) And Not ((MapData(X - 2, y - 1).Blocked And 1) <> 0 And (MapData(X - 2, y + 2).Blocked And 4) <> 0) Then
                            Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 2 & ", " & y - 1 & " :::: FALTA BLOQUEO TOTAL"
                        End If
                    
                        If ObjData(MapData(X, y).OBJInfo.objindex).Cerrada = 1 Then
                                                                           
                            If (MapData(X - 1, y - 1).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y - 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y - 1).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y - 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                                                                                    
                            If (MapData(X + 1, y - 1).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y - 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X + 1, y).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                        Else

                            If (MapData(X, y - 1).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y - 1 & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y - 1).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y - 1 & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X + 1, y - 1).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y - 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X + 1, y).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: HAY BLOQUEO PARCIAL"
                            End If
                        End If
                    End If
                End If

                If ObjData(MapData(X, y).OBJInfo.objindex).ObjType = 6 And ObjData(MapData(X, y).OBJInfo.objindex).Subtipo = 3 Then
                
                    If X > (XMinMapSize + BordeX) And X < (XMaxMapSize - BordeX) And y > (YMinMapSize + BordeY) And y < (YMaxMapSize - BordeY) Then
                    
                        If Not IsBlock(X + 2, y) And Not ((MapData(X + 2, y).Blocked And 1) <> 0 And (MapData(X + 2, y).Blocked And 4) <> 0) Then
                            Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 2 & ", " & y & " :::: FALTA BLOQUEO TOTAL"
                        End If
                        
                        If Not IsBlock(X - 2, y) And Not ((MapData(X - 2, y).Blocked And 1) <> 0 And (MapData(X - 2, y).Blocked And 4) <> 0) Then
                            Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 2 & ", " & y & " :::: FALTA BLOQUEO TOTAL"
                        End If
                    
                        If ObjData(MapData(X, y).OBJInfo.objindex).Cerrada = 1 Then
                                                                           
                            If (MapData(X - 1, y).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                                                                                    
                            If (MapData(X + 1, y).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y + 1).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y + 1).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X + 1, y + 1).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                        Else

                            If (MapData(X, y).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y + 1).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y + 1&; " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X + 1, y).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X + 1, y + 1).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y + 1).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y + 1 & " :::: HAY BLOQUEO PARCIAL"
                            End If
                        End If
                    End If
                End If

                If ObjData(MapData(X, y).OBJInfo.objindex).ObjType = 6 And ObjData(MapData(X, y).OBJInfo.objindex).Subtipo = 4 Then
                    If X > (XMinMapSize + BordeX) And X < (XMaxMapSize - BordeX) And y > (YMinMapSize + BordeY) And y < (YMaxMapSize - BordeY) Then
                    
                        If Not IsBlock(X + 1, y) And Not ((MapData(X + 1, y).Blocked And 1) <> 0 And (MapData(X + 1, y).Blocked And 4) <> 0) Then
                            Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y & " :::: FALTA BLOQUEO TOTAL"
                        End If
                        
                        If Not IsBlock(X - 1, y) And Not ((MapData(X - 1, y).Blocked And 1) <> 0 And (MapData(X - 1, y).Blocked And 4) <> 0) Then
                            Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y & " :::: FALTA BLOQUEO TOTAL"
                        End If
                    
                        If ObjData(MapData(X, y).OBJInfo.objindex).Cerrada = 1 Then
                            
                            If (MapData(X, y).Blocked And 1) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y + 1).Blocked And 4) = 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                        Else

                            If (MapData(X, y).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y & " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X - 1, y + 1).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X - 1 & ", " & y + 1&; " :::: HAY BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X + 1, y).Blocked And 1) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X + 1, y + 1).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X + 1 & ", " & y + 1 & " :::: FALTA BLOQUEO PARCIAL"
                            End If
                            
                            If (MapData(X, y + 1).Blocked And 4) <> 0 Then
                                Print #handle, frmMain.MapPest(4).Caption & " :::: Posici�n: " & X & ", " & y + 1 & " :::: HAY BLOQUEO PARCIAL"
                            End If
                        End If
                    End If
                    
                End If

            End If
        Next
    Next

    If Not frmMain.MapPest(5).Visible Then
        SaveAllErrores.Enabled = False
        Close #handle
        Exit Sub
    End If

    Call frmMain.NextMap
    
End Sub
