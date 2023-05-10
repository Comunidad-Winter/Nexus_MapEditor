Attribute VB_Name = "modPaneles"
Option Explicit

Private DrawBuffer As cDIBSection

''
' Activa/Desactiva el Estado de la Funcion en el Panel Superior
'
' @param Numero Especifica en numero de funcion
' @param Activado Especifica si esta o no activado

Public Sub EstSelectPanel(ByVal Numero As Byte, ByVal Activado As Boolean)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 30/05/06
    '*************************************************
    If Activado = True Then
        frmMain.SelectPanel(Numero).GradientMode = lv_Bottom2Top
        frmMain.SelectPanel(Numero).HoverBackColor = frmMain.SelectPanel(Numero).GradientColor

        If frmMain.mnuVerAutomatico.Checked = True Then

            Select Case Numero

                Case 0
                    If frmSuperficies.cCapas.Text = 4 Then
                        frmMain.mnuVerCapa4.Tag = CInt(frmMain.mnuVerCapa4.Checked)
                        frmMain.mnuVerCapa4.Checked = True
                        
                    ElseIf frmSuperficies.cCapas.Text = 3 Then
                        frmMain.mnuVerCapa3.Tag = CInt(frmMain.mnuVerCapa3.Checked)
                        frmMain.mnuVerCapa3.Checked = True
                        
                    ElseIf frmSuperficies.cCapas.Text = 2 Then
                        frmMain.mnuVerCapa2.Tag = CInt(frmMain.mnuVerCapa2.Checked)
                        frmMain.mnuVerCapa2.Checked = True

                    End If
                    
                Case 2
                    frmBloqueos.cVerBloqueos.Tag = CInt(frmBloqueos.cVerBloqueos.value)
                    frmBloqueos.cVerBloqueos.value = True
                    frmMain.mnuVerBloqueos.Checked = frmBloqueos.cVerBloqueos.value
                    
                Case 6
                    frmTriggers.cVerTriggers.Tag = CInt(frmTriggers.cVerTriggers.value)
                    frmTriggers.cVerTriggers.value = True
                    frmMain.mnuVerTriggers.Checked = frmTriggers.cVerTriggers.value
                    
            End Select

        End If

    Else
        frmMain.SelectPanel(Numero).HoverBackColor = frmMain.SelectPanel(Numero).BackColor
        frmMain.SelectPanel(Numero).GradientMode = lv_NoGradient

        If frmMain.mnuVerAutomatico.Checked = True Then

            Select Case Numero

                Case 0
                    If frmSuperficies.cCapas.Text = 4 Then
                        If LenB(frmMain.mnuVerCapa3.Tag) <> 0 Then frmMain.mnuVerCapa4.Checked = CBool(frmMain.mnuVerCapa4.Tag)
                    ElseIf frmSuperficies.cCapas.Text = 3 Then

                        If LenB(frmMain.mnuVerCapa3.Tag) <> 0 Then frmMain.mnuVerCapa3.Checked = CBool(frmMain.mnuVerCapa3.Tag)
                    ElseIf frmSuperficies.cCapas.Text = 2 Then

                        If LenB(frmMain.mnuVerCapa2.Tag) <> 0 Then frmMain.mnuVerCapa2.Checked = CBool(frmMain.mnuVerCapa2.Tag)

                    End If
                    
                Case 2

                    If LenB(frmBloqueos.cVerBloqueos.Tag) = 0 Then frmBloqueos.cVerBloqueos.Tag = 0
                    frmBloqueos.cVerBloqueos.value = CBool(frmBloqueos.cVerBloqueos.Tag)
                    frmMain.mnuVerBloqueos.Checked = frmBloqueos.cVerBloqueos.value
                    
                Case 6

                    If LenB(frmTriggers.cVerTriggers.Tag) = 0 Then frmTriggers.cVerTriggers.Tag = 0
                    frmTriggers.cVerTriggers.value = CBool(frmTriggers.cVerTriggers.Tag)
                    frmMain.mnuVerTriggers.Checked = frmTriggers.cVerTriggers.value
                    
            End Select

        End If

    End If

End Sub

''
' Muestra los controles que componen a la funcion seleccionada del Panel
'
' @param Numero Especifica el numero de Funcion

Public Sub VerFuncion(ByVal Numero As Byte)
    
    '*************************************************
    'Author: Lorwik
    'Last modified: 208/04/2023
    '*************************************************
    
    On Error GoTo VerFuncion_Err

    Select Case Numero

        Case 0 ' Superficies

            If Not frmSuperficies.Visible Then
                frmSuperficies.Show , frmMain
            Else
                frmSuperficies.Visible = False

            End If
            
        Case 1 ' Traslados

            If Not frmTraslados.Visible Then
                frmTraslados.Show , frmMain
            Else
                frmTraslados.Visible = False

            End If

        Case 2 ' Bloqueos

            If Not frmBloqueos.Visible Then
                frmBloqueos.Show , frmMain
            Else
                frmBloqueos.Visible = False

            End If
            
        Case 3 ' NPCs

            If Not frmNpcs.Visible Then
                frmNpcs.Show , frmMain
            Else
                frmNpcs.Visible = False

            End If

        Case 4 ' OBJs

            If Not frmObjetos.Visible Then
                frmObjetos.Show , frmMain
            Else
                frmObjetos.Visible = False

            End If

        Case 5 ' Triggers

            If Not frmTriggers.Visible Then
                frmTriggers.Show , frmMain

            Else
                frmTriggers.Visible = False

            End If
            
        Case 6 ' Particulas

            If Not frmParticulas.Visible Then
                frmParticulas.Show , frmMain
            Else
                frmParticulas.Visible = False

            End If
    
        Case 7 ' Luces

            If Not frmLuces.Visible Then
                frmLuces.Show , frmMain
            Else
                frmLuces.Visible = False

            End If
            
        Case 8 'Bordes

            If Not frmCopiarBordes.Visible Then
                frmCopiarBordes.Show , frmMain
            Else
                frmCopiarBordes.Visible = False

            End If
            
        Case 9 'Información del MApa

            If Not frmMapInfo.Visible Then
                frmMapInfo.Show , frmMain
                frmMapInfo.Visible = True
            Else
                frmMapInfo.Visible = False

            End If
            
        Case 10 'Preview
            If Not frmPreview.Visible Then
                frmPreview.Show , frmMain
            Else
                frmPreview.Visible = False
                
            End If
            
    End Select
    
    Exit Sub

VerFuncion_Err:
    Call LogError(Err.Number, Err.Description, "modPaneles.VerFuncion", Erl)

    Resume Next
    
End Sub

Public Function DameGrhIndex(ByVal GrhIn As Long) As Long
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    DameGrhIndex = SupData(GrhIn).Grh
    
    If SupData(GrhIn).Width > 0 Then
        frmConfigSup.MOSAICO.value = vbChecked
        frmConfigSup.mAncho.Text = SupData(GrhIn).Width
        frmConfigSup.mLargo.Text = SupData(GrhIn).Height
        
    Else
        frmConfigSup.MOSAICO.value = vbUnchecked
        frmConfigSup.mAncho.Text = "0"
        frmConfigSup.mLargo.Text = "0"
        
    End If

End Function

Public Sub fPreviewGrh(ByVal GrhIn As Long)
    '*************************************************
    'Author: Unkwown
    'Last modified: 22/05/06
    '*************************************************
    
    On Error GoTo fPreviewGrh_Err

    If Val(GrhIn) < 1 Then
        frmSuperficies.cGrh.Text = grhCount
        Exit Sub

    End If

    If Val(GrhIn) > grhCount Then
        frmSuperficies.cGrh.Text = 1
        Exit Sub

    End If

    'Change CurrentGrh
    CurrentGrh.GrhIndex = GrhIn
    CurrentGrh.Started = 1
    CurrentGrh.FrameCounter = 1

    Exit Sub

fPreviewGrh_Err:
    Call LogError(Err.Number, Err.Description, "modPaneles.fPreviewGrh", Erl)

    Resume Next
    
End Sub

Public Sub RenderPreview(Optional ByVal SinMosaico As Boolean = False)
    '*************************************************
    'Author: Lorwik
    'Last modified: 29/04/2023
    '*************************************************
    
    On Error Resume Next

    Dim destRect As RECT
    
    Dim i        As Integer

    Dim j        As Integer

    Dim ww       As Integer

    Dim hh       As Integer

    Dim Cont     As Integer
    
    With destRect
        .Bottom = frmPreview.PreviewGrh.Height
        .Right = frmPreview.PreviewGrh.Width

    End With
    
    'Si el Render no esta activo, salimos
    If Not frmPreview.PreviewGrh.Visible Then Exit Sub
    
    'Clear the inventory window
    Call Engine_BeginScene
    
    If frmConfigSup.MOSAICO = vbUnchecked Or SinMosaico Then
        Call Draw_GrhIndex(CurrentGrh.GrhIndex, frmPreview.PreviewGrh.Height / 2, frmPreview.PreviewGrh.Width - 50, 1, Normal_RGBList(), 0)
        
    Else
    
        hh = Val(frmConfigSup.mLargo)
        ww = Val(frmConfigSup.mAncho)
        
        For i = 1 To hh
            For j = 1 To ww
            
                Call Draw_GrhIndex(CurrentGrh.GrhIndex, j * 32, i * 32, 0, Normal_RGBList())

                If Cont < hh * ww Then Cont = Cont + 1
                CurrentGrh.GrhIndex = CurrentGrh.GrhIndex + 1
            Next
        Next
        
        CurrentGrh.GrhIndex = CurrentGrh.GrhIndex - Cont

    End If
    
    frmPreview.PreviewGrh.AutoRedraw = False

    Call Engine_EndScene(destRect, frmPreview.PreviewGrh.hwnd)

    Call DrawBuffer.LoadPictureBlt(frmPreview.PreviewGrh.hDC)

    frmPreview.PreviewGrh.AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmPreview.PreviewGrh.hDC, 0, 0, frmPreview.PreviewGrh.Width, frmPreview.PreviewGrh.Height, 0, 0, vbSrcCopy)

End Sub

Public Sub PonerAlAzar(ByVal n As Integer, T As Byte)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 by GS
    '*************************************************
    
    On Error GoTo PonerAlAzar_Err
    
    Dim objindex As Long
    Dim NPCIndex As Long
    Dim X, y, i
    Dim Head    As Integer
    Dim Body    As Integer
    Dim Heading As Byte
    Dim Leer    As New clsIniManager
    i = n

    modEdicion.Deshacer_Add "Aplicar " & IIf(T = 0, "Objetos", "NPCs") & " al Azar" ' Hago deshacer

    Do While i > 0
        X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
        y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
        Select Case T

            Case 0

                If MapData(X, y).OBJInfo.objindex = 0 Then
                    i = i - 1

                    If frmBloqueos.cInsertarBloqueo.value = True Then
                        MapData(X, y).Blocked = 1
                    Else
                        MapData(X, y).Blocked = 0

                    End If

                    If frmObjetos.cCantidad.Text > 0 Then
                        objindex = frmObjetos.cCantidad.Text
                        InitGrh MapData(X, y).ObjGrh, ObjData(objindex).GrhIndex
                        MapData(X, y).OBJInfo.objindex = objindex
                        MapData(X, y).OBJInfo.Amount = Val(frmObjetos.cCantidad.Text)

                        Select Case ObjData(objindex).ObjType ' GS

                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, y).Graphic(3) = MapData(X, y).ObjGrh

                        End Select

                    End If

                End If

            Case 1

                If (MapData(X, y).Blocked And &HF) <> &HF Then
                    i = i - 1

                    If frmNpcs.cNPC.Text > 0 Then
                        NPCIndex = frmNpcs.cNPC.Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call Char_Make(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(y), 2, 2, 2, 0, 0)
                        MapData(X, y).NPCIndex = NPCIndex

                    End If

                End If

            Case 2

                If (MapData(X, y).Blocked And &HF) <> &HF Then
                    i = i - 1

                    If frmNpcs.cNPC.Text >= 0 Then
                        NPCIndex = frmNpcs.cNPC.Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call Char_Make(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(y), 2, 2, 2, 0, 0)
                        MapData(X, y).NPCIndex = NPCIndex

                    End If

                End If

        End Select

        DoEvents
    Loop

    
    Exit Sub

PonerAlAzar_Err:
    Call LogError(Err.Number, Err.Description, "modPaneles.PonerAlAzar", Erl)
    Resume Next
    
End Sub

