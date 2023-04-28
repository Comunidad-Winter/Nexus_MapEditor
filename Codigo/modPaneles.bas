Attribute VB_Name = "modPaneles"
Option Explicit

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
            If Not frmBordes.Visible Then
                frmBordes.Show
            Else
                frmBordes.Visible = False
            End If
            
    End Select
    
    Exit Sub

VerFuncion_Err:
    Call LogError(Err.Number, Err.Description, "modPaneles.VerFuncion", Erl)
    Resume Next
    
End Sub
