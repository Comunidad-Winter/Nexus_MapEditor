Attribute VB_Name = "mDx8_Clima"
Option Explicit

'***************************************************
'Autor: Lorwik
'Descripci�n: Este sistema es una adaptaci�n del que hice en
'las versiones anteriores de Imperium que posteriormente mejore en
'AODrag. El sistema fue adaptado al que trae AOLibre que a su vez
'se basaba en el de Blisse.
'***************************************************

Public Enum e_estados
    Amanecer = 0
    MedioDia
    Tarde
    Noche
End Enum

Public Estados(0 To 8) As D3DCOLORVALUE
Public Estado_Actual As D3DCOLORVALUE
Public Estado_Custom As D3DCOLORVALUE

Private m_Hora_Actual As Long
Private m_Last_Hora_Actual As Long

Public Sub Init_MeteoEngine()
'***************************************************
'Author: Standelf
'Last Modification: 15/05/10
'Initializate
'***************************************************
    With Estados(e_estados.Amanecer)
        .a = 255
        .B = 230
        .R = 200
        .G = 200
    End With
    
    With Estados(e_estados.MedioDia)
        .a = 255
        .R = 255
        .G = 255
        .B = 255
    End With
    
    With Estados(e_estados.Tarde)
        .a = 255
        .B = 200
        .R = 230
        .G = 200
    End With
  
    With Estados(e_estados.Noche)
        .a = 255
        .B = 170
        .R = 170
        .G = 170
    End With
    
End Sub

Public Sub Actualizar_Estado()

    '***************************************************
    'Author: Lorwik
    'Last Modification: 09/08/2020
    'Actualiza el estado del clima y del dia
    '***************************************************
    Dim X  As Byte, Y As Byte
    Dim tR As Byte
    Dim tG As Byte
    Dim tB As Byte
    
    '�El mapa tiene su propia luz?
    If MapInfo.LuzBase <> 0 Then
        
        Call ConvertLongToRGB(MapInfo.LuzBase, tR, tG, tB)
                    
        With Estado_Custom
            .a = 255
            .R = tR
            .G = tG
            .B = tB

        End With
        
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Custom)
            Next Y
        Next X
            
        Call LightRenderAll
            
        Exit Sub

    End If
            
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X
    
    Call LightRenderAll

End Sub
