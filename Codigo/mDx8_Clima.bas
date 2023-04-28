Attribute VB_Name = "mDx8_Clima"
Option Explicit

'***************************************************
'Autor: Lorwik
'Descripción: Este sistema es una adaptación del que hice en
'las versiones anteriores de Imperium que posteriormente mejore en
'AODrag. El sistema fue adaptado al que trae AOLibre que a su vez
'se basaba en el de Blisse.
'***************************************************

Public Enum e_estados
    Amanecer = 0
    MedioDia
    Tarde
    noche
    Lluvia
    Niebla
    FogLluvia 'Niebla mas lluvia
End Enum

Public Estados(0 To 8) As D3DCOLORVALUE
Public Estado_Actual As D3DCOLORVALUE
Public Estado_Actual_Date As Byte

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
        .b = 230
        .r = 200
        .g = 200
    End With
    
    With Estados(e_estados.MedioDia)
        .a = 255
        .r = 255
        .g = 255
        .b = 255
    End With
    
    With Estados(e_estados.Tarde)
        .a = 255
        .b = 200
        .r = 230
        .g = 200
    End With
  
    With Estados(e_estados.noche)
        .a = 255
        .b = 170
        .r = 170
        .g = 170
    End With
    
    With Estados(e_estados.Lluvia)
        .a = 255
        .r = 200
        .g = 200
        .b = 200
    End With
    
    Estado_Actual_Date = 1
    
End Sub

Public Sub Actualizar_Estado(ByVal Estado As Byte)
'***************************************************
'Author: Lorwik
'Last Modification: 09/08/2020
'Actualiza el estado del clima y del dia
'***************************************************
    Dim X As Byte, y As Byte

    'Primero actualizamos la imagen del frmmain
    'Call ActualizarImgClima

    '¿El mapa tiene su propia luz?
'    If MapInfo.LuzBase <> -1 Then
'
'        For X = XMinMapSize To XMaxMapSize
'            For y = YMinMapSize To YMaxMapSize
'                Call Engine_Long_To_RGB_List(MapData(X, y).Engine_Light(), MapInfo.LuzBase)
'            Next y
'        Next X
'
'        Call LightRenderAll
'
'        Exit Sub
'    End If

    '¿Es un estado invalido?
    Estado = e_estados.noche
        
    Estado_Actual = Estados(Estado)
    Estado_Actual_Date = Estado
        
    For X = XMinMapSize To XMaxMapSize
        For y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, y).Engine_Light(), Estado_Actual)
        Next y
    Next X
        
    Call LightRenderAll

End Sub
