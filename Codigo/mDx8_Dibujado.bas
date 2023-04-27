Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

Private DrawBuffer As cDIBSection

Sub DrawGrhtoHdc(ByRef Pic As PictureBox, _
                 ByVal GrhIndex As Long, _
                 ByRef DestRect As RECT)

    '*****************************************************************
    'Draws a Grh's portion to the given area of any Device Context
    '*****************************************************************
         
    DoEvents
    
    Pic.AutoRedraw = False
        
    'Clear the inventory window
    Call Engine_BeginScene
        
    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, Normal_RGBList())
        
    Call Engine_EndScene(DestRect, Pic.hWnd)
    
    Call DrawBuffer.LoadPictureBlt(Pic.hDC)

    Pic.AutoRedraw = True

    Call DrawBuffer.PaintPicture(Pic.hDC, 0, 0, Pic.Width, Pic.Height, 0, 0, vbSrcCopy)

    Pic.Picture = Pic.Image
        
End Sub

Public Sub PrepareDrawBuffer()
    Set DrawBuffer = New cDIBSection
    'El tamanio del buffer es arbitrario = 1024 x 1024
    Call DrawBuffer.Create(1024, 1024)
End Sub

Public Sub CleanDrawBuffer()
    Set DrawBuffer = Nothing
End Sub
