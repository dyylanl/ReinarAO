Attribute VB_Name = "Module2"
 
 
Option Explicit
 
Public Desvanecio As Boolean
Public Opacidad As Byte
 
        Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, _
ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, _
ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer
Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
 
 Sub SurfaceConColor(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByVal Rojo As Integer, ByVal Verde As Integer, ByVal azul As Integer, Optional ByVal KillAnim As Integer = 0)
 
Dim iGrhIndex As Integer
Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean
 
 
If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
 
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx )
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
 
                        End If
                    End If
               End If
            End If
        End If
    End If
End If
 
If Grh.GrhIndex = 0 Then Exit Sub
 
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
 
With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
 
Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim modo As Long
 
Set Src = SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum)
Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest
With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
   
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With
 
Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
Src.Lock SourceRect, ddsdSrc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
Surface.Lock rDest, ddsdDest, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
 
Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()
       
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
  modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
  modo = 565
Else
'  MsgBox "Modo de vídeo no esta en 555 o 565 o algo falló."
  End
End If
 
 
 
                       
            If Desvanecio = True Then
                If Val(Opacidad) = 20 Then
                    Desvanecio = False
                Else
                    Opacidad = Val(Opacidad) - 20
                End If
            End If
            If Desvanecio = False Then
                If Val(Opacidad) = 220 Then
                    Desvanecio = True
                Else
                    Opacidad = Val(Opacidad) + 20
                End If
            End If
           
 Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), Opacidad, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Rojo, Verde, azul)
Surface.Unlock rDest
Src.Unlock SourceRect
 
Exit Sub
 
HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest
 
End Sub
