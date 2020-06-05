Attribute VB_Name = "Dx8_Effects"
'/////////////////////////////Motor Grafico en DirectX 8///////////////////////////////
'////////////////////////Extraccion de varios motores por ShaFTeR//////////////////////
'///////////////////ORE - VBGORE - GSZAO - KKAO y algunos ejemplos de webs/////////////
'**************************************************************************************
Option Explicit


Public WeatherEffectIndex As Integer
Public LastWeather As Byte


Public Type Projectile

    X As Single
    Y As Single
    tX As Single
    tY As Single
    RotateSpeed As Byte
    Rotate As Single
    Grh As Grh

End Type


Public Type TLVERTEX2

    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    tu As Single
    tv As Single

End Type

'Blood list

Public Type BloodData

    V(0 To 5) As TLVERTEX2
    Life As Long
    TileX As Byte
    TileY As Byte

End Type

Public LastBlood As Long
Public BloodList() As BloodData
Public ProjectileList() As Projectile
Public LastProjectile As Integer  'Last projectile index used


Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEn ... ne_TPtoSPX" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
    Engine_TPtoSPX = X * 32 - ScreenMinX * 32 + OffsetCounterX - 16
End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEn ... ne_TPtoSPY" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
    Engine_TPtoSPY = Y * 32 - ScreenMinY * 32 + OffsetCounterY - 16

End Function

Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEn ... _PixelPosX" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************

    Engine_PixelPosX = (X - 1) * TilePixelWidth

End Function

Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEn ... _PixelPosY" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************

    Engine_PixelPosY = (Y - 1) * TilePixelWidth

End Function

Public Function Engine_SPtoTPX(ByVal X As Long) As Long
'************************************************************
'Screen Position to Tile Position
'Takes the screen pixel position and returns the tile position
'************************************************************
    Engine_SPtoTPX = UserPos.X + X \ TilePixelWidth - WindowTileWidth \ 2
End Function

Public Function Engine_SPtoTPY(ByVal Y As Long) As Long
'************************************************************
'Screen Position to Tile Position
'Takes the screen pixel position and returns the tile position
'************************************************************
    Engine_SPtoTPY = UserPos.Y + Y \ TilePixelHeight - WindowTileHeight \ 2
End Function

Sub Engine_Blood_Create(ByVal X As Single, ByVal Y As Single, ByVal size As Byte)

'*****************************************************************
'Creates a puddle of blood on the ground
'*****************************************************************

    Dim TileX As Integer
    Dim TileY As Integer

    Const texwidth As Single = 64
    Const TexHeight As Single = 64
    Const NumLarge As Long = 2
    Const NumMedium As Long = 12
    Const NumSmall As Long = 8
    Const Large1X As Single = 0
    Const Large1Y As Single = 0
    Const Large1W As Single = 32
    Const Large1H As Single = 16
    Const Large2X As Single = 0
    Const Large2Y As Single = 17
    Const Large2W As Single = 32
    Const Large2H As Single = 16
    Const Med1X As Single = 0
    Const Med1Y As Single = 34
    Const Med1W As Single = 14
    Const Med1H As Single = 6
    Const Med2X As Single = 0
    Const Med2Y As Single = 41
    Const Med2W As Single = 12
    Const Med2H As Single = 7
    Const Med3X As Single = 0
    Const Med3Y As Single = 49
    Const Med3W As Single = 11
    Const Med3H As Single = 8
    Const Med4X As Single = 15
    Const Med4Y As Single = 34
    Const Med4W As Single = 12
    Const Med4H As Single = 5
    Const Med5X As Single = 15
    Const Med5Y As Single = 40
    Const Med5W As Single = 8
    Const Med5H As Single = 9
    Const Med6X As Single = 12
    Const Med6Y As Single = 50
    Const Med6W As Single = 9
    Const Med6H As Single = 7
    Const Med7X As Single = 22
    Const Med7Y As Single = 50
    Const Med7W As Single = 10
    Const Med7H As Single = 7
    Const Med8X As Single = 33
    Const Med8Y As Single = 0
    Const Med8W As Single = 16
    Const Med8H As Single = 7
    Const Med9X As Single = 33
    Const Med9Y As Single = 8
    Const Med9W As Single = 14
    Const Med9H As Single = 7
    Const Med10X As Single = 33
    Const Med10Y As Single = 29
    Const Med10W As Single = 17
    Const Med10H As Single = 8
    Const Med11X As Single = 33
    Const Med11Y As Single = 38
    Const Med11W As Single = 15
    Const Med11H As Single = 9
    Const Med12X As Single = 33
    Const Med12Y As Single = 48
    Const Med12W As Single = 11
    Const Med12H As Single = 6
    Const Small1X As Single = 28
    Const Small1Y As Single = 34
    Const Small1W As Single = 4
    Const Small1H As Single = 6
    Const Small2X As Single = 24
    Const Small2Y As Single = 41
    Const Small2W As Single = 6
    Const Small2H As Single = 4
    Const Small3X As Single = 33
    Const Small3Y As Single = 16
    Const Small3W As Single = 10
    Const Small3H As Single = 3
    Const Small4X As Single = 44
    Const Small4Y As Single = 16
    Const Small4W As Single = 4
    Const Small4H As Single = 5
    Const Small5X As Single = 33
    Const Small5Y As Single = 20
    Const Small5W As Single = 8
    Const Small5H As Single = 4
    Const Small6X As Single = 42
    Const Small6Y As Single = 22
    Const Small6W As Single = 4
    Const Small6H As Single = 3
    Const Small7X As Single = 33
    Const Small7Y As Single = 25
    Const Small7W As Single = 8
    Const Small7H As Single = 3
    Const Small8X As Single = 42
    Const Small8Y As Single = 26
    Const Small8W As Single = 5
    Const Small8H As Single = 2

    Dim BloodIndex As Integer
    Dim i As Long
    Dim L As Long

    'Find the tile
    TileX = ((X - 288) \ 32) + 1
    TileY = ((Y - 288) \ 32) + 1

    If TileX < 1 Then TileX = 1
    If TileX > 100 Then TileX = 100
    If TileY < 1 Then TileY = 1
    If TileY > 100 Then TileY = 100

    'Check if there is too much blood on this tile already

    If MapData(TileX, TileY).Blood > 40 Then Exit Sub

    'Get the next open blood slot

    Do
        BloodIndex = BloodIndex + 1

        'Update LastBlood if we go over the size of the current array

        If BloodIndex > LastBlood Then
            LastBlood = BloodIndex
            ReDim Preserve BloodList(1 To LastBlood)

            Exit Do

        End If

    Loop While BloodList(BloodIndex).Life > 0

    'Set the blood's lfie
    BloodList(BloodIndex).Life = GetTickCount + 7000

    'Get a random size if none is specified

    If size < 1 Or size > 3 Then
        size = Int(Rnd * (NumLarge + NumSmall + NumMedium)) + 1

        If size <= NumLarge Then
            size = 3
        ElseIf size <= NumLarge + NumMedium Then
            size = 2
        Else
            size = 1
        End If
    End If

    With BloodList(BloodIndex)

        'Set up the general blood information

        For L = 0 To 5
            .V(L).Color = -1
            .V(L).rhw = 1
            .V(L).X = X
            .V(L).Y = Y

        Next L

        ' 3____4
        ' 0|\\ | 0 = 3
        ' | \\ | 1 = 5
        ' | \\ |
        ' | \\ |
        ' 2|____\\|
        ' 1 5
        'Large blood

        If size = 3 Then
            i = Int(Rnd * NumLarge) + 1

            Select Case i

                Case 1
                    .V(4).X = X + Large1W
                    .V(2).Y = Y + Large1H
                    .V(0).tu = Large1X / texwidth
                    .V(0).tv = Large1Y / TexHeight
                    .V(5).tu = (Large1X + Large1W) / texwidth
                    .V(5).tv = (Large1Y + Large1H) / TexHeight

                Case 2
                    .V(4).X = X + Large2W
                    .V(2).Y = Y + Large2H
                    .V(0).tu = Large2X / texwidth
                    .V(0).tv = Large2Y / TexHeight
                    .V(5).tu = (Large2X + Large2W) / texwidth
                    .V(5).tv = (Large2Y + Large2H) / TexHeight
            End Select

            'Medium blood
        ElseIf size = 2 Then
            i = Int(Rnd * NumMedium) + 1

            Select Case i

                Case 1
                    .V(4).X = X + Med1W
                    .V(2).Y = Y + Med1H
                    .V(0).tu = Med1X / texwidth
                    .V(0).tv = Med1Y / TexHeight
                    .V(5).tu = (Med1X + Med1W) / texwidth
                    .V(5).tv = (Med1Y + Med1H) / TexHeight

                Case 2
                    .V(4).X = X + Med2W
                    .V(2).Y = Y + Med2H
                    .V(0).tu = Med2X / texwidth
                    .V(0).tv = Med2Y / TexHeight
                    .V(5).tu = (Med2X + Med2W) / texwidth
                    .V(5).tv = (Med2Y + Med2H) / TexHeight

                Case 3
                    .V(4).X = X + Med3W
                    .V(2).Y = Y + Med3H
                    .V(0).tu = Med3X / texwidth
                    .V(0).tv = Med3Y / TexHeight
                    .V(5).tu = (Med3X + Med3W) / texwidth
                    .V(5).tv = (Med3Y + Med3H) / TexHeight

                Case 4
                    .V(4).X = X + Med4W
                    .V(2).Y = Y + Med4H
                    .V(0).tu = Med4X / texwidth
                    .V(0).tv = Med4Y / TexHeight
                    .V(5).tu = (Med4X + Med4W) / texwidth
                    .V(5).tv = (Med4Y + Med4H) / TexHeight

                Case 5
                    .V(4).X = X + Med5W
                    .V(2).Y = Y + Med5H
                    .V(0).tu = Med5X / texwidth
                    .V(0).tv = Med5Y / TexHeight
                    .V(5).tu = (Med5X + Med5W) / texwidth
                    .V(5).tv = (Med5Y + Med5H) / TexHeight

                Case 6
                    .V(4).X = X + Med6W
                    .V(2).Y = Y + Med6H
                    .V(0).tu = Med6X / texwidth
                    .V(0).tv = Med6Y / TexHeight
                    .V(5).tu = (Med6X + Med6W) / texwidth
                    .V(5).tv = (Med6Y + Med6H) / TexHeight

                Case 7
                    .V(4).X = X + Med7W
                    .V(2).Y = Y + Med7H
                    .V(0).tu = Med7X / texwidth
                    .V(0).tv = Med7Y / TexHeight
                    .V(5).tu = (Med7X + Med7W) / texwidth
                    .V(5).tv = (Med7Y + Med7H) / TexHeight

                Case 8
                    .V(4).X = X + Med8W
                    .V(2).Y = Y + Med8H
                    .V(0).tu = Med8X / texwidth
                    .V(0).tv = Med8Y / TexHeight
                    .V(5).tu = (Med8X + Med8W) / texwidth
                    .V(5).tv = (Med8Y + Med8H) / TexHeight

                Case 9
                    .V(4).X = X + Med9W
                    .V(2).Y = Y + Med9H
                    .V(0).tu = Med9X / texwidth
                    .V(0).tv = Med9Y / TexHeight
                    .V(5).tu = (Med9X + Med9W) / texwidth
                    .V(5).tv = (Med9Y + Med9H) / TexHeight

                Case 10
                    .V(4).X = X + Med10W
                    .V(2).Y = Y + Med10H
                    .V(0).tu = Med10X / texwidth
                    .V(0).tv = Med10Y / TexHeight
                    .V(5).tu = (Med10X + Med10W) / texwidth
                    .V(5).tv = (Med10Y + Med10H) / TexHeight

                Case 11
                    .V(4).X = X + Med11W
                    .V(2).Y = Y + Med11H
                    .V(0).tu = Med11X / texwidth
                    .V(0).tv = Med11Y / TexHeight
                    .V(5).tu = (Med11X + Med11W) / texwidth
                    .V(5).tv = (Med11Y + Med11H) / TexHeight

                Case 12
                    .V(4).X = X + Med12W
                    .V(2).Y = Y + Med12H
                    .V(0).tu = Med12X / texwidth
                    .V(0).tv = Med12Y / TexHeight
                    .V(5).tu = (Med12X + Med12W) / texwidth
                    .V(5).tv = (Med12Y + Med12H) / TexHeight
            End Select

            'Small blood
        Else
            i = Int(Rnd * NumSmall) + 1

            Select Case i

                Case 1
                    .V(4).X = X + Small1W
                    .V(2).Y = Y + Small1H
                    .V(0).tu = Small1X / texwidth
                    .V(0).tv = Small1Y / TexHeight
                    .V(5).tu = (Small1X + Small1W) / texwidth
                    .V(5).tv = (Small1Y + Small1H) / TexHeight

                Case 2
                    .V(4).X = X + Small2W
                    .V(2).Y = Y + Small2H
                    .V(0).tu = Small2X / texwidth
                    .V(0).tv = Small2Y / TexHeight
                    .V(5).tu = (Small2X + Small2W) / texwidth
                    .V(5).tv = (Small2Y + Small2H) / TexHeight

                Case 3
                    .V(4).X = X + Small3W
                    .V(2).Y = Y + Small3H
                    .V(0).tu = Small3X / texwidth
                    .V(0).tv = Small3Y / TexHeight
                    .V(5).tu = (Small3X + Small3W) / texwidth
                    .V(5).tv = (Small3Y + Small3H) / TexHeight

                Case 4
                    .V(4).X = X + Small4W
                    .V(2).Y = Y + Small4H
                    .V(0).tu = Small4X / texwidth
                    .V(0).tv = Small4Y / TexHeight
                    .V(5).tu = (Small4X + Small4W) / texwidth
                    .V(5).tv = (Small4Y + Small4H) / TexHeight

                Case 5
                    .V(4).X = X + Small5W
                    .V(2).Y = Y + Small5H
                    .V(0).tu = Small5X / texwidth
                    .V(0).tv = Small5Y / TexHeight
                    .V(5).tu = (Small5X + Small5W) / texwidth
                    .V(5).tv = (Small5Y + Small5H) / TexHeight

                Case 6
                    .V(4).X = X + Small6W
                    .V(2).Y = Y + Small6H
                    .V(0).tu = Small6X / texwidth
                    .V(0).tv = Small6Y / TexHeight
                    .V(5).tu = (Small6X + Small6W) / texwidth
                    .V(5).tv = (Small6Y + Small6H) / TexHeight

                Case 7
                    .V(4).X = X + Small7W
                    .V(2).Y = Y + Small7H
                    .V(0).tu = Small7X / texwidth
                    .V(0).tv = Small7Y / TexHeight
                    .V(5).tu = (Small7X + Small7W) / texwidth
                    .V(5).tv = (Small7Y + Small7H) / TexHeight

                Case 8
                    .V(4).X = X + Small8W
                    .V(2).Y = Y + Small8H
                    .V(0).tu = Small8X / texwidth
                    .V(0).tv = Small8Y / TexHeight
                    .V(5).tu = (Small8X + Small8W) / texwidth
                    .V(5).tv = (Small8Y + Small8H) / TexHeight
            End Select

        End If

        'These variables are the same no blood used
        .V(4).tu = .V(5).tu
        .V(4).tv = .V(0).tv
        .V(2).tu = .V(0).tu
        .V(2).tv = .V(5).tv
        .V(5).X = .V(4).X
        .V(5).Y = .V(2).Y
        .V(3) = .V(0)
        .V(1) = .V(5)
        'Find the blood tile location
        .TileX = TileX
        .TileY = TileY
        MapData(.TileX, .TileY).Blood = MapData(.TileX, .TileY).Blood + 1
    End With

End Sub

Sub Engine_Blood_Erase(ByVal BloodIndex As Long)

'*****************************************************************
'Erases a blood splatter by index
'*****************************************************************

    With BloodList(BloodIndex)
        'Set the life to 0 to not use it
        BloodList(BloodIndex).Life = 0

        'Erase the blood from the tile

        If .TileX > 0 Then
            If .TileY > 0 Then
                If .TileX <= 92 Then
                    If .TileY <= 92 Then
                        MapData(.TileX, .TileY).Blood = MapData(.TileX, .TileY).Blood - 1
                    End If
                End If
            End If
        End If

    End With

    'Resize the array if needed

    If BloodIndex = LastBlood Then

        Do Until BloodList(LastBlood).Life > 0
            LastBlood = LastBlood - 1

            If LastBlood = 0 Then Exit Do
        Loop

        If LastBlood <> BloodIndex Then
            If LastBlood <> 0 Then
                ReDim Preserve BloodList(1 To LastBlood)
            Else
                Erase BloodList
            End If
        End If
    End If

End Sub

Public Sub Engine_Render_Blood()

'*****************************************************************
'Batch render the blood on the ground
'*****************************************************************

    Dim BloodVB As Direct3DVertexBuffer8    'Vertex buffer
    Dim BloodVL() As TLVERTEX2   'Vertex list
    Dim BloodCount As Long
    Dim Alpha As Long
    Dim i As Long
    Dim j As Long
    Dim Tex As Direct3DTexture8
    Dim SRDesc As D3DSURFACE_DESC

    'Check if Render Blood option is enabled

    'If Config.renderBlood = 0 Then Exit Sub

    'Check for any blood

    If LastBlood = 0 Then Exit Sub

    Set Tex = D3DX.CreateTextureFromFileEx(d3ddevice, App.Path & "\RECURSOS\Graficos\Grh\25005.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
    Tex.GetLevelDesc 0, SRDesc

    'Create the vertex list
    ReDim BloodVL(1 To LastBlood * 6) As TLVERTEX2

    For i = 1 To LastBlood

        If BloodList(i).Life <> 0 Then
            If BloodList(i).Life > GetTickCount Then
                If BloodList(i).Life - GetTickCount > 3000 Then
                    Alpha = 255
                Else
                    Alpha = (BloodList(i).Life - GetTickCount) / 7

                    If Alpha > 255 Then Alpha = 255
                End If

                For j = 1 To 6
                    BloodVL((BloodCount * 6) + j) = BloodList(i).V(j - 1)

                    With BloodVL((BloodCount * 6) + j)
                        .X = .X - ParticleOffsetX
                        .Y = .Y - ParticleOffsetY
                        .Color = D3DColorARGB(Alpha, 255, 255, 255)
                    End With

                Next j

                BloodCount = BloodCount + 1
            Else
                Engine_Blood_Erase i
            End If
        End If

    Next i

    d3ddevice.SetTexture 0, Tex
    d3ddevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(Alpha, 0, 0, 0)

    'Check if any blood was found in use

    If BloodCount = 0 Then Exit Sub

    'Create the vertex buffer
    Set BloodVB = d3ddevice.CreateVertexBuffer(28 * BloodCount * 6, 0, FVF, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData BloodVB, 0, 28 * BloodCount * 6, 0, BloodVL(1)

    'Draw the blood
    d3ddevice.SetStreamSource 0, BloodVB, 28

    d3ddevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, BloodCount * 2

End Sub
Public Function Engine_GetAngle(ByVal CenterX As Integer, _
                                ByVal CenterY As Integer, _
                                ByVal TargetX As Integer, _
                                ByVal TargetY As Integer) As Single

'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetAngle
'************************************************************

    Dim SideA As Single
    Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)

    If CenterY = TargetY Then

        'Check for going right (90 degrees)

        If CenterX < TargetX Then
            Engine_GetAngle = 90
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function

        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)

    If CenterX = TargetX Then

        'Check for going up (360 degrees)

        If CenterY > TargetY Then
            Engine_GetAngle = 360
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function

        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)

    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360

    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    Exit Function

    'Check for error
ErrOut:
    'Return a 0 saying there was an error
    Engine_GetAngle = 0

    Exit Function

End Function

Public Sub Engine_Projectile_Create(ByVal AttackerIndex As Integer, _
                                    ByVal TargetIndex As Integer, _
                                    ByVal GrhIndex As Long, _
                                    ByVal Rotation As Byte)

'*****************************************************************
'Creates a projectile for a ranged weapon
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Projectile_Create
'*****************************************************************

    Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If TargetIndex = 0 Then Exit Sub
    If AttackerIndex > UBound(CharList) Then Exit Sub
    If TargetIndex > UBound(CharList) Then Exit Sub

    'Get the next open projectile slot

    Do
        ProjectileIndex = ProjectileIndex + 1

        'Update LastProjectile if we go over the size of the current array

        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)

            Exit Do

        End If

    Loop While ProjectileList(ProjectileIndex).Grh.GrhIndex > 0

    'Figure out the initial rotation value
    ProjectileList(ProjectileIndex).Rotate = Engine_GetAngle(CharList(AttackerIndex).Pos.X, CharList(AttackerIndex).Pos.Y, CharList(TargetIndex).Pos.X, CharList(TargetIndex).Pos.Y)

    'Fill in the values
    ProjectileList(ProjectileIndex).tX = CharList(TargetIndex).Pos.X * 32    '+ charlist(TargetIndex).MoveOffsetX
    ProjectileList(ProjectileIndex).tY = CharList(TargetIndex).Pos.Y * 32  '+ charlist(TargetIndex).MoveOffsetY
    ProjectileList(ProjectileIndex).RotateSpeed = Rotation
    ProjectileList(ProjectileIndex).X = CharList(AttackerIndex).Pos.X * 32    ' * 32 '+ charlist(AttackerIndex).MoveOffsetX
    ProjectileList(ProjectileIndex).Y = CharList(AttackerIndex).Pos.Y * 32 - 10    ' * 32 '+ charlist(AttackerIndex).MoveOffset

    InitGrh ProjectileList(ProjectileIndex).Grh, GrhIndex

End Sub

Public Sub Engine_Projectile_Erase(ByVal ProjectileIndex As Integer)

'*****************************************************************
'Erase a projectile by the projectile index
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Projectile_Erase
'*****************************************************************

'Clear the selected index
    ProjectileList(ProjectileIndex).Grh.FrameCounter = 0
    ProjectileList(ProjectileIndex).Grh.GrhIndex = 0
    ProjectileList(ProjectileIndex).X = 0
    ProjectileList(ProjectileIndex).Y = 0
    ProjectileList(ProjectileIndex).tX = 0
    ProjectileList(ProjectileIndex).tY = 0
    ProjectileList(ProjectileIndex).Rotate = 0
    ProjectileList(ProjectileIndex).RotateSpeed = 0

    'Update LastProjectile

    If ProjectileIndex = LastProjectile Then

        Do Until ProjectileList(ProjectileIndex).Grh.GrhIndex > 1
            'Move down one projectile
            LastProjectile = LastProjectile - 1

            If LastProjectile = 0 Then Exit Do
        Loop

        If ProjectileIndex <> LastProjectile Then

            'We still have projectiles, resize the array to end at the last used slot

            If LastProjectile > 0 Then
                ReDim Preserve ProjectileList(1 To LastProjectile)
            Else
                Erase ProjectileList
            End If

        End If

    End If

End Sub

Public Function Engine_UTOV_Particle(ByVal UserIndex As Integer, _
                                     ByVal VictimIndex As Integer, _
                                     ByVal Particle_ID As Integer) As Integer

    Dim X As Long
    Dim Y As Long
    Dim TempIndex As Integer
    Dim RetNum As Integer

    Select Case Particle_ID

        Case 1              'Teleport
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y) + 18
            TempIndex = Effect_Fire_Begin(X, Y, 1, 150, 180, 1)
            effect(RetNum).BindSpeed = 12
        Case 3          ' Tormenta de Fuego
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Torch_Begin(X, Y, 1, 150, 179, -5000)
            effect(RetNum).BindToChar = VictimIndex
            effect(RetNum).BindSpeed = 12
            effect(RetNum).KillWhenAtTarget = True
        Case 4          ' Curar heridas Graves
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Bless_Begin(X, Y, 3, 50, 16, 7)
            effect(RetNum).BindToChar = VictimIndex
            'Effect(RetNum).BindSpeed = 10
            effect(RetNum).KillWhenAtTarget = True
        Case 5          ' Misil Magico
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Misile_Begin(X, Y, 1, 16, 100)    ' 2, 100)
            effect(RetNum).BindToChar = VictimIndex
            effect(RetNum).BindSpeed = 12
            effect(RetNum).KillWhenAtTarget = True
        Case 6          '  Descarga Electrica
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Ray_Begin(X, Y, 2, 150, 179, -5000)
            effect(RetNum).BindToChar = VictimIndex
            effect(RetNum).BindSpeed = 10
            effect(RetNum).KillWhenAtTarget = True
        Case 7          '  Inmovilizar
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Curse_Begin(X, Y, 1, 300, 179, 200)
            effect(RetNum).BindToChar = VictimIndex
            effect(RetNum).BindSpeed = 8
            effect(RetNum).KillWhenAtTarget = True
        Case 8         '  Apocalipsis
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Necro_Begin(X, Y, 1, 300, 179, 200)
            effect(RetNum).BindToChar = VictimIndex
            effect(RetNum).BindSpeed = 10
            effect(RetNum).KillWhenAtTarget = True
        Case 9         '  Dardo Magico
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Spell_Begin(X, Y, 1, 50, 10, 3, 0.5, 0.5, 0.5, 179, 200)
            effect(RetNum).BindToChar = VictimIndex
            effect(RetNum).BindSpeed = 10
            effect(RetNum).KillWhenAtTarget = True
        Case 10         '  Fuerza
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Strengthen_Begin(X, Y, 12, 50, 16, 7)
            effect(RetNum).BindToChar = VictimIndex
            'Effect(RetNum).BindSpeed = 10
            'Effect(RetNum).KillWhenAtTarget = True
        Case 11        '  Celeridad
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Strengthen_Begin(X, Y, 11, 50, 16, 7, True)
            effect(RetNum).BindToChar = VictimIndex
            'Effect(RetNum).BindSpeed = 10
            'Effect(RetNum).KillWhenAtTarget = True
        Case 12         ' Flecha electrica
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Spell_Begin(X, Y, 2, 100, 12, 6, 0.5, 0.5, 1, 179, 200)
            effect(RetNum).BindToChar = VictimIndex
            effect(RetNum).BindSpeed = 10
            effect(RetNum).KillWhenAtTarget = True
        Case 13         ' Curar
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Bless_Begin(X, Y, 3, 5, 8, 7)
            ' Effect(RetNum).BindToChar = VictimIndex
            'Effect(RetNum).BindSpeed = 12
            effect(RetNum).KillWhenAtTarget = True
        Case 14         ' Resucitar
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Holy_Begin(X, Y, 3, 20, 179, 10)
            ' Effect(RetNum).BindToChar = VictimIndex
            'Effect(RetNum).BindSpeed = 12
            effect(RetNum).KillWhenAtTarget = True
        Case 15         ' Paralizar
            X = Engine_TPtoSPX(CharList(UserIndex).Pos.X)
            Y = Engine_TPtoSPY(CharList(UserIndex).Pos.Y)
            RetNum = Effect_Implode_Begin(X, Y, 1, 200, 1)
            effect(RetNum).BindToChar = VictimIndex
            effect(RetNum).BindSpeed = 10
            'Effect(RetNum).KillWhenAtTarget = True

    End Select

    Engine_UTOV_Particle = TempIndex
End Function

Public Sub Effect_Die_Connect()

    If Colorinicial <> 139 And UserMinHP <> 0 Then
        Colorinicial = Colorinicial + 1
        base_light = ARGB(Colorinicial, Colorinicial, Colorinicial, 255)
    ElseIf Colorinicial = 139 Then
        If YaPrendioLuces = False Then
            YaPrendioLuces = True
            SwitchMapNew UserMap, True
            Light.Light_Render_All
        End If
    End If

    If ColorMuerto <> 139 And UserMinHP = 0 Then
        If YaPrendioLuces = True Then    'apago las luces
            YaPrendioLuces = False
            Light.Light_Remove_All
        End If
        
        If Colorinicial > 40 Then
                Colorinicial = Colorinicial - 1
        ElseIf Colorinicial < 39 Then
                Colorinicial = Colorinicial + 1
        End If
        
        ColorMuerto = ColorMuerto + 1
        base_light = ARGB(ColorMuerto, Colorinicial, Colorinicial, 255)
        'Light.Light_Reset_Color ColorMuerto, Colorinicial, True

    ElseIf UserMinHP <> 0 And ColorMuerto = 139 Then    'esta vivo
        ColorMuerto = 1
        Colorinicial = 40    ' me garantizo que se resucito
    End If
    
End Sub

Sub Engine_Weather_Update()

' / Author: Emanuel Matías (Dunkan)
' / Note: Actualiza el clima (Nieve, lluvia, niebla, lluvia + niebla)

    Select Case LastWeather

        Case 1  '// Lluvia
            If WeatherEffectIndex <= 0 Then
                WeatherEffectIndex = Effect_Rain_Begin(9, Opciones.bGraphics)
            ElseIf effect(WeatherEffectIndex).EffectNum <> EffectNum_Rain Then
                Effect_Kill WeatherEffectIndex
                WeatherEffectIndex = Effect_Rain_Begin(9, Opciones.bGraphics)
            ElseIf Not effect(WeatherEffectIndex).Used Then
                WeatherEffectIndex = Effect_Rain_Begin(9, Opciones.bGraphics)
            End If
            WeatherDoFog = 0

        Case 2  '// Nieve
            If WeatherEffectIndex <= 0 Then
                WeatherEffectIndex = Effect_Snow_Begin(15, Opciones.bGraphics)
            ElseIf effect(WeatherEffectIndex).EffectNum <> EffectNum_Snow Then
                Effect_Kill WeatherEffectIndex
                WeatherEffectIndex = Effect_Snow_Begin(15, Opciones.bGraphics)
            ElseIf Not effect(WeatherEffectIndex).Used Then
                WeatherEffectIndex = Effect_Snow_Begin(15, Opciones.bGraphics)
            End If
            WeatherDoFog = 0

        Case 3  '// Niebla
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
            End If
            WeatherEffectIndex = 0
            WeatherDoFog = 10

        Case 4    '// Niebla + Lluvia
            If WeatherEffectIndex <= 0 Then
                WeatherEffectIndex = Effect_Rain_Begin(9, Opciones.bGraphics)
            ElseIf effect(WeatherEffectIndex).EffectNum <> EffectNum_Rain Then
                Effect_Kill WeatherEffectIndex
                WeatherEffectIndex = Effect_Rain_Begin(9, Opciones.bGraphics)
            ElseIf Not effect(WeatherEffectIndex).Used Then
                WeatherEffectIndex = Effect_Rain_Begin(9, Opciones.bGraphics)
            End If
            WeatherDoFog = 2

        Case Else   'None
            If WeatherEffectIndex > 0 Then  'Kill the weather effect if used
                If effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
            End If
            WeatherEffectIndex = 0
            WeatherDoFog = 0

    End Select


    'Update fog
    If WeatherDoFog Then Engine_Weather_UpdateFog

    'If WeatherDoLightning Then Engine_Weather_UpdateLightning

End Sub

