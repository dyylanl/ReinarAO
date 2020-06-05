Attribute VB_Name = "modCarga"
'Parra: Modulo donde se encuentran todos los subs para cargar los recursos & info adicional

Option Explicit


Private Numheads As Integer
Private NumFxs As Integer
Private NumWeaponAnims As Integer
Sub CargarCabezas()
    On Error Resume Next
    Dim n As Integer, i As Integer, Numheads As Integer

    Dim Miscabezas() As tIndiceCabeza

    n = FreeFile
    Open App.Path & "\RECURSOS\init\Cabezas.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , Numheads


    ReDim HeadData(0 To Numheads + 1) As HeadData
    ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
        InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
        InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
        InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
    Next i

    Close #n

End Sub

Sub CargarCascos()
    On Error Resume Next
    Dim n As Integer, i As Integer, NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza

    n = FreeFile
    Open App.Path & "\RECURSOS\init\Cascos.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , NumCascos


    ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
    ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

    For i = 1 To NumCascos
        Get #n, , Miscabezas(i)
        InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
        InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
        InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
        InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
    Next i

    Close #n

End Sub

Sub CargarCuerpos()
    On Error Resume Next
    Dim n As Integer, i As Integer
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo

    n = FreeFile
    Open App.Path & "\RECURSOS\init\Personajes.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , NumCuerpos


    ReDim BodyData(0 To NumCuerpos + 1) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
        InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
        InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
        InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
        BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
        BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
    Next i

    Close #n

End Sub
Sub CargarFxs()
    On Error Resume Next
    Dim n As Integer, i As Integer
    Dim NumFxs As Integer
    Dim MisFxs() As tIndiceFx

    n = FreeFile
    Open App.Path & "\RECURSOS\init\Fxs.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , NumFxs


    ReDim FxData(0 To NumFxs + 1) As FxData
    ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

    For i = 1 To NumFxs
        Get #n, , MisFxs(i)
        Call InitGrh(FxData(i).FX, MisFxs(i).Animacion, 1)
        FxData(i).OffSetX = MisFxs(i).OffSetX
        FxData(i).OffSetY = MisFxs(i).OffSetY
    Next i

    Close #n

End Sub
Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
    On Error Resume Next
    If GrhIndex = 0 Then Exit Sub
    Grh.GrhIndex = GrhIndex

    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        Grh.Started = Started
    End If

    Grh.FrameCounter = 1

    If Grh.GrhIndex <> 0 Then
        'If Grh.GrhIndex = 4719 Or Grh.GrhIndex = 10296 Or Grh.GrhIndex = 10294 Or Grh.GrhIndex = 10297 Or Grh.GrhIndex = 10295 Or Grh.GrhIndex = 10270 Or Grh.GrhIndex = 10268 Or Grh.GrhIndex = 10271 Or Grh.GrhIndex = 10269 Or Grh.GrhIndex = 4723 Or Grh.GrhIndex = 4724 Or Grh.GrhIndex = 4722 Or Grh.GrhIndex = 4721 Or Grh.GrhIndex = 4718 Or Grh.GrhIndex = 4720 Or Grh.GrhIndex = 4723 Or Grh.GrhIndex = 4725 Then
        '    Grh.SpeedCounter = 0
        'Else
            Grh.SpeedCounter = GrhData(Grh.GrhIndex).speed
        'End If
    End If
End Sub
Sub LoadGrhData()
    On Error GoTo ErrorHandler

    Dim Grh As Integer
    Dim Frame As Integer
    Dim tempint As Integer


    ReDim GrhData(1 To 32000) As GrhData

    Open IniPath & "Graficos.ind" For Binary Access Read As #1
    Seek #1, 1

    Get #1, , MiCabecera
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint

    Get #1, , Grh

    Do Until Grh <= 0


        Get #1, , GrhData(Grh).NumFrames
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler

        If GrhData(Grh).NumFrames > 1 Then


            For Frame = 1 To GrhData(Grh).NumFrames

                Get #1, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > 32000 Then
                    GoTo ErrorHandler
                End If

            Next Frame
            
            Dim a As Integer
            Get #1, , a
            GrhData(Grh).speed = a
       
            ñoñal Grh
            
            If GrhData(Grh).speed <= 0 Then GoTo ErrorHandler
            

            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler

            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler

        Else


            Get #1, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler

            Get #1, , GrhData(Grh).sX
            If GrhData(Grh).sX < 0 Then GoTo ErrorHandler

            Get #1, , GrhData(Grh).sY
            If GrhData(Grh).sY < 0 Then GoTo ErrorHandler

            Get #1, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            Get #1, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler


            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth

            GrhData(Grh).Frames(1) = Grh

        End If


        Get #1, , Grh

    Loop


    Close #1

    Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub
Sub ñoñal(Grh As Integer)
 
GrhData(Grh).speed = ((GrhData(Grh).speed * 1000) / 18)
 
End Sub
Sub CrearGrh(GrhIndex As Integer, Index As Integer)
    ReDim Preserve Grh(1 To Index) As Grh
    Grh(Index).FrameCounter = 1
    Grh(Index).GrhIndex = GrhIndex
    'Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
    Grh(Index).Started = 1
End Sub
Sub CargarAnimsExtra()
    Call CrearGrh(6580, 1)
    Call CrearGrh(534, 2)
End Sub
Sub CargarAnimArmas()

    On Error Resume Next

    Dim loopc As Integer
    Dim arch As String
    arch = App.Path & "\RECURSOS\init\" & "armas.dat"
    DoEvents

    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))

    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc

End Sub
Sub CargarAnimEscudos()
    On Error Resume Next

    Dim loopc As Integer
    Dim arch As String
    arch = App.Path & "\RECURSOS\init\" & "escudos.dat"
    DoEvents

    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))

    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData

    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc

End Sub

Sub SwitchMapNew(map As Integer, crealuz As Boolean)
    Dim Y As Integer
    Dim X As Integer
    Dim tempint As Integer
    Dim infotile As Byte
    Dim i As Integer

    Particle_Group_Remove_All
    Light.Light_Remove_All

    
    Open App.Path & "\RECURSOS\MAPS\Mapa" & map & ".mcl" For Binary As #1
    Seek #1, 1

    Get #1, , tempint

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            'MapData(X, Y).ObjGrh.GrhIndex = 0
            Get #1, , infotile

            MapData(X, Y).Blocked = (infotile And 1)    ' osea lo que está haciendo acá es diciendo que si infotile vale 1 hay un bloqueo en ese tile

            Get #1, , MapData(X, Y).Graphic(1).GrhIndex
            If crealuz = False Then
                InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            End If


            For i = 2 To 4
                If infotile And (2 ^ (i - 1)) Then
                    Get #1, , MapData(X, Y).Graphic(i).GrhIndex
                    If crealuz = False Then
                        InitGrh MapData(X, Y).Graphic(i), MapData(X, Y).Graphic(i).GrhIndex
                    End If
                Else
                    If crealuz = False Then
                        MapData(X, Y).Graphic(i).GrhIndex = 0
                    End If
                End If
            Next

            If crealuz = False Then
                MapData(X, Y).Trigger = 0
            End If
            
            If (infotile And 16) Then _
               Get #1, , MapData(X, Y).Trigger

            If (infotile And 32) Then
                'Dim particula As Integer
                Get #1, , MapData(X, Y).parti_index
                'If crealuz = False Then
                MapData(X, Y).particle_group_index = General_Particle_Create(MapData(X, Y).parti_index, X, Y, -1)
                'End If
            End If

                If (infotile And 64) Then
                    Dim TempLNG As Long
                    Dim TempByte1 As Byte
                    Dim TempByte2 As Byte
                    Dim TempByte3 As Byte
                    Get #1, , TempLNG
                    Get #1, , TempByte1
                    Get #1, , TempByte2
                    Get #1, , TempByte3
                    Call Light.Light_Create(X, Y, TempLNG, , TempByte1, TempByte2, TempByte3)
                End If
                
            Dim aux As Integer
            If (infotile And 128) Then
                
                If crealuz = False Then
                    Get #1, , MapData(X, Y).ObjGrh.GrhIndex
                    InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
                Else
                    Get #1, , aux
                    'MapData(X, Y).ObjGrh.GrhIndex = 0
                End If
            End If
            
            
            If crealuz = False Then
                If MapData(X, Y).CharIndex > 0 Then Call EraseChar(MapData(X, Y).CharIndex)
            End If
        Next X
    Next Y

    Close #1

    CurMap = map


    Call DibujarPuntoMinimap
    Call DibujarMinimap

    If map = 205 Then
        Nombres = False
    Else
        Nombres = True
    End If

    'niebla
    If map = 214 Or map = 216 Then
        Niebla = True
    Else
        Niebla = False
    End If

    'fin niebla
    'Cargar_Luces_Particulas map
    If YaPrendioLuces Then
        Light.Light_Render_All
    End If

End Sub
