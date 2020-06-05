Attribute VB_Name = "modGeneralCharFunctions"

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer, ByVal Alas As Integer, ByVal EsPremium As Byte)

    If CharIndex > LastChar Then LastChar = CharIndex

    If arma = 0 Then arma = 2
    If escudo = 0 Then escudo = 2
    If casco = 0 Then casco = 2

    CharList(CharIndex).Head = HeadData(Head)

    CharList(CharIndex).Body = BodyData(Body)

    If Body > 83 And Body < 88 Then
        CharList(CharIndex).Navegando = 1
    Else: CharList(CharIndex).Navegando = 0
    End If

    CharList(CharIndex).arma = WeaponAnimData(arma)

    CharList(CharIndex).escudo = ShieldAnimData(escudo)
    CharList(CharIndex).casco = CascoAnimData(casco)
    CharList(CharIndex).Alas = BodyData(Alas)
    CharList(CharIndex).EsPremium = EsPremium

    CharList(CharIndex).Heading = Heading
    CharList(CharIndex).Moving = 0
    CharList(CharIndex).MoveOffset.X = 0
    CharList(CharIndex).MoveOffset.Y = 0


    CharList(CharIndex).Pos.X = X
    CharList(CharIndex).Pos.Y = Y


    CharList(CharIndex).active = 1


    MapData(X, Y).CharIndex = CharIndex
    If CharList(CharIndex).ParticleIndex <> 0 Then
        effect(CharList(CharIndex).ParticleIndex).Progression = 1
        CharList(CharIndex).ParticleIndex = 0
    End If
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    CharList(CharIndex).active = 0
    CharList(CharIndex).Criminal = 0
    CharList(CharIndex).FX = 0
    CharList(CharIndex).FxLoopTimes = 0
    CharList(CharIndex).invisible = False
    CharList(CharIndex).Moving = 0
    CharList(CharIndex).muerto = False
    CharList(CharIndex).Nombre = ""
    CharList(CharIndex).pie = False
    CharList(CharIndex).Pos.X = 0
    CharList(CharIndex).Pos.Y = 0
    CharList(CharIndex).UsandoArma = False
    If CharList(CharIndex).ParticleIndex <> 0 Then
        effect(CharList(CharIndex).ParticleIndex).Progression = 1
        CharList(CharIndex).ParticleIndex = 0
    End If
End Sub
Function NextOpenChar()
    Dim loopc As Integer

    loopc = 1

    Do While CharList(loopc).active
        loopc = loopc + 1
    Loop

    NextOpenChar = loopc

End Function
Sub EraseChar(ByVal CharIndex As Integer)
    On Error Resume Next

    CharList(CharIndex).active = 0


    If CharIndex = LastChar Then
        Do Until CharList(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If




    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

    Call ResetCharInfo(CharIndex)

End Sub
Sub MoveCharByHead(CharIndex As Integer, nheading As Byte)

    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y


    Select Case nheading

        Case NORTH
            addy = -1

        Case EAST
            addx = 1

        Case SOUTH
            addy = 1

        Case WEST
            addx = -1

    End Select

    nX = X + addx
    nY = Y + addy

    MapData(nX, nY).CharIndex = CharIndex
    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY
    MapData(X, Y).CharIndex = 0

    CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addx)
    CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addy)

    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nheading

    CharList(CharIndex).scrollDirectionX = addx
    CharList(CharIndex).scrollDirectionY = addy

    If UserEstado <> 1 Then Call DoPasosFx(CharIndex)


End Sub


Function EstaPCarea(ByVal Index2 As Integer) As Boolean

    Dim X As Integer, Y As Integer

    For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
        For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1

            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If

        Next X
    Next Y

    EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
    Static TickCount(200) As Integer
    If Cont = 999 Then Exit Function
    TickCount(Cual) = TickCount(Cual) + 1
    If TickCount(Cual) < Cont Then
        TickON = False
    Else
        TickCount(Cual) = 0
        TickON = True
    End If
End Function
Sub MoveCharByPosAndHead(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)

    On Error Resume Next

    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer



    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y

    MapData(X, Y).CharIndex = 0

    addx = nX - X
    addy = nY - Y




    MapData(nX, nY).CharIndex = CharIndex


    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY

    CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addx)
    CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addy)

    CharList(CharIndex).scrollDirectionX = Sgn(addx)
    CharList(CharIndex).scrollDirectionY = Sgn(addy)

    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nheading


End Sub
Sub MoveCharByPos(CharIndex As Integer, nX As Integer, nY As Integer)
    On Error Resume Next

    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nheading As Byte

    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y

    MapData(X, Y).CharIndex = 0

    addx = nX - X
    addy = nY - Y


    If Sgn(addx) = 1 Then nheading = EAST
    If Sgn(addx) = -1 Then nheading = WEST
    If Sgn(addy) = -1 Then nheading = NORTH
    If Sgn(addy) = 1 Then nheading = SOUTH

    MapData(nX, nY).CharIndex = CharIndex

    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY

    CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addx)
    CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addy)

    CharList(CharIndex).scrollDirectionX = Sgn(addx)
    CharList(CharIndex).scrollDirectionY = Sgn(addy)

    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nheading

End Sub
Sub MoveCharByPosConHeading(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
    On Error Resume Next

    If InMapBounds(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y) Then MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

    MapData(nX, nY).CharIndex = CharIndex

    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY

    CharList(CharIndex).Moving = 0
    CharList(CharIndex).MoveOffset.X = 0
    CharList(CharIndex).MoveOffset.Y = 0

    CharList(CharIndex).Heading = nheading

End Sub
Sub MoveMe(Direction As Byte)

    If CONGELADO Then Exit Sub
    If Stoppeado = True Then Exit Sub
    If Cartel Then Cartel = False

    If ProxLegalPos(Direction) And Not UserMeditar And Not UserParalizado Then
        If TiempoTranscurrido(LastPaso) >= IntervaloPaso Then
            Call SendData("M" & Direction)

            LastPaso = Timer
            If Not UserDescansar Then
                Call EliminarChars(Direction)
                Call MoveCharByHead(UserCharIndex, Direction)
                Call MoveScreen(Direction)
                Call DoFogataFx
            End If
        End If
    ElseIf CharList(UserCharIndex).Heading <> Direction Then Call SendData("CHEA" & Direction)
    End If

    Call DibujarPuntoMinimap
    frmMain.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.X & " - " & UserPos.Y & "]"

End Sub
Function ProxLegalPos(Direction As Byte) As Boolean

    Select Case Direction
        Case NORTH
            ProxLegalPos = LegalPos(UserPos.X, UserPos.Y - 1)
        Case SOUTH
            ProxLegalPos = LegalPos(UserPos.X, UserPos.Y + 1)
        Case WEST
            ProxLegalPos = LegalPos(UserPos.X - 1, UserPos.Y)
        Case EAST
            ProxLegalPos = LegalPos(UserPos.X + 1, UserPos.Y)
    End Select

End Function

Sub MoveScreen(Heading As Byte)

    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer

    Select Case Heading

        Case NORTH
            Y = -1

        Case EAST
            X = 1

        Case SOUTH
            Y = 1

        Case WEST
            X = -1

    End Select


    tX = UserPos.X + X
    tY = UserPos.Y + Y


    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1

        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                     MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                     MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                     MapData(UserPos.X, UserPos.Y).Trigger = 5 Or _
                     MapData(UserPos.X, UserPos.Y).Trigger = 6 Or _
                     MapData(UserPos.X, UserPos.Y).Trigger = 7, True, False)

    End If



End Sub
Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function
Sub RefreshAllChars()
    Dim loopc As Integer

    For loopc = 1 To LastChar
        If CharList(loopc).active = 1 Then

        End If
    Next loopc

End Sub
Function LegalPos(X As Integer, Y As Integer) As Boolean

    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        LegalPos = False
        Exit Function
    End If

    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If

    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If

    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If

    LegalPos = True

End Function
Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function
Function HayAgua(X As Integer, Y As Integer) As Boolean

    If MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
       MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
       MapData(X, Y).Graphic(2).GrhIndex = 0 Then
        HayAgua = True
    Else
        HayAgua = False
    End If

End Function
Sub EliminarChars(Direction As Byte)
    Dim X(2) As Integer
    Dim Y(2) As Integer

    Select Case Direction
        Case NORTH, SOUTH
            X(1) = UserPos.X - MinXBorder - 2
            X(2) = UserPos.X + MinXBorder + 2
        Case EAST, WEST
            Y(1) = UserPos.Y - MinYBorder - 2
            Y(2) = UserPos.Y + MinYBorder + 2
    End Select

    Select Case Direction
        Case NORTH
            Y(1) = UserPos.Y - MinYBorder - 3
            If Y(1) < 1 Then Y(1) = 1
            Y(2) = Y(1)
        Case EAST
            X(1) = UserPos.X + MinXBorder + 3
            If X(1) > 99 Then X(1) = 99
            X(2) = X(1)
        Case SOUTH
            Y(1) = UserPos.Y + MinYBorder + 3
            If Y(1) > 99 Then Y(1) = 99
            Y(2) = Y(1)
        Case WEST
            X(1) = UserPos.X - MinXBorder - 3
            If X(1) < 1 Then X(1) = 1
            X(2) = X(1)
    End Select

    For Y(0) = Y(1) To Y(2)
        For X(0) = X(1) To X(2)
            If X(0) > 6 And X(0) < 95 And Y(0) > 6 And Y(0) < 95 Then
                If MapData(X(0), Y(0)).CharIndex > 0 Then
                    CharList(MapData(X(0), Y(0)).CharIndex).Pos.X = 0
                    CharList(MapData(X(0), Y(0)).CharIndex).Pos.Y = 0
                    MapData(X(0), Y(0)).CharIndex = 0
                End If
            End If
        Next
    Next

End Sub
Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub
Sub DoPasosFx(ByVal CharIndex As Integer)
Static TerrenoDePaso As TipoPaso
 
    With CharList(CharIndex)
        If Not UserNavegando Then
            If Not .muerto And EstaPCarea(CharIndex) Then 'And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
             
                    If Not Char_Big_Get(CharIndex) Then
                        TerrenoDePaso = GetTerrenoDePaso(.Pos.X, .Pos.Y)
                    Else
                        TerrenoDePaso = CONST_PESADO
                    End If
             
                    If .pie = 0 Then
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(1), , Sound.Calculate_Volume(.Pos.X, .Pos.Y), Sound.Calculate_Pan(.Pos.X, .Pos.Y))
                    Else
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(2), , Sound.Calculate_Volume(.Pos.X, .Pos.Y), Sound.Calculate_Pan(.Pos.X, .Pos.Y))
                    End If
            End If
        Else
    ' TODO : Actually we would have to check if the CharIndex char is in the water or not....
            Call Sound.Sound_Play(SND_NAVEGANDO)
        End If
    End With
End Sub
 
Private Function GetTerrenoDePaso(ByVal X As Byte, ByVal Y As Byte) As TipoPaso
    With MapData(X, Y).Graphic(1)
        If .GrhIndex >= 6000 And .GrhIndex <= 6307 Then
            GetTerrenoDePaso = CONST_BOSQUE
            Exit Function
        ElseIf .GrhIndex >= 7501 And .GrhIndex <= 7507 Or .GrhIndex >= 7508 And .GrhIndex <= 2508 Then
            GetTerrenoDePaso = CONST_DUNGEON
            Exit Function
        'ElseIf (TerrainFileNum >= 5000 And TerrainFileNum <= 5004) Then
        '    GetTerrenoDePaso = CONST_NIEVE
        '    Exit Function
        Else
            GetTerrenoDePaso = CONST_PISO
        End If
    End With
End Function
 Private Function Char_Check(ByVal char_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (CharList(char_index).Heading > 0)
    End If
    
End Function
Public Function Char_Big_Get(ByVal CharIndex As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
   On Error GoTo ErrorHandler
 
 
   'Make sure it's a legal char_index
    If Char_Check(CharIndex) Then
        Char_Big_Get = (GrhData(CharList(CharIndex).Body.Walk(CharList(CharIndex).Heading).GrhIndex).TileWidth > 4)
    End If
 
    Exit Function
 
ErrorHandler:
 
End Function
