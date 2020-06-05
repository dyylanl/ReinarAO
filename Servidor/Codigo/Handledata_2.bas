Attribute VB_Name = "Handledata_2"
Public Sub HandleData2(UserIndex As Integer, rdata As String, Procesado As Boolean)
    Dim i As Integer, TIndex As Integer, N As Integer, X As Integer, Y As Integer, tInt As Integer, LoopCla As Integer
    Dim nPos As WorldPos
    Dim tStr As String
    Dim tLong As Long

    Procesado = True

    Select Case Left$(UCase$(rdata), 2)
        Case "#\"
            rdata = Right$(rdata, Len(rdata) - 2)
            TIndex = NameIndex(rdata)
            If TIndex Then
                If UserList(TIndex).flags.Privilegios < 2 Then
                    Call SendData(ToIndex, TIndex, 0, "||Fuiste congelado. No podras moverte ni salir del juego." & FONTTYPE_INFO)
                    Call SendData(ToIndex, TIndex, 0, "XX")
                    UserList(TIndex).flags.Congelado = 1

                Else

                    Call SendData(ToIndex, TIndex, 0, "||No tienes los privilegios necesarios." & FONTTYPE_INFO)

                End If
            Else
                Call SendData(ToIndex, TIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)

            End If
            Exit Sub

        Case "#º"
            rdata = Right$(rdata, Len(rdata) - 2)
            TIndex = NameIndex(rdata)

            If TIndex Then
                If UserList(TIndex).flags.Privilegios < 2 Then
                    If UserList(TIndex).flags.Congelado = 1 Then

                        Call SendData(ToIndex, TIndex, 0, "||Fuiste Descongelado. Podras moverte y salir del juego." & FONTTYPE_INFO)
                        Call SendData(ToIndex, TIndex, 0, "XX")
                        UserList(TIndex).flags.Congelado = 0

                    Else

                        Call SendData(ToIndex, TIndex, 0, "||No tienes los privilegios necesarios." & FONTTYPE_INFO)

                    End If
                Else
                    Call SendData(ToIndex, TIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)

                End If
            End If
            Exit Sub
        Case "#*"
            rdata = Right$(rdata, Len(rdata) - 2)
            TIndex = NameIndex(rdata)
            If TIndex Then
                If UserList(TIndex).flags.Privilegios < 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "||El jugador " & UserList(TIndex).name & " se encuentra online." & FONTTYPE_INFO)
                Else: Call SendData(ToIndex, UserIndex, 0, "1A")
                End If
            Else: Call SendData(ToIndex, UserIndex, 0, "1A")
            End If
            Exit Sub
        Case "#]"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call TirarRuleta(UserIndex, rdata)

            Exit Sub


        Case "^A"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).name & ": " & rdata & FONTTYPE_FIGHT)
            Exit Sub

        Case "#$"
            rdata = Right$(rdata, Len(rdata) - 2)
            If UserList(UserIndex).flags.Privilegios < 2 Then Exit Sub
            X = ReadField(1, rdata, 44)
            Y = ReadField(2, rdata, 44)
            N = MapaPorUbicacion(X, Y)
            If N Then Call WarpUserChar(UserIndex, N, 50, 50, True)
            Call LogGM(UserList(UserIndex).name, "Se transporto por mapa a Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
            Exit Sub


        Case "#:"
            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If
            If Not UserList(UserIndex).flags.Meditando And UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then Exit Sub
            Call SendData(ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
                Call SendData(ToIndex, UserIndex, 0, "7M")
            Else
                Call SendData(ToIndex, UserIndex, 0, "D9")
            End If
            UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando

            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = Timer
                Call SendData(ToIndex, UserIndex, 0, "8M" & TIEMPO_INICIOMEDITAR)


                UserList(UserIndex).Char.loops = LoopAdEternum

                If UserList(UserIndex).Stats.ELV < 15 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & 0 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARCHICO
                ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & "," & FXMEDITARMEDIANO & "," & 0 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV <= 49 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & 0 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARGRANDE
                ElseIf UserList(UserIndex).Stats.ELV < 100 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGIGANTE & "," & 0 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARGIGANTE
                ElseIf UserList(UserIndex).Stats.ELV >= 100 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFM" & UserList(UserIndex).Char.CharIndex & "," & 11 & "," & 1)
                    UserList(UserIndex).Char.FX = 0
                End If
            Else
                UserList(UserIndex).Char.FX = 0
                UserList(UserIndex).Char.loops = 0
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFM" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0 & "," & 0)
            End If
            Exit Sub

        Case "#A"
            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If
            If Not UserList(UserIndex).flags.Meditando And UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then Exit Sub
            Call SendData(ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
                Call SendData(ToIndex, UserIndex, 0, "7M")
            Else
                Call SendData(ToIndex, UserIndex, 0, "D9")
            End If
            UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando

            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = Timer
                Call SendData(ToIndex, UserIndex, 0, "8M" & TIEMPO_INICIOMEDITAR)


                UserList(UserIndex).Char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < 15 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & 0 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARCHICO
                ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & 0 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV <= 49 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & 0 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARGRANDE
                ElseIf UserList(UserIndex).Stats.ELV < 100 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGIGANTE & "," & 0 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARGIGANTE
                ElseIf UserList(UserIndex).Stats.ELV >= 100 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFM" & UserList(UserIndex).Char.CharIndex & "," & 11 & "," & 1)
                    UserList(UserIndex).Char.FX = 0
                End If
            Else
                UserList(UserIndex).Char.FX = 0
                UserList(UserIndex).Char.loops = 0
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFM" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0 & "," & 0)
            End If
            Exit Sub

        Case "#B"
            If UserList(UserIndex).flags.Congelado Then Exit Sub
            If UserList(UserIndex).pos.Map = 22 Or UserList(UserIndex).pos.Map = 205 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 19 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 7 Or UserList(UserIndex).pos.Map = 79 Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes salir de este mapa, si deseas salir, pideles a los gm's via /SOPORTE " & FONTTYPE_INFO)    '  Brilcair!
                Exit Sub
            End If
            If UserList(UserIndex).flags.Paralizado Then Exit Sub

            If (Not MapInfo(UserList(UserIndex).pos.Map).Pk And TiempoTranscurrido(UserList(UserIndex).Counters.LastRobo) > 10) Or UserList(UserIndex).flags.Privilegios > 1 Then
                Call SendData(ToIndex, UserIndex, 0, "FINOK")
                Call CloseSocket(UserIndex)
                Exit Sub
            End If

            Call Cerrar_Usuario(UserIndex)

            Exit Sub

        Case "#C"
            If CanCreateGuild(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SHOWFUN" & UserList(UserIndex).Faccion.Bando)
            Exit Sub

        Case "#D"
            Call SendData(ToIndex, UserIndex, 0, "7L")
            Exit Sub



        Case "#E"
            Call SendData(ToIndex, UserIndex, 0, "7L")
            Exit Sub

        Case "#F"
            Call SendData(ToIndex, UserIndex, 0, "7L")
            Exit Sub


        Case "#G"

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If

            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
               Or UserList(UserIndex).flags.Muerto Then Exit Sub

            Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
            Exit Sub
        Case "#H"

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If

            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> _
               UserIndex Then Exit Sub
            Npclist(UserList(UserIndex).flags.TargetNpc).Movement = ESTATICO
            Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
            Exit Sub
        Case "#I"

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If

            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> _
               UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNpc)
            Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
            Exit Sub
        Case "#J"

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If

            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNpc)
            Exit Sub
        Case "#K"
            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If
            If HayOBJarea(UserList(UserIndex).pos, FOGATA) Then
                Call SendData(ToIndex, UserIndex, 0, "DOK")
                If Not UserList(UserIndex).flags.Descansar Then
                    Call SendData(ToIndex, UserIndex, 0, "3M")
                Else
                    Call SendData(ToIndex, UserIndex, 0, "4M")
                End If
                UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                If UserList(UserIndex).flags.Descansar Then
                    Call SendData(ToIndex, UserIndex, 0, "4M")

                    UserList(UserIndex).flags.Descansar = False
                    Call SendData(ToIndex, UserIndex, 0, "DOK")
                    Exit Sub
                End If
                Call SendData(ToIndex, UserIndex, 0, "6M")
            End If
            Exit Sub

        Case "#L"

            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_REVIVIR _
               Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If

            Call RevivirUsuarioNPC(UserIndex)
            Call SendData(ToIndex, UserIndex, 0, "RZ")
            Exit Sub
        Case "#M"

            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_REVIVIR _
               Or UserList(UserIndex).flags.Muerto Then Exit Sub
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
            Call SendUserHP(UserIndex)
            Exit Sub
        Case "#N"
            If UserList(UserIndex).flags.Muerto Then Exit Sub
            Call EnviarSubclase(UserIndex)
            Exit Sub
        Case "#O"
            If PuedeRecompensa(UserIndex) And Not UserList(UserIndex).flags.Muerto Then _
               Call SendData(ToIndex, UserIndex, 0, "RELON" & UserList(UserIndex).Clase & "," & PuedeRecompensa(UserIndex))
            Exit Sub
        Case "#P"
            If UserList(UserIndex).flags.Privilegios > 0 Then
                For LoopC = 1 To LastUser
                    If Len(UserList(LoopC).name) > 0 Then
                        tStr = tStr & UserList(LoopC).name & ", "
                    End If
                Next
                If Len(tStr) > 0 Then
                    tStr = Left$(tStr, Len(tStr) - 2)
                    Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "4L" & NumUsers + NumBots)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "6L")
                End If
                '        Else
                '           Call SendData(ToIndex, UserIndex, 0, "||Este comando ya no está disponible. La cantidad de users online está abajo de la pantalla." & FONTTYPE_INFO)
            End If
            Exit Sub

        Case "#Q"
            Call SendUserSTAtsTxt(UserIndex, UserIndex)
            Exit Sub
        Case "#R"
            If UserList(UserIndex).Counters.Pena Then
                Call SendData(ToIndex, UserIndex, 0, "9M" & CalcularTiempoCarcel(UserIndex))
            Else
                Call SendData(ToIndex, UserIndex, 0, "2N")
            End If
            Exit Sub
        Case "#S"
            If UserList(UserIndex).flags.TargetUser Then
                If MapData(UserList(UserList(UserIndex).flags.TargetUser).pos.Map, UserList(UserList(UserIndex).flags.TargetUser).pos.X, UserList(UserList(UserIndex).flags.TargetUser).pos.Y).OBJInfo.OBJIndex > 0 And _
                   UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto Then
                    Call SendData(ToAdmins, 0, 0, "8T" & UserList(UserIndex).name & "," & UserList(UserList(UserIndex).flags.TargetUser).name)
                    Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "!!Fuiste echado por mantenerte sobre un item estando muerto.")
                    Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "FINOK")
                    Call CloseSocket(UserList(UserIndex).flags.TargetUser)
                End If
            End If
            Exit Sub

        Case "#T"
            If entorneo Then
                Dim jugadores As Integer
                jugadores = val(GetVar(App.path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD"))
                Dim jugador As Integer
                For jugador = 1 To jugadores
                    If UCase$(GetVar(App.path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugador)) = UCase$(UserList(UserIndex).name) Then Exit Sub
                Next
                Call WriteVar(App.path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD", jugadores + 1)
                Call WriteVar(App.path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugadores + 1, UserList(UserIndex).name)
                Call SendData(ToIndex, UserIndex, 0, "9T")
                Call SendData(ToAdmins, 0, 0, "2U" & UserList(UserIndex).name)
                PTorneo = PTorneo - 1
                If PTorneo = 0 Then
                    Call SendData(ToAll, 0, 0, "||CUPO ALCANZADO, Ya estan elegidos los participantes del Torneo!." & FONTTYPE_TALK)
                    entorneo = 0
                    Exit Sub
                End If
            End If
            Exit Sub

        Case "#U"
            Dim NpcIndex As Integer
            Dim theading As Byte
            Dim atra1 As Integer
            Dim atra2 As Integer
            Dim atra3 As Integer
            Dim atra4 As Integer

            If Not LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X - 1, UserList(UserIndex).pos.Y) And _
               Not LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X + 1, UserList(UserIndex).pos.Y) And _
               Not LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1) And _
               Not LegalPos(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1) Then
                If UserList(UserIndex).flags.Muerto Then
                    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X - 1, UserList(UserIndex).pos.Y).NpcIndex Then
                        atra1 = MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X - 1, UserList(UserIndex).pos.Y).NpcIndex
                        theading = WEST
                        Call MoveNPCChar(atra1, theading)
                    End If
                    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X + 1, UserList(UserIndex).pos.Y).NpcIndex Then
                        atra2 = MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X + 1, UserList(UserIndex).pos.Y).NpcIndex
                        theading = EAST
                        Call MoveNPCChar(atra2, theading)
                    End If
                    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).NpcIndex Then
                        atra3 = MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).NpcIndex
                        theading = NORTH
                        Call MoveNPCChar(atra3, theading)
                    End If
                    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1).NpcIndex Then
                        atra4 = MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1).NpcIndex
                        theading = SOUTH
                        Call MoveNPCChar(atra4, theading)
                    End If
                End If
            End If
            Exit Sub

        Case "#V"

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If
            If UserList(UserIndex).flags.Privilegios = 1 Then
                Exit Sub
            End If

            If UserList(UserIndex).flags.TargetNpc Then

                If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNpc).Desc) > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "3Q" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(ToIndex, UserIndex, 0, "DL")
                    Exit Sub
                End If

                Call IniciarComercioNPC(UserIndex)


            ElseIf UserList(UserIndex).flags.TargetUser Then


                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto Then
                    Call SendData(ToIndex, UserIndex, 0, "4U")
                    Exit Sub
                End If

                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(ToIndex, UserIndex, 0, "5U")
                    Exit Sub
                End If

                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).pos, UserList(UserIndex).pos) > 3 Then
                    Call SendData(ToIndex, UserIndex, 0, "DL")
                    Exit Sub
                End If

                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando And _
                   UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(ToIndex, UserIndex, 0, "6U")
                    Exit Sub
                End If

                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).name
                UserList(UserIndex).ComUsu.Cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False


                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)

            Else
                Call SendData(ToIndex, UserIndex, 0, "ZP")
            End If
            Exit Sub


        Case "#W"

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If

            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If

            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 3 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub

            Call IniciarDeposito(UserIndex)

            Exit Sub

        Case "#Y"


            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(UserIndex).flags.Muerto Then Exit Sub

            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If

            If ClaseBase(UserList(UserIndex).Clase) Or ClaseTrabajadora(UserList(UserIndex).Clase) Then Exit Sub

            Call Enlistar(UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion)

            Exit Sub
        Case "#1"

            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(UserIndex).flags.Muerto Or Not Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion Then Exit Sub
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If

            If UserList(UserIndex).Faccion.Bando <> Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion Then
                Call SendData(ToIndex, UserIndex, 0, Mensajes(Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion, 16) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call Recompensado(UserIndex)
            Exit Sub
        Case "#5"
            rdata = Right$(rdata, Len(rdata) - 3)

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "M4")
                Exit Sub
            End If

            If Not AsciiValidos(rdata) Then
                Call SendData(ToIndex, UserIndex, 0, "7U")
                Exit Sub
            End If

            If Len(rdata) > 80 Then
                Call SendData(ToIndex, UserIndex, 0, "||La descripción debe tener menos de 80 cáracteres de largo." & FONTTYPE_INFO)
                Exit Sub
            End If

            UserList(UserIndex).Desc = rdata
            Call SendData(ToIndex, UserIndex, 0, "8U")
            Exit Sub

        Case "#6 "
            rdata = Right$(rdata, Len(rdata) - 3)
            Call ComputeVote(UserIndex, rdata)
            Exit Sub

        Case "#7"
            Call SendData(ToIndex, UserIndex, 0, "||Este comando ya no anda, para hablar por tu clan presiona la tecla 3 y habla normalmente." & FONTTYPE_INFO)
            Exit Sub

        Case "#8"
            Call SendData(ToIndex, UserIndex, 0, "||Este comando ya no se usa, pon /PASSWORD para cambiar tu password." & FONTTYPE_INFO)
            Exit Sub

        Case "#!"
            If PuedeFaccion(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "4&")
            Exit Sub

        Case "#9"
            rdata = Right$(rdata, Len(rdata) - 3)
            tLong = CLng(val(rdata))
            If tLong > 32000 Then tLong = 32000
            N = tLong
            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
            ElseIf UserList(UserIndex).flags.TargetNpc = 0 Then

                Call SendData(ToIndex, UserIndex, 0, "ZP")
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
            ElseIf Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_APOSTADOR Then
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            ElseIf N < 1 Then
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            ElseIf N > 5000 Then
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            ElseIf UserList(UserIndex).Stats.GLD < N Then
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                    Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

                    Apuestas.Ganancias = Apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                    Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

                    Apuestas.Perdidas = Apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))

                Call SendUserORO(UserIndex)
            End If
            Exit Sub




        Case "#,"    ' /irdesafio
            If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then Exit Sub

            If UserList(UserIndex).flags.Muerto Then    ' muerto?
                Call SendData(ToIndex, UserIndex, 0, "||Estás muerto!" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).pos.Map = 22 Then    ' ya entro?
                Call SendData(ToIndex, UserIndex, 0, "||Ya estás en Desafio." & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).pos.Map = 66 Then    'Carcel?
                Call SendData(ToIndex, UserIndex, 0, "||Estas en la carcel, no podes desafiar." & FONTTYPE_INFO)
                Exit Sub
            End If
            If DesaFiante(1) <> 0 Then    ' alguien?
                Call SendData(ToIndex, UserIndex, 0, "||Ya hay un desafio, si quieres participar escribe /DESAFIAR." & FONTTYPE_INFO)
                Exit Sub
            ElseIf DesaFiante(2) <> 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Ya está peleando con otro usuario." & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.GLD < 200000 Then
                Call SendData(ToIndex, UserIndex, 0, "||Para crear un desafio necesitas 200.000 monedas de oro." & FONTTYPE_SERVER)
                Exit Sub
            End If

            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 200000
            Call WarpUserChar(UserIndex, 22, 83, 36, True)    ' para /desafiar
            Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).name & " [Clase: " & ListaClases(UserList(UserIndex).Clase) & " - Frags: " & UserList(UserIndex).flags.tdead & " - Nivel: " & UserList(UserIndex).Stats.ELV & "] Desafia a cualquier usuario a un Duelo. Si quieres participar Escribe /DESAFIAR." & FONTTYPE_DESAFIO)
            UserList(UserIndex).flags.Esperando = True
            DeFenZas = 0
            DesaFiante(1) = UserIndex
            Call SendUserORO(DesaFiante(1))
            Exit Sub
        Case "#-"    '/desafiar
            If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then Exit Sub

            If UserList(UserIndex).flags.Muerto Then    ' muerto?
                Call SendData(ToIndex, UserIndex, 0, "||Estás muerto!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).pos.Map = 22 Then    ' ya entro?
                Call SendData(ToIndex, UserIndex, 0, "||Ya estás en Desafio." & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).pos.Map = 66 Then    'Carcel?
                Call SendData(ToIndex, UserIndex, 0, "||Estas en la carcel, no podes desafiar." & FONTTYPE_INFO)
                Exit Sub
            End If
            If MapInfo(22).NumUsers = 2 Then  ' alguien?
                Call SendData(ToIndex, UserIndex, 0, "||La sala de desafíos está ocupada." & FONTTYPE_INFO)
                Exit Sub
            End If
            If MapInfo(22).NumUsers = 0 Then  ' alguien?
                Call SendData(ToIndex, UserIndex, 0, "||No hay nadie en desafio, si quieres participar escribe  /IRDESAFIO" & FONTTYPE_INFO)
                Exit Sub
            End If


            If UserList(DesaFiante(1)).flags.Esperando Then
                Call WarpUserChar(DesaFiante(1), 22, 83, 36, False)  'Modifiquen el mapa, x, y
                Call SendData(ToIndex, DesaFiante(1), 0, "||Nombre: " & UserList(UserIndex).name & " - (Clase: " & ListaClases(UserList(UserIndex).Clase) & ") - (Frags: " & UserList(UserIndex).flags.tdead & ") - (Nivel: " & UserList(UserIndex).Stats.ELV & ") - Entro a la sala de desafios." & FONTTYPE_FENIX)
                UserList(DesaFiante(1)).Stats.MinHP = UserList(DesaFiante(1)).Stats.MaxHP
                UserList(DesaFiante(1)).Stats.MinMAN = UserList(DesaFiante(1)).Stats.MaxMAN
                Call SendUserStatsBox(DesaFiante(1))
                Call WarpUserChar(UserIndex, 22, 83, 92, True)    ' entro al desafio
                Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).name & " entró al desafio" & FONTTYPE_DESAFIO)
                UserList(UserIndex).flags.Desafiando = True
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 30000
                Call SendUserORO(UserIndex)
                DesaFiante(2) = UserIndex
                Exit Sub
            End If

        Case "#Z"

            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes teletransportarte si te encuentras paralizado!!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then
                Call SendData(ToIndex, UserIndex, 0, "||No podes viajar en zona insegura." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).Stats.GLD < 10000 Then
                Call SendData(ToIndex, UserIndex, 0, "||Para ir a esta ciudad necesitas 10.000 monedas de oro" & "~0~255~255~1~0")

                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 66 Then

                Call SendData(ToIndex, UserIndex, 0, "|| No podes ir a la ciudad si estas en la carcel." & FONTTYPE_INFO)

                Exit Sub

            End If
            If UserList(UserIndex).pos.Map = 60 Then

                Call SendData(ToIndex, UserIndex, 0, "|| No podes viajar si sos Newbie." & FONTTYPE_INFO)
                Exit Sub
            End If

            Call WarpUserChar(UserIndex, 1, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
            Call SendUserStatsBox(UserIndex)

            Exit Sub

        Case "#X"

            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes teletransportarte si te encuentras paralizado!!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then
                Call SendData(ToIndex, UserIndex, 0, "||No podes viajar en zona insegura." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).Stats.GLD < 10000 Then

                Call SendData(ToIndex, UserIndex, 0, "||Para ir a esta ciudad necesitas 10.000 monedas de oro" & "~0~255~255~1~0")

                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 66 Then

                Call SendData(ToIndex, UserIndex, 0, "|| No podes ir a la ciudad si estas en la carcel." & FONTTYPE_INFO)

                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 60 Then

                Call SendData(ToIndex, UserIndex, 0, "|| No podes viajar si sos Newbie." & FONTTYPE_INFO)
                Exit Sub
            End If

            Call WarpUserChar(UserIndex, 28, 50, 52, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
            Call SendUserStatsBox(UserIndex)

            Exit Sub

        Case "#®"

            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes teletransportarte si te encuentras paralizado!!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then
                Call SendData(ToIndex, UserIndex, 0, "||No podes viajar en zona insegura." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).Stats.GLD < 10000 Then

                Call SendData(ToIndex, UserIndex, 0, "||Para ir a esta ciudad necesitas 10.000 monedas de oro" & "~0~255~255~1~0")

                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 66 Then

                Call SendData(ToIndex, UserIndex, 0, "|| No podes ir a la ciudad si estas en la carcel." & FONTTYPE_INFO)

                Exit Sub

            End If
            If UserList(UserIndex).pos.Map = 60 Then

                Call SendData(ToIndex, UserIndex, 0, "|| No podes viajar si sos Newbie." & FONTTYPE_INFO)
                Exit Sub
            End If

            Call WarpUserChar(UserIndex, 32, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
            Call SendUserStatsBox(UserIndex)

            Exit Sub

        Case "#¥"

            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes teletransportarte si te encuentras paralizado!!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then
                Call SendData(ToIndex, UserIndex, 0, "||No podes viajar en zona insegura." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).Stats.GLD < 10000 Then

                Call SendData(ToIndex, UserIndex, 0, "||Para ir a esta ciudad necesitas 10.000 monedas de oro" & "~0~255~255~1~0")

                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 66 Then
                Call SendData(ToIndex, UserIndex, 0, "|| No podes ir a la ciudad si estas en la carcel." & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).pos.Map = 60 Then
                Call SendData(ToIndex, UserIndex, 0, "|| No podes viajar si sos Newbie." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call WarpUserChar(UserIndex, 43, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
            Call SendUserStatsBox(UserIndex)
            Exit Sub

        Case "#Ø"
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes teletransportarte si te encuentras paralizado!!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then
                Call SendData(ToIndex, UserIndex, 0, "||No podes viajar en zona insegura." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).Stats.GLD < 10000 Then
                Call SendData(ToIndex, UserIndex, 0, "||Para ir a esta ciudad necesitas 10.000 monedas de oro" & "~0~255~255~1~0")
                Exit Sub
            End If

            If UserList(UserIndex).pos.Map = 66 Then

                Call SendData(ToIndex, UserIndex, 0, "|| No podes ir a la ciudad si estas en la carcel." & FONTTYPE_INFO)

                Exit Sub

            End If

            If UserList(UserIndex).pos.Map = 60 Then
                Call SendData(ToIndex, UserIndex, 0, "|| No podes viajar si sos Newbie." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call WarpUserChar(UserIndex, 18, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
            Call SendUserStatsBox(UserIndex)
            Exit Sub

        Case "#0"
            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If UserList(UserIndex).flags.Muerto Then Exit Sub

            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub

            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If

            rdata = Right$(rdata, Len(rdata) - 3)

            If val(rdata) > 0 Then
                If val(rdata) > UserList(UserIndex).Stats.Banco Then rdata = UserList(UserIndex).Stats.Banco
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rdata)
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
            End If
            Call SendUserORO(UserIndex)
            Exit Sub

        Case "#Ñ"
            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "MU")
                Exit Sub
            End If
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Or UserList(UserIndex).flags.Muerto Then Exit Sub
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 3)
            If CLng(val(rdata)) > 0 Then
                If CLng(val(rdata)) > UserList(UserIndex).Stats.GLD Then rdata = UserList(UserIndex).Stats.GLD
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rdata)
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
            End If
            Call SendUserORO(UserIndex)
            Exit Sub

        Case "#~"    '/INSCRIBIR
            rdata = Right$(rdata, Len(rdata) - 3)
            TIndex = NameIndex(rdata)
            If TIndex <= 0 Then Exit Sub
            If MegaTorneo = False Then
                Call SendData(ToIndex, UserIndex, 0, "||No Disponible!" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).pos.Map = 14 Then
                Call SendData(ToIndex, UserIndex, 0, "||No se Permite Mandar mas de una Vez!" & FONTTYPE_INFO)
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "||Has sido enviado a la Zona de Espera del Torneo!" & FONTTYPE_WARNING)
            iPosicionTorneo = iPosicionTorneo + 1
            If iPosicionTorneo = 1 Then
                Call WarpUserChar(UserIndex, 14, 56, 86, True)
                Call WarpUserChar(TIndex, 14, 54, 86, True)
                Exit Sub
            End If
            If iPosicionTorneo = 2 Then
                Call WarpUserChar(UserIndex, 14, 52, 86, True)
                Call WarpUserChar(TIndex, 14, 50, 86, True)
                Exit Sub
            End If
            If iPosicionTorneo = 3 Then
                Call WarpUserChar(UserIndex, 14, 48, 86, True)
                Call WarpUserChar(TIndex, 14, 46, 86, True)
                Exit Sub
            End If
            If iPosicionTorneo = 4 Then
                Call WarpUserChar(UserIndex, 14, 44, 86, True)
                Call WarpUserChar(TIndex, 14, 42, 86, True)
                Exit Sub
            End If
            If iPosicionTorneo = 5 Then
                Call WarpUserChar(UserIndex, 14, 40, 86, True)
                Call WarpUserChar(TIndex, 14, 38, 86, True)
                Exit Sub
            End If
            If iPosicionTorneo = 6 Then
                Call WarpUserChar(UserIndex, 14, 36, 86, True)
                Call WarpUserChar(TIndex, 14, 34, 86, True)
                Exit Sub
            End If
            If iPosicionTorneo = 7 Then
                Call WarpUserChar(UserIndex, 14, 32, 86, True)
                Call WarpUserChar(TIndex, 14, 30, 86, True)
                Exit Sub
            End If
            If iPosicionTorneo = 8 Then
                Call WarpUserChar(UserIndex, 14, 28, 86, True)
                Call WarpUserChar(TIndex, 14, 26, 86, True)
                Call SendData(ToAll, UserIndex, 0, "||Torneo Lleno" & FONTTYPE_FENIX)
                AutoTorneo = 0
                PosicionTorneo = 0
                Call SendData(ToAll, UserIndex, 0, "||Modo Torneo Desactivado." & FONTTYPE_WARNING)
                Exit Sub
            End If
            Exit Sub

        Case "#("    '/ENTRAR
            If AutoTorneo = False Then
                Call SendData(ToIndex, UserIndex, 0, "||No Disponible!" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).pos.Map = 14 Then
                Call SendData(ToIndex, UserIndex, 0, "||No se Permite Mandar mas de una Vez!" & FONTTYPE_INFO)
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "||Has sido enviado a la Zona de Espera del Torneo!" & FONTTYPE_WARNING)
            PosicionTorneo = PosicionTorneo + 1
            If PosicionTorneo = 1 Then
                Call WarpUserChar(UserIndex, 14, 55, 86, True)
                Exit Sub
            End If
            If PosicionTorneo = 2 Then
                Call WarpUserChar(UserIndex, 14, 51, 86, True)
                Exit Sub
            End If
            If PosicionTorneo = 3 Then
                Call WarpUserChar(UserIndex, 14, 47, 86, True)
                Exit Sub
            End If
            If PosicionTorneo = 4 Then
                Call WarpUserChar(UserIndex, 14, 43, 86, True)
                Exit Sub
            End If
            If PosicionTorneo = 5 Then
                Call WarpUserChar(UserIndex, 14, 39, 86, True)
                Exit Sub
            End If
            If PosicionTorneo = 6 Then
                Call WarpUserChar(UserIndex, 14, 35, 86, True)
                Exit Sub
            End If
            If PosicionTorneo = 7 Then
                Call WarpUserChar(UserIndex, 14, 31, 86, True)
                Exit Sub
            End If
            If PosicionTorneo = 8 Then
                Call WarpUserChar(UserIndex, 14, 27, 86, True)
                Call SendData(ToAll, UserIndex, 0, "||Torneo Lleno" & FONTTYPE_FENIX)
                AutoTorneo = 0
                PosicionTorneo = 0
                Call SendData(ToAll, UserIndex, 0, "||Modo Torneo Desactivado." & FONTTYPE_WARNING)
                Exit Sub
            End If
            Exit Sub

        Case "#2"
            If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then
                If UserList(UserIndex).GuildInfo.EsGuildLeader And UserList(UserIndex).flags.Privilegios < 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "4V")
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "5V")
                Exit Sub
            End If
            Call SendData(ToGuildMembers, UserIndex, 0, "6V" & UserList(UserIndex).name)
            Call SendData(ToIndex, UserIndex, 0, "7V")
            Dim oGuild As cGuild
            Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
            If oGuild Is Nothing Then Exit Sub
            For i = 1 To LastUser
                If UserList(i).GuildInfo.GuildName = oGuild.GuildName Then UserList(i).flags.InfoClanEstatica = 0
            Next
            UserList(UserIndex).GuildInfo.GuildPoints = 0
            UserList(UserIndex).GuildInfo.GuildName = ""
            Call oGuild.RemoveMember(UserList(UserIndex).name)
            Call UpdateUserChar(UserIndex)
            Exit Sub

        Case "#·"
            If Len(UserList(UserIndex).GuildInfo.GuildName) = 0 Then Exit Sub
            With UserList(UserIndex).GuildInfo
                .Seguro = Not .Seguro
                If .Seguro Then
                    Call SendData(ToIndex, UserIndex, 0, "|| Seguro de clanes activado." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "|| Seguro de clanes desactivado." & FONTTYPE_INFO)
                End If
            End With
            Exit Sub

        Case "#4"
            If UserList(UserIndex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "ZP")
                Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(UserIndex).flags.Muerto Or Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then Exit Sub
            If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
                Call SendData(ToIndex, UserIndex, 0, "DL")
                Exit Sub
            End If
            If UserList(UserIndex).Faccion.Bando <> Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion Then Exit Sub
            If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then
                Call SendData(ToIndex, UserIndex, 0, Mensajes(UserList(UserIndex).Faccion.Bando, 23) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, Mensajes(Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion, 18) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            UserList(UserIndex).Faccion.Bando = Neutral
            UserList(UserIndex).Faccion.Jerarquia = 0
            Call UpdateUserChar(UserIndex)
            Exit Sub

        Case "#3"
            If Len(UserList(UserIndex).GuildInfo.GuildName) = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "5V")
                Exit Sub
            End If
            For LoopC = 1 To LastUser
                If UserList(LoopC).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName Then
                    tStr = tStr & UserList(LoopC).name & ", "
                End If
            Next
            If Len(tStr) > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                Call SendData(ToIndex, UserIndex, 0, "||Miembros de tu clan online:" & tStr & "." & FONTTYPE_GUILD)
            Else: Call SendData(ToIndex, UserIndex, 0, "8V")
            End If
            Exit Sub


        Case "#)"    'espada
            Dim TengoA As Integer
            Dim Superoro As Obj


            TengoA = 0
            If TieneObjetos(920, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(887, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(873, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(413, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(407, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(959, 1, UserIndex) Then TengoA = TengoA + 1

            If TengoA = 8 Then
                Call QuitarObjetos(920, 1, UserIndex)
                Call QuitarObjetos(887, 1, UserIndex)
                Call QuitarObjetos(873, 1, UserIndex)
                Call QuitarObjetos(407, 1, UserIndex)
                Call QuitarObjetos(413, 1, UserIndex)
                Call QuitarObjetos(959, 1, UserIndex)

                Superoro.Amount = 1  'Cantidad de copas
                Superoro.OBJIndex = 922  'Numero de item

                If Not MeterItemEnInventario(UserIndex, Superoro) Then Call TirarItemAlPiso(UserList(UserIndex).pos, Superoro)
                Call SendData(ToIndex, UserIndex, 0, "||Contruiste la Espada de Orsula." & FONTTYPE_FENIX)
                Exit Sub
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No tienes uno de los objetos requeridos para obtener la Espada de Orsula." & FONTTYPE_FIGHT)
                Exit Sub
            End If

        Case "#?"    'escudo

            TengoA = 0
            If TieneObjetos(859, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(906, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(884, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(410, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(411, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(959, 1, UserIndex) Then TengoA = TengoA + 1

            If TengoA = 8 Then
                Call QuitarObjetos(859, 1, UserIndex)
                Call QuitarObjetos(906, 1, UserIndex)
                Call QuitarObjetos(884, 1, UserIndex)
                Call QuitarObjetos(410, 1, UserIndex)
                Call QuitarObjetos(411, 1, UserIndex)
                Call QuitarObjetos(959, 1, UserIndex)
                Superoro.Amount = 1  'Cantidad de copas
                Superoro.OBJIndex = 921  'Numero de item
                If Not MeterItemEnInventario(UserIndex, Superoro) Then Call TirarItemAlPiso(UserList(UserIndex).pos, Superoro)
                Call SendData(ToIndex, UserIndex, 0, "||Contruiste el Escudo de Geryon." & FONTTYPE_FENIX)
                Exit Sub
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No tienes uno de los objetos requeridos para obtener el Escudo de Geryon." & FONTTYPE_FIGHT)
                Exit Sub
            End If


        Case "#¡"    'armadura
            TengoA = 0
            If TieneObjetos(920, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(887, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(873, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(413, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(407, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(959, 1, UserIndex) Then TengoA = TengoA + 1

            If TengoA = 8 Then
                Call QuitarObjetos(920, 1, UserIndex)
                Call QuitarObjetos(887, 1, UserIndex)
                Call QuitarObjetos(873, 1, UserIndex)
                Call QuitarObjetos(407, 1, UserIndex)
                Call QuitarObjetos(413, 1, UserIndex)
                Call QuitarObjetos(959, 1, UserIndex)
                Superoro.Amount = 1  'Cantidad de copas
                Superoro.OBJIndex = 923  'Numero de item
                If Not MeterItemEnInventario(UserIndex, Superoro) Then Call TirarItemAlPiso(UserList(UserIndex).pos, Superoro)
                Call SendData(ToIndex, UserIndex, 0, "||Contruiste la Armadura de Baal." & FONTTYPE_FENIX)
                Exit Sub
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No tienes uno de los objetos requeridos para obtener la Armadura de Baal." & FONTTYPE_FIGHT)
                Exit Sub
            End If

        Case "#¬"    'anillo
            TengoA = 0
            If TieneObjetos(413, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(864, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(882, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(408, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(406, 1, UserIndex) Then TengoA = TengoA + 1
            If TieneObjetos(959, 1, UserIndex) Then TengoA = TengoA + 1

            If TengoA = 8 Then
                Call QuitarObjetos(413, 1, UserIndex)
                Call QuitarObjetos(864, 1, UserIndex)
                Call QuitarObjetos(882, 1, UserIndex)
                Call QuitarObjetos(408, 1, UserIndex)
                Call QuitarObjetos(406, 1, UserIndex)
                Call QuitarObjetos(959, 1, UserIndex)
                Superoro.Amount = 1  'Cantidad de copas
                Superoro.OBJIndex = 924  'Numero de item
                If Not MeterItemEnInventario(UserIndex, Superoro) Then Call TirarItemAlPiso(UserList(UserIndex).pos, Superoro)
                Call SendData(ToIndex, UserIndex, 0, "||Contruiste el Anillo de Pyro." & FONTTYPE_INFO)
                Exit Sub
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No tienes uno de los objetos requeridos para obtener el Anillo de Pyro." & FONTTYPE_FIGHT)
                Exit Sub
            End If

        Case "#°"
            TengoA = 0

            If UserList(UserIndex).Stats.ELV < 60 Then    ' si es 50 te deja
                Call SendData(ToIndex, UserIndex, 0, "||Debes ser al menos Nivel 60 para hacerte noble!!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).flags.EsNoble = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||Ya sos Noble." & FONTTYPE_INFO)
                Exit Sub
            End If

            For i = 921 To 924
                If TieneObjetos(i, 1, UserIndex) Then TengoA = TengoA + 1
            Next i
            If TengoA = 4 Then
                Call QuitarObjetos(921, 1, UserIndex)
                Call QuitarObjetos(922, 1, UserIndex)
                Call QuitarObjetos(923, 1, UserIndex)
                Call QuitarObjetos(924, 1, UserIndex)
                UserList(UserIndex).flags.EsNoble = 1
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Te hiciste Noble!! ¡¡Felicitaciones!! Ahora podras tener todas las ventajas de un Noble.~125~255~125~1~1")
                SendData ToAll, 0, 0, "||" & UserList(UserIndex).name & " es nuevo Noble de estas tierras." & FONTTYPE_FENIX
                Superoro.Amount = 1  'Cantidad
                Superoro.OBJIndex = 878  'Numero de item
                If Not MeterItemEnInventario(UserIndex, Superoro) Then Call TirarItemAlPiso(UserList(UserIndex).pos, Superoro)
                Exit Sub
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Necesitas los objetos de noble!!" & FONTTYPE_INFO)
                Exit Sub
            End If

    End Select
    Procesado = False
End Sub
