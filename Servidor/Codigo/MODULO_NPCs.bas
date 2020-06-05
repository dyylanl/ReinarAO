Attribute VB_Name = "NPCs"
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Option Explicit
Public rdata As String
Sub QuitarMascota(UserIndex As Integer, ByVal NpcIndex As Integer)
    Dim i As Integer

    UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1

    For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
            UserList(UserIndex).MascotasIndex(i) = 0
            UserList(UserIndex).MascotasType(i) = 0
            Exit For
        End If
    Next

End Sub
Sub QuitarMascotaNpc(Maestro As Integer, ByVal Mascota As Integer)

    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

End Sub
Sub MuereNpc(ByVal NpcIndex As Integer, UserIndex As Integer)
    On Error GoTo errhandler
    Dim Exp As Long
    Dim MiNPC As Npc
    MiNPC = Npclist(NpcIndex)



    If MiNPC.Stats.MinHP < 1 Then MiNPC.Stats.MinHP = MiNPC.Stats.MaxHP

    Call QuitarNPC(NpcIndex)

    If (MataGuardiasEmperador <> 3) And ((MiNPC.Numero = 642) Or (MiNPC.Numero = 643)) Then
        MataGuardiasEmperador = MataGuardiasEmperador + 1
        If MataGuardiasEmperador = 3 Then
            Call SendData(ToAll, 0, 0, "||El Emperador es vulnerable a los ataques..." & FONTTYPE_FIGHT)
        End If
        Exit Sub
    End If
    If MiNPC.Numero = 244 Then
        Call SendData(ToAll, 0, 0, "||El Emperador ha muerto y reencarnará en busca de venganza por los que lo han matado..." & FONTTYPE_INFO)
        MataGuardiasEmperador = 0
        frmMain.TEmperador.Enabled = True
    End If

    If MiNPC.pos.Map = mapa_castilloNorte And UCase$(MiNPC.name) = "REY DEL CASTILLO" And UserList(UserIndex).GuildInfo.GuildName <> "" Then
        CastilloNorte = UserList(UserIndex).GuildInfo.GuildName

        DateNorte = Date
        HoraNorte = Time
        Call SendData(ToAll, 0, 0, "||El clan " & UserList(UserIndex).GuildInfo.GuildName & " ha conquistado el castillo Norte." & FONTTYPE_GUILD)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "castillo1", UserList(UserIndex).GuildInfo.GuildName)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "date1", Date)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "hora1", Time)
        Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
        Dim SumRey As WorldPos
        SumRey.Map = mapa_castilloNorte
        SumRey.X = MapaX
        SumRey.Y = MapaY
        Call SpawnNpc(NpcRey, SumRey, True, False)
        If CastilloNorte = CastilloSur Then
            Call SendData(ToGuildMembers, UserIndex, 0, "||Ahora tu clan posee la Fortaleza de Clan, donde encontraras criaturas especiales para quests y teleports a los dungeons de este mundo. Escribe /FORTALEZA para ingresar." & FONTTYPE_FENIX)
        End If
    End If

    If MiNPC.pos.Map = mapa_castilloSur And UCase$(MiNPC.name) = "REY DEL CASTILLO" And UserList(UserIndex).GuildInfo.GuildName <> "" Then
        CastilloSur = UserList(UserIndex).GuildInfo.GuildName

        DateSur = Date
        HoraSur = Time
        Call SendData(ToAll, 0, 0, "||El clan " & UserList(UserIndex).GuildInfo.GuildName & " ha conquistado el castillo Sur." & FONTTYPE_GUILD)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "castillo2", UserList(UserIndex).GuildInfo.GuildName)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "date2", Date)
        Call WriteVar(IniPath & "castillos.txt", "INIT", "hora2", Time)
        Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
        SumRey.Map = mapa_castilloSur
        SumRey.X = MapaX
        SumRey.Y = MapaY
        Call SpawnNpc(NpcRey, SumRey, True, False)
        If CastilloNorte = CastilloSur Then
            Call SendData(ToGuildMembers, UserIndex, 0, "||Ahora tu clan posee la Fortaleza de Clan, donde encontraras criaturas especiales para quests y teleports a los dungeons de este mundo. Escribe /FORTALEZA para ingresar." & FONTTYPE_FENIX)
        End If
    End If



    If HayGuerra Then
        If MiNPC.pos.Map = CiudadGuerra And MiNPC.Numero = NPC1 Then
            TerminaGuerra "Caos"
        ElseIf MiNPC.pos.Map = CiudadGuerra And MiNPC.Numero = NPC2 Then
            TerminaGuerra "Real"
        End If
    End If

    If MiNPC.MaestroUser = 0 Then
        If UserIndex Then Call NPCTirarOro(MiNPC, UserIndex)
        If UserIndex Then Call NPC_TIRAR_ITEMS(MiNPC)
    End If

    If UserIndex > 0 Then Call SubirSkill(UserIndex, Supervivencia, 40)
    Call ReSpawnNpc(MiNPC)
    With UserList(UserIndex)
        If MiNPC.Numero = .Quest.IndexNPC Then
            .Quest.NPCs = .Quest.NPCs - 1
            If .Quest.NPCs <= 0 And .Quest.Users <= 0 Then
                SendData ToIndex, UserIndex, 0, "||Has terminado la quest!!" & FONTTYPE_FENIX
                'reset flags y damos recompensa
                .Quest.IndexNPC = 0
                .Quest.NPCs = 0
                .Quest.Users = 0

                'ahora con index de su quest le damos la recompensa y reseteamos
                If UserList(UserIndex).flags.EsNoble = 1 Then
                    .Stats.GLD = .Stats.GLD + Quest(.Quest.Index).Oro
                    .Stats.Exp = .Stats.Exp + Quest(.Quest.Index).Exp
                    .Faccion.Quests = .Faccion.Quests + Quest(.Quest.Index).Canje
                    .Stats.GLD = .Stats.GLD + Quest(.Quest.Index).Oro
                    .Stats.Exp = .Stats.Exp + Quest(.Quest.Index).Exp
                    .Faccion.Quests = .Faccion.Quests + Quest(.Quest.Index).Canje
                    .Quest.Index = 0
                    CheckUserLevel UserIndex
                    SendUserStatsBox UserIndex
                    Exit Sub
                Else
                    .Stats.GLD = .Stats.GLD + Quest(.Quest.Index).Oro
                    .Stats.Exp = .Stats.Exp + Quest(.Quest.Index).Exp
                    .Faccion.Quests = .Faccion.Quests + Quest(.Quest.Index).Canje
                    .Quest.Index = 0
                    CheckUserLevel UserIndex
                    SendUserStatsBox UserIndex
                    Exit Sub
                End If
                'cerramos su indice

                'le avisamos qe termino

            End If
        End If
    End With

errhandler:
    If Err.Number <> 0 Then
        Call LogError("Error en MuereNpc " & Err.Description)
    End If
End Sub
Function NPCListable(NpcIndex As Integer) As Boolean

    NPCListable = (Npclist(NpcIndex).Attackable And Not Npclist(NpcIndex).flags.Respawn)

End Function
Sub QuitarNPC(ByVal NpcIndex As Integer)
    On Error GoTo errhandler
    Dim i As Integer

    If Npclist(NpcIndex).NPCtype = NPCTYPE_BOT Then
        Call QuitarNPCBOT(NpcIndex)
        Exit Sub
    End If

    Npclist(NpcIndex).flags.NPCActive = False

    If NPCListable(NpcIndex) Then Call QuitarNPCDeLista(Npclist(NpcIndex).Numero, Npclist(NpcIndex).pos.Map)

    Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).pos.Map, "QDL" & Npclist(NpcIndex).Char.CharIndex)

    If InMapBounds(Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y) Then Call EraseNPCChar(ToMap, 0, Npclist(NpcIndex).pos.Map, NpcIndex)

    If Npclist(NpcIndex).MaestroUser Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
    If Npclist(NpcIndex).MaestroNpc Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)

    Npclist(NpcIndex) = NpcNoIniciado

    For i = LastNPC To 1 Step -1
        If Npclist(i).flags.NPCActive Then
            LastNPC = i
            Exit For
        End If
    Next

    If NumNPCs Then NumNPCs = NumNPCs - 1

    Exit Sub

errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC-" & Err.Description)

End Sub
Sub QuitarNPCBOT(ByVal NpcIndex As Integer)
    On Error GoTo errhandler
    Dim i As Integer
    If Not Npclist(NpcIndex).NPCtype = NPCTYPE_BOT Then Exit Sub

    Npclist(NpcIndex).flags.NPCActive = False

    If NPCListable(NpcIndex) Then Call QuitarNPCDeLista(Npclist(NpcIndex).Numero, Npclist(NpcIndex).pos.Map)

    Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).pos.Map, "QDL" & Npclist(NpcIndex).Char.CharIndex)

    If InMapBounds(Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y) Then Call EraseNPCChar(ToMap, 0, Npclist(NpcIndex).pos.Map, NpcIndex)

    If Npclist(NpcIndex).MaestroUser Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
    If Npclist(NpcIndex).MaestroNpc Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)

    Npclist(NpcIndex) = NpcNoIniciado

    For i = LastNPC To 1 Step -1
        If Npclist(i).flags.NPCActive Then
            LastNPC = i
            Exit For
        End If
    Next

    If NumNPCs Then NumNPCs = NumNPCs - 1
    If NumBots Then NumBots = NumBots - 1


    Exit Sub

errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPCBOT-" & Err.Description)

End Sub
Function TestSpawnTrigger(pos As WorldPos) As Boolean

    If Not InMapBounds(pos.X, pos.Y) Or Not MapaValido(pos.Map) Then Exit Function

    TestSpawnTrigger = _
    MapData(pos.Map, pos.X, pos.Y).trigger <> 3 And _
                       MapData(pos.Map, pos.X, pos.Y).trigger <> 2 And _
                       MapData(pos.Map, pos.X, pos.Y).trigger <> 1

End Function
Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)


    Dim pos As WorldPos
    Dim newpos As WorldPos
    Dim nIndex As Integer
    Dim PosicionValida As Boolean
    Dim Iteraciones As Long
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    On Error GoTo Error

    nIndex = OpenNPC(NroNPC)

    If nIndex > MAXNPCS Then Exit Sub


    If InMapBounds(OrigPos.X, OrigPos.Y) Then

        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).pos = OrigPos

    Else

        pos.Map = mapa

        Do While Not PosicionValida
            DoEvents

            pos.X = CInt(Rnd * 100 + 1)
            pos.Y = CInt(Rnd * 100 + 1)

            Call ClosestLegalPos(pos, newpos, Npclist(nIndex).flags.AguaValida = 1)


            If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida = 1) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then

                Npclist(nIndex).pos.Map = newpos.Map
                Npclist(nIndex).pos.X = newpos.X
                Npclist(nIndex).pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0

            End If


            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                Call QuitarNPC(nIndex)
                Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                Exit Sub
            End If
        Loop


        Map = newpos.Map
        X = Npclist(nIndex).pos.X
        Y = Npclist(nIndex).pos.Y
    End If


    Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)

    If NPCListable(nIndex) Then Call AgregarNPC(Npclist(nIndex).Numero, mapa)
    Exit Sub
Error:

    Call LogError("Error en CrearNPC." & Map & " " & X & " " & Y & " " & nIndex)
End Sub
Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer, Map As Integer, X As Integer, Y As Integer)
    Dim CharIndex As Integer

    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If

    MapData(Map, X, Y).NpcIndex = NpcIndex
    If Npclist(NpcIndex).Char.Aura > 0 Then Call SendData(sndRoute, sndIndex, sndMap, ("CRA" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Char.Aura))

    Call SendData(sndRoute, sndIndex, sndMap, ("CC" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & X & "," & Y))    ' & ",0,0,0,0,0,0,0,0,0,0"))

End Sub

Sub ChangeNPCChar(NpcIndex As Integer, Body As Integer, Head As Integer, ByVal Heading As Byte)

    If Npclist(NpcIndex).Char.Body = Body And _
       Npclist(NpcIndex).Char.Head = Head And _
       Npclist(NpcIndex).Char.Heading = Heading Then Exit Sub
    If NpcIndex Then
        Npclist(NpcIndex).Char.Body = Body
        Npclist(NpcIndex).Char.Head = Head
        Npclist(NpcIndex).Char.Heading = Heading
        Call SendData(ToNPCAreaG, NpcIndex, Npclist(NpcIndex).pos.Map, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
    End If

End Sub

Sub EraseNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer)

    If Npclist(NpcIndex).Char.CharIndex Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

    If Npclist(NpcIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar < 1 Then Exit Do
        Loop
    End If


    MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y).NpcIndex = 0


    Call SendData(ToMap, 0, Npclist(NpcIndex).pos.Map, "BP" & THeDEnCripTe(Npclist(NpcIndex).Char.CharIndex, "mHlzsJxIQi"))


    Npclist(NpcIndex).Char.CharIndex = 0



    NumChars = NumChars - 1


End Sub
Sub MoveNPCChar(NpcIndex As Integer, ByVal nHeading As Byte)
    On Error GoTo errh
    Dim nPos As WorldPos

    If Npclist(NpcIndex).AutoCurar = 1 Then Exit Sub

    nPos = Npclist(NpcIndex).pos
    Call HeadtoPos(nHeading, nPos)

    If (Npclist(NpcIndex).MaestroUser And LegalPos(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y)) Or LegalPosNPC(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida = 1) Then
        If (Npclist(NpcIndex).flags.AguaValida = 0 And MapData(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y).Agua = 1) Or (Npclist(NpcIndex).flags.TierraInvalida = 1 And MapData(Npclist(NpcIndex).pos.Map, nPos.X, nPos.Y).Agua = 0) Then Exit Sub

        Call SendData(ToNPCAreaG, NpcIndex, Npclist(NpcIndex).pos.Map, "MP" & THeDEnCripTe(Npclist(NpcIndex).Char.CharIndex & "," & (nPos.X) & "," & (nPos.Y), "mHlzsJxIQi"))


        MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y).NpcIndex = 0
        Npclist(NpcIndex).pos = nPos
        Npclist(NpcIndex).Char.Heading = nHeading
        MapData(Npclist(NpcIndex).pos.Map, Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y).NpcIndex = NpcIndex
    Else
        If Npclist(NpcIndex).Movement = NPC_PATHFINDING Then Npclist(NpcIndex).PFINFO.PathLenght = 0
    End If

    Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)

End Sub
Function Bin(N)

    Dim S As String, i As Integer, uu, T

    uu = Int(Log(N) / Log(2))

    For i = 0 To uu
        S = (N Mod 2) & S
        T = N / 2
        N = Int(T)
    Next
    Bin = S

End Function
Function NextOpenNPC() As Integer
    On Error GoTo errhandler

    Dim LoopC As Integer

    For LoopC = 1 To MAXNPCS + 1
        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next

    NextOpenNPC = LoopC

    Exit Function
errhandler:
    Call LogError("Error en NextOpenNPC")
End Function
Sub NpcEnvenenarUser(UserIndex As Integer)
    Dim N As Integer

    N = RandomNumber(1, 10)

    If N < 3 Then
        UserList(UserIndex).flags.Envenenado = 1
        UserList(UserIndex).flags.EstasEnvenenado = Timer
        UserList(UserIndex).Counters.Veneno = Timer
        Call SendData(ToIndex, UserIndex, 0, "1P")
    End If

End Sub
Function SpawnNpc(NpcIndex As Integer, pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
    On Error GoTo Error
    Dim newpos As WorldPos
    Dim nIndex As Integer
    Dim PosicionValida As Boolean
    Dim Map As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim it As Integer

    nIndex = OpenNPC(NpcIndex, Respawn)

    If nIndex > MAXNPCS Then
        SpawnNpc = nIndex
        Exit Function
    End If

    Do While Not PosicionValida
        Call ClosestLegalPos(pos, newpos)

        If LegalPos(newpos.Map, newpos.X, newpos.Y) Then
            Npclist(nIndex).pos.Map = newpos.Map
            Npclist(nIndex).pos.X = newpos.X
            Npclist(nIndex).pos.Y = newpos.Y
            PosicionValida = True
        Else
            newpos.X = 0
            newpos.Y = 0
        End If

        it = it + 1

        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(nIndex)
            SpawnNpc = MAXNPCS
            Call LogError("Más de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & pos.Map & " Index:" & NpcIndex)
            Exit Function
        End If
    Loop

    Map = newpos.Map
    X = Npclist(nIndex).pos.X
    Y = Npclist(nIndex).pos.Y

    Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)

    If NPCListable(nIndex) Then Call AgregarNPC(Npclist(nIndex).Numero, pos.Map)

    If FX Then
        Call SendData(ToNPCArea, nIndex, Npclist(NpcIndex).pos.Map, "TW" & SND_WARP & "," & Npclist(NpcIndex).pos.X & "," & Npclist(NpcIndex).pos.Y)
        Call SendData(ToNPCArea, nIndex, Npclist(NpcIndex).pos.Map, "CFM" & Npclist(nIndex).Char.CharIndex & "," & FXWARP & "," & 1)
    End If

    SpawnNpc = nIndex

    Exit Function
Error:
    Call LogError("Error en SpawnNPC: " & Err.Description & " " & nIndex & " " & X & " " & Y)
End Function
Sub ReSpawnNpc(MiNPC As Npc)

    If (MiNPC.flags.Respawn = 0) Then Exit Sub
    Call CrearNPC(MiNPC.Numero, MiNPC.pos.Map, MiNPC.Orig)

End Sub
Function NPCHostiles(Map As Integer) As Integer
    Dim i As Integer
    Dim cont As Integer

    cont = 0

    For i = 1 To UBound(MapInfo(Map).NPCsTeoricos)
        cont = cont + MapInfo(Map).NPCsReales(i).Cantidad
    Next

    NPCHostiles = cont

End Function
Sub NPCTirarOro(MiNPC As Npc, UserIndex As Integer)
    Dim i As Integer, MiembroIndex As Integer, OroDado As Long

    If MiNPC.GiveGLD Then
        If UserList(UserIndex).PartyIndex = 0 Then
            If MiNPC.GiveGLD + UserList(UserIndex).Stats.GLD <= MAXORO Then
                If UserList(UserIndex).flags.EsNoble = 1 Then
                    OroDado = (MiNPC.GiveGLD * CantidadOro) * 2
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + OroDado
                    Call SendUserORO(UserIndex)
                ElseIf UserList(UserIndex).flags.EsNoble = 0 Then
                    OroDado = (MiNPC.GiveGLD) * CantidadOro
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + OroDado
                    Call SendUserORO(UserIndex)
                End If
            End If
        Else
            For i = 1 To Party(UserList(UserIndex).PartyIndex).NroMiembros
                MiembroIndex = Party(UserList(UserIndex).PartyIndex).MiembrosIndex(i)
                If MiNPC.GiveGLD + UserList(MiembroIndex).Stats.GLD <= MAXORO Then
                    OroDado = (MiNPC.GiveGLD * CantidadOro) / Party(UserList(MiembroIndex).PartyIndex).NroMiembros
                    UserList(MiembroIndex).Stats.GLD = UserList(MiembroIndex).Stats.GLD + OroDado
                    Call SendUserORO(MiembroIndex)
                End If
            Next
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & PonerPuntos(OroDado) & " monedas de oro al matar a la criatura." & FONTTYPE_FIGHT)
    End If

End Sub
Function NameNpc(Number As Integer) As String
    Dim a As Long, S As Long

    If Number > 499 Then
        a = Anpc_host
    Else
        a = ANpc
    End If

    S = INIBuscarSeccion(a, "NPC" & Number)

    NameNpc = INIDarClaveStr(a, S, "Name")

End Function
Function OpenNPC(NPCNumber As Integer, Optional ByVal Respawn = True) As Integer
    On Local Error Resume Next
    Dim NpcIndex As Integer

    Dim a As Long, S As Long

    If NPCNumber > 499 Then

        a = Anpc_host
    Else

        a = ANpc
    End If

    S = INIBuscarSeccion(a, "NPC" & NPCNumber)

    NpcIndex = NextOpenNPC

    If NpcIndex > MAXNPCS Then
        OpenNPC = NpcIndex
        Exit Function
    End If

    Npclist(NpcIndex).Numero = NPCNumber






    If S >= 0 Then
        Npclist(NpcIndex).name = INIDarClaveStr(a, S, "Name")
        Npclist(NpcIndex).Desc = INIDarClaveStr(a, S, "Desc")

        Npclist(NpcIndex).Movement = INIDarClaveInt(a, S, "Movement")
        Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

        Npclist(NpcIndex).flags.AguaValida = INIDarClaveInt(a, S, "AguaValida")
        Npclist(NpcIndex).flags.TierraInvalida = INIDarClaveInt(a, S, "TierraInValida")
        Npclist(NpcIndex).flags.Faccion = INIDarClaveInt(a, S, "Faccion")

        Npclist(NpcIndex).NPCtype = INIDarClaveInt(a, S, "NpcType")

        Npclist(NpcIndex).Char.Body = INIDarClaveInt(a, S, "Body")
        Npclist(NpcIndex).Char.Aura = INIDarClaveInt(a, S, "Aura")

        Npclist(NpcIndex).Char.Head = INIDarClaveInt(a, S, "Head")
        Npclist(NpcIndex).Char.Heading = INIDarClaveInt(a, S, "Heading")

        Npclist(NpcIndex).Attackable = INIDarClaveInt(a, S, "Attackable")
        Npclist(NpcIndex).Comercia = INIDarClaveInt(a, S, "Comercia")
        Npclist(NpcIndex).Hostile = INIDarClaveInt(a, S, "Hostile")
        Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile
        Npclist(NpcIndex).GiveEXP = INIDarClaveInt(a, S, "GiveEXP")

        Npclist(NpcIndex).InmuneParalisis = INIDarClaveInt(a, S, "InmuneParalisis")

        Npclist(NpcIndex).Veneno = INIDarClaveInt(a, S, "Veneno")

        Npclist(NpcIndex).flags.Domable = INIDarClaveInt(a, S, "Domable")

        Npclist(NpcIndex).MaxRecom = INIDarClaveInt(a, S, "MaxRecom")
        Npclist(NpcIndex).MinRecom = INIDarClaveInt(a, S, "MinRecom")
        Npclist(NpcIndex).Probabilidad = INIDarClaveInt(a, S, "Probabilidad")
        Npclist(NpcIndex).GiveGLD = INIDarClaveInt(a, S, "GiveGLD")



        Npclist(NpcIndex).PoderAtaque = INIDarClaveInt(a, S, "PoderAtaque")
        Npclist(NpcIndex).PoderEvasion = INIDarClaveInt(a, S, "PoderEvasion")

        Npclist(NpcIndex).AutoCurar = INIDarClaveInt(a, S, "Autocurar")
        Npclist(NpcIndex).Stats.MaxHP = INIDarClaveInt(a, S, "MaxHP")
        Npclist(NpcIndex).Stats.MinHP = INIDarClaveInt(a, S, "MinHP")
        Npclist(NpcIndex).Stats.MaxHit = INIDarClaveInt(a, S, "MaxHIT")
        Npclist(NpcIndex).Stats.MinHit = INIDarClaveInt(a, S, "MinHIT")
        Npclist(NpcIndex).Stats.Def = INIDarClaveInt(a, S, "DEF")
        Npclist(NpcIndex).Stats.Alineacion = INIDarClaveInt(a, S, "Alineacion")
        Npclist(NpcIndex).Stats.ImpactRate = INIDarClaveInt(a, S, "ImpactRate")
        Npclist(NpcIndex).InvReSpawn = INIDarClaveInt(a, S, "InvReSpawn")
        Npclist(NpcIndex).Bot = INIDarClaveInt(a, S, "Bot")


        Dim LoopC As Integer
        Dim ln As String
        Npclist(NpcIndex).Invent.NroItems = INIDarClaveInt(a, S, "NROITEMS")


        For LoopC = 1 To Minimo(30, Npclist(NpcIndex).Invent.NroItems)
            ln = INIDarClaveStr(a, S, "Obj" & LoopC)
            Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
            Npclist(NpcIndex).Invent.Object(LoopC).ProbTirar = val(ReadField(3, ln, 45))
            Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
        Next


        If Npclist(NpcIndex).InvReSpawn Or Npclist(NpcIndex).Comercia = 0 Then

            For LoopC = 1 To Minimo(30, Npclist(NpcIndex).Invent.NroItems)
                ln = INIDarClaveStr(a, S, "Obj" & LoopC)
                Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
                Npclist(NpcIndex).Invent.Object(LoopC).ProbTirar = val(ReadField(3, ln, 45))
                Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            Next

        End If

        Npclist(NpcIndex).flags.LanzaSpells = INIDarClaveInt(a, S, "LanzaSpells")
        If Npclist(NpcIndex).flags.LanzaSpells Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
        For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
            Npclist(NpcIndex).Spells(LoopC) = INIDarClaveInt(a, S, "Sp" & LoopC)
        Next


        If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
            Npclist(NpcIndex).NroCriaturas = INIDarClaveInt(a, S, "NroCriaturas")
            ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
            For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
                Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = INIDarClaveInt(a, S, "CI" & LoopC)
                Npclist(NpcIndex).Criaturas(LoopC).NpcName = INIDarClaveStr(a, S, "CN" & LoopC)

            Next
        End If
        If Npclist(NpcIndex).NPCtype = NPCTYPE_BOT Then
            NumBots = NumBots + 1
            frmMain.CantUsuarios.Caption = NumUsers + NumBots
            Call SendData(ToAll, 0, 0, "NON" & NumUsers + NumBots)
        End If



        Npclist(NpcIndex).Inflacion = INIDarClaveInt(a, S, "Inflacion")

        Npclist(NpcIndex).flags.NPCActive = True
        Npclist(NpcIndex).flags.UseAINow = False

        If Respawn Then
            Npclist(NpcIndex).flags.Respawn = INIDarClaveInt(a, S, "ReSpawn")
        Else
            Npclist(NpcIndex).flags.Respawn = 1
        End If

        Npclist(NpcIndex).flags.RespawnOrigPos = INIDarClaveInt(a, S, "OrigPos")
        Npclist(NpcIndex).flags.AfectaParalisis = INIDarClaveInt(a, S, "AfectaParalisis")
        Npclist(NpcIndex).flags.GolpeExacto = INIDarClaveInt(a, S, "GolpeExacto")
        Npclist(NpcIndex).flags.Apostador = INIDarClaveInt(a, S, "Apostador")
        Npclist(NpcIndex).flags.PocaParalisis = INIDarClaveInt(a, S, "PocaParalisis")
        Npclist(NpcIndex).flags.NoMagia = INIDarClaveInt(a, S, "NoMagia")
        Npclist(NpcIndex).VeInvis = INIDarClaveInt(a, S, "VerInvis")

        Npclist(NpcIndex).flags.Snd1 = INIDarClaveInt(a, S, "Snd1")
        Npclist(NpcIndex).flags.Snd2 = INIDarClaveInt(a, S, "Snd2")
        Npclist(NpcIndex).flags.Snd3 = INIDarClaveInt(a, S, "Snd3")
        Npclist(NpcIndex).flags.Snd4 = INIDarClaveInt(a, S, "Snd4")



        Dim AUX As Long
        AUX = INIDarClaveInt(a, S, "NROEXP")
        Npclist(NpcIndex).NroExpresiones = (AUX)

        If AUX Then
            ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
            For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
                Npclist(NpcIndex).Expresiones(LoopC) = INIDarClaveStr(a, S, "Exp" & LoopC)
            Next
        End If




        Npclist(NpcIndex).TipoItems = INIDarClaveInt(a, S, "TipoItems")
    End If


    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1



    OpenNPC = NpcIndex

End Function


Function OpenNPC_Viejo(NPCNumber As Integer, Optional ByVal Respawn = True) As Integer

    Dim NpcIndex As Integer
    Dim npcfile As String

    If NPCNumber > 499 Then
        npcfile = DatPath & "NPCs-HOSTILES.dat"
    Else
        npcfile = DatPath & "NPCs.dat"
    End If


    NpcIndex = NextOpenNPC

    If NpcIndex > MAXNPCS Then
        OpenNPC_Viejo = NpcIndex
        Exit Function
    End If

    Npclist(NpcIndex).Numero = NPCNumber
    Npclist(NpcIndex).name = GetVar(npcfile, "NPC" & NPCNumber, "Name")
    Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NPCNumber, "Desc")

    Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NPCNumber, "Movement"))
    Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

    Npclist(NpcIndex).flags.AguaValida = val(GetVar(npcfile, "NPC" & NPCNumber, "AguaValida"))
    Npclist(NpcIndex).flags.TierraInvalida = val(GetVar(npcfile, "NPC" & NPCNumber, "TierraInValida"))
    Npclist(NpcIndex).flags.Faccion = val(GetVar(npcfile, "NPC" & NPCNumber, "Faccion"))

    Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NPCNumber, "NpcType"))

    Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NPCNumber, "Body"))
    Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NPCNumber, "Head"))
    Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NPCNumber, "Heading"))

    Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NPCNumber, "Attackable"))
    Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NPCNumber, "Comercia"))
    Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NPCNumber, "Hostile"))
    Npclist(NpcIndex).InmuneParalisis = val(GetVar(npcfile, "NPC" & NPCNumber, "InmuneParalisis"))
    Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile


    Npclist(NpcIndex).MaxRecom = val(GetVar(npcfile, "NPC" & NPCNumber, "MaxRecom"))
    Npclist(NpcIndex).MinRecom = val(GetVar(npcfile, "NPC" & NPCNumber, "MinRecom"))
    Npclist(NpcIndex).Probabilidad = val(GetVar(npcfile, "NPC" & NPCNumber, "Probabilidad"))


    Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NPCNumber, "GiveEXP"))


    Npclist(NpcIndex).Veneno = val(GetVar(npcfile, "NPC" & NPCNumber, "Veneno"))

    Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NPCNumber, "Domable"))


    Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NPCNumber, "GiveGLD"))

    Npclist(NpcIndex).PoderAtaque = val(GetVar(npcfile, "NPC" & NPCNumber, "PoderAtaque"))
    Npclist(NpcIndex).PoderEvasion = val(GetVar(npcfile, "NPC" & NPCNumber, "PoderEvasion"))

    Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NPCNumber, "InvReSpawn"))
    Npclist(NpcIndex).Bot = val(GetVar(npcfile, "NPC" & NPCNumber, "Bot"))
    Npclist(NpcIndex).AutoCurar = val(GetVar(npcfile, "NPC" & NPCNumber, "autocurar"))


    Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NPCNumber, "MaxHP"))
    Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NPCNumber, "MinHP"))
    Npclist(NpcIndex).Stats.MaxHit = val(GetVar(npcfile, "NPC" & NPCNumber, "MaxHIT"))
    Npclist(NpcIndex).Stats.MinHit = val(GetVar(npcfile, "NPC" & NPCNumber, "MinHIT"))
    Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NPCNumber, "DEF"))
    Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NPCNumber, "Alineacion"))
    Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NPCNumber, "ImpactRate"))


    Dim LoopC As Integer
    Dim ln As String
    Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NPCNumber, "NROITEMS"))
    For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
        ln = GetVar(npcfile, "NPC" & NPCNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ProbTirar = val(ReadField(3, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))

    Next

    Npclist(NpcIndex).flags.LanzaSpells = val(GetVar(npcfile, "NPC" & NPCNumber, "LanzaSpells"))
    If Npclist(NpcIndex).flags.LanzaSpells Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
    For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
        Npclist(NpcIndex).Spells(LoopC) = val(GetVar(npcfile, "NPC" & NPCNumber, "Sp" & LoopC))
    Next


    If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
        Npclist(NpcIndex).NroCriaturas = val(GetVar(npcfile, "NPC" & NPCNumber, "NroCriaturas"))
        ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
        For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
            Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = GetVar(npcfile, "NPC" & NPCNumber, "CI" & LoopC)
            Npclist(NpcIndex).Criaturas(LoopC).NpcName = GetVar(npcfile, "NPC" & NPCNumber, "CN" & LoopC)
        Next
    End If


    Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NPCNumber, "Inflacion"))

    Npclist(NpcIndex).flags.NPCActive = True
    Npclist(NpcIndex).flags.UseAINow = False

    If Respawn Then
        Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NPCNumber, "ReSpawn"))
    Else
        Npclist(NpcIndex).flags.Respawn = 0
    End If

    Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NPCNumber, "OrigPos"))
    Npclist(NpcIndex).flags.AfectaParalisis = val(GetVar(npcfile, "NPC" & NPCNumber, "AfectaParalisis"))
    Npclist(NpcIndex).flags.GolpeExacto = val(GetVar(npcfile, "NPC" & NPCNumber, "GolpeExacto"))
    Npclist(NpcIndex).flags.PocaParalisis = val(GetVar(npcfile, "NPC" & NPCNumber, "PocaParalisis"))
    Npclist(NpcIndex).VeInvis = val(GetVar(npcfile, "NPC" & NPCNumber, "veinvis"))



    Npclist(NpcIndex).flags.Snd1 = val(GetVar(npcfile, "NPC" & NPCNumber, "Snd1"))
    Npclist(NpcIndex).flags.Snd2 = val(GetVar(npcfile, "NPC" & NPCNumber, "Snd2"))
    Npclist(NpcIndex).flags.Snd3 = val(GetVar(npcfile, "NPC" & NPCNumber, "Snd3"))
    Npclist(NpcIndex).flags.Snd4 = val(GetVar(npcfile, "NPC" & NPCNumber, "Snd4"))



    Dim AUX As String
    AUX = GetVar(npcfile, "NPC" & NPCNumber, "NROEXP")
    If Len(AUX) = 0 Then
        Npclist(NpcIndex).NroExpresiones = 0
    Else
        Npclist(NpcIndex).NroExpresiones = val(AUX)
        ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
        For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
            Npclist(NpcIndex).Expresiones(LoopC) = GetVar(npcfile, "NPC" & NPCNumber, "Exp" & LoopC)
        Next
    End If




    Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NPCNumber, "TipoItems"))


    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1



    OpenNPC_Viejo = NpcIndex

End Function

Sub EnviarListaCriaturas(UserIndex As Integer, NpcIndex)
    Dim SD As String
    Dim k As Integer
    SD = SD & Npclist(NpcIndex).NroCriaturas & ","
    For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
    Next
    SD = "LSTCRI" & SD
    Call SendData(ToIndex, UserIndex, 0, SD)
End Sub


Sub DoFollow(NpcIndex As Integer, UserIndex As Integer)

    If Npclist(NpcIndex).flags.Follow Then
        Npclist(NpcIndex).flags.AttackedBy = 0
        Npclist(NpcIndex).flags.Follow = False
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    Else
        Npclist(NpcIndex).flags.AttackedBy = UserIndex
        Npclist(NpcIndex).flags.Follow = True
        Npclist(NpcIndex).Movement = 4
        Npclist(NpcIndex).Hostile = 0
    End If

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).flags.Follow = True
    Npclist(NpcIndex).Movement = SIGUE_AMO
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNpc = 0

End Sub
