Attribute VB_Name = "modHechizos"
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
Sub NpcLanzaSpellSobreUser(NpcIndex As Integer, UserIndex As Integer, Spell As Integer)

    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub

    If UserList(UserIndex).flags.Privilegios Then Exit Sub

    Npclist(NpcIndex).CanAttack = 0
    Dim Daño As Integer

    If Hechizos(Spell).SubeHP = 1 Then
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

        Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + Daño
        If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP

        Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, Resistencia)
    ElseIf Hechizos(Spell).SubeHP = 2 Then
        Daño = (RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)) * 3

        If Npclist(NpcIndex).MaestroUser = 0 Then Daño = Daño * (1 - UserList(UserIndex).Stats.UserSkills(Resistencia) / 200)

        If UserList(UserIndex).Invent.CascoEqpObjIndex Then
            Dim Obj As ObjData
            Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
            If Obj.Gorro = 1 Then
                Dim absorbido As Integer
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                absorbido = absorbido
                Daño = Maximo(1, Daño - absorbido)
            End If
        End If

        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TX" & Npclist(NpcIndex).Char.CharIndex & "," & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).EffectIndex & "," & Hechizos(Spell).loops & "," & Hechizos(Spell).WAV)

        If Not UserList(UserIndex).flags.Quest And UserList(UserIndex).flags.Privilegios = 0 Then
            UserList(UserIndex).Stats.MinHP = Maximo(0, UserList(UserIndex).Stats.MinHP - Daño)
            Call SendUserHP(UserIndex)
        End If

        Call SendData(ToIndex, UserIndex, 0, "%A" & Npclist(NpcIndex).name & "," & Daño)
        Call SubirSkill(UserIndex, Resistencia)

        If UserList(UserIndex).Stats.MinHP = 0 Then Call UserDie(UserIndex)

    End If

    If Hechizos(Spell).Paraliza > 0 Then
        If UserList(UserIndex).flags.Paralizado = 0 Then
            If UserList(UserIndex).Clase = PIRATA And UserList(UserIndex).Recompensas(3) = 1 Then Exit Sub
            UserList(UserIndex).flags.Paralizado = 1
            UserList(UserIndex).Counters.Paralisis = Timer - 15 * (UserList(UserIndex).Clase = GUERRERO And UserList(UserIndex).Recompensas(3))
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
            Call SendData(ToIndex, UserIndex, 0, ("P9"))
            Call SendData(ToIndex, UserIndex, 0, "PU" & DesteEncripTE(UserList(UserIndex).pos.X & "," & UserList(UserIndex).pos.Y))
        End If
    End If

    If Hechizos(Spell).Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 1
        UserList(UserIndex).Counters.Ceguera = Timer
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        Call SendData(ToIndex, UserIndex, 0, "CEGU")
        Call SendData(ToIndex, UserIndex, 0, "%B")
    End If

    If Hechizos(Spell).RemoverParalisis = 1 Then
        If Npclist(NpcIndex).flags.Paralizado Then
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
        End If
    End If

    Call DecirPalabrasMagicas(Hechizos(Spell).PalabrasMagicas, UserIndex, True, NpcIndex)

End Sub
Function TieneHechizo(ByVal i As Integer, UserIndex As Integer) As Boolean

    On Error GoTo errhandler

    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

    Exit Function
errhandler:

End Function
Sub AgregarHechizo(UserIndex As Integer, Slot As Byte)
    Dim hIndex As Integer, j As Integer

    hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).OBJIndex).HechizoIndex

    If Not TieneHechizo(hIndex, UserIndex) Then
        For j = 1 To MAXUSERHECHIZOS
            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next

        If UserList(UserIndex).Stats.UserHechizos(j) Then
            Call SendData(ToIndex, UserIndex, 0, "%C")
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))

            Call QuitarUnItem(UserIndex, CByte(Slot))
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "%D")
    End If

End Sub
Sub Aprenderhechizo(UserIndex As Integer, ByVal hechizoespecial As Integer)
    Dim hIndex As Integer
    Dim j As Integer
    hIndex = hechizoespecial

    If Not TieneHechizo(hIndex, UserIndex) Then

        For j = 1 To MAXUSERHECHIZOS
            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next

        If UserList(UserIndex).Stats.UserHechizos(j) Then
            Call SendData(ToIndex, UserIndex, 0, "%C")
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))

        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "%D")
    End If

End Sub
Sub DecirPalabrasMagicas(ByVal S As String, UserIndex As Integer, Optional EsNpc As Boolean = False, Optional NpcIndex As Integer)
    On Error GoTo errhandler
    If EsNpc = False Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbCyan & "°" & S & "°" & UserList(UserIndex).Char.CharIndex)
    Else
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbGreen & "°" & S & "°" & str(Npclist(NpcIndex).Char.CharIndex))
    End If

errhandler:
    If Err.Number <> 0 Then
        LogError ("Error en Sub DecirPalabrasMagicas - " & S & " - UI:" & UserIndex & " - EsNpc:" & EsNpc & " - NpcI:" & NpcIndex & " - Description:" & Err.Description)
    End If
End Sub
Function ManaHechizo(UserIndex As Integer, Hechizo As Integer) As Integer

    If UserList(UserIndex).flags.Privilegios > 2 Or UserList(UserIndex).flags.Quest Then Exit Function

    If UserList(UserIndex).Recompensas(3) = 1 And _
       ((UserList(UserIndex).Clase = DRUIDA And Hechizo = 24) Or _
        (UserList(UserIndex).Clase = PALADIN And Hechizo = 10)) Then
        ManaHechizo = 250
    ElseIf UserList(UserIndex).Clase = CLERIGO And UserList(UserIndex).Recompensas(3) = 2 And Hechizo = 11 Then
        ManaHechizo = 1100
    Else: ManaHechizo = Hechizos(Hechizo).ManaRequerido
    End If

End Function
Function PuedeLanzar(UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
    Dim wp2 As WorldPos

    wp2.Map = UserList(UserIndex).flags.TargetMap
    wp2.X = UserList(UserIndex).flags.TargetX
    wp2.Y = UserList(UserIndex).flags.TargetY

    If Not EnPantalla(UserList(UserIndex).pos, wp2, 1) Then Exit Function

    If UserList(UserIndex).flags.Muerto Then
        Call SendData(ToIndex, UserIndex, 0, "MU")
        Exit Function
    End If

    If MapInfo(UserList(UserIndex).pos.Map).NoMagia Then
        Call SendData(ToIndex, UserIndex, 0, "/T")
        Exit Function
    End If

    If UserList(UserIndex).Stats.ELV < Hechizos(HechizoIndex).Nivel Then
        Call SendData(ToIndex, UserIndex, 0, "%%" & Hechizos(HechizoIndex).Nivel)
        Exit Function
    End If

    If UserList(UserIndex).Stats.UserSkills(Magia) < Hechizos(HechizoIndex).MinSkill Then
        Call SendData(ToIndex, UserIndex, 0, "%E")
        Exit Function
    End If

    If UserList(UserIndex).Stats.MinMAN < ManaHechizo(UserIndex, HechizoIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "%F")
        Exit Function
    End If

    If UserList(UserIndex).Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
        Call SendData(ToIndex, UserIndex, 0, "9C")
        Exit Function
    End If

    PuedeLanzar = True

End Function
Sub HechizoInvocacion(UserIndex As Integer, b As Boolean)
    Dim Masc As Integer

    If UserList(UserIndex).pos.Map = 60 Or UserList(UserIndex).pos.Map = 205 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 66 Or UserList(UserIndex).pos.Map = 22 Or UserList(UserIndex).pos.Map = 7 Or UserList(UserIndex).pos.Map = 19 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes invocar criaturas en este mapa" & FONTTYPE_INFO)    ' Brilcair!
        Exit Sub
    End If

    If Not MapInfo(UserList(UserIndex).pos.Map).Pk Then
        Call SendData(ToIndex, UserIndex, 0, "A&")
        Exit Sub
    End If

    If Not UserList(UserIndex).flags.Quest And UserList(UserIndex).NroMascotas >= 3 Then Exit Sub
    If UserList(UserIndex).NroMascotas >= MAXMASCOTAS Then Exit Sub

    Dim h As Integer, j As Integer, ind As Integer, Index As Integer
    Dim TargetPos As WorldPos

    TargetPos.Map = UserList(UserIndex).flags.TargetMap
    TargetPos.X = UserList(UserIndex).flags.TargetX
    TargetPos.Y = UserList(UserIndex).flags.TargetY

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(j) Then
            If Npclist(UserList(UserIndex).MascotasIndex(j)).Numero = Hechizos(h).NumNPC Then Masc = Masc + 1
        End If
    Next

    If (Hechizos(h).NumNPC = 103 And Masc >= 2 And Not UserList(UserIndex).flags.Quest) Or (Hechizos(h).NumNPC = 94 And Masc >= 1) Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes invocar más mascotas de este tipo." & FONTTYPE_FIGHT)
        Exit Sub
    End If

    For j = 1 To Hechizos(h).Cant
        If (UserList(UserIndex).NroMascotas < 3 Or UserList(UserIndex).flags.Quest) And UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
            ind = SpawnNpc(Hechizos(h).NumNPC, TargetPos, True, False)
            If ind < MAXNPCS Then

                UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1

                Index = FreeMascotaIndex(UserIndex)

                UserList(UserIndex).MascotasIndex(Index) = ind
                UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero

                If UserList(UserIndex).Clase = DRUIDA And UserList(UserIndex).Recompensas(3) = 2 Then
                    If Hechizos(h).NumNPC >= 92 And Hechizos(h).NumNPC <= 94 Then
                        Npclist(ind).Stats.MaxHP = Npclist(ind).Stats.MaxHP + 75
                        Npclist(ind).Stats.MinHP = Npclist(ind).Stats.MaxHP
                    End If
                End If

                If Npclist(ind).Numero = 103 And UserList(UserIndex).Raza <> ELFO_OSCURO Then
                    Npclist(ind).Stats.MaxHP = Npclist(ind).Stats.MaxHP - 200
                    Npclist(ind).Stats.MinHP = Npclist(ind).Stats.MinHP - 200
                End If

                Npclist(ind).MaestroUser = UserIndex
                Npclist(ind).Contadores.TiempoExistencia = Timer
                Npclist(ind).GiveGLD = 0

                Call FollowAmo(ind)
            End If
        Else: Exit For
        End If
    Next

    Call InfoHechizo(UserIndex)
    b = True

End Sub
Sub HechizoTerrenoEstado(UserIndex As Integer, b As Boolean)
    Dim PosCasteada As WorldPos
    Dim TU As Integer
    Dim h As Integer
    Dim i As Integer

    PosCasteada.X = UserList(UserIndex).flags.TargetX
    PosCasteada.Y = UserList(UserIndex).flags.TargetY
    PosCasteada.Map = UserList(UserIndex).flags.TargetMap

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    If Hechizos(h).Invisibilidad = 2 Then
        For i = 1 To MapInfo(UserList(UserIndex).pos.Map).NumUsers
            TU = MapInfo(UserList(UserIndex).pos.Map).UserIndex(i)
            If EnPantalla(PosCasteada, UserList(TU).pos, -1) And UserList(TU).flags.Invisible = 1 And UserList(TU).flags.AdminInvisible = 0 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(TU).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & 0 & "," & Hechizos(h).loops)
            End If
        Next
        b = True
    End If

    Call InfoHechizo(UserIndex)

End Sub
Sub HandleHechizoTerreno(UserIndex As Integer, ByVal uh As Integer)
    Dim b As Boolean

    Select Case Hechizos(uh).Tipo
        Case uInvocacion
            Call HechizoInvocacion(UserIndex, b)
        Case uRadial
            Call HechizoTerrenoEstado(UserIndex, b)
        Case uMaterializa    'matute
            Call HechizoMaterializar(UserIndex, b)
    End Select

    If b Then
        Call SubirSkill(UserIndex, Magia)
        Call QuitarSta(UserIndex, Hechizos(uh).StaRequerido)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizo(UserIndex, uh)
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call SendUserMANASTA(UserIndex)
    End If

End Sub
Sub HandleHechizoUsuario(UserIndex As Integer, ByVal uh As Integer)
    Dim b As Boolean
    Dim tempChr As Integer
    Dim TU, tN As Integer

    tempChr = UserList(UserIndex).flags.TargetUser

    If UserList(tempChr).flags.Protegido = 1 Or UserList(tempChr).flags.Protegido = 2 Then Exit Sub

    Select Case Hechizos(uh).Tipo
        Case uTerreno
            Call HechizoInvocacion(UserIndex, b)
        Case uEstado
            Call HechizoEstadoUsuario(UserIndex, b)
        Case uPropiedades
            Call HechizoPropUsuario(UserIndex, b)
    End Select

    If b Then
        Call SubirSkill(UserIndex, Magia)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizo(UserIndex, uh)
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call QuitarSta(UserIndex, Hechizos(uh).StaRequerido)
        Call SendUserMANASTA(UserIndex)
        Call SendUserHPSTA(UserList(UserIndex).flags.TargetUser)
        UserList(UserIndex).flags.TargetUser = 0
    End If

End Sub
Sub HandleHechizoNPC(UserIndex As Integer, ByVal uh As Integer)
    Dim b As Boolean

    If Npclist(UserList(UserIndex).flags.TargetNpc).flags.NoMagia = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "/U")
        Exit Sub
    End If

    If UserList(UserIndex).flags.Protegido > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar NPC's mientrás estas siendo protegido." & FONTTYPE_FIGHT)
        Exit Sub
    End If

    Select Case Hechizos(uh).Tipo
        Case uEstado
            Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNpc, uh, b, UserIndex)
        Case uPropiedades
            Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNpc, UserIndex, b)
    End Select

    If b Then
        Call SubirSkill(UserIndex, Magia)
        UserList(UserIndex).flags.TargetNpc = 0
        Call QuitarSta(UserIndex, Hechizos(uh).StaRequerido)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizo(UserIndex, uh)
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call SendUserMANASTA(UserIndex)
    End If

End Sub
Sub LanzarHechizo(Index As Integer, UserIndex As Integer)
    Dim uh As Integer
    Dim exito As Boolean

    If UserList(UserIndex).flags.Protegido = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podés tirar hechizos mientras estás siendo protegido por un GM." & FONTTYPE_FIGHT)
        Exit Sub
    ElseIf UserList(UserIndex).flags.Protegido = 2 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podés tirar hechizos tan pronto al conectarte." & FONTTYPE_FIGHT)
        Exit Sub
    End If

    uh = UserList(UserIndex).Stats.UserHechizos(Index)

    If (UserList(UserIndex).pos.Map = 148 Or UserList(UserIndex).pos.Map = 150) And (Hechizos(uh).Invoca > 0 Or Hechizos(uh).SubeHP = 2 Or Hechizos(uh).Invisibilidad = 1 Or Hechizos(uh).Paraliza > 0 Or Hechizos(uh).Estupidez = 1) Then
        Call SendData(ToIndex, UserIndex, 0, "||Una extraña energía te impide lanzar este hechizo..." & FONTTYPE_INFO)
        Exit Sub
    End If

    If TiempoTranscurrido(UserList(UserIndex).Counters.LastHechizo) < IntervaloUserPuedeCastear Then Exit Sub
    If TiempoTranscurrido(UserList(UserIndex).Counters.LastGolpe) < IntervaloUserPuedeGolpeHechi Then Exit Sub
    UserList(UserIndex).Counters.LastHechizo = Timer
    Call SendData(ToIndex, UserIndex, 0, "LH")

    If Hechizos(uh).Baculo > 0 And (UserList(UserIndex).Clase = DRUIDA Or UserList(UserIndex).Clase = MAGO Or UserList(UserIndex).Clase = NIGROMANTE) Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo < Hechizos(uh).Baculo Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "BN")
            Else: Call SendData(ToIndex, UserIndex, 0, "||Debes equiparte un báculo de mayor rango para lanzar este hechizo." & FONTTYPE_INFO)
            End If
            Exit Sub
        End If
    End If

    If PuedeLanzar(UserIndex, uh) Then
        Select Case Hechizos(uh).Target

            Case uUsuarios
                If UserList(UserIndex).flags.TargetUser Then
                    If UserList(UserList(UserIndex).flags.TargetUser).pos.Y - UserList(UserIndex).pos.Y >= 7 Then
                        Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos para lanzar ese hechizo." & FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
                End If

            Case uNPC
                If UserList(UserIndex).flags.TargetNpc Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
                End If

            Case uUsuariosYnpc
                If UserList(UserIndex).flags.TargetUser Then
                    If UserList(UserList(UserIndex).flags.TargetUser).pos.Y - UserList(UserIndex).pos.Y >= 7 Then
                        Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos para lanzar ese hechizo." & FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                    Call HandleHechizoUsuario(UserIndex, uh)
                ElseIf UserList(UserIndex).flags.TargetNpc Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Target invalido." & FONTTYPE_INFO)
                End If

            Case uTerreno
                Call HandleHechizoTerreno(UserIndex, uh)

            Case uArea
                Call HandleHechizoArea(UserIndex, uh)

        End Select
    End If

End Sub
Sub HandleHechizoArea(UserIndex As Integer, ByVal uh As Integer)
    On Error GoTo Error
    Dim TargetPos As WorldPos
    Dim X2 As Integer, Y2 As Integer
    Dim UI As Integer
    Dim b As Boolean

    TargetPos.Map = UserList(UserIndex).flags.TargetMap
    TargetPos.X = UserList(UserIndex).flags.TargetX
    TargetPos.Y = UserList(UserIndex).flags.TargetY

    For X2 = TargetPos.X - Hechizos(uh).RadioX To TargetPos.X + Hechizos(uh).RadioX
        For Y2 = TargetPos.Y - Hechizos(uh).RadioY To TargetPos.Y + Hechizos(uh).RadioY
            UI = MapData(TargetPos.Map, X2, Y2).UserIndex
            If UI > 0 Then
                UserList(UserIndex).flags.TargetUser = UI
                Select Case Hechizos(uh).Tipo
                    Case uEstado
                        Call HechizoEstadoUsuario(UserIndex, b)
                    Case uPropiedades
                        Call HechizoPropUsuario(UserIndex, b)
                End Select
            End If
        Next
    Next

    If b Then
        Call SubirSkill(UserIndex, Magia)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizo(UserIndex, uh)
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Call QuitarSta(UserIndex, Hechizos(uh).StaRequerido)
        Call SendUserMANASTA(UserIndex)
        UserList(UserIndex).flags.TargetUser = 0
    End If

    Exit Sub
Error:
    Call LogError("Error en HandleHechizoArea")
End Sub
Public Function Amigos(UserIndex As Integer, UI As Integer) As Boolean

    Amigos = (((UserList(UserIndex).Faccion.Bando = UserList(UI).Faccion.Bando) Or (EsNewbie(UI)) Or (EsNewbie(UserIndex)))) Or (UserList(UserIndex).pos.Map = 190) Or (UserList(UserIndex).Faccion.Bando = Neutral)

End Function
Sub HechizoEstadoUsuario(UserIndex As Integer, b As Boolean)
    Dim h As Integer, TU As Integer, HechizoBueno As Boolean

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    TU = UserList(UserIndex).flags.TargetUser

    HechizoBueno = Hechizos(h).RemoverParalisis Or Hechizos(h).CuraVeneno Or Hechizos(h).Invisibilidad Or Hechizos(h).Revivir Or Hechizos(h).Flecha Or Hechizos(h).Estupidez = 2 Or Hechizos(h).Transforma

    If HechizoBueno Then
        If Not Amigos(UserIndex, TU) Then
            Call SendData(ToIndex, UserIndex, 0, "2F")
            Exit Sub
        End If
    Else
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserList(UserIndex).flags.Invisible Then
            Call QuitarInvisible(UserIndex)
            Call SendData(ToIndex, UserIndex, 0, "||Has sido descubierto atacando invisible!" & FONTTYPE_FIGHT)
            Exit Sub
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
    End If

    If Hechizos(h).Envenena Then
        UserList(TU).flags.Envenenado = Hechizos(h).Envenena
        UserList(TU).flags.EstasEnvenenado = Timer
        UserList(TU).Counters.Veneno = Timer
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If

    If Hechizos(h).Maldicion = 1 Then
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If

    If Hechizos(h).Paraliza > 0 Then
        If UserList(TU).flags.Paralizado = 0 Then
            If (UserList(TU).Clase = MINERO And UserList(TU).Recompensas(2) = 1) Or (UserList(TU).Clase = PIRATA And UserList(TU).Recompensas(3) = 1) Then
                Call SendData(ToIndex, UserIndex, 0, "%&")
                Exit Sub
            End If

            UserList(TU).flags.QuienParalizo = UserIndex
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = Timer - 15 * Buleano(UserList(TU).Clase = GUERRERO And UserList(TU).Recompensas(3) = 2)
            Call SendData(ToIndex, TU, 0, "PU" & DesteEncripTE(UserList(TU).pos.X & "," & UserList(TU).pos.Y))
            Call SendData(ToIndex, TU, 0, ("P9"))
            Call InfoHechizo(UserIndex)
            b = True
            Exit Sub
        End If
    End If

    If Hechizos(h).Ceguera = 1 Then
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = Timer
        Call SendData(ToIndex, TU, 0, "CEGU")
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If

    If Hechizos(h).Estupidez = 1 Then
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Estupidez = Timer
        Call SendData(ToIndex, TU, 0, "DUMB")
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If

    If Hechizos(h).Transforma = 1 Then
        If UserList(TU).flags.Transformado = 0 Then
            If UserList(TU).Stats.ELV > 39 And UserList(TU).Raza = ELFO And UserList(TU).Clase = DRUIDA Then
                Call DoMetamorfosis(UserIndex)
            Else
                Call SendData(ToIndex, UserIndex, 0, "{E")
            End If
            Call InfoHechizo(UserIndex)
            b = True
            Exit Sub
        End If
    End If

    If Hechizos(h).Revivir = 1 Then
        If UserList(TU).flags.Muerto Then
            Call RevivirUsuario(UserIndex, TU, UserList(UserIndex).Clase = CLERIGO And UserList(UserIndex).Recompensas(3) = 2)
            Call InfoHechizo(UserIndex)
            b = True
            Exit Sub
        End If
    End If

    If UserList(TU).flags.Muerto Then
        Call SendData(ToIndex, UserIndex, 0, "8C")
        Exit Sub
    End If

    If Hechizos(h).Estupidez = 2 Then
        If UserList(TU).flags.Estupidez = 1 Then
            UserList(TU).flags.Estupidez = 0
            UserList(TU).Counters.Estupidez = 0
            Call SendData(ToIndex, TU, 0, "NESTUP")
            Call InfoHechizo(UserIndex)
            b = True
            Exit Sub
        End If
    End If

    If Hechizos(h).Flecha = 1 Then
        If TU <> UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "||Este hechizo solo puedes usarlo sobre ti mismo." & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(TU).flags.BonusFlecha = True
        UserList(TU).Counters.BonusFlecha = Timer
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If

    If Hechizos(h).RemoverParalisis = 1 Then
        If UserList(TU).flags.Paralizado Then
            Call SendData(ToIndex, TU, 0, "P8")
            UserList(TU).flags.Paralizado = 0
            UserList(TU).flags.QuienParalizo = 0
            Call InfoHechizo(UserIndex)
            b = True
            Exit Sub
        End If
    End If

    If Hechizos(h).Invisibilidad = 1 Then
        If UserList(UserIndex).pos.Map = 3 Or UserList(UserIndex).pos.Map = 4 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 6 Or UserList(UserIndex).pos.Map = 7 Or UserList(UserIndex).pos.Map = 205 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 190 Or UserList(UserIndex).pos.Map = 22 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes hacerte invisible en este mapa" & FONTTYPE_INFO)    'BrIlCaIR!
            Exit Sub
        End If
        If UserList(TU).flags.Invisible Then Exit Sub
        UserList(TU).flags.Invisible = 1
        UserList(TU).Counters.Invisibilidad = Timer
        Call SendData(ToMap, 0, UserList(TU).pos.Map, ("V3" & DesteEncripTE(UserList(TU).Char.CharIndex & ",1")))
        Call SendData(ToIndex, UserIndex, 0, "INVI")
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If

    If Hechizos(h).CuraVeneno = 1 Then
        If UserList(TU).flags.Envenenado = 1 Then
            UserList(TU).flags.Envenenado = 0
            Call InfoHechizo(UserIndex)
            b = True
            Exit Sub
        Else
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no está envenenado." & FONTTYPE_FIGHT)
            Exit Sub
        End If
    End If

    If Hechizos(h).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If

    If Hechizos(h).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If

End Sub
Sub HechizoEstadoNPC(NpcIndex As Integer, ByVal hIndex As Integer, b As Boolean, UserIndex As Integer)

    If Npclist(NpcIndex).Attackable = 0 Then Exit Sub

    If Hechizos(hIndex).Invisibilidad = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Invisible = 1
        b = True
    End If

    If Hechizos(hIndex).Envenena = 1 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "NO")
            Exit Sub
        End If
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Envenenado = 1
        b = True
    End If

    If Hechizos(hIndex).CuraVeneno = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Envenenado = 0
        b = True
    End If

    If Hechizos(hIndex).Maldicion = 1 Then
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "NO")
            Exit Sub
        End If
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Maldicion = 1
        b = True
    End If

    If Hechizos(hIndex).RemoverMaldicion = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Maldicion = 0
        b = True
    End If

    If Hechizos(hIndex).Bendicion = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Bendicion = 1
        b = True
    End If

    If Hechizos(hIndex).Paraliza Then
        If Npclist(NpcIndex).InmuneParalisis = 1 Then Call SendData(ToIndex, UserIndex, 0, "||Esta criatura es inmune a la paralisis." & FONTTYPE_FIGHT): Exit Sub
        If Npclist(NpcIndex).flags.QuienParalizo <> 0 And Npclist(NpcIndex).flags.QuienParalizo <> UserIndex Then Exit Sub
        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = Hechizos(hIndex).Paraliza
            Npclist(NpcIndex).flags.QuienParalizo = UserIndex
            If Npclist(NpcIndex).flags.PocaParalisis = 1 Then
                Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado / 4
            Else: Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            End If
            b = True
        Else: Call SendData(ToIndex, UserIndex, 0, "7D")
        End If
    End If

    If Hechizos(hIndex).RemoverParalisis = 1 Then
        If Npclist(NpcIndex).flags.QuienParalizo = UserIndex Or Npclist(NpcIndex).MaestroUser = UserIndex Then
            If Npclist(NpcIndex).flags.Paralizado Then
                Call InfoHechizo(UserIndex)
                Npclist(NpcIndex).flags.Paralizado = 0
                Npclist(NpcIndex).Contadores.Paralisis = 0
                Npclist(NpcIndex).flags.QuienParalizo = 0
                b = True
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "8D")
        End If
    End If

End Sub
Sub VerNPCMuere(ByVal NpcIndex As Integer, ByVal Daño As Long, ByVal UserIndex As Integer)

    If Npclist(NpcIndex).Numero = 245 Then Exit Sub

    If Npclist(NpcIndex).AutoCurar = 0 Then Npclist(NpcIndex).Stats.MinHP = Maximo(0, Npclist(NpcIndex).Stats.MinHP - Daño)

    If Npclist(NpcIndex).Stats.MinHP <= 0 Then
        If Npclist(NpcIndex).flags.Snd3 Then Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).pos.Map, "TW" & Npclist(NpcIndex).flags.Snd3 & "," & Npclist(NpcIndex).pos.X & "," & Npclist(NpcIndex).pos.Y)

        If UserIndex Then
            If UserList(UserIndex).NroMascotas Then
                Dim T As Integer
                For T = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
                    If UserList(UserIndex).MascotasIndex(T) Then
                        If Npclist(UserList(UserIndex).MascotasIndex(T)).TargetNpc = NpcIndex Then Call FollowAmo(UserList(UserIndex).MascotasIndex(T))
                    End If
                Next
            End If
            Call AddtoVar(UserList(UserIndex).Stats.NPCsMuertos, 1, 32000)

            UserList(UserIndex).flags.TargetNpc = 0
            UserList(UserIndex).flags.TargetNpcTipo = 0
        End If

        Call MuereNpc(NpcIndex, UserIndex)
    End If

End Sub
Sub ExperienciaPorGolpe(UserIndex As Integer, ByVal NpcIndex As Integer, Daño As Integer)
    Dim ExpDada As Double
    Daño = Minimo(Daño, Npclist(NpcIndex).Stats.MinHP)

    If Npclist(NpcIndex).Numero = 245 Then
        ExpDada = 0
    Else
        ExpDada = ((Npclist(NpcIndex).GiveEXP) * (Daño / 100)) * CantidadEXP
    End If

    'If Daño >= Npclist(NpcIndex).Stats.MinHP Then ExpDada = ExpDada + Npclist(NpcIndex).GiveEXP / 2

    If ModoQuest Then ExpDada = ExpDada / 2

    If UserList(UserIndex).flags.EsNoble = 1 Then
        ExpDada = ExpDada * 2
    End If

    If UserList(UserIndex).flags.Party = 0 Then
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpDada
        If Daño >= Npclist(NpcIndex).Stats.MinHP Then
            Call SendData(ToIndex, UserIndex, 0, "EL" & ExpDada)
        Else: Call SendData(ToIndex, UserIndex, 0, "EX" & ExpDada)
        End If
        Call SendUserEXP(UserIndex)
        Call CheckUserLevel(UserIndex)
        Exit Sub
    Else: Call RepartirExp(UserIndex, ExpDada, Daño >= Npclist(NpcIndex).Stats.MinHP)
    End If

End Sub
Sub HechizoPropNPC(ByVal hIndex As Integer, NpcIndex As Integer, UserIndex As Integer, b As Boolean)
    Dim Daño As Integer

    If Npclist(NpcIndex).Attackable = 0 Then Exit Sub
    If Hechizos(hIndex).SubeHP = 1 Then
        Daño = DañoHechizo(UserIndex, hIndex)

        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbGreen & "°+" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))
        Call InfoHechizo(UserIndex)
        Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, Daño, Npclist(NpcIndex).Stats.MaxHP)
        Call SendData(ToIndex, UserIndex, 0, "CU" & Daño)
        b = True
    ElseIf Hechizos(hIndex).SubeHP = 2 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).name = "Báculo de los Dioses" Then
            Daño = 1.3 * Daño
        ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo = Hechizos(hIndex).Baculo Then
            Daño = 0.95 * Daño
        End If
        If Npclist(NpcIndex).Numero = 244 Then
            If MataGuardiasEmperador <> 3 Then
                Call SendData(ToIndex, UserIndex, 0, "||Primero debes acabar con mis Comandantes y mi general antes de querer destriuirme." & FONTTYPE_FIGHT)
                Exit Sub
            End If
        End If
        If UCase$(Npclist(NpcIndex).name) = "REY DEL CASTILLO" Then
            If (Npclist(NpcIndex).pos.Map = mapa_castilloNorte) Then
                If UserList(UserIndex).GuildInfo.GuildName = "" Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡No tienes clan!" & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                If UserList(UserIndex).GuildInfo.GuildName = CastilloNorte Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡No puedes atacar tu castillo!" & FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If
        End If
        If UCase$(Npclist(NpcIndex).name) = "REY DEL CASTILLO" Then
            If (Npclist(NpcIndex).pos.Map = mapa_castilloSur) Then
                If UserList(UserIndex).GuildInfo.GuildName = "" Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡No tienes clan!" & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                If UserList(UserIndex).GuildInfo.GuildName = CastilloSur Then
                    Call SendData(ToIndex, UserIndex, 0, "||¡No puedes atacar tu castillo!" & FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If
        End If
        If Npclist(NpcIndex).Attackable = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "NO")
            Exit Sub
        End If

        If UserList(UserIndex).Faccion.Bando <> Neutral And Npclist(NpcIndex).MaestroUser Then
            If Not PuedeAtacarMascota(UserIndex, (Npclist(NpcIndex).MaestroUser)) Then Exit Sub
        End If

        If UserList(UserIndex).Faccion.Bando <> Neutral And UserList(UserIndex).Faccion.Bando = Npclist(NpcIndex).flags.Faccion Then
            Call SendData(ToIndex, UserIndex, 0, Mensajes(Npclist(NpcIndex).flags.Faccion, 19))
            Exit Sub
        End If

        Daño = DañoHechizo(UserIndex, hIndex)

        b = True
        Call NpcAtacado(NpcIndex, UserIndex)

        If Npclist(NpcIndex).flags.Snd2 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2 & "," & Npclist(NpcIndex).pos.X & "," & Npclist(NpcIndex).pos.Y)

        Call SendData(ToIndex, UserIndex, 0, "X2" & Daño)

        Call ExperienciaPorGolpe(UserIndex, NpcIndex, Daño)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbRed & "°-" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))

        Call InfoHechizo(UserIndex)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFF" & UserList(UserIndex).Char.CharIndex & "," & Npclist(NpcIndex).Char.CharIndex)
        Call VerNPCMuere(NpcIndex, Daño, UserIndex)
    
    End If
End Sub
Sub InfoHechizo(UserIndex As Integer)
    Dim h As Integer
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, UserIndex)
    

    If UserList(UserIndex).flags.TargetUser Then
        Call SendData(ToPCArea, UserIndex, UserList(UserList(UserIndex).flags.TargetUser).pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).EffectIndex & "," & Hechizos(h).loops)
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(h).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).name & FONTTYPE_ATACO)
            Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).name & " " & Hechizos(h).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(h).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).flags.TargetNpc Then
        Call SendData(ToPCArea, UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).EffectIndex & "," & Hechizos(h).loops)
        Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(h).HechizeroMsg & " " & "la criatura." & FONTTYPE_ATACO)
    End If
    
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(h).WAV & "," & UserList(UserIndex).flags.TargetX & "," & UserList(UserIndex).flags.TargetY)
End Sub
Function DañoHechizo(UserIndex As Integer, Hechizo As Integer) As Integer

    DañoHechizo = RandomNumber(Hechizos(Hechizo).MinHP + 5 * Buleano(UserList(UserIndex).Clase = BARDO And UserList(UserIndex).Recompensas(3) = 2 And (Hechizo = 23 Or Hechizo = 25)) _
                               + 10 * Buleano(UserList(UserIndex).Clase = NIGROMANTE And UserList(UserIndex).Recompensas(3) = 1) _
                               + 20 * Buleano(UserList(UserIndex).Clase = CLERIGO And UserList(UserIndex).Recompensas(3) = 1 And Hechizo = 5) _
                               + 10 * Buleano(UserList(UserIndex).Clase = MAGO And UserList(UserIndex).Recompensas(3) = 2 And Hechizo = 25), _
                               Hechizos(Hechizo).MaxHP + 5 * Buleano(UserList(UserIndex).Clase = BARDO And UserList(UserIndex).Recompensas(3) = 2 And (Hechizo = 23 Or Hechizo = 25)) _
                               + 20 * Buleano(UserList(UserIndex).Clase = CLERIGO And UserList(UserIndex).Recompensas(3) = 1 And Hechizo = 5) _
                               + 10 * Buleano(UserList(UserIndex).Clase = MAGO And UserList(UserIndex).Recompensas(3) = 1 And Hechizo = 23))
    If UserList(UserIndex).Stats.ELV <= 50 Then
        DañoHechizo = DañoHechizo + Porcentaje(DañoHechizo, 3 * UserList(UserIndex).Stats.ELV)
    Else
        DañoHechizo = DañoHechizo + Porcentaje(DañoHechizo, 3 * 50)
    End If

End Function
Sub HechizoPropUsuario(UserIndex As Integer, b As Boolean)
    Dim h As Integer
    Dim Daño As Integer
    Dim tempChr As Integer
    Dim reducido As Integer
    Dim HechizoBueno As Boolean
    Dim msg As String

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tempChr = UserList(UserIndex).flags.TargetUser

    HechizoBueno = Hechizos(h).SubeHam = 1 Or Hechizos(h).SubeSed = 1 Or Hechizos(h).SubeHP = 1 Or Hechizos(h).SubeAgilidad = 1 Or Hechizos(h).SubeFuerza = 1 Or Hechizos(h).SubeFuerza = 3 Or Hechizos(h).SubeMana = 1 Or Hechizos(h).SubeSta = 1

    If HechizoBueno And Not Amigos(UserIndex, tempChr) Then
        Call SendData(ToIndex, UserIndex, 0, "2F")
        Exit Sub
    ElseIf Not HechizoBueno Then
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        If UserList(UserIndex).flags.Invisible Then Call BajarInvisible(UserIndex)
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If

    If Hechizos(h).Revivir = 0 And UserList(tempChr).flags.Muerto Then Exit Sub

    If Hechizos(h).SubeHam = 1 Then

        Daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
        Call InfoHechizo(UserIndex)

        Call AddtoVar(UserList(tempChr).Stats.MinHam, Daño, UserList(tempChr).Stats.MaxHam)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de hambre a " & UserList(tempChr).name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
        End If

        Call EnviarHyS(tempChr)
        b = True

    ElseIf Hechizos(h).SubeHam = 2 Then
        Daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
        UserList(tempChr).Stats.MinHam = Maximo(0, UserList(tempChr).Stats.MinHam - Daño)

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de hambre a " & UserList(tempChr).name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
        End If
        If UserList(tempChr).Stats.MinHam = 0 Then UserList(tempChr).flags.Hambre = 1
        Call EnviarHyS(tempChr)
        b = True
    End If


    If Hechizos(h).SubeSed = 1 Then

        Call AddtoVar(UserList(tempChr).Stats.MinAGU, Daño, UserList(tempChr).Stats.MaxAGU)
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de sed a " & UserList(tempChr).name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
        End If

        b = True

    ElseIf Hechizos(h).SubeSed = 2 Then
        Daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
        UserList(tempChr).Stats.MinAGU = Maximo(0, UserList(tempChr).Stats.MinAGU - Daño)
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de sed a " & UserList(tempChr).name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
        End If

        If UserList(tempChr).Stats.MinAGU = 0 Then UserList(tempChr).flags.Sed = 1
        b = True
    ElseIf Hechizos(h).SubeSed = 3 Then

        UserList(tempChr).Stats.MinAGU = 0
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).Stats.MinSta = 0
        UserList(tempChr).flags.Sed = 1
        UserList(tempChr).flags.Hambre = 1

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "S3" & UserList(tempChr).name)
            Call SendData(ToIndex, tempChr, 0, "S4" & UserList(UserIndex).name)
        Else
            Call SendData(ToIndex, UserIndex, 0, "S5")
        End If
        Call SendData(ToIndex, tempChr, 0, "2G")

        b = True
    End If


    If Hechizos(h).SubeAgilidad = 1 Then
        Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)

        UserList(tempChr).flags.DuracionEfecto = Timer
        Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2, MAXATRIBUTOS))
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        UserList(tempChr).flags.TomoPocion = True
        b = True

    ElseIf Hechizos(h).SubeAgilidad = 2 Then
        UserList(tempChr).flags.TomoPocion = True
        Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call RestVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, MINATRIBUTOS)
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        b = True
    ElseIf Hechizos(h).SubeAgilidad = 3 Then
        UserList(tempChr).flags.TomoPocion = True
        Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call RestVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, MINATRIBUTOS)
        Call RestVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, MINATRIBUTOS)
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        b = True
    End If


    If Hechizos(h).SubeFuerza = 1 Then
        Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call AddtoVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(fuerza) * 2, MAXATRIBUTOS))
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        UserList(tempChr).flags.TomoPocion = True
        b = True
    ElseIf Hechizos(h).SubeFuerza = 2 Then
        UserList(tempChr).flags.TomoPocion = True
        Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call RestVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, MINATRIBUTOS)
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        b = True
    ElseIf Hechizos(h).SubeFuerza = 3 Then
        Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Timer
        Call AddtoVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(fuerza) * 2, MAXATRIBUTOS))
        Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2, MAXATRIBUTOS))
        Call InfoHechizo(UserIndex)
        Call UpdateFuerzaYAg(tempChr)
        UserList(tempChr).flags.TomoPocion = True
        b = True
    End If


    If Hechizos(h).SubeHP = 1 Then
        If UserList(tempChr).flags.Muerto = 1 Then Exit Sub

        If UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP Then
            Call SendData(ToIndex, UserIndex, 0, "9D")
            Exit Sub
        End If
        Daño = DañoHechizo(UserIndex, h)

        Call AddtoVar(UserList(tempChr).Stats.MinHP, Daño, UserList(tempChr).Stats.MaxHP)
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "R3" & Daño & "," & UserList(tempChr).name)
            Call SendData(ToIndex, tempChr, 0, "R4" & UserList(UserIndex).name & "," & Daño)
        Else
            Call SendData(ToIndex, UserIndex, 0, "R5" & Daño)
        End If
        b = True
    ElseIf Hechizos(h).SubeHP = 2 Then
        Daño = DañoHechizo(UserIndex, h)

        If Hechizos(h).Baculo > 0 And (UserList(UserIndex).Clase = DRUIDA Or UserList(UserIndex).Clase = MAGO Or UserList(UserIndex).Clase = NIGROMANTE) Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo < Hechizos(h).Baculo Then
                Call SendData(ToIndex, UserIndex, 0, "BN")
                Exit Sub
            Else
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).name = "Báculo de los Dioses" Then
                    Daño = 1.1 * Daño
                ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Baculo = Hechizos(h).Baculo Then
                    Daño = 0.95 * Daño
                End If
            End If
        End If

        If UserList(tempChr).Invent.CascoEqpObjIndex Then
            Dim Obj As ObjData
            Obj = ObjData(UserList(tempChr).Invent.CascoEqpObjIndex)
            If Obj.Gorro = 1 Then Daño = Maximo(1, (1 - (RandomNumber(Obj.MinDef, Obj.MaxDef) / 100)) * Daño)
            Daño = Maximo(1, Daño)
        End If

        If Not UserList(tempChr).flags.Quest Then UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - Daño
        Call InfoHechizo(UserIndex)

        Call SendData(ToIndex, UserIndex, 0, "6B" & Daño & "," & UserList(tempChr).name)
        Call SendData(ToIndex, tempChr, 0, "7B" & Daño & "," & UserList(UserIndex).name)

        If UserList(tempChr).Stats.MinHP > 0 Then
            Call SubirSkill(tempChr, Resistencia)
            If UserList(tempChr).pos.Map = UserList(UserIndex).pos.Map Then Call SendData(ToPCArea, tempChr, UserList(tempChr).pos.Map, "CFF" & UserList(UserIndex).Char.CharIndex & "," & UserList(tempChr).Char.CharIndex)
        Else
            Call ContarMuerte(tempChr, UserIndex)
            UserList(tempChr).Stats.MinHP = 0
            Call ActStats(tempChr, UserIndex)
        End If

        b = True
    End If


    If Hechizos(h).SubeMana = 1 Then
        Call AddtoVar(UserList(tempChr).Stats.MinMAN, Daño, UserList(tempChr).Stats.MaxMAN)
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
        End If

        b = True

    ElseIf Hechizos(h).SubeMana = 2 Then

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
        End If

        UserList(tempChr).Stats.MinMAN = Maximo(0, UserList(tempChr).Stats.MinMAN - Daño)
        b = True

    End If


    If Hechizos(h).SubeSta = 1 Then
        Call AddtoVar(UserList(tempChr).Stats.MinSta, Daño, UserList(tempChr).Stats.MaxSta)

        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
        End If
        b = True
    ElseIf Hechizos(h).SubeSta = 2 Then
        Call InfoHechizo(UserIndex)

        If UserIndex <> tempChr Then
            Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
        End If
        Call QuitarSta(tempChr, Daño)
        b = True
    End If

End Sub
Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, UserIndex As Integer, Slot As Byte)
    Dim LoopC As Byte

    If Not UpdateAll Then
        If UserList(UserIndex).Stats.UserHechizos(Slot) Then
            Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
        Else
            Call ChangeUserHechizo(UserIndex, Slot, 0)
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "6H")
        For LoopC = 1 To MAXUSERHECHIZOS
            If UserList(UserIndex).Stats.UserHechizos(LoopC) Then
                Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
            End If
        Next
    End If

End Sub
Sub ChangeUserHechizo(UserIndex As Integer, Slot As Byte, ByVal Hechizo As Integer)

    UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo

    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
        Call SendData(ToIndex, UserIndex, 0, "SHS" & DesteEncripTE(Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre))
    Else
        Call SendData(ToIndex, UserIndex, 0, "SHS" & DesteEncripTE(Slot & "," & "0" & "," & "Nada"))
    End If

End Sub
Public Sub DesplazarHechizo(UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Byte)

    If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
    If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

    Dim TempHechizo As Integer

    If Dire = 1 Then
        If CualHechizo = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "%G")
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

            Call UpdateUserHechizos(False, UserIndex, CualHechizo - 1)
        End If
    Else
        If CualHechizo = MAXUSERHECHIZOS Then
            Call SendData(ToIndex, UserIndex, 0, "%G")
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

            Call UpdateUserHechizos(False, UserIndex, CualHechizo + 1)
        End If
    End If

    Call UpdateUserHechizos(False, UserIndex, CualHechizo)

End Sub

Sub HechizoMaterializar(UserIndex As Integer, b As Boolean)

    Dim TU As Integer
    Dim h As Integer
    Dim i As Integer

    Dim PosTIROTELEPORT As WorldPos    'matute

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    If Hechizos(h).Nombre = "Portal Luminoso" Then
        If MapInfo(UserList(UserIndex).pos.Map).Pk = False Or UserList(UserIndex).pos.Map = 191 And UserList(UserIndex).pos.Map = 190 Or _
           UserList(UserIndex).pos.Map = 66 Or UserList(UserIndex).pos.Map = 2 Or UserList(UserIndex).pos.Map = 3 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 205 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes invocar un portal luminoso desde aquí." & FONTTYPE_FIGHT)
            Exit Sub
        End If
        If UserList(UserIndex).flags.TiroPortalL = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Tienes un portal invocado" & FONTTYPE_FIGHT)
            Exit Sub
        End If
        'If UserList(UserIndex).Counters.TimeTeleport <> 0 Then Exit Sub 'Ya invocó.
    End If

    If Hechizos(h).Materializa = 1 Then    'matute

        'If UserList(UserIndex).flags.TiroPortalL = True Then Exit Sub

        PosTIROTELEPORT.X = UserList(UserIndex).flags.TargetX
        PosTIROTELEPORT.Y = UserList(UserIndex).flags.TargetY
        PosTIROTELEPORT.Map = UserList(UserIndex).flags.TargetMap

        UserList(UserIndex).flags.DondeTiroMap = PosTIROTELEPORT.Map
        UserList(UserIndex).flags.DondeTiroX = PosTIROTELEPORT.X
        UserList(UserIndex).flags.DondeTiroY = PosTIROTELEPORT.Y

        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).OBJInfo.OBJIndex Then Exit Sub
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).NpcIndex Then Exit Sub
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Agua Then Exit Sub
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).UserIndex Then Exit Sub
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).TileExit.Map Then Exit Sub
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then Exit Sub
        If Not MapaValido(UserList(UserIndex).pos.Map) Or Not InMapBounds(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then Exit Sub
        '    Dim ET As Obj
        '    ET.Amount = 1
        '    ET.OBJIndex = 1017 'veamos asd - [Primer FX que se ve en la imagen 1] -VER OBJ.DAT

        Call SendData(ToIndex, UserIndex, 0, "||Concentras tus energías mágicas y en 5 segundos abrirás un portal hacia Althalos." & FONTTYPE_INFO)
        '    Call MakeObj(ToMap, UserIndex, UserList(UserIndex).Pos.Map, ET, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY)
        Call SendData(ToMap, UserIndex, UserList(UserIndex).pos.Map, "CXX" & UserList(UserIndex).flags.TargetX & "," & UserList(UserIndex).flags.TargetY & "," & 1)
        b = True

        UserList(UserIndex).Counters.TimeTeleport = 0
        UserList(UserIndex).Counters.CreoTeleport = True
        UserList(UserIndex).flags.TiroPortalL = 1
    End If
    UserList(UserIndex).Stats.MinMAN = 0    'le consumimos todo el mana
    Call DecirPalabrasMagicas(Hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)).PalabrasMagicas, UserIndex)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW" & Hechizos(h).WAV & "," & UserList(UserIndex).flags.TargetX & "," & UserList(UserIndex).flags.TargetY)

End Sub
