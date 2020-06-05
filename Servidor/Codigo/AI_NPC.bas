Attribute VB_Name = "AI"
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

Public Const ESTATICO = 1
Public Const MUEVE_AL_AZAR = 2
Public Const MOVEMENT_GUARDIA = 3
Public Const NPCDEFENSA = 4
Public Const SIGUE_AMO = 8
Public Const NPC_ATACA_NPC = 9
Public Const NPC_PATHFINDING = 10
Public Sub QuitarNPCDeLista(NPCNumber As Integer, Map As Integer)
    Dim i As Integer

    For i = 1 To 10
        If MapInfo(Map).NPCsReales(i).Numero = NPCNumber Then
            MapInfo(Map).NPCsReales(i).Cantidad = MapInfo(Map).NPCsReales(i).Cantidad - 1
            If MapInfo(Map).NPCsReales(i).Cantidad = 0 Then MapInfo(Map).NPCsReales(i).Numero = 0
            Exit Sub
        End If
    Next

End Sub
Public Sub AgregarNPC(NPCNumber As Integer, Map As Integer)
    Dim i As Integer

    For i = 1 To UBound(MapInfo(Map).NPCsReales)
        If MapInfo(Map).NPCsReales(i).Numero = NPCNumber Then
            MapInfo(Map).NPCsReales(i).Cantidad = MapInfo(Map).NPCsReales(i).Cantidad + 1
            Exit Sub
        ElseIf MapInfo(Map).NPCsReales(i).Numero = 0 Then
            MapInfo(Map).NPCsReales(i).Numero = NPCNumber
            MapInfo(Map).NPCsReales(i).Cantidad = 1
            Exit Sub
        End If
    Next

End Sub
Public Function UltimoNpc(Map As Integer) As Integer
    Dim i As Integer

    For i = 1 To UBound(MapInfo(Map).NPCsTeoricos)
        If MapInfo(Map).NPCsTeoricos(i).Numero = 0 Then
            UltimoNpc = i
            Exit Function
        End If
    Next

End Function
Public Sub AgregarNPCTeorico(NPCNumber As Integer, Map As Integer)
    Dim i As Integer

    For i = 1 To 10
        If MapInfo(Map).NPCsTeoricos(i).Numero = NPCNumber Then
            MapInfo(Map).NPCsTeoricos(i).Cantidad = MapInfo(Map).NPCsTeoricos(i).Cantidad + 1
            Exit Sub
        ElseIf MapInfo(Map).NPCsTeoricos(i).Numero = 0 Then
            MapInfo(Map).NPCsTeoricos(i).Numero = NPCNumber
            MapInfo(Map).NPCsTeoricos(i).Cantidad = 1
            Exit Sub
        End If
    Next

End Sub
Public Sub NPCAtacaAI(NpcIndex As Integer)
    On Error GoTo Error
    Dim nPos As WorldPos
    Dim HeadingLoop As Byte
    Dim UI As Integer

    For HeadingLoop = NORTH To WEST
        nPos = Npclist(NpcIndex).pos
        Call HeadtoPos(HeadingLoop, nPos)
        If InMapBounds(nPos.X, nPos.Y) Then
            UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
            If UI Then
                If Perseguible(UI, NpcIndex, True) Then
                    If Npclist(NpcIndex).flags.LanzaSpells Then
                        Dim k As Integer
                        k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                        Call NpcLanzaUnSpell(NpcIndex, UI)
                    End If
                    If Npclist(NpcIndex).MaestroUser = 0 Then
                        Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop)
                        Call NpcAtacaUser(NpcIndex, UI)
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next

    If Npclist(NpcIndex).Movement <> NPC_ATACA_NPC Then Call RestoreOldMovement(NpcIndex)
    Exit Sub
Error:
    Call LogError("Error en NPCAtacaAI: " & Npclist(NpcIndex).name & " " & UserList(UI).name & " " & Err.Description)
End Sub
Public Sub NPCAtacaAlFrente(NpcIndex As Integer)
    Dim nPos As WorldPos, UI As Integer, i As Integer

    For i = 1 To MapInfo(Npclist(NpcIndex).pos.Map).NumUsers
        UI = MapInfo(Npclist(NpcIndex).pos.Map).UserIndex(i)
        If Perseguible(UI, NpcIndex, True) Then
            If AtacableEnLinea(UI, NpcIndex) Then
                If Npclist(NpcIndex).flags.LanzaSpells Then Call NpcLanzaUnSpell(NpcIndex, UI)
            End If
        End If
    Next

    nPos = Npclist(NpcIndex).pos
    Call HeadtoPos(Npclist(NpcIndex).Char.Heading, nPos)
    If InMapBounds(nPos.X, nPos.Y) Then
        UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
        If UI Then
            If Perseguible(UI, NpcIndex, True) Then
                Call NpcAtacaUser(NpcIndex, UI)
                Exit Sub
            End If
        End If
    End If

    Call RestoreOldMovement(NpcIndex)

End Sub
Function AtacableEnLinea(UserIndex As Integer, NpcIndex As Integer) As Boolean
    Dim X As Integer, Y As Integer

    Select Case Npclist(NpcIndex).Char.Heading
        Case NORTH
            AtacableEnLinea = (Npclist(NpcIndex).pos.X = UserList(UserIndex).pos.X) And MinYBorder > Npclist(NpcIndex).pos.Y - UserList(UserIndex).pos.Y And Npclist(NpcIndex).pos.Y - UserList(UserIndex).pos.Y > 0
        Case SOUTH
            AtacableEnLinea = (Npclist(NpcIndex).pos.X = UserList(UserIndex).pos.X) And MinYBorder > UserList(UserIndex).pos.Y - Npclist(NpcIndex).pos.Y And UserList(UserIndex).pos.Y - Npclist(NpcIndex).pos.Y > 0
        Case WEST
            AtacableEnLinea = (Npclist(NpcIndex).pos.Y = UserList(UserIndex).pos.Y) And MinXBorder > Npclist(NpcIndex).pos.X - UserList(UserIndex).pos.X And Npclist(NpcIndex).pos.X - UserList(UserIndex).pos.X > 0
        Case EAST
            AtacableEnLinea = (Npclist(NpcIndex).pos.Y = UserList(UserIndex).pos.Y) And MinXBorder > UserList(UserIndex).pos.X - Npclist(NpcIndex).pos.X And UserList(UserIndex).pos.X - Npclist(NpcIndex).pos.X > 0
    End Select

End Function
Public Sub HostilMalvadoAIParalizado(NpcIndex As Integer)
    Dim nPos As WorldPos
    Dim HeadingLoop As Byte
    Dim theading As Byte
    Dim Y As Integer
    Dim X As Integer
    Dim UI As Integer

    For HeadingLoop = NORTH To WEST
        nPos = Npclist(NpcIndex).pos
        Call HeadtoPos(HeadingLoop, nPos)
        If InMapBounds(nPos.X, nPos.Y) Then
            UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
            If UI Then
                If UserList(UI).flags.Muerto = 0 Then
                    Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop)
                    Call NpcAtacaUser(NpcIndex, MapData(nPos.Map, nPos.X, nPos.Y).UserIndex)
                    Exit Sub
                End If
            End If
        End If
    Next

    Call RestoreOldMovement(NpcIndex)

End Sub
Private Function HayarUser(NpcIndex As Integer) As Integer
    Dim ElegidoChar As Integer
    Dim TempChar As Integer
    Dim i As Integer


    For i = 1 To MapInfo(Npclist(NpcIndex).pos.Map).NumUsers
        TempChar = MapInfo(Npclist(NpcIndex).pos.Map).UserIndex(i)
        If Perseguible(TempChar, NpcIndex) Then ElegidoChar = PrimerUser(ElegidoChar, TempChar, NpcIndex)
    Next

    HayarUser = ElegidoChar

End Function
Public Function Perseguible(UserIndex As Integer, NpcIndex As Integer, Optional Atacando As Boolean) As Boolean

    Perseguible = EnPantalla(UserList(UserIndex).pos, Npclist(NpcIndex).pos, 3) And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Ignorar = 0 And UserList(UserIndex).flags.Protegido = 0

    If Perseguible Then
        If Not Atacando Then Perseguible = Perseguible And UserList(UserIndex).flags.Invisible = 0

        If Npclist(NpcIndex).flags.Faccion <> Neutral Then Perseguible = Perseguible And (UserList(UserIndex).Faccion.Bando = Enemigo(Npclist(NpcIndex).flags.Faccion) Or UserList(UserIndex).Faccion.Ataco(Npclist(NpcIndex).flags.Faccion) > 0 Or UserList(UserIndex).Faccion.BandoOriginal <> UserList(UserIndex).Faccion.Bando)
        If Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion = 0 Then Perseguible = Perseguible And (UserIndex = Npclist(NpcIndex).flags.AttackedBy)
    End If

End Function
Private Function PrimerUser(UserIndex1 As Integer, UserIndex2 As Integer, NpcIndex As Integer) As Integer


    If UserIndex1 = 0 Then
        PrimerUser = UserIndex2
        Exit Function
    End If

    If Distancia(UserList(UserIndex1).pos, Npclist(NpcIndex).pos) < Distancia(UserList(UserIndex2).pos, Npclist(NpcIndex).pos) Then
        PrimerUser = UserIndex1
    Else
        PrimerUser = UserIndex2
    End If

End Function
Private Sub IrUsuarioCercano(NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim UI As Integer

    UI = HayarUser(NpcIndex)

    If UI Then
        If Distancia(Npclist(NpcIndex).pos, UserList(UI).pos) > 1 Then
            Call MoveNPCChar(NpcIndex, FindDirection(Npclist(NpcIndex).pos, UserList(UI).pos))
            If Npclist(NpcIndex).flags.LanzaSpells Then Call NpcLanzaUnSpell(NpcIndex, UI)
        End If
    Else
        Call RestoreOldMovement(NpcIndex)
    End If

    Exit Sub

ErrorHandler:
    Call LogError("Ir UsuarioCercano " & Npclist(NpcIndex).name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).pos.Map & " x:" & Npclist(NpcIndex).pos.X & " y:" & Npclist(NpcIndex).pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNpc)
    Call QuitarNPC(NpcIndex)

End Sub
Private Sub SeguirAgresor(NpcIndex As Integer)
    Dim nPos As WorldPos
    Dim HeadingLoop As Byte
    Dim UI As Integer



    UI = Npclist(NpcIndex).flags.AttackedBy

    If UserList(UI).flags.UserLogged And EnPantalla(Npclist(NpcIndex).pos, UserList(UI).pos, 3) And UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
        If Npclist(NpcIndex).flags.LanzaSpells Then
            Dim k As Integer
            k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
            Call NpcLanzaUnSpell(NpcIndex, UI)
        End If
        Call MoveNPCChar(NpcIndex, FindDirection(Npclist(NpcIndex).pos, UserList(UI).pos))
    Else
        Call RestoreOldMovement(NpcIndex)
    End If

End Sub
Public Sub RestoreOldMovement(NpcIndex As Integer)

    If Npclist(NpcIndex).MaestroUser = 0 Then
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
        Npclist(NpcIndex).flags.AttackedBy = 0
    End If

End Sub
Private Sub SeguirAmo(NpcIndex As Integer)
    Dim nPos As WorldPos
    Dim HeadingLoop As Byte
    Dim UI As Integer

    UI = Npclist(NpcIndex).MaestroUser

    If UI = 0 Then Exit Sub

    If UserList(UI).flags.UserLogged And EnPantalla(Npclist(NpcIndex).pos, UserList(UI).pos, 3) And UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
        Call MoveNPCChar(NpcIndex, FindDirection(Npclist(NpcIndex).pos, UserList(UI).pos))
    Else
        Call RestoreOldMovement(NpcIndex)
    End If

End Sub
Private Sub AiNpcAtacaNpc(NpcIndex As Integer)
    Dim nPos As WorldPos
    Dim HeadingLoop As Byte
    Dim NI As Integer

    NI = Npclist(NpcIndex).TargetNpc

    If NI = 0 Then Exit Sub

    If EnPantalla(Npclist(NpcIndex).pos, Npclist(NI).pos, 3) Then
        If Npclist(NpcIndex).MaestroUser > 0 Then
            Call MoveNPCChar(NpcIndex, FindDirection(Npclist(NpcIndex).pos, Npclist(NI).pos))
            Call NpcAtacaNpc(NpcIndex, NI)
        ElseIf Distancia(Npclist(NpcIndex).pos, Npclist(NI).pos) <= 1 Then
            Call NpcAtacaNpc(NpcIndex, NI)
        End If
    ElseIf Npclist(NpcIndex).MaestroUser Then
        Call FollowAmo(NpcIndex)
    Else
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    End If

End Sub
Function NPCMovementAI(NpcIndex As Integer)
    On Error GoTo ErrorHandler

    If Npclist(NpcIndex).MaestroUser = 0 And (Npclist(NpcIndex).Hostile = 1 Or Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS) Then Call NPCAtacaAI(NpcIndex)

    Select Case Npclist(NpcIndex).Movement
        Case MUEVE_AL_AZAR
            If Int(RandomNumber(1, 12)) = 3 Then
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
            Else
                If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then Call IrUsuarioCercano(NpcIndex)
            End If

        Case MOVEMENT_GUARDIA
            Call IrUsuarioCercano(NpcIndex)

        Case NPCDEFENSA
            Call SeguirAgresor(NpcIndex)

        Case SIGUE_AMO
            Call SeguirAmo(NpcIndex)
            If Int(RandomNumber(1, 12)) = 3 Then Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))

        Case NPC_ATACA_NPC
            Call AiNpcAtacaNpc(NpcIndex)

        Case NPC_PATHFINDING
            If ReCalculatePath(NpcIndex) Then
                Call PathFindingAI(NpcIndex)
                If Npclist(NpcIndex).PFINFO.NoPath Then
                    Call MoveNPCChar(NpcIndex, Int(RandomNumber(1, 4)))
                End If
            Else
                If Not PathEnd(NpcIndex) Then
                    Call FollowPath(NpcIndex)
                Else
                    Npclist(NpcIndex).PFINFO.PathLenght = 0
                End If
            End If

    End Select

    Exit Function


ErrorHandler:
    Call LogError("NPCMovementAI " & Npclist(NpcIndex).name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).pos.Map & " x:" & Npclist(NpcIndex).pos.X & " y:" & Npclist(NpcIndex).pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNpc & " " & Err.Description)
    Dim MiNPC As Npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)

End Function
Function UserNear(NpcIndex As Integer) As Boolean

    UserNear = Not Int(Distance(Npclist(NpcIndex).pos.X, Npclist(NpcIndex).pos.Y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).pos.X, UserList(Npclist(NpcIndex).PFINFO.TargetUser).pos.Y)) > 1

End Function
Function ReCalculatePath(NpcIndex As Integer) As Boolean

    If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf (Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1) Then
        ReCalculatePath = True
    End If

End Function
Function SimpleAI(NpcIndex As Integer) As Boolean
    Dim nPos As WorldPos
    Dim HeadingLoop As Byte
    Dim theading As Byte
    Dim Y As Integer
    Dim X As Integer

    For Y = Npclist(NpcIndex).pos.Y - 5 To Npclist(NpcIndex).pos.Y + 5
        For X = Npclist(NpcIndex).pos.X - 5 To Npclist(NpcIndex).pos.X + 5

            If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then

                If MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex Then

                    theading = FindDirection(Npclist(NpcIndex).pos, UserList(MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex).pos)
                    MoveNPCChar NpcIndex, theading

                    Exit Function
                End If
            End If
        Next
    Next

End Function
Function PathEnd(NpcIndex As Integer) As Boolean

    PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght

End Function
Function FollowPath(NpcIndex As Integer) As Boolean
    Dim tmpPos As WorldPos
    Dim theading As Byte

    tmpPos.Map = Npclist(NpcIndex).pos.Map
    tmpPos.X = Npclist(NpcIndex).PFINFO.path(Npclist(NpcIndex).PFINFO.CurPos).Y
    tmpPos.Y = Npclist(NpcIndex).PFINFO.path(Npclist(NpcIndex).PFINFO.CurPos).X

    theading = FindDirection(Npclist(NpcIndex).pos, tmpPos)

    MoveNPCChar NpcIndex, theading

    Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1

End Function
Function PathFindingAI(NpcIndex As Integer) As Boolean
    Dim nPos As WorldPos
    Dim HeadingLoop As Byte
    Dim theading As Byte
    Dim Y As Integer
    Dim X As Integer

    For Y = Npclist(NpcIndex).pos.Y - 10 To Npclist(NpcIndex).pos.Y + 10
        For X = Npclist(NpcIndex).pos.X - 10 To Npclist(NpcIndex).pos.X + 10


            If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then


                If MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex Then

                    Dim tmpUserIndex As Integer
                    tmpUserIndex = MapData(Npclist(NpcIndex).pos.Map, X, Y).UserIndex
                    If UserList(tmpUserIndex).flags.Muerto = 0 Then



                        Npclist(NpcIndex).PFINFO.Target.X = UserList(tmpUserIndex).pos.Y
                        Npclist(NpcIndex).PFINFO.Target.Y = UserList(tmpUserIndex).pos.X
                        Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                        Call SeekPath(NpcIndex)
                        Exit Function
                    End If
                End If

            End If

        Next
    Next
End Function
Sub NpcLanzaUnSpell(NpcIndex As Integer, UserIndex As Integer)
    Dim k As Integer

    If Not EnPantalla(Npclist(NpcIndex).pos, UserList(UserIndex).pos) Then Exit Sub
    If UserList(UserIndex).flags.Invisible And Npclist(NpcIndex).VeInvis = 0 Then Exit Sub

    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))

End Sub
