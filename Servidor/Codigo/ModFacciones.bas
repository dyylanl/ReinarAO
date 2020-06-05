Attribute VB_Name = "ModFacciones"
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
Public Sub Recompensado(UserIndex As Integer)
    Dim Fuerzas As Byte
    Dim MiObj As Obj

    Fuerzas = UserList(UserIndex).Faccion.Bando
    MiObj.Amount = 1

    If UserList(UserIndex).Faccion.Jerarquia = 0 Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 11))
        Exit Sub
    End If

    If UserList(UserIndex).Faccion.Jerarquia = 1 Then
        If UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) < 100 Then    ' Cantidad de matados para primera jerarquía en mi caso 20.
            Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 12) & 100)    ' Cambiamos para que te diga la cantidad necesaria cuando clikeamos en el npc.
            Exit Sub
        End If

        UserList(UserIndex).Faccion.Jerarquia = 2
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 15) & Titulo(UserIndex))

        If Fuerzas = 1 Then    'alianza
            MiObj.OBJIndex = 969
        ElseIf Fuerzas = 2 Then    'hordas
            MiObj.OBJIndex = 971
        End If
        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

    ElseIf UserList(UserIndex).Faccion.Jerarquia = 2 Then
        If UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) < 250 Then    ' Número de matados necesarios
            Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 12) & 250)    ' Cambiamos para que te diga la cantidad necesaria cuando clikeamos en el npc.
            Exit Sub
        End If

        UserList(UserIndex).Faccion.Jerarquia = 3    '
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 15) & Titulo(UserIndex))

        If Fuerzas = 1 Then    'alianza
            MiObj.OBJIndex = 972
        ElseIf Fuerzas = 2 Then    'hordas
            MiObj.OBJIndex = 973
        End If
        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

    ElseIf UserList(UserIndex).Faccion.Jerarquia = 3 Then
        If UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) < 500 Then    ' Matados necesarios para enlistarse.
            Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 12) & 500)    ' Cambiamos para que te diga la cantidad necesaria cuando clikeamos en el npc.
            Exit Sub
        End If

        UserList(UserIndex).Faccion.Jerarquia = 4
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 15) & Titulo(UserIndex))

        If Fuerzas = 1 Then    'alianza
            MiObj.OBJIndex = 974
        ElseIf Fuerzas = 2 Then    'hordas
            MiObj.OBJIndex = 975
        End If
        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)


    End If


    If Not UserList(UserIndex).Faccion.Jerarquia < 4 Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 22) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    End If

End Sub
Public Sub Expulsar(UserIndex As Integer)

    Call SendData(ToIndex, UserIndex, 0, Mensajes(UserList(UserIndex).Faccion.Bando, 8))
    UserList(UserIndex).Faccion.Bando = Neutral
    Call UpdateUserChar(UserIndex)

End Sub
Public Sub Enlistar(UserIndex As Integer, ByVal Fuerzas As Byte)
    Dim MiObj As Obj

    If UserList(UserIndex).Faccion.Bando = Neutral Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 1) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If

    If UserList(UserIndex).Faccion.Bando = Enemigo(Fuerzas) Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 2) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If

    Dim oGuild As cGuild

    Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

    If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 Then
        If oGuild.Bando <> Fuerzas Then
            Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 3) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
    End If

    If UserList(UserIndex).Faccion.Jerarquia Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 4) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If

    If UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) < 30 Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 5) & UserList(UserIndex).Faccion.Matados(Enemigo(Fuerzas)) & "!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If

    If UserList(UserIndex).Stats.ELV < 25 Then
        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 6) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If

    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 7) & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

    UserList(UserIndex).Faccion.Jerarquia = 1

    MiObj.Amount = 1

    If Fuerzas = 1 Then    'alianza
        MiObj.OBJIndex = 967
    ElseIf Fuerzas = 2 Then    'hordas
        MiObj.OBJIndex = 968
    End If
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)

    Call LogBando(Fuerzas, UserList(UserIndex).name)

End Sub
Public Function Titulo(UserIndex As Integer) As String

    Select Case UserList(UserIndex).Faccion.Bando
        Case Real
            If UserList(UserIndex).flags.EsConseReal = 1 Then
                Titulo = "Consejero de la Alianza Imperial"
                Exit Function
            End If
            Select Case UserList(UserIndex).Faccion.Jerarquia
                Case 0
                    Titulo = "Alianza Imperial"
                Case 1
                    Titulo = "Primera Jerarquia de la Alianza Imperial"
                Case 2
                    Titulo = "Segunda Jerarquia de la Alianza Imperial"
                Case 3
                    Titulo = "Tercera Jerarquia de la Alianza Imperial"
                Case 4
                    Titulo = "Maxima Jerarquia de la Alianza Imperial"
            End Select
        Case Caos
            If UserList(UserIndex).flags.EsConseCaos = 1 Then
                Titulo = "Consejero de la Horda del Mal"
                Exit Function
            End If
            Select Case UserList(UserIndex).Faccion.Jerarquia
                Case 0
                    Titulo = "Horda del Mal"
                Case 1
                    Titulo = "Primera Jerarquia de la Horda del Mal"
                Case 2
                    Titulo = "Segunda Jerarquia de la Horda del Mal"
                Case 3
                    Titulo = "Tercera Jerarquia de la Horda del Mal"
                Case 4
                    Titulo = "Maxima Jerarquia de la Horda del Mal"
            End Select
    End Select

End Function
