Attribute VB_Name = "General"
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

Global ANpc As Long
Global Anpc_host As Long

Option Explicit
Public Function MapaPorUbicacion(X As Integer, Y As Integer) As Integer
    Dim i As Integer

    For i = 1 To NumMaps
        If MapInfo(i).LeftPunto = X And MapInfo(i).TopPunto = Y And MapInfo(i).Zona <> Dungeon Then
            MapaPorUbicacion = i
            Exit Function
        End If
    Next

End Function
Public Sub WriteBIT(Variable As Byte, pos As Byte, value As Byte)

    If ReadBIT(Variable, pos) = value Then Exit Sub

    If value = 0 Then
        Variable = Variable - 2 ^ (pos - 1)
    Else: Variable = Variable + 2 ^ (pos - 1)
    End If

End Sub
Public Function Valorcito(Variable As Byte, pos As Byte, Valor As Byte) As Byte

    Call WriteBIT(Variable, pos, Valor)
    Valorcito = Variable

End Function
Public Function ReadBIT(Variable As Byte, pos As Byte) As Byte
    Dim i As Integer

    ReadBIT = Variable

    For i = 7 To pos Step -1
        ReadBIT = ReadBIT Mod 2 ^ i
    Next

    ReadBIT = ReadBIT \ 2 ^ (pos - 1)

End Function
Public Function Enemigo(ByVal Bando As Byte) As Byte

    Select Case Bando
        Case Neutral
            Enemigo = 3
        Case Real
            Enemigo = Caos
        Case Caos
            Enemigo = Real
    End Select

End Function
Sub DarCuerpoDesnudo(UserIndex As Integer)

    If UserList(UserIndex).flags.Navegando Then
        UserList(UserIndex).Char.Head = 0
        UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
        Exit Sub
    End If

    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head

    Select Case UserList(UserIndex).Raza
        Case HUMANO
            Select Case UserList(UserIndex).Genero
                Case HOMBRE
                    UserList(UserIndex).Char.Body = 21
                Case MUJER
                    UserList(UserIndex).Char.Body = 39
            End Select
        Case ELFO_OSCURO
            Select Case UserList(UserIndex).Genero
                Case HOMBRE
                    UserList(UserIndex).Char.Body = 32
                Case MUJER
                    UserList(UserIndex).Char.Body = 40
            End Select
        Case ENANO
            Select Case UserList(UserIndex).Genero
                Case HOMBRE
                    UserList(UserIndex).Char.Body = 53
                Case MUJER
                    UserList(UserIndex).Char.Body = 60
            End Select
        Case GNOMO
            Select Case UserList(UserIndex).Genero
                Case HOMBRE
                    UserList(UserIndex).Char.Body = 53
                Case MUJER
                    UserList(UserIndex).Char.Body = 60
            End Select

        Case Else
            Select Case UserList(UserIndex).Genero
                Case HOMBRE
                    UserList(UserIndex).Char.Body = 21
                Case MUJER
                    UserList(UserIndex).Char.Body = 39
            End Select

    End Select

    UserList(UserIndex).flags.Desnudo = 1

End Sub
Public Function PuedeDestrabarse(UserIndex As Integer) As Boolean
    Dim i As Byte, nPos As WorldPos

    If (UserList(UserIndex).flags.Muerto = 0) Or (Not MapInfo(UserList(UserIndex).pos.Map).Pk And UserList(UserIndex).pos.Map <> 37) Then Exit Function

    For i = NORTH To WEST
        nPos = UserList(UserIndex).pos
        Call HeadtoPos(i, nPos)
        If InMapBounds(nPos.X, nPos.Y) Then
            If LegalPos(nPos.Map, nPos.X, nPos.Y, CBool(UserList(UserIndex).flags.Navegando)) Then Exit Function
        End If
    Next

    PuedeDestrabarse = True

End Function
Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, X As Integer, Y As Integer, b As Byte)

    Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)

End Sub

Public Sub LimpiarItemsMundo()
    Dim MapaActual, Xnn, Ynn, UserIndex As Integer
    MapaActual = 1
    Call SendData(ToAll, 0, 0, "||Realizando Limpieza del Mundo" & FONTTYPE_FENIX)
    For MapaActual = 1 To NumMaps
        For Ynn = YMinMapSize To YMaxMapSize
            For Xnn = XMinMapSize To XMaxMapSize
                If MapData(MapaActual, Xnn, Ynn).OBJInfo.OBJIndex > 0 And MapData(MapaActual, Xnn, Ynn).Blocked = 0 Then
                    If Not ItemEsDeMapa(val(MapaActual), val(Xnn), val(Ynn)) Then
                        Call EraseObj(ToMap, UserIndex, MapaActual, 10000, val(MapaActual), val(Xnn), val(Ynn))
                    End If
                End If
            Next Xnn
        Next Ynn
    Next MapaActual

    Call SendData(ToAll, 0, 0, "||Limpieza del Mundo finalizada!" & FONTTYPE_FENIX)

    If frmMain.Tlimpiar.Enabled = True Then
        frmMain.Tlimpiar.Enabled = False
    End If

End Sub

Sub LimpiarMundo()
    On Error Resume Next
    Dim i As Integer

    For i = 1 To TrashCollector.Count
        Dim d As cGarbage
        Set d = TrashCollector(1)
        Call EraseObj(ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
        Call TrashCollector.Remove(1)
        Set d = Nothing
    Next

End Sub
Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
    #If UsarQueSocket = 0 Then

        Obj.AddressFamily = AF_INET
        Obj.protocol = IPPROTO_IP
        Obj.SocketType = SOCK_STREAM
        Obj.Binary = False
        Obj.Blocking = False
        Obj.BufferSize = 1024
        Obj.LocalPort = Port
        Obj.backlog = 5
        Obj.listen

    #End If
End Sub
Sub EnviarSpawnList(UserIndex As Integer)
    Dim k As Integer, SD As String
    SD = "SPL" & UBound(SpawnList) & ","

    For k = 1 To UBound(SpawnList)
        SD = SD & SpawnList(k).NpcName & ","
    Next

    Call SendData(ToIndex, UserIndex, 0, SD)
End Sub
Sub EstablecerRecompensas()

    Recompensas(MINERO, 1, 1).SubeHP = 120

    Recompensas(MAGO, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
    Recompensas(MAGO, 1, 1).Obj(1).Amount = 1000
    Recompensas(MAGO, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(MAGO, 1, 2).Obj(1).Amount = 1000
    Recompensas(MAGO, 2, 1).SubeHP = 10

    Recompensas(NIGROMANTE, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
    Recompensas(NIGROMANTE, 1, 1).Obj(1).Amount = 1000
    Recompensas(NIGROMANTE, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(NIGROMANTE, 1, 2).Obj(1).Amount = 1000
    Recompensas(NIGROMANTE, 2, 1).SubeHP = 15
    Recompensas(NIGROMANTE, 2, 2).SubeMP = 40

    Recompensas(PALADIN, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
    Recompensas(PALADIN, 1, 1).Obj(1).Amount = 1000
    Recompensas(PALADIN, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(PALADIN, 1, 2).Obj(1).Amount = 1000
    Recompensas(PALADIN, 2, 1).SubeHP = 5
    Recompensas(PALADIN, 2, 1).SubeMP = 10
    Recompensas(PALADIN, 2, 2).SubeMP = 30

    Recompensas(CLERIGO, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
    Recompensas(CLERIGO, 1, 1).Obj(1).Amount = 1000
    Recompensas(CLERIGO, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(CLERIGO, 1, 2).Obj(1).Amount = 1000
    Recompensas(CLERIGO, 2, 1).SubeHP = 10
    Recompensas(CLERIGO, 2, 2).SubeMP = 50

    Recompensas(BARDO, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
    Recompensas(BARDO, 1, 1).Obj(1).Amount = 1000
    Recompensas(BARDO, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(BARDO, 1, 2).Obj(1).Amount = 1000
    Recompensas(BARDO, 2, 1).SubeHP = 10
    Recompensas(BARDO, 2, 2).SubeMP = 50

    Recompensas(DRUIDA, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
    Recompensas(DRUIDA, 1, 1).Obj(1).Amount = 1000
    Recompensas(DRUIDA, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(DRUIDA, 1, 2).Obj(1).Amount = 1000
    Recompensas(DRUIDA, 2, 1).SubeHP = 15
    Recompensas(DRUIDA, 2, 2).SubeMP = 40

    Recompensas(ASESINO, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
    Recompensas(ASESINO, 1, 1).Obj(1).Amount = 1000
    Recompensas(ASESINO, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(ASESINO, 1, 2).Obj(1).Amount = 1000
    Recompensas(ASESINO, 2, 1).SubeHP = 10
    Recompensas(ASESINO, 2, 2).SubeMP = 30

    Recompensas(CAZADOR, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
    Recompensas(CAZADOR, 1, 1).Obj(1).Amount = 1000
    Recompensas(CAZADOR, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(CAZADOR, 1, 2).Obj(1).Amount = 1000
    Recompensas(CAZADOR, 2, 1).SubeHP = 10
    Recompensas(CAZADOR, 2, 2).SubeMP = 50

    Recompensas(ARQUERO, 1, 1).Obj(1).OBJIndex = Flecha
    Recompensas(ARQUERO, 1, 1).Obj(1).Amount = 1500
    Recompensas(ARQUERO, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(ARQUERO, 1, 2).Obj(1).Amount = 1000
    Recompensas(ARQUERO, 2, 1).SubeHP = 10

    Recompensas(GUERRERO, 1, 1).Obj(1).OBJIndex = PocionVerdeNoCae
    Recompensas(GUERRERO, 1, 1).Obj(1).Amount = 80
    Recompensas(GUERRERO, 1, 1).Obj(2).OBJIndex = PocionAmarillaNoCae
    Recompensas(GUERRERO, 1, 1).Obj(2).Amount = 100
    Recompensas(GUERRERO, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
    Recompensas(GUERRERO, 1, 2).Obj(1).Amount = 1000
    Recompensas(GUERRERO, 2, 1).SubeHP = 5

    Recompensas(PIRATA, 1, 1).SubeHP = 20
    Recompensas(PIRATA, 2, 2).SubeHP = 40
End Sub
Sub EstablecerRestas()

    Resta(CIUDADANO) = 3
    AumentoHit(CIUDADANO) = 3
    Resta(TRABAJADOR) = 2.5
    AumentoHit(TRABAJADOR) = 3
    Resta(EXPERTO_MINERALES) = 2.5
    AumentoHit(EXPERTO_MINERALES) = 3
    Resta(MINERO) = 2.5
    AumentoHit(MINERO) = 2
    Resta(HERRERO) = 2.5
    AumentoHit(HERRERO) = 2
    Resta(EXPERTO_MADERA) = 2.5
    AumentoHit(EXPERTO_MADERA) = 3
    Resta(TALADOR) = 2.5
    AumentoHit(TALADOR) = 2
    Resta(CARPINTERO) = 2.5
    AumentoHit(CARPINTERO) = 2
    Resta(PESCADOR) = 2.5
    AumentoHit(PESCADOR) = 1
    Resta(SASTRE) = 2.5
    AumentoHit(SASTRE) = 2
    Resta(ALQUIMISTA) = 2.5
    AumentoHit(ALQUIMISTA) = 2
    Resta(LUCHADOR) = 3
    AumentoHit(LUCHADOR) = 3
    Resta(CON_MANA) = 3
    AumentoHit(CON_MANA) = 3
    Resta(HECHICERO) = 3
    AumentoHit(HECHICERO) = 3
    Resta(MAGO) = 3
    AumentoHit(MAGO) = 1
    Resta(NIGROMANTE) = 3
    AumentoHit(NIGROMANTE) = 1
    Resta(ORDEN_SAGRADA) = 1.5
    AumentoHit(ORDEN_SAGRADA) = 3
    Resta(PALADIN) = 0.5
    AumentoHit(PALADIN) = 3
    Resta(CLERIGO) = 1.5
    AumentoHit(CLERIGO) = 2
    Resta(NATURALISTA) = 2.5
    AumentoHit(NATURALISTA) = 3
    Resta(BARDO) = 1.5
    AumentoHit(BARDO) = 2
    Resta(DRUIDA) = 3
    AumentoHit(DRUIDA) = 2
    Resta(SIGILOSO) = 1.5
    AumentoHit(SIGILOSO) = 3
    Resta(ASESINO) = 1.5
    AumentoHit(ASESINO) = 3
    Resta(CAZADOR) = 0.5
    AumentoHit(CAZADOR) = 3
    Resta(SIN_MANA) = 2
    AumentoHit(SIN_MANA) = 2
    AumentoHit(ARQUERO) = 3
    AumentoHit(GUERRERO) = 3
    AumentoHit(CABALLERO) = 3
    AumentoHit(BANDIDO) = 2
    Resta(PIRATA) = 1.5
    AumentoHit(PIRATA) = 2
    Resta(LADRON) = 2.5
    AumentoHit(LADRON) = 2

End Sub
Sub LoadMensajes()

    Mensajes(Real, 1) = "||&HFF8080°¡No eres fiel al rey!°"
    Mensajes(Real, 2) = "||&HFF8080°¡¡Maldito insolente!! ¡Los seguidores de Horda Infernal no tienen lugar en nuestro ejército!°"
    Mensajes(Real, 3) = "||&HFF8080°Tu Clan no responde a la Alianza del Lhirius AO, debes retirarte de él para poder enlistarte.°"
    Mensajes(Real, 4) = "||&HFF8080°¡Ya perteneces a las tropas reales! ¡Ve a combatir criminales!°"
    Mensajes(Real, 5) = "||&HFF8080°¡Para unirte a nuestras fuerzas debes matar al menos 150 criminales, solo has matado "
    Mensajes(Real, 6) = "||&HFF8080°¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!°"
    Mensajes(Real, 7) = "||&HFF8080°¡¡Bienvenido a al Ejército Imperial!! Si demuestras fidelidad al rey y destreza en las peleas, podrás aumentar de jerarquía.°"
    Mensajes(Real, 8) = "%4"
    Mensajes(Real, 9) = "5&"
    Mensajes(Real, 10) = "8&"
    Mensajes(Real, 11) = "N0"
    Mensajes(Real, 12) = "L0"
    Mensajes(Real, 13) = "J0"
    Mensajes(Real, 14) = "K0"
    Mensajes(Real, 15) = "M0"
    Mensajes(Real, 16) = "||&HFF8080°¡No perteneces a las tropas reales!°"
    Mensajes(Real, 17) = "||&HFF8080°Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa.°"
    Mensajes(Real, 18) = "||&HFF8080°¿Has decidido abandonarnos? Bien, ya nunca volveremos a aceptarte como ciudadano.°"
    Mensajes(Real, 19) = "1W"
    Mensajes(Real, 20) = "||Si ambos juraron fidelidad a la Alianza tienen que estar en clanes enemigos para poder atacarse." & FONTTYPE_FIGHT
    Mensajes(Real, 21) = "/E"
    Mensajes(Real, 22) = "||&HFF8080°¡Ya haz alcanzado la jerarquia más alta en las filas de la Alianza del Lhirius AO!"
    Mensajes(Real, 23) = "||&HFF8080°¡No puedes abandonar la Alianza del Lhirius AO! Perteneces a un clan ya, debes abandonarlo primero.°"

    Mensajes(Caos, 1) = "||&H8080FF°¡No eres fiel a Horda Infernal!°"
    Mensajes(Caos, 2) = "||&H8080FF°¡¡Maldito insolente!! ¡Los seguidores del rey no tienen lugar en nuestro ejército!°"
    Mensajes(Caos, 3) = "||&H8080FF°Tu Clan no responde al Ejército de Horda Infernal, debes retirarte de él para poder enlistarte.°"
    Mensajes(Caos, 4) = "||&H8080FF°¡Ya perteneces a las tropas del mal! ¡Ve a combatir ciudadanos!°"
    Mensajes(Caos, 5) = "||&H8080FF°¡Para unirte a nuestras fuerzas debes matar al menos 150 ciudadanos, solo has matado "
    Mensajes(Caos, 6) = "||&H8080FF°¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!°"
    Mensajes(Caos, 7) = "||&H8080FF°¡Bienvenido al Ejército de Horda Infernal! Si demuestras tu fidelidad y destreza en las peleas, podrás aumentar de jerarquía.°"
    Mensajes(Caos, 8) = "%5"
    Mensajes(Caos, 9) = "6&"
    Mensajes(Caos, 10) = "9&"
    Mensajes(Caos, 11) = "R0"
    Mensajes(Caos, 12) = "P0"
    Mensajes(Caos, 13) = "Ñ0"
    Mensajes(Caos, 14) = "O0"
    Mensajes(Caos, 15) = "Q0"
    Mensajes(Caos, 16) = "||&H8080FF°¡No perteneces al Ejército de Horda Infernal!°"
    Mensajes(Caos, 17) = "||&H8080FF°Tu deber es combatir ciudadanos, cada 100 ciudadanos que derrotes te dare una recompensa.°"
    Mensajes(Caos, 18) = "||&H8080FF°¡Traidor! ¡Jamás podrás volver con nosotros!°"
    Mensajes(Caos, 19) = "2&"
    Mensajes(Caos, 20) = "||Si ambos son seguidores de Horda Infernal tienen que estar en clanes enemigos para poder atacarse." & FONTTYPE_FIGHT
    Mensajes(Caos, 21) = "/D"
    Mensajes(Caos, 22) = "||&H8080FF°¡Ya haz alcanzado la jerarquia más alta en las filas del Ejército de Horda Infernal!°"
    Mensajes(Caos, 23) = "||&H8080FF°¡No puedes abandonar el Ejército de Horda Infernal! Perteneces a un clan ya, debes abandonarlo primero.°"

End Sub
Sub RevisarCarpetas()

    If Not FileExist(App.path & "\Logs", vbDirectory) Then Call MkDir$(App.path & "\Logs")
    If Not FileExist(App.path & "\Logs\Consejeros", vbDirectory) Then Call MkDir$(App.path & "\Logs\Consejeros")
    If Not FileExist(App.path & "\Logs\Data", vbDirectory) Then Call MkDir$(App.path & "\Logs\Data")
    If Not FileExist(App.path & "\Foros", vbDirectory) Then Call MkDir$(App.path & "\Foros")
    If Not FileExist(App.path & "\Guilds", vbDirectory) Then Call MkDir$(App.path & "\Guilds")
    If Not FileExist(App.path & "\WorldBackUp", vbDirectory) Then Call MkDir$(App.path & "\WorldBackUp")
    If FileExist(App.path & "\Logs\NPCs.log", vbNormal) Then Call Kill(App.path & "\Logs\NPCs.log")

End Sub
Sub Listas()
    Dim i As Integer

    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97

    For i = 40 To 100
        LevelSkill(i).LevelValue = 100
    Next i


    ELUs(1) = 300
    EFrags(50) = 20
    Dim j As Long

    For i = 2 To 10
        ELUs(i) = ELUs(i - 1) * 1.5
    Next

    For i = 11 To 24
        ELUs(i) = ELUs(i - 1) * 1.3
    Next i

    For i = 25 To 49
        ELUs(i) = ELUs(i - 1) * 1.4
    Next i
    For i = 50 To 100
        ELUs(i) = 0
    Next i

    For j = 1 To 49
        EFrags(j) = 20
    Next j

    For j = 51 To STAT_MAXELV - 1
        EFrags(j) = EFrags(j - 1) + 20
    Next j

    EFrags(100) = 0

    ReDim ListaRazas(1 To NUMRAZAS) As String
    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Enano"
    ListaRazas(3) = "Elfo"
    ListaRazas(4) = "Elfo oscuro"
    ListaRazas(5) = "Gnomo"

    ReDim ListaBandos(0 To 2) As String
    ListaBandos(0) = "Neutral"
    ListaBandos(1) = "Alianza del Lhirius AO"
    ListaBandos(2) = "Ejército de Horda Infernal"

    ReDim ListaClases(1 To NUMCLASES) As String
    ListaClases(1) = "Ciudadano"
    ListaClases(2) = "Trabajador"
    ListaClases(3) = "Experto en minerales"
    ListaClases(4) = "Minero"
    ListaClases(8) = "Herrero"
    ListaClases(13) = "Experto en uso de madera"
    ListaClases(14) = "Leñador"
    ListaClases(18) = "Carpintero"
    ListaClases(23) = "Pescador"
    ListaClases(27) = "Sastre"
    ListaClases(31) = "Alquimista"
    ListaClases(35) = "Luchador"
    ListaClases(36) = "Con uso de mana"
    ListaClases(37) = "Hechicero"
    ListaClases(38) = "Mago"
    ListaClases(39) = "Nigromante"
    ListaClases(40) = "Orden sagrada"
    ListaClases(41) = "Paladin"
    ListaClases(42) = "Clerigo"
    ListaClases(43) = "Naturalista"
    ListaClases(44) = "Bardo"
    ListaClases(45) = "Druida"
    ListaClases(46) = "Sigiloso"
    ListaClases(47) = "Asesino"
    ListaClases(48) = "Cazador"
    ListaClases(49) = "Sin uso de mana"
    ListaClases(50) = "Arquero"
    ListaClases(51) = "Guerrero"
    ListaClases(52) = "Caballero"
    ListaClases(53) = "Bandido"
    ListaClases(55) = "Pirata"
    ListaClases(56) = "Ladron"

    ReDim SkillsNames(1 To NUMSKILLS) As String

    SkillsNames(1) = "Magia"
    SkillsNames(2) = "Robar"
    SkillsNames(3) = "Tacticas de combate"
    SkillsNames(4) = "Combate con armas"
    SkillsNames(5) = "Meditar"
    SkillsNames(6) = "Destreza con dagas"
    SkillsNames(7) = "Ocultarse"
    SkillsNames(8) = "Supervivencia"
    SkillsNames(9) = "Talar árboles"
    SkillsNames(10) = "Defensa con escudos"
    SkillsNames(11) = "Pesca"
    SkillsNames(12) = "Mineria"
    SkillsNames(13) = "Carpinteria"
    SkillsNames(14) = "Herreria"
    SkillsNames(15) = "Liderazgo"
    SkillsNames(16) = "Domar animales"
    SkillsNames(17) = "Armas de proyectiles"
    SkillsNames(18) = "Wresterling"
    SkillsNames(19) = "Navegacion"
    SkillsNames(20) = "Sastrería"
    SkillsNames(21) = "Comercio"
    SkillsNames(22) = "Resistencia Mágica"


    ReDim UserSkills(1 To NUMSKILLS) As Integer

    ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
    ReDim AtributosNames(1 To NUMATRIBUTOS) As String
    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"

End Sub

Public Sub LoadGuildsNew()
    Dim NumGuilds As Integer, GuildNum As Integer
    Dim i As Integer, Num As Integer
    Dim a As Long, S As Long
    Dim NewGuild As cGuild

    If Not FileExist(App.path & "\Guilds\GuildsInfo.inf", vbNormal) Then Exit Sub

    a = INICarga(App.path & "\Guilds\GuildsInfo.inf")
    Call INIConf(a, 0, "", 0)

    S = INIBuscarSeccion(a, "INIT")
    NumGuilds = INIDarClaveInt(a, S, "NroGuilds")

    For GuildNum = 1 To NumGuilds

        S = INIBuscarSeccion(a, "Guild" & GuildNum)

        If S >= 0 Then
            Set NewGuild = New cGuild
            With NewGuild
                .GuildName = INIDarClaveStr(a, S, "GuildName")
                .Founder = INIDarClaveStr(a, S, "Founder")
                .FundationDate = INIDarClaveStr(a, S, "Date")
                .CVCsGanados = INIDarClaveStr(a, S, "CVCSGANADOS")
                .Description = INIDarClaveStr(a, S, "Desc")

                .Codex = INIDarClaveStr(a, S, "Codex")

                .Leader = INIDarClaveStr(a, S, "Leader")
                .Gold = INIDarClaveInt(a, S, "Gold")
                .URL = INIDarClaveStr(a, S, "URL")
                .GuildExperience = INIDarClaveInt(a, S, "Exp")
                .DaysSinceLastElection = INIDarClaveInt(a, S, "DaysLast")
                .GuildNews = INIDarClaveStr(a, S, "GuildNews")
                .Bando = INIDarClaveInt(a, S, "Bando")

                Num = INIDarClaveInt(a, S, "NumAliados")

                For i = 1 To Num
                    Call .AlliedGuilds.Add(INIDarClaveStr(a, S, "Aliado" & i))
                Next

                Num = INIDarClaveInt(a, S, "NumEnemigos")

                For i = 1 To Num
                    Call .EnemyGuilds.Add(INIDarClaveStr(a, S, "Enemigo" & i))
                Next

                Num = INIDarClaveInt(a, S, "NumMiembros")

                For i = 1 To Num
                    Call .Members.Add(INIDarClaveStr(a, S, "Miembro" & i))
                Next

                Num = INIDarClaveInt(a, S, "NumSolicitudes")

                Dim sol As cSolicitud

                For i = 1 To Num
                    Set sol = New cSolicitud
                    sol.UserName = ReadField(1, INIDarClaveStr(a, S, "Solicitud" & i), 172)
                    sol.Desc = ReadField(2, INIDarClaveStr(a, S, "Solicitud" & i), 172)
                    Call .Solicitudes.Add(sol)
                Next

                Num = INIDarClaveInt(a, S, "NumProposiciones")

                For i = 1 To Num
                    Set sol = New cSolicitud
                    sol.UserName = ReadField(1, INIDarClaveStr(a, S, "Proposicion" & i), 172)
                    sol.Desc = ReadField(2, INIDarClaveStr(a, S, "Proposicion" & i), 172)
                    Call .PeacePropositions.Add(sol)
                Next

                Call Guilds.Add(NewGuild)
            End With
        End If
    Next

End Sub
Sub Main()
    On Error Resume Next

    'Shell "regsvr32 -s GeniuXSVAC.dll"
    'LoadAntiCheat

    Call Randomize(Timer)

    ChDir App.path
    ChDrive App.path

    Call RevisarCarpetas
    Call LoadMotd

    TiempoEmperador = 0
    MataGuardiasEmperador = 0
    TiempoGuerra = 0
    GuerrasAutomaticas = True

    EmperadorPos.Map = 177
    EmperadorPos.X = 55
    EmperadorPos.Y = 19


    Prision.Map = 66
    Libertad.Map = 1

    Prision.X = 50
    Prision.Y = 49
    Libertad.X = 50
    Libertad.Y = 50

    ReDim Resta(1 To NUMCLASES) As Single
    ReDim Recompensas(1 To NUMCLASES, 1 To 3, 1 To 2) As Recompensa
    ReDim AumentoHit(1 To NUMCLASES) As Byte
    Call EstablecerRestas

    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")

    ReDim Npclist(1 To MAXNPCS) As Npc
    ReDim CharList(1 To MAXCHARS) As Integer

    IniPath = App.path & "\"
    DatPath = App.path & "\Dat\"
    MapPath = App.path & "\Maps\"
    MapDatFile = MapPath & "Info.dat"

    Call Listas

    frmCargando.Show

    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    ENDL = Chr$(13) & Chr$(10)
    ENDC = Chr$(1)
    IniPath = App.path & "\"
    CharPath = App.path & "\charfile\"

    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    DoEvents

    Call LoadBans
    Call LoadSoportes
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    Call LoadGuildsNew
    Call CargarMods
    Call CargarSpawnList
    Call CargarForbidenWords
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    Call LoadSini
    Call LoadQuest
    Call CargarPremiosList
    Call CargaNpcsDat
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    Call LoadOBJData
    Call LoadTops(Nivel)
    Call LoadTops(Muertos)
    Call LoadMensajes
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    Call LoadEscudosHerreria
    Call LoadCascosHerreria
    Call LoadObjCarpintero
    Call LoadObjSastre
    Call LoadVentas
    Call EstablecerRecompensas
    Call LoadCasino

    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapDataNew
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    End If

    Dim LoopC As Integer

    NpcNoIniciado.name = "NPC SIN INICIAR"
    UserOffline.ConnID = -1
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
    Next

    If ClientsCommandsQueue = 1 Then
        frmMain.CmdExec.Enabled = True
    Else
        frmMain.CmdExec.Enabled = False
    End If

    #If UsarQueSocket = 1 Then
        If LastSockListen >= 0 Then Call ApiCloseSocket(LastSockListen)    'Cierra el socket de escucha
        Call IniciaWsApi
        SockListen = ListenForConnect(Puerto, hWndMsg, "")
        If SockListen <> -1 Then
            Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", CStr(SockListen))    ' Guarda el socket escuchando
        Else
            MsgBox "Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly
        End If
        SockListen = ListenForConnect(Puerto, hWndMsg, "")
    #ElseIf UsarQueSocket = 0 Then

        frmCargando.Label1(2).Caption = "Configurando Sockets"

        frmMain.Socket2(0).AddressFamily = AF_INET
        frmMain.Socket2(0).protocol = IPPROTO_IP
        frmMain.Socket2(0).SocketType = SOCK_STREAM
        frmMain.Socket2(0).Binary = False
        frmMain.Socket2(0).Blocking = False
        frmMain.Socket2(0).BufferSize = 2048

        Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    #End If

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

    Call NpcCanAttack(True)
    Call NpcAITimer(True)
    Call AutoTimer(True)

    Unload frmCargando

    Call LogMain("Server iniciado.")

    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If

    tInicioServer = Timer
End Sub
Public Sub ApagarSistema()
    On Error GoTo Terminar
    Dim UI As Integer

    Call WorldSave
    Call SaveGuildsNew
    For UI = 1 To LastUser
        Call CloseSocket(UI)
    Next
    Call DescargaNpcsDat
    Call SaveSoportes
    Call NpcCanAttack(False)
    Call NpcAITimer(False)
    Call AutoTimer(False)

Terminar:
    End

End Sub
Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

    FileExist = Len(Dir$(file, FileType))

End Function
Public Function Tilde(Data As String) As String

    Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(Data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")

End Function
Public Function ReadField(pos As Integer, Text As String, SepASCII As Integer) As String
    Dim i As Integer, LastPos As Integer, FieldNum As Integer

    For i = 1 To Len(Text)
        If mid$(Text, i, 1) = Chr$(SepASCII) Then
            FieldNum = FieldNum + 1
            If FieldNum = pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(SepASCII), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next

    If FieldNum + 1 = pos Then ReadField = mid$(Text, LastPos + 1)

End Function
Function MapaValido(Map As Integer) As Boolean

    MapaValido = Map >= 1 And Map <= NumMaps

End Function
Public Sub LogCriticEvent(Desc As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile
    Open App.path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub
Public Sub LogBando(Bando As Byte, Desc As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile
    Select Case Bando
        Case Real
            Open App.path & "\logs\EjercitoReal.log" For Append Shared As #nfile
        Case Caos
            Open App.path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    End Select
    Print #nfile, Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub
Public Sub LogMain(Desc As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile
    Open App.path & "\Logs\Main.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time, Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub
Public Sub Logear(Archivo As String, Desc As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile
    Open App.path & "\Logs\" & Archivo & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time, Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogBalance(Desc As String, Nombre As String, ip As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile
    Open App.path & "\BALANCE DEL SERVIDOR!!.log" For Append Shared As #nfile
    Print #nfile, "Descripcion:" & Desc & " - Nick:" & Nombre & " - IP:" & ip
    Close #nfile

    Exit Sub

errhandler:

End Sub
Public Sub LogErrorUrgente(Desc As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile
    Open App.path & "\ErroresUrgentes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub
Public Sub LogError(Desc As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile
    Open App.path & "\ErroresUrgentes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogGM(Nombre As String, Texto As String, Consejero As Boolean)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile



    If Consejero Then
        Open App.path & "\logs\consejeros\" & Nombre & ".log" For Append Shared As #nfile
    Else
        Open App.path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
    End If
    Print #nfile, Date & " " & Time & " " & Texto
    Close #nfile

    Exit Sub

errhandler:

End Sub

Public Sub LogVentaCasa(ByVal Texto As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile

    Open App.path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & Texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errhandler:

End Sub
Public Sub LogHackAttemp(Texto As String)
    On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile
    Open App.path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & Texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errhandler:

End Sub
Function ValidInputNP(cad As String) As Boolean
    Dim Arg As String, i As Integer

    For i = 1 To 33
        Arg = ReadField(i, cad, 44)
        If Len(Arg) = 0 Then Exit Function
    Next

    ValidInputNP = True

End Function
Sub Restart()

    On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next

    LastUser = 0
    NumUsers = 0
    NumBots = 0
    NumNoGMs = 0

    ReDim Npclist(1 To MAXNPCS) As Npc
    ReDim CharList(1 To MAXCHARS) As Integer

    Call LoadSini
    Call LoadOBJData
    Call LoadMapDataNew

    Call CargarHechizos

    #If UsarQueSocket = 0 Then
        frmMain.Socket1.Cleanup
        frmMain.Socket1.Startup

        frmMain.Socket2(0).Cleanup
        frmMain.Socket2(0).Startup


        frmMain.Socket1.AddressFamily = AF_INET
        frmMain.Socket1.protocol = IPPROTO_IP
        frmMain.Socket1.SocketType = SOCK_STREAM
        frmMain.Socket1.Binary = False
        frmMain.Socket1.Blocking = False
        frmMain.Socket1.BufferSize = 1024

        frmMain.Socket2(0).AddressFamily = AF_INET
        frmMain.Socket2(0).protocol = IPPROTO_IP
        frmMain.Socket2(0).SocketType = SOCK_STREAM
        frmMain.Socket2(0).Blocking = False
        frmMain.Socket2(0).BufferSize = 2048


        frmMain.Socket1.LocalPort = val(Puerto)
        frmMain.Socket1.listen
    #End If

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."


    Call LogMain(" Servidor reiniciado.")



    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If


End Sub
Public Function Intemperie(UserIndex As Integer) As Boolean

    If MapInfo(UserList(UserIndex).pos.Map).Zona <> "DUNGEON" Then
        Intemperie = MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).trigger <> 1 And _
                     MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).trigger <> 2 And _
                     MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).trigger <> 4
    End If

End Function
Sub Desmontar(UserIndex As Integer)
    Dim Posss As WorldPos

    UserList(UserIndex).flags.Montado = 0
    Call Tilelibre(UserList(UserIndex).pos, Posss)
    Call TraerCaballo(UserIndex, UserList(UserIndex).flags.CaballoMontado + 1, Posss.X, Posss.Y, Posss.Map)
    UserList(UserIndex).flags.CaballoMontado = -1
    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
    Else
        Call DarCuerpoDesnudo(UserIndex)
    End If
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then _
       UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
    If UserList(UserIndex).Invent.WeaponEqpObjIndex Then _
       UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
    If UserList(UserIndex).Invent.CascoEqpObjIndex Then _
       UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, RopaEquitacion(UserIndex), UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Alas)
    Call SendData(ToIndex, UserIndex, 0, "MONTA0")
End Sub
Function RopaEquitacion(UserIndex As Integer) As Integer

    If RazaBaja(UserIndex) Then
        RopaEquitacion = ROPA_DE_EQUITACION_ENANO
    Else
        RopaEquitacion = ROPA_DE_EQUITACION_NORMAL
    End If

End Function
Public Sub TraerCaballo(UserIndex As Integer, ByVal Num As Integer, Optional X As Integer, Optional Y As Integer, Optional Map As Integer)
    Dim NPCNN As Integer
    Dim Poss As WorldPos
    If Map Then
        Poss.Map = Map
        Poss.X = X
        Poss.Y = Y
    Else
        Poss = Ubicar(UserList(UserIndex).pos)
    End If

    NPCNN = SpawnNpc(108, Poss, False, False)

    UserList(UserIndex).Caballos.NpcNum(Num - 1) = NPCNN
    UserList(UserIndex).Caballos.pos(Num - 1) = Npclist(NPCNN).pos

End Sub
Public Function Ubicar(pos As WorldPos) As WorldPos
    On Error GoTo errhandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    Call Tilelibre(pos, NuevaPos)
    If NuevaPos.X <> 0 And NuevaPos.Y Then
        Ubicar = NuevaPos
    End If

    Exit Function

errhandler:
End Function
Sub QuitarCaballos(UserIndex As Integer)
    Dim i As Integer

    For i = 0 To UserList(UserIndex).Caballos.Num - 1
        QuitarNPC (UserList(UserIndex).Caballos.NpcNum(i))
    Next

End Sub
Public Sub CargaNpcsDat()
    Dim npcfile As String

    npcfile = DatPath & "NPCs.dat"
    ANpc = INICarga(npcfile)
    Call INIConf(ANpc, 0, "", 0)

    npcfile = DatPath & "NPCs-HOSTILES.dat"
    Anpc_host = INICarga(npcfile)
    Call INIConf(Anpc_host, 0, "", 0)

End Sub
Public Sub DescargaNpcsDat()

    If ANpc Then Call INIDescarga(ANpc)
    If Anpc_host Then Call INIDescarga(Anpc_host)

End Sub
Sub GuardarUsuarios()
    Dim i As Integer

    Call SendData(ToAll, 0, 0, "2R")

    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr")
    Next

    Call SendData(ToAll, 0, 0, "3R")

End Sub

Public Sub LoadAntiCheatZ()
    Dim i As Integer
    Lac_Camina = CLng(val(GetVar$(App.path & "\AntiCheats.ini", "INTERVALOS", "Caminar")))
    Lac_Lanzar = CLng(val(GetVar$(App.path & "\AntiCheats.ini", "INTERVALOS", "Lanzar")))
    Lac_Usar = CLng(val(GetVar$(App.path & "\AntiCheats.ini", "INTERVALOS", "Usar")))
    Lac_Tirar = CLng(val(GetVar$(App.path & "\AntiCheats.ini", "INTERVALOS", "Tirar")))
    Lac_Pociones = CLng(val(GetVar$(App.path & "\AntiCheats.ini", "INTERVALOS", "Pociones")))
    Lac_Pegar = CLng(val(GetVar$(App.path & "\AntiCheats.ini", "INTERVALOS", "Pegar")))
    For i = 1 To MaxUsers
        ResetearLac i
    Next
End Sub
Public Sub ResetearLac(UserIndex As Integer)
    With UserList(UserIndex).Lac
        .LCaminar.init Lac_Camina
        .LPociones.init Lac_Pociones
        .LUsar.init Lac_Usar
        .LPegar.init Lac_Pegar
        .LLanzar.init Lac_Lanzar
        .LTirar.init Lac_Tirar
    End With
End Sub
Public Sub CargaLac(UserIndex As Integer)
    With UserList(UserIndex).Lac
        Set .LCaminar = New Cls_InterGTC
        Set .LLanzar = New Cls_InterGTC
        Set .LPegar = New Cls_InterGTC
        Set .LPociones = New Cls_InterGTC
        Set .LTirar = New Cls_InterGTC
        Set .LUsar = New Cls_InterGTC
        .LCaminar.init Lac_Camina
        .LPociones.init Lac_Pociones
        .LUsar.init Lac_Usar
        .LPegar.init Lac_Pegar
        .LLanzar.init Lac_Lanzar
        .LTirar.init Lac_Tirar
    End With
End Sub
Public Sub DescargaLac(UserIndex As Integer)
    Exit Sub
    With UserList(UserIndex).Lac
        Set .LCaminar = Nothing
        Set .LLanzar = Nothing
        Set .LPegar = Nothing
        Set .LPociones = Nothing
        Set .LTirar = Nothing
        Set .LUsar = Nothing
    End With
End Sub

Public Sub GuardarRanking(ByVal TipoRanking As String, ByVal UserIndex As Integer, Optional ByVal ClanIndex As Integer)

    Dim Archivo As String
    Archivo = App.path & "\Ranking.ini"

    Dim CantDuelosCheck As Integer
    Dim CantParejaCheck As Integer
    Dim CantRDesafioCheck As Integer
    Dim CantTorneoCheck As Integer
    Dim CantUserFragsCheck As Integer
    Dim CantClanCvcsCheck As Integer

    CantDuelosCheck = GetVar(Archivo, "RANKING", "CANTDUELOS")

    CantParejaCheck = GetVar(Archivo, "RANKING", "CANTPAREJA")

    CantRDesafioCheck = GetVar(Archivo, "RANKING", "CANTRDESAFIO")

    CantTorneoCheck = GetVar(Archivo, "RANKING", "CANTTORNEO")

    CantUserFragsCheck = GetVar(Archivo, "RANKING", "CANTUSERFRAGS")

    CantClanCvcsCheck = GetVar(Archivo, "RANKING", "CANTCLANCVC")

    Select Case TipoRanking

        Case "Duelo"

            If UserList(UserIndex).Ranking.DuelosGanados > CantDuelosCheck Then

                Call WriteVar(Archivo, "RANKING", "USERDUELOS", UserList(UserIndex).name)
                Call WriteVar(Archivo, "RANKING", "CANTDUELOS", val(UserList(UserIndex).Ranking.DuelosGanados))

                SendData ToAll, 0, 0, "||Ranking Actualizado, Usuario con mas duelos ganados: " & UserList(UserIndex).name & " (" & UserList(UserIndex).Ranking.DuelosGanados & " Duelos Ganados)" & FONTTYPE_SERVER

            End If

        Case "Duelos_Pareja"

            If UserList(UserIndex).Ranking.DuelosParejaGanados > CantParejaCheck Then

                Call WriteVar(Archivo, "RANKING", "USERPAREJA", UserList(UserIndex).name)
                Call WriteVar(Archivo, "RANKING", "CANTPAREJA", val(UserList(UserIndex).Ranking.DuelosParejaGanados))

                SendData ToAll, 0, 0, "||Ranking Actualizado, Usuario con mas duelos en pareja ganados: " & UserList(UserIndex).name & " (" & UserList(UserIndex).Ranking.DuelosParejaGanados & " Duelos Ganados)" & FONTTYPE_SERVER

            End If

        Case "Desafio"

            If UserList(UserIndex).Ranking.MaxRondasDesafio > CantRDesafioCheck Then    'Rondas Desafio

                Call WriteVar(Archivo, "RANKING", "USERDESAFIO", UserList(UserIndex).name)
                Call WriteVar(Archivo, "RANKING", "CANTRDESAFIO", val(UserList(UserIndex).Ranking.MaxRondasDesafio))

                SendData ToAll, 0, 0, "||Ranking Actualizado, Usuario con mas rondas ganadas en desafio: " & UserList(UserIndex).name & " (" & UserList(UserIndex).Ranking.MaxRondasDesafio & " Rondas Ganadas)" & FONTTYPE_SERVER

            End If

        Case "Torneo"


            If UserList(UserIndex).Ranking.TorneosGanados > CantTorneoCheck Then

                Call WriteVar(Archivo, "RANKING", "USERTORNEO", UserList(UserIndex).name)
                Call WriteVar(Archivo, "RANKING", "CANTTORNEO", val(UserList(UserIndex).Ranking.TorneosGanados))

                SendData ToAll, 0, 0, "||Ranking Actualizado, Usuario con mas torneos ganados: " & UserList(UserIndex).name & " (" & UserList(UserIndex).Ranking.TorneosGanados & " Torneos Ganados)" & FONTTYPE_SERVER

            End If

        Case "CVC"
            If UserList(UserIndex).GuildRef.CVCsGanados > CantClanCvcsCheck Then
                Call WriteVar(Archivo, "RANKING", "CLANCVC", UserList(UserIndex).GuildInfo.GuildName)
                Call WriteVar(Archivo, "RANKING", "CANTCLANCVC", val(UserList(UserIndex).GuildRef.CVCsGanados))

                SendData ToAll, 0, 0, "||Ranking Actualizado, Clan con más CVCs ganados: " & UserList(UserIndex).GuildInfo.GuildName & " (" & UserList(UserIndex).GuildRef.CVCsGanados & " CVCs Ganados)" & FONTTYPE_SERVER
            End If

    End Select

End Sub

Public Sub LeerRanking(ByVal UserIndex As Integer)

    Dim PathRank As String
    PathRank = App.path & "\Ranking.ini"

    Dim UserDuelos As String
    Dim CantDuelos As Integer
    Dim UserPareja As String
    Dim CantPareja As Integer
    Dim UserDesafio As String
    Dim CantRDesafio As Integer
    Dim UserTorneo As String
    Dim CantTorneo As Integer
    Dim UserFrags As String
    Dim CantUserFrags As Integer
    Dim ClanCvcs As String
    Dim CantClanCvcs As Integer

    UserDuelos = GetVar(PathRank, "RANKING", "USERDUELOS")
    CantDuelos = GetVar(PathRank, "RANKING", "CANTDUELOS")

    UserPareja = GetVar(PathRank, "RANKING", "USERPAREJA")
    CantPareja = GetVar(PathRank, "RANKING", "CANTPAREJA")

    UserDesafio = GetVar(PathRank, "RANKING", "USERDESAFIO")
    CantRDesafio = GetVar(PathRank, "RANKING", "CANTRDESAFIO")

    UserTorneo = GetVar(PathRank, "RANKING", "USERTORNEO")
    CantTorneo = GetVar(PathRank, "RANKING", "CANTTORNEO")
    ClanCvcs = GetVar(PathRank, "RANKING", "CLANCVC")
    CantClanCvcs = GetVar(PathRank, "RANKING", "CANTCLANCVC")

    SendData ToIndex, UserIndex, 0, "||Usuario con más duelos ganados: " & UserDuelos & " (Duelos: " & CantDuelos & ")." & FONTTYPE_SERVER
    SendData ToIndex, UserIndex, 0, "||Usuario con más duelos en pareja ganados: " & UserPareja & " (Duelos: " & CantPareja & ")." & FONTTYPE_SERVER
    SendData ToIndex, UserIndex, 0, "||Usuario con más rondas ganadas en desafio: " & UserDesafio & " (Rondas: " & CantRDesafio & ")." & FONTTYPE_SERVER
    SendData ToIndex, UserIndex, 0, "||Usuario con más torneos ganados: " & UserTorneo & " (Torneos: " & CantTorneo & ")." & FONTTYPE_SERVER
    SendData ToIndex, UserIndex, 0, "||Clan con mas CVCs ganados: " & ClanCvcs & "(CVCs: " & CantClanCvcs & ")." & FONTTYPE_SERVER
End Sub

Sub SendQuest(ByVal User As Integer)
'Envia las quests :A

    Dim i As Long
    Dim tmpS As String
    For i = 1 To MaxQuest
        tmpS = tmpS & Quest(i).Namex & ","
    Next i

    SendData ToIndex, User, 0, "RQN" & MaxQuest & "," & tmpS
End Sub
Public Sub GuardarPremium(ByVal PasoDia As Boolean, Optional ByVal Nick As String, Optional ByVal Dias As Integer)
'Guarda el nick y los dias en la lista premium.

    Dim PremiumFile As String, CantUsers As Integer, i As Long, Existe As Boolean, pos As Long, DiasParciales As Integer

    PremiumFile = App.path & "\Premium.ini"
    CantUsers = val(GetVar(PremiumFile, "INIT", "Usuarios"))

    If PasoDia = False Then
        'primero me fijo que el usuario no esté guardado
        If CantUsers > 0 Then    'no es el primer usuario
            For i = 1 To CantUsers
                If UCase$(GetVar(PremiumFile, "USUARIO" & i, "Nick")) = UCase$(Nick) Then
                    Existe = True
                    pos = i    'guardo la posicion en donde está el usuario
                End If
            Next i
            If Existe = True Then    'ya fue guardado
                DiasParciales = val(GetVar(PremiumFile, "USUARIO" & pos, "Dias"))
                DiasParciales = DiasParciales + Dias
                Call WriteVar(PremiumFile, "USUARIO" & pos, "Dias", str$(DiasParciales))
                Exit Sub
            Else    'NO EXISTE
                Call WriteVar(PremiumFile, "USUARIO" & CantUsers + 1, "Nick", UCase$(Nick))
                Call WriteVar(PremiumFile, "USUARIO" & CantUsers + 1, "Dias", str$(Dias))
                Call WriteVar(PremiumFile, "INIT", "Usuarios", CantUsers + 1)
                Exit Sub
            End If
        Else    ' si es el primer usuario
            Call WriteVar(PremiumFile, "USUARIO" & CantUsers + 1, "Nick", UCase$(Nick))
            Call WriteVar(PremiumFile, "USUARIO" & CantUsers + 1, "Dias", str$(Dias))
            Call WriteVar(PremiumFile, "INIT", "Usuarios", CantUsers + 1)
            Exit Sub
        End If
    Else
        If CantUsers > 0 Then
            For i = 1 To CantUsers
                DiasParciales = val(GetVar(PremiumFile, "USUARIO" & i, "Dias"))
                If DiasParciales <> 1 Then    'si todavia le quedan dias
                    DiasParciales = DiasParciales - 1
                    Call WriteVar(PremiumFile, "USUARIO" & i, "Dias", str$(DiasParciales))
                ElseIf DiasParciales = 1 Then    'si ya no le quedan mas dias
                    If NameIndex(UCase$(GetVar(PremiumFile, "USUARIO" & i, "Nick"))) = 0 Then
                        Call WriteVar(App.path & "\CHARFILE\" & UCase$(GetVar(PremiumFile, "USUARIO" & i, "Nick")) & ".chr", "STATS", "PuntosDonador", 0)
                    Else
                        UserList(NameIndex(UCase$(GetVar(PremiumFile, "USUARIO" & i, "Nick")))).Stats.PuntosDonador = 0
                    End If
                End If
            Next i
            Exit Sub
        End If
    End If

End Sub

Public Sub LoadQuest()

    Dim QuestFile As String
    QuestFile = App.path & "\Dat\Quests.dat"

    MaxQuest = val(GetVar(QuestFile, "NUM", "NumeroQuest"))

    ReDim Quest(MaxQuest) As Tques

    Dim i As Long

    For i = 1 To MaxQuest

        With Quest(i)

            .Exp = val(GetVar(QuestFile, "QUEST" & i, "ExpRecompensa"))
            .Oro = val(GetVar(QuestFile, "QUEST" & i, "OroRecompensa"))
            .Namex = GetVar(QuestFile, "QUEST" & i, "Nombre")
            .Users = val(GetVar(QuestFile, "QUEST" & i, "Usuarios"))
            .NPCs = val(GetVar(QuestFile, "QUEST" & i, "Npcs"))
            .iNPCs = val(GetVar(QuestFile, "QUEST" & i, "NumeroNPC"))
            .Recompense = GetVar(QuestFile, "QUEST" & i, "Recompensa")
            .Canje = GetVar(QuestFile, "QUEST" & i, "Canje")
            .Premium = val(GetVar(QuestFile, "QUEST" & i, "Premium"))
        End With


    Next i
End Sub

Sub BanPC(UserIndex As Integer, TIndex As Integer)

    If TIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(UserList(TIndex).name) = "DYLAN" Then
        Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).name & " Intento darte ban. " & FONTTYPE_FIGHT)
        Call SendData(ToIndex, UserIndex, 0, "||Dylan fue avisado por esta acción. Traidor!!" & FONTTYPE_FIGHT)
        Exit Sub
    End If

    If TIndex Then
        Call SendData(ToIndex, TIndex, 0, "JHT")    'ban registro
        'ban disco
        BanHDs.Add UserList(TIndex).HD
        Call SendData(ToIndex, UserIndex, 0, "||Has baneado al disco duro: " & UserList(TIndex).HD & " del usuario " & UserList(TIndex).name & "." & FONTTYPE_INFO)
        Dim numHD As Integer
        numHD = val(GetVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad"))
        If FileExist(App.path & "\Logs\BanHDs.dat", vbNormal) Then
            Call WriteVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad", numHD + 1)
            Call WriteVar(App.path & "\Logs\BanHDs.dat", "BANS", "HD" & numHD + 1, UserList(TIndex).HD)
            Call LogGM(UserList(UserIndex).name, "/BanHD " & UserList(TIndex).name & " " & UserList(TIndex).HD, False)
        Else
            Call WriteVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad", 1)
            Call WriteVar(App.path & "\Logs\BanHDs.dat", "BANS", "HD1", UserList(TIndex).HD)
        End If    'ban disco

        'banip
        Dim BanIP As String
        BanIP = UserList(TIndex).ip
        BanIps.Add BanIP
        'banip
    End If

    Call LogBan(TIndex, UserIndex, "Personaje baneado Tolerancia 0")
    Call LogGM(UserList(UserIndex).name, "/BanPC " & rdata, False)
    Call SendData(ToAdmins, 0, 0, "|| " & UserList(UserIndex).name & " Baneo la PC a " & rdata & "." & FONTTYPE_FIGHT)
    UserList(TIndex).flags.Ban = 1
    Call CloseSocket(TIndex)
End Sub
