Attribute VB_Name = "Mod_Guerras"

'******************************************************************************
Option Explicit

Public HayGuerra As Boolean    'Temporal: Hay Guerra o No?
Public CiudadGuerra As Integer    'Temporal: En que ciudad es la Guerra?
Public TiempoGuerra As Integer    'Temporal: Tiempo Transcurrido
Public GuerrasAutomaticas As Boolean    'Temporal: Guerras Automaticas
Private PosicionNPC As WorldPos    'Temporal: Posicion del NPC
Private NPCGuerra As Integer    'Temporal: NPC Usado en Guerra
Private NPCGuerra1 As Integer    'Temporal: NPC Usado en Guerra

'Facccion Real:
Public Const NPC1 As Integer = 259    'NPC de La Faccion Real
Private Const MapaGuerra1 As Integer = 207    'Mapa de la Faccion Real
Private Const MapaGuerra1X As Byte = 26    'X del Mapa de la Faccion Real
Private Const MapaGuerra1Y As Byte = 10    'Y del Mapa de la Faccion Caos

'Faccion Caos:
Public Const NPC2 As Integer = 260    'NPC de La Faccion Caos
Private Const MapaGuerra2 As Integer = 207    'Mapa de la Faccion Caos
Private Const MapaGuerra2X As Byte = 75    'X del Mapa de la Faccion Real
Private Const MapaGuerra2Y As Byte = 10    'Y del Mapa de la Faccion Caos

Public Const TiempoEntreGuerra As Byte = 90    'Duración de entre una Guerra y otra (Minutos)
Private Const DuracionGuerra As Byte = 15    'Duración de Guerra (Minutos)

Private Const OroRecompenza As Long = 200000    'Oro de Recompenz
Private Const PuntosRecompensa As Byte = 10    'puntos de canjes de recompensa

Public Const FONTGUERRA As String = "~255~0~0~1~0"


Public Sub IniciarGuerra(ByVal UserIndex As Integer)
    If UserIndex <> 0 Then
        If HayGuerra Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya hay una Guerra Actualmente." & FONTGUERRA)
            Exit Sub
        End If
    End If

    CiudadGuerra = MapaGuerra1
    HayGuerra = True
    TiempoGuerra = 0
    MapInfo(MapaGuerra1).Pk = True

    NPCGuerra = NPC1
    With PosicionNPC
        .Map = MapaGuerra1
        .X = MapaGuerra1X
        .Y = MapaGuerra1Y
    End With
    SpawnNpc NPC1, PosicionNPC, True, False

    'parte caos
    NPCGuerra1 = NPC2
    With PosicionNPC
        .Map = MapaGuerra1
        .X = MapaGuerra2X
        .Y = MapaGuerra2Y
    End With
    SpawnNpc NPC2, PosicionNPC, True, False

    Call SendData(ToAll, UserIndex, 0, "||Se ha desatado una guerra devastadora entre facciones de éstas tierras, si quieres participar para defender tu reino escribe /GUERRA y acaba con todos tus enemigos ¡Apresurate!" & FONTGUERRA)
End Sub

Public Sub TerminaGuerra(ByVal FaccionGanadora As String)
    Dim UI As Integer, X As Integer, Y As Integer

    If FaccionGanadora = "Real" Then
        Call SendData(ToAll, 0, 0, "||La Guerra ha terminado, la facción ganadora es la Armada Real, los miembros de esta faccion reciben a cambio " & PonerPuntos(OroRecompenza) & " Monedas de oro y " & PuntosRecompensa & " puntos de canjes." & FONTGUERRA)
    ElseIf FaccionGanadora = "Caos" Then
        Call SendData(ToAll, 0, 0, "||La Guerra ha terminado, la facción ganadora es la Legion Oscura, los miembros de esta faccion reciben a cambio " & PonerPuntos(OroRecompenza) & " Monedas de oro y " & PuntosRecompensa & " puntos de canjes." & FONTGUERRA)
    ElseIf FaccionGanadora = "NONE" Then
        Call SendData(ToAll, 0, 0, "||La Guerra ha terminado en un empate ya que ningun jefe de cada facción ha fallecido..." & FONTGUERRA)
    End If

    For UI = 1 To LastUser
        If UserList(UI).flags.Guerra = True Then
            If FaccionGanadora = "Caos" And UserList(UI).Faccion.Bando = 2 Then
                UserList(UI).Stats.GLD = UserList(UI).Stats.GLD + OroRecompenza
                UserList(UI).Faccion.Quests = UserList(UI).Faccion.Quests + PuntosRecompensa
                SendUserStatsBox UI

            ElseIf FaccionGanadora = "Real" And UserList(UI).Faccion.Bando = 1 Then
                UserList(UI).Stats.GLD = UserList(UI).Stats.GLD + OroRecompenza
                UserList(UI).Faccion.Quests = UserList(UI).Faccion.Quests + PuntosRecompensa
                SendUserStatsBox UI
            End If
            WarpUserChar UI, 1, RandomNumber(52, 58), RandomNumber(53, 62), True
            SendData ToIndex, UI, 0, "||Has sido teletransportado." & FONTTYPE_INFO
            UserList(UI).flags.Guerra = False
        End If
    Next UI

    For Y = 1 To 100
        For X = 1 To 100
            If MapData(CiudadGuerra, X, Y).NpcIndex > 0 Then
                If Npclist(MapData(CiudadGuerra, X, Y).NpcIndex).Numero = NPC1 Then
                    Call QuitarNPC(MapData(CiudadGuerra, X, Y).NpcIndex)
                End If
            End If
        Next X
    Next Y

    For Y = 1 To 100
        For X = 1 To 100
            If MapData(CiudadGuerra, X, Y).NpcIndex > 0 Then
                If Npclist(MapData(CiudadGuerra, X, Y).NpcIndex).Numero = NPC2 Then
                    Call QuitarNPC(MapData(CiudadGuerra, X, Y).NpcIndex)
                End If
            End If
        Next X
    Next Y

    MapInfo(CiudadGuerra).Pk = False
    HayGuerra = False
    TiempoGuerra = 0
    Exit Sub
End Sub

Public Sub TimeGuerra()
    TiempoGuerra = TiempoGuerra + 1

    If Not HayGuerra And GuerrasAutomaticas Then
        If val(TiempoGuerra) = TiempoEntreGuerra Then
            IniciarGuerra 0
            Exit Sub
        End If
    End If

    If HayGuerra Then
        If DuracionGuerra - val(TiempoGuerra) > 1 Then
            Call SendData(ToAll, 0, 0, "||Quedan " & DuracionGuerra - val(TiempoGuerra) & " minutos de guerra. Escribe /GUERRA para participar." & FONTTYPE_INFO)
        Else
            Call SendData(ToAll, 0, 0, "||Queda " & DuracionGuerra - val(TiempoGuerra) & " minuto de guerra. Escribe /GUERRA para participar." & FONTTYPE_INFO)
        End If
        If val(TiempoGuerra) = DuracionGuerra Then
            TerminaGuerra "NONE"
        End If
    End If
    Exit Sub
End Sub

Public Sub EntrarGuerra(ByVal UserIndex As Integer)
    If Not HayGuerra Then
        Call SendData(ToIndex, UserIndex, 0, "||No hay guerras faccionarias en curso." & FONTGUERRA)
        Exit Sub
    End If

    If UserList(UserIndex).flags.Guerra = True Then
        Call SendData(ToIndex, UserIndex, 0, "||Ya estas participando de la Guerra." & FONTGUERRA)
        Exit Sub
    End If

    If UserList(UserIndex).Faccion.Bando = 0 Then Exit Sub
    If UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 66 Or UserList(UserIndex).pos.Map = 79 Or UserList(UserIndex).Stats.ELV < 49 Then Exit Sub

    If UserList(UserIndex).Faccion.Bando = 1 Then
        WarpUserChar UserIndex, MapaGuerra1, 20, RandomNumber(80, 90), True
    ElseIf UserList(UserIndex).Faccion.Bando = 2 Then
        WarpUserChar UserIndex, MapaGuerra2, 80, RandomNumber(80, 90), True
    End If

    Call SendData(ToIndex, UserIndex, 0, "||La Guerra ha comenzado para ti, defiende a tu facción para recibir una recompensa." & FONTGUERRA)
    UserList(UserIndex).flags.Guerra = True
    Exit Sub
End Sub

Public Sub GuerrasAuto(ByVal UserIndex As Integer, OnOff As Integer)
    If OnOff = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||Las Guerras Automaticas han sido Ativadas." & FONTGUERRA)
        GuerrasAutomaticas = True
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Las Guerras Automaticas han sido Desativadas." & FONTGUERRA)
        GuerrasAutomaticas = False
    End If
    Exit Sub
End Sub
