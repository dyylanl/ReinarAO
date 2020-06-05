Attribute VB_Name = "Module1"
Option Explicit
Public Const NPCRey As Integer = 645 ' aca tienen q pner el npc del bichos.dat o como se llame
'Public Const DaPuntosHonor As Integer = 500
Public Const CastilloMap As Byte = 23 ' aca el mapa donde respawnea
Public Const CastilloX As Byte = 51 ' aca x donde respawnea
Public Const CastilloY As Byte = 35 'aca Y donde respawnea
Public GolpesRey As Byte
Public HayRey As Byte
 
Public Sub MuereRey(ByVal UserIndex As Integer)
    UserList(UserIndex).Faccion.Quests = UserList(UserIndex).Faccion.Quests + 1
    Call SendData(ToAll, 0, 0, "|| El Clan " & UserList(UserIndex).GuildInfo.GuildName & " Ha conquistado el castillo Divino!" & FONTTYPE_TALK)
    Call WriteVar(IniPath & "\Dat\Castillitos.Siam", "CASTILLOS", "ClanCastillo", UserList(UserIndex).GuildInfo.GuildName)
    Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
    Call SendData(ToIndex, UserIndex, 0, "||Has matado al rey" & FONTTYPE_INFO)
    HayRey = 0
End Sub
 
Public Sub DarPremioCastillos()
 
' esto no estoy muy seguro q lo quieran tener, en litio se usaba honor, pero cambienlo por lo q quieran
 
Dim ClanCastillo As String
Dim LoopC As Integer
Dim Puntos As Byte
 
Puntos = 15
ClanCastillo = GetVar(App.Path & "\Dat\Castillitos.ini", "CASTILLOS", "ClanCastillo")
 
Call SendData(ToAll, 0, 0, "|| Repartiendo Premios de Castillo a Clanes." & FONTTYPE_INFO)
 
    For LoopC = 1 To LastUser
 
        If UserList(LoopC).GuildInfo.GuildName <> "" Then
   
            If UserList(LoopC).GuildInfo.GuildName = ClanCastillo Then
            UserList(LoopC).Faccion.Quests = UserList(LoopC).Faccion.Quests + Puntos ' esto lo acoto,  por lo q explique arriba
            End If
        End If
    Next LoopC
   
    Call SendData(ToAll, 0, 0, "|| Premios de Castillo Repartidos." & FONTTYPE_INFO)
 
End Sub
 
