Attribute VB_Name = "ModQuest"
 
Option Explicit
  
Public Sub Misiones(UserIndex As Integer)
Dim UserMision As Integer
  
  
If UserList(UserIndex).flags.EnMision = 1 Then
Call SendData(ToIndex, UserIndex, 0, "||Ya estás en quest. No es posible hacer otra." & FONTTYPE_INFO)
Exit Sub
End If
  
  
UserMision = 1 Or 2 Or 3
  
If UserMision = 1 Then
UserList(UserIndex).flags.EnMision = 1
Call SendData(ToIndex, UserIndex, 0, "||Tu misión es derrotar a 3 dragones rojos que andan dando vueltas en el Dungeon Makial." & FONTTYPE_INFO)
Call SendData(ToIndex, UserIndex, 0, "||La recompenza es 1000000 monedas de oro y experiencia." & FONTTYPE_INFO)
Call Mision1(UserIndex)
Exit Sub
End If
  
If UserMision = 2 Then
UserList(UserIndex).flags.EnMision = 1
Call SendData(ToIndex, UserIndex, 0, "||Tu misión es derrotar a 10 Zombies que andan dando vueltas en el bosque." & FONTTYPE_INFO)
Call SendData(ToIndex, UserIndex, 0, "||La recompenza es 30000 monedas de oro y experiencia." & FONTTYPE_INFO)
Call Mision2(UserIndex)
Exit Sub
End If
  
If UserMision = 3 Then
UserList(UserIndex).flags.EnMision = 1
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu misión es derrotar a 10 usuarios." & FONTTYPE_INFO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La recompenza es 250000 monedas de oro y experiecia." & FONTTYPE_INFO)
Call Mision3(UserIndex)
Exit Sub
End If
  
'If UserMision = 4 Then
'Call SendData(ToIndex, UserIndex, 0, "||No hay quest en este momento." & FONTTYPE_INFO)
'Exit Sub
'End If
  
End Sub
Public Sub Mision1(UserIndex As Integer)
UserList(UserIndex).flags.OroM = 1000000
UserList(UserIndex).flags.ExpM = 50000
UserList(UserIndex).flags.NpcKillM = 547
UserList(UserIndex).flags.NpcKillerM = 3
MD = "Tu misión es derrotar a 3 dragones rojos que andan dando vuelta en el Dungeon Makial."
MATAUSER = 0
MATANPC = 1
End Sub
Public Sub Mision2(UserIndex As Integer)
UserList(UserIndex).flags.OroM = 30000
UserList(UserIndex).flags.ExpM = 10000
UserList(UserIndex).flags.NpcKillM = 507
UserList(UserIndex).flags.NpcKillerM = 10
MD = "Tu misión es derrotar a 10 Zombies que andan dando vuelta en el bosque."
MATAUSER = 0
MATANPC = 1
End Sub
Public Sub Mision3(UserIndex As Integer)
UserList(UserIndex).flags.OroM = 25000
UserList(UserIndex).flags.ExpM = 20000
UserList(UserIndex).flags.UsersKillM = 0
UserList(UserIndex).flags.UsersKillerM = 10
MD = "Tu misión es derrotar a 10 usuarios."
MATAUSER = 1
MATANPC = 0
End Sub
Public Sub TerminoQuest(UserIndex As Integer)
Dim DD As Obj
If UserList(UserIndex).flags.EnMision = 1 And MISION = 1 Then
If UserList(UserIndex).flags.RecompensaM <> 0 Then
DD.Amount = UserList(UserIndex).flags.CantidadRM
DD.OBJIndex = UserList(UserIndex).flags.RecompensaM
Call MeterItemEnInventario(UserIndex, DD)
End If
If UserList(UserIndex).flags.ExpM <> 0 Then UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).flags.OroM
If UserList(UserIndex).flags.OroM <> 0 Then UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + UserList(UserIndex).flags.ExpM
UserList(UserIndex).flags.OroM = 0
UserList(UserIndex).flags.ExpM = 0
UserList(UserIndex).flags.NpcKillM = 0
UserList(UserIndex).flags.NpcKillerM = 0
MISION = 0
MD = 0
UserList(UserIndex).flags.EnMision = 0
Call SendData(ToIndex, UserIndex, 0, "||Has completado la quest." & FONTTYPE_INFO)
Call SendUserStatsBox(UserIndex)
Call EnviarMiniEstadisticas(UserIndex)
Call CheckUserLevel(UserIndex)
Else
Call SendData(ToIndex, UserIndex, 0, "||No terminaste la quest." & FONTTYPE_INFO)
End If
End Sub
  
  
  
