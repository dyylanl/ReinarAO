Attribute VB_Name = "Quest"
Type NpcQuest
    Numero As Integer
    Cant As Integer
End Type
 
Type UserQuest
    QuestIndex      As Integer
    Npc             As NpcQuest
    UsersAmatar     As Integer
End Type
 
Type tQuest
    Premio As Integer
    Nivel As Byte
    Npc As NpcQuest
    UsersAmatar As Integer
End Type
 
Public MaxQuest As Byte
Public Quest() As tQuest
 
Sub CargarQuest()
 
    Dim path As String
    path = App.path & "\Dat\Quests.Siam"
    MaxQuest = val(GetVar(path, "Init", "MaxQuest"))
   
    If MaxQuest <= 0 Then Exit Sub
   
    ReDim Quest(1 To MaxQuest) As tQuest
    Dim i As Integer
   
    For i = 1 To MaxQuest
   
        With Quest(i)
            .Premio = val(GetVar(path, CStr(i), "Premio"))
            .Nivel = val(GetVar(path, CStr(i), "Nivel"))
            .UsersAmatar = val(GetVar(path, CStr(i), "UsersAmatar"))
            .Npc.Cant = val(GetVar(path, CStr(i), "NPCCant"))
            .Npc.Numero = val(GetVar(path, CStr(i), "NPCNumero"))
        End With
   
    Next
   
   
 
End Sub
 
 
Sub EnviarQuest(UI As Integer)
Dim i As Integer
 
If MaxQuest <= 0 Then Exit Sub
 
 SendData ToIndex, UI, 0, "ECQ" & MaxQuest
 
    For i = 1 To MaxQuest
   
        SendData ToIndex, UI, 0, "QQT" & i _
        & "," & Quest(i).Nivel _
        & "," & Quest(i).Premio _
        & "," & Quest(i).Npc.Cant & "," & NombreNPC(Quest(i).Npc.Numero) _
        & "," & Quest(i).UsersAmatar
       
   
    Next i
   
 
End Sub
 
Function NombreNPC(NPCNumber As Integer) As String
 
 
 
Dim A As Long, S As Long
 
If NPCNumber > 499 Then
 
    A = Anpc_host
Else
 
    A = ANpc
End If
 
S = INIBuscarSeccion(A, "NPC" & NPCNumber)
 
    NombreNPC = INIDarClaveStr(A, S, "Name")
   
   
End Function
Sub AceptarQuest(UI As Integer, Qi As Integer)
 
    With UserList(UI).Quest
   
        If .QuestIndex <> 0 Then SendData ToIndex, UI, 0, "||Deves terminar tu quest o rechazarla para poder comenzar otra" & FONTTYPE_INFO: Exit Sub
       
        If Qi > MaxQuest Then Exit Sub
        If Qi < 1 Then Exit Sub
       
        If UserList(UI).Stats.ELV < Quest(Qi).Nivel Then
            SendData ToIndex, UI, 0, "||Tu nivel no es adecuado para realizar esta quest." & FONTTYPE_INFO
            Exit Sub
        End If
       
       
       
        .QuestIndex = Qi
        .UsersAmatar = Quest(Qi).UsersAmatar
        .Npc = Quest(Qi).Npc
       
       
       
        SendData ToIndex, UI, 0, "||Has aceptado la mision " & .QuestIndex & FONTTYPE_TALK
    End With
       
       
 
End Sub
 
Sub FinalizarQuest(UI As Integer)
 
If UserList(UI).Quest.QuestIndex = 0 Then Exit Sub
 
    Dim QQ As UserQuest
   
    UserList(UI).Quest = QQ
   
   
    SendData ToIndex, UI, 0, "||Quest Rechazada" & FONTTYPE_INFO
   
End Sub
 
Sub VerificarFinQuest(UI As Integer)
 
If UserList(UI).Quest.QuestIndex <= 0 Or UserList(UI).Quest.QuestIndex > MaxQuest Then Exit Sub
 
 
    If CumpliQuest(UI) Then
   
        OtorgarPremioQuest UI
        SendData ToIndex, UI, 0, "||Quest Completada" & FONTTYPE_FENIX
        FinalizarQuest UI
   
    End If
 
End Sub
Sub OtorgarPremioQuest(UI As Integer)
 
   UserList(UI).flags.Quest = UserList(UI).flags.Quest + Quest(UserList(UI).Quest.QuestIndex).Premio
 
End Sub
Function CumpliQuest(UI As Integer) As Boolean
 
   
        With UserList(UI).Quest
       
       
        If .QuestIndex <= 0 Then Exit Function
        If .UsersAmatar <> 0 Then Exit Function
        If .Npc.Cant > 0 Then Exit Function
       
               
        End With
       
   
CumpliQuest = True
 
 
End Function
 
Sub DescontarUserQuest(UI As Integer)
 
If UserList(UI).Quest.QuestIndex <> 0 Then
 
Dim i As Integer
 
    If Quest(UserList(UI).Quest.QuestIndex).UsersAmatar > 0 Then
           
            If UserList(UI).Quest.UsersAmatar > 0 Then _
            UserList(UI).Quest.UsersAmatar = UserList(UI).Quest.UsersAmatar - 1
           
            VerificarFinQuest UI
   
    End If
End If
End Sub
 
Sub DescontarNPCQuest(UI As Integer, Npc As Integer)
 
 
If UserList(UI).Quest.QuestIndex > 0 Then
 
       
            If Npc = UserList(UI).Quest.Npc.Numero Then
           
                If UserList(UI).Quest.Npc.Cant > 0 Then UserList(UI).Quest.Npc.Cant = UserList(UI).Quest.Npc.Cant - 1
           
            End If
           
            VerificarFinQuest UI
       
   
End If
 
End Sub
 
Sub SendInfoQuest(Userindex As Integer)
 
 
    If UserList(Userindex).Quest.QuestIndex < 1 Then UserList(Userindex).Quest.QuestIndex = 0: Exit Sub
    If UserList(Userindex).Quest.QuestIndex > MaxQuest Then UserList(Userindex).Quest.QuestIndex = 0: Exit Sub
   
    SendData ToIndex, Userindex, 0, "|| Quest N°" & UserList(Userindex).Quest.QuestIndex & FONTTYPE_TALK
   
    If Quest(UserList(Userindex).Quest.QuestIndex).UsersAmatar > 0 Then _
    SendData ToIndex, Userindex, 0, "|| Usuarios restantes: " & UserList(Userindex).Quest.UsersAmatar & FONTTYPE_TALK
 
    If Quest(UserList(Userindex).Quest.QuestIndex).Npc.Numero > 0 Then _
    SendData ToIndex, Userindex, 0, "||" & NombreNPC(UserList(Userindex).Quest.Npc.Numero) & ": " & UserList(Userindex).Quest.Npc.Cant
 
    SendData ToIndex, Userindex, 0, "||Puntos De Torneo: " & Quest(UserList(Userindex).Quest.QuestIndex).Premio
   
End Sub
 
 

