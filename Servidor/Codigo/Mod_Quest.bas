Attribute VB_Name = "Mod_Quest"
'Amra
'Argentum Online 0.11.2.1
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'
'**************************************************************
' modQuests.bas - Realiza todos los handles para el sistema de
' Quests dentro del juego.
'
' Escrito y dise�ado por Hern�n Gurmendi a.k.a. Amraphen
' (hgurmen@hotmail.com)
'**************************************************************
Option Explicit

Public Type tQuest
    Nombre As String
    Descripcion As String
    NivelRequerido As Integer
    
    NpcKillIndex As Integer
    CantNPCs As Integer

    OBJIndex As Integer
    CantOBJs As Integer
    
    GLDReward As Long
    PuntosTorneoReward As Long
    EXPReward As Long
    
    OBJRewardIndex As Integer
    CantOBJsReward As Integer
    
    Redoable As Byte
End Type

Public Type tUserQuest
    QuestIndex As Integer
    NPCsKilled As Integer
End Type

Public Const MAXUSERQUESTS As Byte = 10
Public QuestList() As tQuest

Public Sub LoadQuests()
'**************************************************************
'Author: Hern�n Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Carga el archivo QUESTS.DAT.
'**************************************************************
Dim QuestFile As clsIniReader
Dim tmpInt As Integer

    Set QuestFile = New clsIniReader
    Call QuestFile.Initialize(App.path & "\DAT\QUESTS.DAT")
       
    ReDim QuestList(1 To QuestFile.GetValue("INIT", "NumQuests"))
    
    For tmpInt = 1 To UBound(QuestList)
        QuestList(tmpInt).Nombre = QuestFile.GetValue("QUEST" & tmpInt, "Nombre")
        QuestList(tmpInt).Descripcion = QuestFile.GetValue("QUEST" & tmpInt, "Descripcion")
        
        QuestList(tmpInt).NpcKillIndex = QuestFile.GetValue("QUEST" & tmpInt, "NpcKillIndex")
        QuestList(tmpInt).CantNPCs = QuestFile.GetValue("QUEST" & tmpInt, "CantNPCs")
        
        QuestList(tmpInt).OBJIndex = QuestFile.GetValue("QUEST" & tmpInt, "OBJIndex")
        QuestList(tmpInt).CantOBJs = QuestFile.GetValue("QUEST" & tmpInt, "CantOBJs")
        
        QuestList(tmpInt).GLDReward = QuestFile.GetValue("QUEST" & tmpInt, "GLDReward")
        QuestList(tmpInt).PuntosTorneoReward = QuestFile.GetValue("QUEST" & tmpInt, "PuntosTorneoReward")
        QuestList(tmpInt).EXPReward = QuestFile.GetValue("QUEST" & tmpInt, "EXPReward")
                
        QuestList(tmpInt).OBJRewardIndex = QuestFile.GetValue("QUEST" & tmpInt, "OBJRewardIndex")
        QuestList(tmpInt).CantOBJsReward = QuestFile.GetValue("QUEST" & tmpInt, "CantOBJsReward")
        QuestList(tmpInt).Redoable = QuestFile.GetValue("QUEST" & tmpInt, "Redoable")
    Next tmpInt
End Sub

Public Function UserTieneQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer) As Integer
'**************************************************************
'Author: Hern�n Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Devuelve 0 si no tiene la quest especificada en QuestNumber, o
'el numero de slot en el que tiene la quest.
'**************************************************************
Dim tmpInt As Integer

    For tmpInt = 1 To MAXUSERQUESTS
        If UserList(UserIndex).Stats.UserQuests(tmpInt).QuestIndex = QuestNumber Then
            UserTieneQuest = tmpInt
            Exit Function
        End If
    Next tmpInt
    
    UserTieneQuest = 0
End Function

Public Sub UserFinishQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer)
'**************************************************************
'Author: Hern�n Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Realiza el handle de /QUEST en caso de que el personaje ya
'tenga la quest.
'**************************************************************
Dim UTQ As Integer 'Determina el valor de UserTieneQuest
Dim tmpObj As Obj
Dim tmpInt As Integer

    If QuestList(QuestNumber).OBJIndex Then
        If TieneObjetos(QuestList(QuestNumber).OBJIndex, QuestList(QuestNumber).CantOBJs, UserIndex) = False Then
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "�" & "Debes traerme los objetos que te he pedido antes de poder terminar la misi�n." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
    End If
    
    UTQ = UserTieneQuest(UserIndex, QuestNumber)
    
    If QuestList(QuestNumber).NpcKillIndex Then
        If UserList(UserIndex).Stats.UserQuests(UTQ).NPCsKilled < QuestList(QuestNumber).CantNPCs Then
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "�" & "Debes matar las criaturas que te he pedido antes de poder terminar la misi�n." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
    End If
    
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "�" & "Gracias por ayudarme, noble aventurero, he aqu� tu recompensa." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Call SendData(ToIndex, UserIndex, 0, "||Has completado la misi�n " & Chr(34) & QuestList(QuestNumber).Nombre & Chr(34) & "." & FONTTYPE_INFO)
        
    If QuestList(QuestNumber).OBJIndex Then
        For tmpInt = 1 To MAX_INVENTORY_SLOTS
            If UserList(UserIndex).Invent.Object(tmpInt).OBJIndex = QuestList(QuestNumber).OBJIndex Then
                Call QuitarUserInvItem(UserIndex, CByte(tmpInt), QuestList(QuestNumber).CantOBJs)
                Exit For
            End If
        Next tmpInt
    End If
        
    If QuestList(QuestNumber).EXPReward Then
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + QuestList(QuestNumber).EXPReward
        Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & QuestList(QuestNumber).EXPReward & " puntos de experiencia como recompensa." & FONTTYPE_INFO)
    End If
    
    If QuestList(QuestNumber).GLDReward Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + QuestList(QuestNumber).GLDReward
        Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & QuestList(QuestNumber).GLDReward & " monedas de oro como recompensa." & FONTTYPE_INFO)
    End If
    If QuestList(QuestNumber).PuntosTorneoReward Then
    If UserList(UserIndex).flags.EsNoble = True Then
        Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & QuestList(QuestNumber).PuntosTorneoReward * 2 & " puntos de canje como recompensa." & FONTTYPE_INFO)
UserList(UserIndex).Faccion.Quests = UserList(UserIndex).Faccion.Quests + QuestList(QuestNumber).PuntosTorneoReward * 2
Else
        Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & QuestList(QuestNumber).PuntosTorneoReward & " puntos de canje como recompensa." & FONTTYPE_INFO)
UserList(UserIndex).Faccion.Quests = UserList(UserIndex).Faccion.Quests + QuestList(QuestNumber).PuntosTorneoReward
End If
    End If
    If QuestList(QuestNumber).OBJRewardIndex Then
    If UserList(UserIndex).flags.EsNoble = True Then
        tmpObj.OBJIndex = QuestList(QuestNumber).OBJRewardIndex
        tmpObj.Amount = QuestList(QuestNumber).CantOBJsReward * 2
        
        If MeterItemEnInventario(UserIndex, tmpObj) = False Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, tmpObj)
            Call SendData(ToIndex, UserIndex, 0, "||Has recibido " & QuestList(QuestNumber).CantOBJsReward * 2 & " " & ObjData(QuestList(QuestNumber).OBJRewardIndex).Name & " como recompensa." & FONTTYPE_INFO)
        Else
        
        tmpObj.OBJIndex = QuestList(QuestNumber).OBJRewardIndex
        tmpObj.Amount = QuestList(QuestNumber).CantOBJsReward
        
        If MeterItemEnInventario(UserIndex, tmpObj) = False Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, tmpObj)
            Call SendData(ToIndex, UserIndex, 0, "||Has recibido " & QuestList(QuestNumber).CantOBJsReward & " " & ObjData(QuestList(QuestNumber).OBJRewardIndex).Name & " como recompensa." & FONTTYPE_INFO)
End If
End If
        End If
    End If
    
    UserList(UserIndex).Stats.UserQuests(UTQ).QuestIndex = 0
    UserList(UserIndex).Stats.UserQuests(UTQ).NPCsKilled = 0
    UserList(UserIndex).Stats.UserQuestsDone = UserList(UserIndex).Stats.UserQuestsDone & QuestNumber & "-"
    
    Call UpdateUserInv(True, UserIndex, 0)
    Call CheckUserLevel(UserIndex)
    Call SendUserStatsBox(UserIndex)
End Sub

Public Sub UserAceptarQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer)
'**************************************************************
'Author: Hern�n Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Realiza el handle de /QUEST en caso de que el personaje no
'tenga la quest.
'**************************************************************
Dim UFQS As Integer

    UFQS = UserFreeQuestSlot(UserIndex)
    
    If QuestList(QuestNumber).Redoable = 0 Then
        If UserHizoQuest(UserIndex, QuestNumber) = True Then
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "�" & "Ya has hecho la misi�n." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
    End If
    
    If UFQS = 0 Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "�" & "Debes terminar o cancelar alguna misi�n antes de poder aceptar otra." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If

    If UserList(UserIndex).Stats.ELV < QuestList(QuestNumber).NivelRequerido Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "�" & "No tienes nivel suficiente como para empezar esta misi�n." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If
    
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "�" & Npclist(UserList(UserIndex).flags.TargetNpc).TalkDuringQuest & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Call SendData(ToIndex, UserIndex, 0, "||Has aceptado la misi�n " & Chr(34) & QuestList(QuestNumber).Nombre & Chr(34) & "." & FONTTYPE_INFO)
    
    UserList(UserIndex).Stats.UserQuests(UFQS).QuestIndex = QuestNumber
    UserList(UserIndex).Stats.UserQuests(UFQS).NPCsKilled = 0
End Sub

Public Function UserFreeQuestSlot(ByVal UserIndex As Integer) As Integer
'**************************************************************
'Author: Hern�n Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Devuelve 0 si no tiene algun slot de quest libre, o el primer
'slot de quest que tiene libre.
'**************************************************************
Dim tmpInt As Integer

    For tmpInt = 1 To MAXUSERQUESTS
        If UserList(UserIndex).Stats.UserQuests(tmpInt).QuestIndex = 0 Then
            UserFreeQuestSlot = tmpInt
            Exit Function
        End If
    Next tmpInt
    
    UserFreeQuestSlot = 0
End Function

Public Function UserHizoQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer) As Boolean
'**************************************************************
'Author: Hern�n Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Devuelve verdadero si el user hizo la quest QuestNumber, o
'falso si el user no la hizo.
'**************************************************************
Dim arrStr() As String
Dim tmpInt As Integer

    arrStr = Split(UserList(UserIndex).Stats.UserQuestsDone, "-")
    
    For tmpInt = 0 To UBound(arrStr)
        If CInt(arrStr(tmpInt)) = QuestNumber Then
            UserHizoQuest = True
            Exit Function
        End If
    Next tmpInt
    
    UserHizoQuest = False
End Function

Public Sub HandleQuest(ByVal UserIndex As Integer)
'**************************************************************
'Author: Hern�n Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Realiza el handle del comando /QUEST.
'**************************************************************
Dim UTQ As Integer 'Determina el valor de la funci�n UserTieneQuest.
Dim QN As Integer 'Determina el valor de la quest que posee el NPC.

    If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 4 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes hablar con el NPC ya que estas demasiado lejos." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(UserIndex).flags.TargetNpc = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar un NPC con el cual hablar." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Muerto Then
        Call SendData(ToIndex, UserIndex, 0, "||Est�s muerto!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    QN = Npclist(UserList(UserIndex).flags.TargetNpc).QuestNumber
    
    If QN = 0 Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "||" & vbWhite & "�" & "No tengo ninguna misi�n para t�." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If
    
    UTQ = UserTieneQuest(UserIndex, QN)
        
    If UTQ Then
        Call UserFinishQuest(UserIndex, QN)
    Else
        Call UserAceptarQuest(UserIndex, QN)
    End If
End Sub

Public Sub SendQuestList(ByVal UserIndex As Integer)
'**************************************************************
'Author: Hern�n Gurmendi (Amraphen)
'Last Modify Date: 23/10/2007
'Env�a a UserIndex la lista de quests.
'**************************************************************
Dim tmpString As String
Dim i As Integer

    For i = 1 To MAXUSERQUESTS
        If UserList(UserIndex).Stats.UserQuests(i).QuestIndex = 0 Then
            tmpString = tmpString & "0-"
        Else
            tmpString = tmpString & QuestList(UserList(UserIndex).Stats.UserQuests(i).QuestIndex).Nombre & "-"
        End If
    Next i
    
    Call SendData(ToIndex, UserIndex, 0, "QL" & Left$(tmpString, Len(tmpString) - 1))
End Sub
'/Amra

