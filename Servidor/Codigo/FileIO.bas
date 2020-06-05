Attribute VB_Name = "ES"
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

Public BalanceCasa As Double
Public Baneos As New Collection
Public SoporteS As New Collection

Option Explicit
Sub LoadUserAccount(ByVal PJinit As String)
    On Error Resume Next
    Dim Nivel As Long
    Dim Clase As String
    Nivel = GetVar(CharPath & PJinit, "STATS", "ELV")
    Clase = UCase$(ListaClases(GetVar(CharPath & PJinit, "INIT", "Clase")))
    PJEnCuenta = Nivel & "," & Clase
End Sub

Public Sub LoadCasino()

    BalanceCasa = val(GetVar(App.path & "\Logs\Casino.log", "INIT", "Balance"))

End Sub
Public Sub SaveCasino()

    Call WriteVar(App.path & "\Logs\Casino.log", "INIT", "Balance", str(BalanceCasa))

End Sub
Public Sub LoadVentas()

    DineroTotalVentas = val(GetVar(App.path & "\Dat\Ventas.dat", "INIT", "Dinero"))
    NumeroVentas = val(GetVar(App.path & "\Dat\Ventas.dat", "INIT", "Numero"))

End Sub


Public Sub CargarPremiosList()

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Lista de items SHOPS."

    Dim P As Long, LoopC As Long, Descuento As Double
    P = val(GetVar(App.path & "\Dat\Premios.dat", "INIT", "NumPremios"))
    Descuento = GetVar(App.path & "\Dat\Premios.dat", "INIT", "Descuento")

    ReDim PremiosList(P) As tPremiosCanjes


    For LoopC = 1 To P
        PremiosList(LoopC).ObjIndexP = val(GetVar(App.path & "\Dat\Premios.dat", "PREMIO" & LoopC, "NumObj"))
        If Descuento > 0 Then
            PremiosList(LoopC).ObjRequiere = ((Descuento * val(GetVar(App.path & "\Dat\Premios.dat", "PREMIO" & LoopC, "Requiere"))))
        Else
            PremiosList(LoopC).ObjRequiere = val(GetVar(App.path & "\Dat\Premios.dat", "PREMIO" & LoopC, "Requiere"))
        End If
        PremiosList(LoopC).ObjPremium = val(GetVar(App.path & "\Dat\Premios.dat", "PREMIO" & LoopC, "Premium"))
        PremiosList(LoopC).ObjDescripcion = GetVar(App.path & "\Dat\Premios.dat", "PREMIO" & LoopC, "Descripcion")
    Next LoopC

    If frmMain.Visible Then frmMain.txStatus.Caption = "La carga de listas Shop ha finalizado con éxito!."
End Sub


Public Sub CargarSpawnList()
    Dim N As Integer, LoopC As Integer

    N = val(GetVar(App.path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))

    ReDim SpawnList(N) As tCriaturasEntrenador

    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(GetVar(App.path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next

End Sub
Function PJQuest(ByVal name As String) As Boolean
    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim Nomb As String

    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "PJsQuest"))

    For WizNum = 1 To NumWizs
        Nomb = UCase$(GetVar(IniPath & "Server.ini", "PJsQuest", "PJQuest" & WizNum))
        If Left$(Nomb, 1) = "*" Or Left$(Nomb, 1) = "+" Then Nomb = Right$(Nomb, Len(Nomb) - 1)
        If UCase$(name) = Nomb Then
            PJQuest = True
            Exit Function
        End If
    Next

End Function
Function PuedeDenunciar(ByVal name As String) As Boolean
    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim Nomb As String

    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SubGMs"))

    CastilloNorte = GetVar(IniPath & "castillos.txt", "INIT", "Castillo1")
    HoraNorte = GetVar(IniPath & "castillos.txt", "INIT", "hora1")
    DateNorte = GetVar(IniPath & "castillos.txt", "INIT", "date1")

    CastilloSur = GetVar(IniPath & "castillos.txt", "INIT", "Castillo2")
    HoraSur = GetVar(IniPath & "castillos.txt", "INIT", "hora2")
    DateSur = GetVar(IniPath & "castillos.txt", "INIT", "date2")

    For WizNum = 1 To NumWizs
        Nomb = UCase$(GetVar(IniPath & "Server.ini", "SubGMs", "SubGM" & WizNum))
        If Left$(Nomb, 1) = "*" Or Left$(Nomb, 1) = "+" Then Nomb = Right$(Nomb, Len(Nomb) - 1)
        If UCase$(name) = Nomb Then
            PuedeDenunciar = True
            Exit Function
        End If
    Next

End Function
Function EsDios(ByVal name As String) As Boolean
    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim Nomb As String

    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))

    For WizNum = 1 To NumWizs
        Nomb = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
        If Left$(Nomb, 1) = "*" Or Left$(Nomb, 1) = "+" Then Nomb = Right$(Nomb, Len(Nomb) - 1)
        If UCase$(name) = Nomb Then
            EsDios = True
            Exit Function
        End If
    Next

End Function
Public Function EsNoble(ByVal UserIndex As Integer) As Boolean
    EsNoble = UserList(UserIndex).flags.EsNoble
End Function
Function EsSemiDios(ByVal name As String) As Boolean
    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim Nomb As String

    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))

    For WizNum = 1 To NumWizs
        Nomb = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
        If Left$(Nomb, 1) = "*" Or Left$(Nomb, 1) = "+" Then Nomb = Right$(Nomb, Len(Nomb) - 1)
        If UCase$(name) = Nomb Then
            EsSemiDios = True
            Exit Function
        End If
    Next

End Function
Function EsConsejero(ByVal name As String) As Boolean
    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim Nomb As String

    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))

    For WizNum = 1 To NumWizs
        Nomb = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
        If Left$(Nomb, 1) = "*" Or Left$(Nomb, 1) = "+" Then Nomb = Right$(Nomb, Len(Nomb) - 1)
        If UCase$(name) = Nomb Then
            EsConsejero = True
            Exit Function
        End If
    Next

End Function
Public Function TxtDimension(ByVal name As String) As Long
    Dim N As Integer, cad As String, Tam As Long

    N = FreeFile(1)

    Open name For Input As #N
    Tam = 0
    Do While Not EOF(N)
        Tam = Tam + 1
        Line Input #N, cad
    Loop
    Close N

    TxtDimension = Tam

End Function
Public Sub CargarForbidenWords()
    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
    Dim N As Integer, i As Integer

    N = FreeFile(1)

    Open DatPath & "NombresInvalidos.txt" For Input As #N
    For i = 1 To UBound(ForbidenNames)
        Line Input #N, ForbidenNames(i)
    Next
    Close N

End Sub
Public Sub CargarHechizos()
    On Error GoTo errhandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

    Dim Hechizo As Integer

    NumeroHechizos = val(GetVar(DatPath & "Hechizos.dat", "INIT", "NumeroHechizos"))
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.value = 0

    For Hechizo = 1 To NumeroHechizos

        Hechizos(Hechizo).Nombre = GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Nombre")
        Hechizos(Hechizo).Desc = GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Desc")
        Hechizos(Hechizo).PalabrasMagicas = GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "PalabrasMagicas")

        Hechizos(Hechizo).HechizeroMsg = GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "HechizeroMsg")
        Hechizos(Hechizo).TargetMsg = GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "TargetMsg")
        Hechizos(Hechizo).PropioMsg = GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "PropioMsg")


        Hechizos(Hechizo).Tipo = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Tipo"))
        Hechizos(Hechizo).WAV = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "WAV"))
        Hechizos(Hechizo).FXgrh = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Fxgrh"))

        Hechizos(Hechizo).loops = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Loops"))

        Hechizos(Hechizo).Resis = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Resis"))
        Hechizos(Hechizo).Baculo = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Baculo"))

        Hechizos(Hechizo).SubeHP = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeHP"))
        Hechizos(Hechizo).MinHP = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinHP"))
        Hechizos(Hechizo).MaxHP = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxHP"))

        Hechizos(Hechizo).SubeMana = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeMana"))
        Hechizos(Hechizo).MiMana = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinMana"))
        Hechizos(Hechizo).MaMana = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxMana"))

        Hechizos(Hechizo).SubeSta = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeSta"))
        Hechizos(Hechizo).MinSta = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinSta"))
        Hechizos(Hechizo).MaxSta = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxSta"))

        Hechizos(Hechizo).SubeHam = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeHam"))
        Hechizos(Hechizo).MinHam = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinHam"))
        Hechizos(Hechizo).MaxHam = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxHam"))

        Hechizos(Hechizo).SubeSed = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeSed"))
        Hechizos(Hechizo).MinSed = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinSed"))
        Hechizos(Hechizo).MaxSed = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxSed"))

        Hechizos(Hechizo).SubeAgilidad = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeAG"))
        Hechizos(Hechizo).MinAgilidad = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinAG"))
        Hechizos(Hechizo).MaxAgilidad = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxAG"))
        Hechizos(Hechizo).EffectIndex = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "EffectIndex"))
        Hechizos(Hechizo).SubeFuerza = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeFU"))
        Hechizos(Hechizo).MinFuerza = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinFU"))
        Hechizos(Hechizo).MaxFuerza = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxFU"))

        Hechizos(Hechizo).SubeCarisma = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeCA"))
        Hechizos(Hechizo).MinCarisma = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinCA"))
        Hechizos(Hechizo).MaxCarisma = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxCA"))

        Hechizos(Hechizo).Invisibilidad = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Invisibilidad"))
        Hechizos(Hechizo).Paraliza = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Paraliza"))

        Hechizos(Hechizo).Transforma = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Transforma"))
        Hechizos(Hechizo).Envenena = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Envenena"))
        Hechizos(Hechizo).Ceguera = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Ceguera"))
        Hechizos(Hechizo).Estupidez = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Estupidez"))

        Hechizos(Hechizo).Revivir = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Revivir"))
        Hechizos(Hechizo).Flecha = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Flecha"))

        Hechizos(Hechizo).Metamorfosis = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Metamorfosis"))
        Hechizos(Hechizo).Maldicion = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Maldicion"))
        Hechizos(Hechizo).Bendicion = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Bendicion"))

        Hechizos(Hechizo).RemoverParalisis = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "RemoverParalisis"))
        Hechizos(Hechizo).CuraVeneno = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "CuraVeneno"))
        Hechizos(Hechizo).RemoverMaldicion = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "RemoverMaldicion"))

        Hechizos(Hechizo).Invoca = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Invoca"))
        Hechizos(Hechizo).NumNPC = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "NumNpc"))
        Hechizos(Hechizo).Cant = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Cant"))

        Hechizos(Hechizo).Materializa = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Materializa"))
        Hechizos(Hechizo).ItemIndex = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "ItemIndex"))

        Hechizos(Hechizo).Nivel = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Nivel"))
        Hechizos(Hechizo).MinSkill = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "MinSkill"))
        Hechizos(Hechizo).ManaRequerido = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "ManaRequerido"))
        Hechizos(Hechizo).StaRequerido = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "StaRequerido"))
        Hechizos(Hechizo).Especial = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Especial"))

        Hechizos(Hechizo).Target = val(GetVar(DatPath & "Hechizos.dat", "Hechizo" & Hechizo, "Target"))
        frmCargando.cargar.value = frmCargando.cargar.value + 1
    Next

    Exit Sub

errhandler:
    Call LogErrorUrgente("Error cargando Hechizos.dat -" & Err.Description & "-" & Hechizo)
End Sub
Sub LoadMotd()
    Dim i As Integer

    DiasSinLluvia = val(GetVar(DatPath & "lluvia.dat", "INIT", "DiasSinLLuvia"))

    MaxLines = val(GetVar(App.path & "\Dat\Motd.ini", "INIT", "NumLines"))

    ReDim MOTD(0 To MaxLines)

    For i = 1 To MaxLines
        MOTD(i).Texto = GetVar(App.path & "\Dat\Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = ""
    Next

End Sub

'soportes Dylan.-
Sub SaveSoportes()
    On Error Resume Next
    Dim Num As Integer

    Kill DatPath & "soportes.dat"
    Call WriteVar(DatPath & "soportes.dat", "INIT", "Numero", SoporteS.Count)

    For Num = 1 To SoporteS.Count
        Call WriteVar(DatPath & "soportes.dat", "INIT", "SOPORTE" & Num, SoporteS.Item(Num))
    Next
End Sub

Sub LoadSoportes()
    Dim i, SoportesX As Integer
    If Not FileExist(DatPath & "soportes.dat", vbNormal) Then Exit Sub
    For i = 1 To SoporteS.Count
        Call SoporteS.Remove(1)
    Next

    SoportesX = val(GetVar(DatPath & "soportes.dat", "INIT", "Numero"))

    For i = 1 To SoportesX
        Call SoporteS.Add(GetVar(DatPath & "soportes.dat", "INIT", "SOPORTE" & i))
    Next
End Sub
'soportes dylan.-


Sub SaveBans()
    Dim Num As Integer

    Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)

    For Num = 1 To Baneos.Count
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & Num, "USER", Baneos(Num).name)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & Num, "BANEADOR", Baneos(Num).Baneador)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & Num, "CAUSA", Baneos(Num).Causa)
    Next

End Sub
Sub SaveBan(Num As Integer)

    Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & Num, "USER", Baneos(Num).name)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & Num, "BANEADOR", Baneos(Num).Baneador)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & Num, "CAUSA", Baneos(Num).Causa)

End Sub
Sub LoadBans()
    Dim BaneosTemporales As Integer
    Dim tBan As tBaneo, i As Integer
    Dim NumHds As Integer
    Dim NumIps As Integer

    NumHds = val(GetVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad"))
    If NumHds > 0 Then
        For i = 1 To NumHds
            BanHDs.Add GetVar(App.path & "\Logs\BanHDs.dat", "BANS", "Disco" & i)
        Next
    End If

    NumIps = val(GetVar(App.path & "\BanIPs.txt", "INIT", "Cantidad"))
    If NumIps > 0 Then
        For i = 1 To NumIps
            BanIps.Add GetVar(App.path & "\BanIPs.txt", "BANS", "IP" & i)
        Next
    End If

    If Not FileExist(DatPath & "baneos.dat", vbNormal) Then Exit Sub

    BaneosTemporales = val(GetVar(DatPath & "baneos.dat", "INIT", "NumeroBans"))

    For i = 1 To BaneosTemporales
        Set tBan = New tBaneo
        With tBan
            .name = GetVar(DatPath & "baneos.dat", "BANEO" & i, "USER")
            '.FechaLiberacion = GetVar(DatPath & "baneos.dat", "BANEO" & i, "FECHA")
            .Causa = GetVar(DatPath & "baneos.dat", "BANEO" & i, "CAUSA")
            .Baneador = GetVar(DatPath & "baneos.dat", "BANEO" & i, "BANEADOR")

            Call Baneos.Add(tBan)
        End With
    Next

End Sub
Public Sub ChekearNPCs()
    Dim Map As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Try As Integer

    For Map = 1 To NumMaps
        For i = 1 To UBound(MapInfo(Map).NPCsTeoricos)
            If MapInfo(Map).NPCsTeoricos(i).Numero > 0 And MapInfo(Map).NPCsTeoricos(i).Cantidad > MapInfo(Map).NPCsReales(i).Cantidad Then
                Do Until MapInfo(Map).NPCsTeoricos(i).Cantidad = MapInfo(Map).NPCsReales(i).Cantidad Or Try >= 100
                    Call CrearNPC(MapInfo(Map).NPCsTeoricos(i).Numero, Map, Npclist(1).Orig)
                    Try = Try + 1
                Loop
                Try = 0
            Else: Exit For
            End If
        Next
    Next

End Sub
Public Sub SaveGuildsNew()
    On Error GoTo errhandler
    Dim j As Integer, file As String, i As Integer

    file = App.path & "\Guilds\" & "GuildsInfo.inf"

    Call WriteVar(file, "INIT", "NroGuilds", str(Guilds.Count))

    For i = 1 To Guilds.Count
        Call WriteVar(file, "GUILD" & i, "GuildName", Guilds(i).GuildName)
        Call WriteVar(file, "GUILD" & i, "Founder", Guilds(i).Founder)
        Call WriteVar(file, "GUILD" & i, "Date", Guilds(i).FundationDate)
        Call WriteVar(file, "GUILD" & i, "CVCSGANADOS", Guilds(i).CVCsGanados)
        Call WriteVar(file, "GUILD" & i, "Desc", Guilds(i).Description)
        Call WriteVar(file, "GUILD" & i, "Codex", Guilds(i).Codex)
        Call WriteVar(file, "GUILD" & i, "Leader", Guilds(i).Leader)
        Call WriteVar(file, "GUILD" & i, "URL", Guilds(i).URL)
        Call WriteVar(file, "GUILD" & i, "GuildExp", str(Guilds(i).GuildExperience))
        Call WriteVar(file, "GUILD" & i, "DaysLast", str(Guilds(i).DaysSinceLastElection))
        Call WriteVar(file, "GUILD" & i, "GuildNews", Guilds(i).GuildNews)
        Call WriteVar(file, "GUILD" & i, "Bando", str(Guilds(i).Bando))



        Call WriteVar(file, "GUILD" & i, "NumAliados", Guilds(i).AlliedGuilds.Count)

        For j = 1 To Guilds(i).AlliedGuilds.Count
            Call WriteVar(file, "GUILD" & i, "Aliado" & j, Guilds(i).AlliedGuilds(j))
        Next

        Call WriteVar(file, "GUILD" & i, "NumEnemigos", Guilds(i).EnemyGuilds.Count)

        For j = 1 To Guilds(i).EnemyGuilds.Count
            Call WriteVar(file, "GUILD" & i, "Enemigo" & j, Guilds(i).EnemyGuilds(j))
        Next

        Call WriteVar(file, "GUILD" & i, "NumMiembros", Guilds(i).Members.Count)

        For j = 1 To Guilds(i).Members.Count
            Call WriteVar(file, "GUILD" & i, "Miembro" & j, Guilds(i).Members(j))
        Next

        Call WriteVar(file, "GUILD" & i, "NumSolicitudes", Guilds(i).Solicitudes.Count)

        For j = 1 To Guilds(i).Solicitudes.Count
            Call WriteVar(file, "GUILD" & i, "Solicitud" & j, Guilds(i).Solicitudes(j).UserName & "¬" & Guilds(i).Solicitudes(j).Desc)
        Next

        Call WriteVar(file, "GUILD" & i, "NumProposiciones", Guilds(i).PeacePropositions.Count)

        For j = 1 To Guilds(i).PeacePropositions.Count
            Call WriteVar(file, "GUILD" & i, "Proposicion" & j, Guilds(i).PeacePropositions(j).UserName & "¬" & Guilds(i).PeacePropositions(j).Desc)
        Next
    Next

    Exit Sub

errhandler:
    Call LogError("Error en SaveGuildsNew: " & Err.Description & "-Clan: " & i & "-" & j)

End Sub
Public Sub DarPremioCastillos()
    On Error GoTo handler

    Dim LoopC As Integer

    For LoopC = 1 To LastUser
        If UserList(LoopC).ConnID > -1 And UserList(LoopC).GuildInfo.GuildName = CastilloNorte Then
            UserList(LoopC).Faccion.Quests = UserList(LoopC).Faccion.Quests + 3
            Call SendData(ToIndex, (LoopC), 0, "||Has Recibido 5 puntos de canje por mantener el castillo Norte." & FONTTYPE_FENIX)
            Call SendUserStatsBox(LoopC)
        End If

    Next LoopC

    For LoopC = 1 To LastUser
        If UserList(LoopC).ConnID > -1 And UserList(LoopC).GuildInfo.GuildName = CastilloSur Then
            UserList(LoopC).Faccion.Quests = UserList(LoopC).Faccion.Quests + 3
            Call SendData(ToIndex, (LoopC), 0, "||Has Recibido 5 puntos de canje por mantener el castillo Sur." & FONTTYPE_FENIX)
            Call SendUserStatsBox(LoopC)
        End If

    Next LoopC

    Exit Sub
handler:
    Call LogError("Error en DarPremioCastillos.")
End Sub
Public Sub DoBackUp(Optional Guilds As Boolean)

    haciendoBK = True

    Call SendData(ToAll, 0, 0, "2P")
    Call SendData(ToAll, 0, 0, "BKW")
    Call SaveSoportes

    If Guilds Then Call SaveGuildsNew
    Call WorldSave
    'BETA
    'Call ChekearNPCs
    Call GuardarUsuarios

    Call SendData(ToAll, 0, 0, "BKW")

    haciendoBK = False

End Sub

Public Sub SaveMapData(ByVal N As Integer)
    Dim SaveAs As String
    Dim Y As Byte
    Dim X As Byte

    SaveAs = App.path & "\WorldBackUP\Map" & N & ".bkp"

    If FileExist(SaveAs, vbNormal) Then Kill SaveAs

    Open SaveAs For Binary As #1
    Seek #1, 1

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(N, X, Y).OBJInfo.OBJIndex Then
                If ObjData(MapData(N, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_FOGATA Then
                    MapData(N, X, Y).OBJInfo.OBJIndex = 0
                    MapData(N, X, Y).OBJInfo.Amount = 0
                ElseIf Not ItemEsDeMapa(N, CInt(X), CInt(Y)) Then
                    Put #1, , X
                    Put #1, , Y
                    Put #1, , MapData(N, X, Y).OBJInfo.OBJIndex
                    Put #1, , MapData(N, X, Y).OBJInfo.Amount
                End If
            End If

        Next
    Next
    Put #1, , CByte(100)
    Close #1

End Sub
Sub LoadArmasHerreria()
    Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

    ReDim Preserve ArmasHerrero(1 To N) As InfoHerre

    For lc = 1 To N
        ArmasHerrero(lc).Index = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
        ArmasHerrero(lc).Recompensa = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Recompensa"))
    Next

End Sub
Sub LoadArmadurasHerreria()
    Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

    ReDim Preserve ArmadurasHerrero(1 To N) As InfoHerre

    For lc = 1 To N
        ArmadurasHerrero(lc).Index = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
        ArmadurasHerrero(lc).Recompensa = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Recompensa"))
    Next

End Sub

Sub LoadEscudosHerreria()

    Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "EscudosHerrero.dat", "INIT", "NumEscudos"))

    ReDim Preserve EscudosHerrero(1 To N) As Integer

    For lc = 1 To N
        EscudosHerrero(lc) = val(GetVar(DatPath & "EscudosHerrero.dat", "Escudo" & lc, "Index"))
    Next

End Sub
Sub LoadCascosHerreria()

    Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "CascosHerrero.dat", "INIT", "NumCascos"))

    ReDim Preserve CascosHerrero(1 To N) As Integer

    For lc = 1 To N
        CascosHerrero(lc) = val(GetVar(DatPath & "CascosHerrero.dat", "Casco" & lc, "Index"))
    Next

End Sub
Sub LoadObjCarpintero()
    Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

    ReDim Preserve ObjCarpintero(1 To N) As InfoHerre

    For lc = 1 To N
        ObjCarpintero(lc).Index = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
        ObjCarpintero(lc).Recompensa = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Recompensa"))
    Next

End Sub



Sub LoadObjSastre()

    Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

    ReDim Preserve ObjSastre(1 To N) As Integer

    For lc = 1 To N
        ObjSastre(lc) = val(GetVar(DatPath & "ObjSastre.dat", "Obj" & lc, "Index"))
    Next

End Sub
Sub LoadOBJData()
    On Error Resume Next
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."



    Dim UserIndex As Integer
    Dim Object As Integer

    Dim a As Long, S As Long

    a = INICarga(DatPath & "Obj.dat")
    Call INIConf(a, 0, "", 0)



    S = INIBuscarSeccion(a, "INIT")
    NumObjDatas = INIDarClaveInt(a, S, "NumOBJs") + 2

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.value = 0


    ReDim ObjData(0 To NumObjDatas) As ObjData


    For Object = 1 To NumObjDatas

        S = INIBuscarSeccion(a, "OBJ" & Object)

        If S >= 0 Then
            ObjData(Object).name = INIDarClaveStr(a, S, "Name")
            ObjData(Object).NoComerciable = INIDarClaveInt(a, S, "NoComerciable")

            ObjData(Object).GrhIndex = INIDarClaveInt(a, S, "GrhIndex")


            ObjData(Object).NoSeCae = INIDarClaveInt(a, S, "NoSeCae") = 1
            ObjData(Object).ObjType = INIDarClaveInt(a, S, "ObjType")
            ObjData(Object).ArbolElfico = INIDarClaveInt(a, S, "ArbolElfico")
            ObjData(Object).SubTipo = INIDarClaveInt(a, S, "Subtipo")
            ObjData(Object).Dosmanos = INIDarClaveInt(a, S, "Dosmanos")
            ObjData(Object).Newbie = INIDarClaveInt(a, S, "Newbie")
            ObjData(Object).Aura = INIDarClaveInt(a, S, "Aura")

            ObjData(Object).SkPociones = INIDarClaveInt(a, S, "SkPociones")
            ObjData(Object).SkSastreria = INIDarClaveInt(a, S, "SkSastreria")
            ObjData(Object).Raices = INIDarClaveInt(a, S, "Raices")
            ObjData(Object).PielLobo = INIDarClaveInt(a, S, "PielLobo")
            ObjData(Object).PielOsoPardo = INIDarClaveInt(a, S, "PielOsoPardo")
            ObjData(Object).PielOsoPolar = INIDarClaveInt(a, S, "PielOsoPolar ")

            If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
                ObjData(Object).ShieldAnim = INIDarClaveInt(a, S, "Anim")
                ObjData(Object).LingH = INIDarClaveInt(a, S, "LingH")
                ObjData(Object).LingP = INIDarClaveInt(a, S, "LingP")
                ObjData(Object).LingO = INIDarClaveInt(a, S, "LingO")

                ObjData(Object).SkHerreria = INIDarClaveInt(a, S, "SkHerreria")
            End If

            If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
                ObjData(Object).CascoAnim = INIDarClaveInt(a, S, "Anim")
                ObjData(Object).LingH = INIDarClaveInt(a, S, "LingH")
                ObjData(Object).Gorro = INIDarClaveInt(a, S, "Gorro")
                ObjData(Object).LingP = INIDarClaveInt(a, S, "LingP")
                ObjData(Object).LingO = INIDarClaveInt(a, S, "LingO")
                ObjData(Object).SkHerreria = INIDarClaveInt(a, S, "SkHerreria")

            End If

            ObjData(Object).Ropaje = INIDarClaveInt(a, S, "NumRopaje")
            ObjData(Object).HechizoIndex = INIDarClaveInt(a, S, "HechizoIndex")

            If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
                ObjData(Object).Baculo = INIDarClaveInt(a, S, "Baculo")
                ObjData(Object).WeaponAnim = INIDarClaveInt(a, S, "Anim")
                ObjData(Object).Apuñala = INIDarClaveInt(a, S, "Apuñala")
                ObjData(Object).Envenena = INIDarClaveInt(a, S, "Envenena")
                ObjData(Object).MaxHit = INIDarClaveInt(a, S, "MaxHIT")
                ObjData(Object).MinHit = INIDarClaveInt(a, S, "MinHIT")
                ObjData(Object).LingH = INIDarClaveInt(a, S, "LingH")
                ObjData(Object).LingP = INIDarClaveInt(a, S, "LingP")
                ObjData(Object).LingO = INIDarClaveInt(a, S, "LingO")
                ObjData(Object).SkHerreria = INIDarClaveInt(a, S, "SkHerreria")
                ObjData(Object).Real = INIDarClaveInt(a, S, "Real")
                ObjData(Object).Caos = INIDarClaveInt(a, S, "Caos")
                ObjData(Object).proyectil = INIDarClaveInt(a, S, "Proyectil")
                ObjData(Object).Municion = INIDarClaveInt(a, S, "Municiones")

            End If

            If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
                ObjData(Object).LingH = INIDarClaveInt(a, S, "LingH")
                ObjData(Object).LingP = INIDarClaveInt(a, S, "LingP")
                ObjData(Object).LingO = INIDarClaveInt(a, S, "LingO")
                ObjData(Object).SkHerreria = INIDarClaveInt(a, S, "SkHerreria")
                ObjData(Object).Real = INIDarClaveInt(a, S, "Real")
                ObjData(Object).Caos = INIDarClaveInt(a, S, "Caos")
                ObjData(Object).Jerarquia = INIDarClaveInt(a, S, "Jerarquia")

            End If

            If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
                ObjData(Object).LingH = INIDarClaveInt(a, S, "LingH")
                ObjData(Object).LingP = INIDarClaveInt(a, S, "LingP")
                ObjData(Object).LingO = INIDarClaveInt(a, S, "LingO")
                ObjData(Object).SkHerreria = INIDarClaveInt(a, S, "SkHerreria")
            End If

            If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
                ObjData(Object).Snd1 = INIDarClaveInt(a, S, "SND1")
                ObjData(Object).Snd2 = INIDarClaveInt(a, S, "SND2")
                ObjData(Object).Snd3 = INIDarClaveInt(a, S, "SND3")
                ObjData(Object).MinInt = INIDarClaveInt(a, S, "MinInt")
            End If

            ObjData(Object).LingoteIndex = INIDarClaveInt(a, S, "LingoteIndex")

            If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
                ObjData(Object).MinSkill = INIDarClaveInt(a, S, "MinSkill")
            End If

            ObjData(Object).MineralIndex = INIDarClaveInt(a, S, "MineralIndex")

            ObjData(Object).MaxHP = INIDarClaveInt(a, S, "MaxHP")
            ObjData(Object).MinHP = INIDarClaveInt(a, S, "MinHP")

            ObjData(Object).MUJER = INIDarClaveInt(a, S, "Mujer")
            ObjData(Object).HOMBRE = INIDarClaveInt(a, S, "Hombre")

            ObjData(Object).SkillCombate = INIDarClaveInt(a, S, "SkCombate")
            ObjData(Object).SkillTacticas = INIDarClaveInt(a, S, "SkTacticas")
            ObjData(Object).SkillProyectiles = INIDarClaveInt(a, S, "SkProyectiles")
            ObjData(Object).SkillApuñalar = INIDarClaveInt(a, S, "SkApuñalar")
            ObjData(Object).SkResistencia = INIDarClaveInt(a, S, "SkResistencia")
            ObjData(Object).SkDefensa = INIDarClaveInt(a, S, "SkEscudos")

            ObjData(Object).MinHam = INIDarClaveInt(a, S, "MinHam")
            ObjData(Object).MinSed = INIDarClaveInt(a, S, "MinAgu")

            ObjData(Object).MinDef = INIDarClaveInt(a, S, "MINDEF")
            ObjData(Object).MaxDef = INIDarClaveInt(a, S, "MAXDEF")

            ObjData(Object).Respawn = INIDarClaveInt(a, S, "ReSpawn")

            ObjData(Object).RazaEnana = INIDarClaveInt(a, S, "RazaEnana")

            ObjData(Object).Valor = INIDarClaveInt(a, S, "Valor")

            ObjData(Object).Crucial = INIDarClaveInt(a, S, "Crucial")

            ObjData(Object).Cerrada = INIDarClaveInt(a, S, "abierta")

            If ObjData(Object).Cerrada = 1 Then
                ObjData(Object).Llave = INIDarClaveInt(a, S, "Llave")
                ObjData(Object).Clave = INIDarClaveInt(a, S, "Clave")
            End If


            If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).ObjType = OBJTYPE_BOTELLALLENA Then
                ObjData(Object).IndexAbierta = INIDarClaveInt(a, S, "IndexAbierta")
                ObjData(Object).IndexCerrada = INIDarClaveInt(a, S, "IndexCerrada")
                ObjData(Object).IndexCerradaLlave = INIDarClaveInt(a, S, "IndexCerradaLlave")
            End If
            If ObjData(Object).ObjType = OBJTYPE_WARP Then
                ObjData(Object).WMapa = INIDarClaveInt(a, S, "WMapa")
                ObjData(Object).WX = INIDarClaveInt(a, S, "WX")
                ObjData(Object).WY = INIDarClaveInt(a, S, "WY")
                ObjData(Object).WI = INIDarClaveInt(a, S, "WI")
            End If

            ObjData(Object).Clave = INIDarClaveInt(a, S, "Clave")

            ObjData(Object).Texto = INIDarClaveStr(a, S, "Texto")
            ObjData(Object).GrhSecundario = INIDarClaveInt(a, S, "VGrande")

            ObjData(Object).Agarrable = INIDarClaveInt(a, S, "Agarrable")
            ObjData(Object).ForoID = INIDarClaveStr(a, S, "ID")
            Dim Num As Integer

            Num = INIDarClaveInt(a, S, "NumClases")

            Dim i As Integer
            For i = 1 To Num
                ObjData(Object).ClaseProhibida(i) = INIDarClaveInt(a, S, "CP" & i)
            Next

            Num = INIDarClaveInt(a, S, "NumRazas")

            Dim d As Integer
            For d = 1 To Num
                ObjData(Object).RazaProhibida(d) = INIDarClaveInt(a, S, "RP" & d)
            Next

            ObjData(Object).Resistencia = INIDarClaveInt(a, S, "Resistencia")


            If ObjData(Object).ObjType = 11 Then
                ObjData(Object).TipoPocion = INIDarClaveInt(a, S, "TipoPocion")
                ObjData(Object).MaxModificador = INIDarClaveInt(a, S, "MaxModificador")
                ObjData(Object).MinModificador = INIDarClaveInt(a, S, "MinModificador")
                ObjData(Object).DuracionEfecto = INIDarClaveInt(a, S, "DuracionEfecto")

            End If
            ObjData(Object).SkCarpinteria = INIDarClaveInt(a, S, "SkCarpinteria")

            If ObjData(Object).SkCarpinteria Then
                ObjData(Object).Madera = INIDarClaveInt(a, S, "Madera")
                ObjData(Object).MaderaElfica = INIDarClaveInt(a, S, "MaderaElfica")
            End If

            If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
                ObjData(Object).MaxHit = INIDarClaveInt(a, S, "MaxHIT")
                ObjData(Object).MinHit = INIDarClaveInt(a, S, "MinHIT")
            End If

            If ObjData(Object).ObjType = OBJTYPE_FLECHAS Then
                ObjData(Object).MaxHit = INIDarClaveInt(a, S, "MaxHIT")
                ObjData(Object).MinHit = INIDarClaveInt(a, S, "MinHIT")
            End If

            ObjData(Object).MinSta = INIDarClaveInt(a, S, "MinST")

            frmCargando.cargar.value = frmCargando.cargar.value + 1
        End If

        DoEvents

    Next

    Call INIDescarga(a)
    Call ExtraObjs

    Exit Sub

errhandler:

    Call INIDescarga(a)

    Call LogErrorUrgente("Error cargando objetos: " & Err.Number & " : " & Err.Description)

End Sub
Function EnPantalla(wp1 As WorldPos, wp2 As WorldPos, Optional Sumar As Integer) As Boolean

    EnPantalla = (wp1.Map = wp2.Map And Abs(wp1.X - wp2.X) < MinXBorder + Sumar And Abs(wp1.Y - wp2.Y) < MinYBorder + Sumar)

End Function
Function GetVar(file As String, Main As String, Var As String) As String
    Dim sSpaces As String

    sSpaces = Space$(5000)

    getprivateprofilestring Main, Var, "", sSpaces, Len(sSpaces), file

    GetVar = RTrim(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function
Public Sub CargarBackUp()
    Dim Map As Integer
    Dim Load As String
    Dim Y As Byte
    Dim X As Byte

    For Map = 1 To NumMaps
        Load = App.path & "\WorldBackUP\Map" & Map & ".bkp"

        If FileExist(Load, vbNormal) Then
            Open Load For Binary As #1
            Seek #1, 1

            Do
                Get #1, , X
                If X = 100 Then Exit Do
                Get #1, , Y
                Get #1, , MapData(Map, X, Y).OBJInfo.OBJIndex
                Get #1, , MapData(Map, X, Y).OBJInfo.Amount
            Loop

            Close #1
        End If
    Next

End Sub
Sub Congela(Optional ByVal Descongelar As Boolean)

    If Descongelar Then
        Call SendData(ToAll, 0, 0, "°¬")
    Else: Call SendData(ToAll, 0, 0, "°°")
    End If

End Sub
Sub LoadMapDats()
    On Error GoTo Error
    Dim a As Long, S As Long, i As Integer

    a = INICarga(MapDatFile)
    Call INIConf(a, 0, "", 0)

    S = INIBuscarSeccion(a, "INIT")
    NumMaps = INIDarClaveInt(a, S, "NumMaps")

    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo

    For i = 1 To NumMaps

        S = INIBuscarSeccion(a, "Mapa" & i)
        If S > 0 Then
            MapInfo(i).name = INIDarClaveStr(a, S, "Name")
            MapInfo(i).Music = INIDarClaveStr(a, S, "MusicNum")

            MapInfo(i).TopPunto = INIDarClaveInt(a, S, "TopPunto")
            MapInfo(i).LeftPunto = INIDarClaveInt(a, S, "LeftPunto")

            MapInfo(i).Pk = (INIDarClaveInt(a, S, "Pk") = 0)
            MapInfo(i).NoMagia = (INIDarClaveInt(a, S, "NoMagia") = 1)

            MapInfo(i).Terreno = INIDarClaveStr(a, S, "Terreno")
            MapInfo(i).Zona = INIDarClaveStr(a, S, "Zona")

            MapInfo(i).Restringir = (INIDarClaveInt(a, S, "Restringir") = 1)
            MapInfo(i).Nivel = INIDarClaveInt(a, S, "Nivel")

            MapInfo(i).BackUp = INIDarClaveInt(a, S, "BackUp")
        End If
    Next
    Exit Sub

Error:
    Call LogErrorUrgente("Error cargando Info.dat-" & Err.Description & "-" & i)
End Sub
Sub LoadMapDataNew()
    On Error GoTo man
    Dim Map As Integer
    Dim LoopC As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim DummyInt As Integer
    Dim TempInt As Integer
    Dim npcfile As String
    Dim InfoTile As Byte
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas."

    Call LoadMapDats

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0

    For Map = 1 To NumMaps
        DoEvents

        Debug.Print Round(Map / NumMaps * 100, 2) & "%"
        frmCargando.Label1(2).Caption = "Cargando mapas... " & Map & "/" & NumMaps

        Open MapPath & "Mapa" & Map & ".msv" For Binary As #1
        Seek #1, 1

        Get #1, , MapInfo(Map).MapVersion

        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize

                Get #1, , InfoTile

                MapData(Map, X, Y).Blocked = (InfoTile And 1)
                MapData(Map, X, Y).Agua = Buleano(InfoTile And 2)

                For LoopC = 2 To 4
                    If (InfoTile And 2 ^ LoopC) Then MapData(Map, X, Y).trigger = MapData(Map, X, Y).trigger Or 2 ^ (LoopC - 2)
                Next

                If InfoTile And 32 Then
                    Get #1, , MapData(Map, X, Y).NpcIndex

                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)

                    If MapData(Map, X, Y).NpcIndex >= 500 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                    Else: npcfile = DatPath & "NPCs.dat"
                    End If

                    Dim fl As Byte

                    If Npclist(MapData(Map, X, Y).NpcIndex).flags.RespawnOrigPos Then
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                    End If

                    Npclist(MapData(Map, X, Y).NpcIndex).pos.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).pos.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).pos.Y = Y


                    If Npclist(MapData(Map, X, Y).NpcIndex).Attackable = 1 And Npclist(MapData(Map, X, Y).NpcIndex).flags.Respawn = 0 Then
                        Call AgregarNPCTeorico(Npclist(MapData(Map, X, Y).NpcIndex).Numero, Map)
                        Call AgregarNPC(Npclist(MapData(Map, X, Y).NpcIndex).Numero, Map)
                    End If

                    Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)

                End If

                If InfoTile And 64 Then
                    Get #1, , MapData(Map, X, Y).OBJInfo.OBJIndex
                    Get #1, , MapData(Map, X, Y).OBJInfo.Amount
                End If

                If MapData(Map, X, Y).OBJInfo.OBJIndex > UBound(ObjData) Then
                    MapData(Map, X, Y).OBJInfo.OBJIndex = 0
                    MapData(Map, X, Y).OBJInfo.Amount = 0
                End If

                If InfoTile And 128 Then
                    Get #1, , MapData(Map, X, Y).TileExit.Map
                    Get #1, , MapData(Map, X, Y).TileExit.X
                    Get #1, , MapData(Map, X, Y).TileExit.Y
                End If

            Next
        Next

        Close #1
        Close #2

        frmCargando.cargar.value = frmCargando.cargar.value + 1

        Dim i As Integer

        Dim nfile As Integer
        nfile = FreeFile
        If MapInfo(Map).NPCsTeoricos(1).Numero Then

            Open App.path & "\Logs\NPCs.log" For Append Shared As #nfile
            Print #nfile, "Mapa " & Map & ": " & MapInfo(Map).name
            For i = 1 To 20
                If MapInfo(Map).NPCsTeoricos(i).Numero Then
                    Print #nfile, MapInfo(Map).NPCsTeoricos(i).Cantidad & " " & NameNpc(MapInfo(Map).NPCsTeoricos(i).Numero)

                Else: Exit For
                End If
            Next
            Print #nfile, ""
            Close #nfile
        End If
    Next

    Exit Sub

man:
    Call LogErrorUrgente("Error durante carga de mapas: " & Map & "-" & X & "-" & Y)
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)


End Sub
Sub LoadSini()
    Dim Temporal As Long
    Dim Temporal1 As Long
    Dim LoopC As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim L As Integer

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

    BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))
    Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
    LastSockListen = val(GetVar(IniPath & "Server.ini", "INIT", "LastSockListen"))
    HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
    MySql = val(GetVar(IniPath & "Server.ini", "INIT", "MySql"))
    AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
    IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))

    UltimaVersion = GetVar(IniPath & "Server.ini", "INIT", "Version")
    AUVersion = GetVar(IniPath & "Server.ini", "INIT", "AUVersion")

    PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
    TiempoPremium = val(GetVar(IniPath & "Server.ini", "PREMIUM", "TiempoPremium"))
    CantidadOro = val(GetVar(IniPath & "Server.ini", "INIT", "ORO"))
    CantidadEXP = val(GetVar(IniPath & "Server.ini", "INIT", "EXP"))

    ClientsCommandsQueue = val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))


    SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

    StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

    SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

    StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

    IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
    FrmInterv.txtIntervaloSed.Text = IntervaloSed

    IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
    FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

    IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
    FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

    IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
    FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

    IntervaloParalizadoUsuario = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizadoUsuario"))

    IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
    FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

    IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
    FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

    IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
    FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

    IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
    FrmInterv.txtInvocacion.Text = IntervaloInvocacion

    IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
    FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion


    MaxUsers2 = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers2"))

    IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo")) / 10
    FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))

    IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores")) / 10

    IntervaloUserPuedePocion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedePocion")) / 100
    IntervaloUserPuedePocionC = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedePocionC")) / 100

    IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar")) / 10
    FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

    IntervaloUserFlechas = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserFlechas")) / 10
    IntervaloUserSH = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserSH"))

    IntervaloUserPuedeGolpeHechi = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeGolpeHechi")) / 10
    IntervaloUserPuedeHechiGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeHechiGolpe")) / 10

    IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))


    ResPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
    ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
    ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))

    recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))


    MaxUsers = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))

    ReDim UserList(1 To MaxUsers) As User
    ReDim Party(1 To (MaxUsers / 2)) As Party

    Althalos.Map = GetVar(DatPath & "Ciudades.dat", "Althalos", "Mapa")
    Althalos.X = GetVar(DatPath & "Ciudades.dat", "Althalos", "X")
    Althalos.Y = GetVar(DatPath & "Ciudades.dat", "Althalos", "Y")

    Hildegard.Map = GetVar(DatPath & "Ciudades.dat", "Hildegard", "Mapa")
    Hildegard.X = GetVar(DatPath & "Ciudades.dat", "Hildegard", "X")
    Hildegard.Y = GetVar(DatPath & "Ciudades.dat", "Hildegard", "Y")

    Lonelerd.Map = GetVar(DatPath & "Ciudades.dat", "Lonelerd", "Mapa")
    Lonelerd.X = GetVar(DatPath & "Ciudades.dat", "Lonelerd", "X")
    Lonelerd.Y = GetVar(DatPath & "Ciudades.dat", "Lonelerd", "Y")

    'LINDOS.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    'LINDOS.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    'LINDOS.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

    ADELAIDE.Map = GetVar(DatPath & "Ciudades.dat", "ADELAIDE", "Mapa")
    ADELAIDE.X = GetVar(DatPath & "Ciudades.dat", "ADELAIDE", "X")
    ADELAIDE.Y = GetVar(DatPath & "Ciudades.dat", "ADELAIDE", "Y")

    ReDim Hush(val(GetVar(IniPath & "Server.ini", "Hash", "HashAceptados")))
    For LoopC = 0 To UBound(Hush)
        Hush(LoopC) = GetVar(IniPath & "Server.ini", "Hash", "HashAceptado" & (LoopC + 1))
    Next
    LoadAntiCheatZ
End Sub
Sub WriteVar(file As String, Main As String, Var As String, value As String)

    writeprivateprofilestring Main, Var, value, file

End Sub
Sub BackUPnPc(NpcIndex As Integer)



    Dim NpcNumero As Integer
    Dim npcfile As String
    Dim LoopC As Integer


    NpcNumero = Npclist(NpcIndex).Numero

    If NpcNumero > 499 Then
        npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    Else
        npcfile = DatPath & "bkNPCs.dat"
    End If


    Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).name)
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.Heading))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "InmuneParalisis", val(Npclist(NpcIndex).InmuneParalisis))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "VeInvis", val(Npclist(NpcIndex).VeInvis))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Pocaparalisis", val(Npclist(NpcIndex).flags.PocaParalisis))

    Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Inflacion", val(Npclist(NpcIndex).Inflacion))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))
    Call WriteVar(npcfile, "npc" & NpcNumero, "Bot", val(Npclist(NpcIndex).Bot))

    Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.Def))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHit))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHit))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))


    Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))


    Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
    If Npclist(NpcIndex).Invent.NroItems Then
        For LoopC = 1 To MAX_NPCINVENTORY_SLOTS
            Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
        Next
    End If

End Sub
Sub CargarNpcBackUp(NpcIndex As Integer, NPCNumber As Integer, UserIndex As Integer)
    On Local Error Resume Next
    Dim npcfile As String

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

    If NPCNumber >= 500 Then
        npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    Else: npcfile = DatPath & "bkNPCs.dat"
    End If

    Npclist(NpcIndex).Numero = NPCNumber
    Npclist(NpcIndex).name = GetVar(npcfile, "NPC" & NPCNumber, "Name")
    Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NPCNumber, "Desc")
    Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NPCNumber, "Movement"))
    Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NPCNumber, "NpcType"))

    Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NPCNumber, "Body"))
    Npclist(NpcIndex).Char.Aura = val(GetVar(npcfile, "NPC" & NPCNumber, "Aura"))

    Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NPCNumber, "Head"))
    Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NPCNumber, "Heading"))

    Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NPCNumber, "Attackable"))
    Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NPCNumber, "Comercia"))
    Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NPCNumber, "Hostile"))
    Npclist(NpcIndex).InmuneParalisis = val(GetVar(npcfile, "NPC" & NPCNumber, "InmuneParalisis"))
    Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NPCNumber, "GiveEXP"))

    Npclist(NpcIndex).VeInvis = val(GetVar(npcfile, "NPC" & NPCNumber, "VeInvis"))
    Npclist(NpcIndex).flags.PocaParalisis = val(GetVar(npcfile, "NPC" & NPCNumber, "pocaparalisis"))
    Npclist(NpcIndex).flags.Apostador = val(GetVar(npcfile, "NPC" & NPCNumber, "Apostador"))

    Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NPCNumber, "GiveGLD"))

    Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NPCNumber, "InvReSpawn"))
    Npclist(NpcIndex).Bot = val(GetVar(npcfile, "npc" & NPCNumber, "Bot"))


    Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NPCNumber, "MaxHP"))
    Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NPCNumber, "MinHP"))
    Npclist(NpcIndex).AutoCurar = val(GetVar(npcfile, "NPC" & NPCNumber, "autocurar"))

    Npclist(NpcIndex).Stats.MaxHit = val(GetVar(npcfile, "NPC" & NPCNumber, "MaxHIT"))
    Npclist(NpcIndex).Stats.MinHit = val(GetVar(npcfile, "NPC" & NPCNumber, "MinHIT"))
    Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NPCNumber, "DEF"))
    Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NPCNumber, "Alineacion"))
    Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NPCNumber, "ImpactRate"))


    Dim LoopC As Integer
    Dim ln As String
    Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NPCNumber, "NROITEMS"))
    If Npclist(NpcIndex).Invent.NroItems Then
        For LoopC = 1 To MAX_NPCINVENTORY_SLOTS
            ln = GetVar(npcfile, "NPC" & NPCNumber, "Obj" & LoopC)
            Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
            Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
        Next
    Else
        For LoopC = 1 To MAX_NPCINVENTORY_SLOTS
            Npclist(NpcIndex).Invent.Object(LoopC).OBJIndex = 0
            Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
        Next
    End If

    Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NPCNumber, "Inflacion"))


    Npclist(NpcIndex).flags.NPCActive = True
    Npclist(NpcIndex).flags.UseAINow = False
    Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NPCNumber, "ReSpawn"))
    Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NPCNumber, "Domable"))
    Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NPCNumber, "OrigPos"))


    Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NPCNumber, "TipoItems"))

End Sub
Sub LogBan(ByVal BannedIndex As Integer, UserIndex As Integer, ByVal Motivo As String)

    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).name, "BannedBy", UserList(UserIndex).name)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).name, "Reason", Motivo)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).name, "IP", UserList(BannedIndex).ip)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).name, "Mail", UserList(BannedIndex).Email)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).name, "Fecha", Format(Now, "dd/mm/yy hh:mm:ss"))

    Dim mifile As Integer
    mifile = FreeFile
    Open App.path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).name
    Close #mifile

End Sub
Sub LogBanOffline(ByVal BannedIndex As String, UserIndex As Integer, ByVal Motivo As String)

    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedIndex, "BannedBy", UserList(UserIndex).name)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedIndex, "Reason", Motivo)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedIndex, "IP", "Ban offline")

    Dim mifile As Integer
    mifile = FreeFile
    Open App.path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedIndex
    Close #mifile

End Sub
Function Criminal()


End Function
