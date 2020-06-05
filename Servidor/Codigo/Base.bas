Attribute VB_Name = "Base"


Public Function CuentaExiste(ByVal Cuenta As String) As Boolean

    CuentaExiste = FileExist(App.path & "\Accounts\" & UCase$(Cuenta) & ".act", vbNormal)

End Function
Public Function ChangePos(UserName As String) As Boolean
    Call WriteVar(CharPath & name & ".chr", "INIT", "Position", Althalos.Map & "-" & Althalos.X & "-" & Althalos.Y)
End Function
Public Function ChangeBan(ByVal name As String, ByVal Baneado As Integer) As Boolean
    Dim Orden As String
    'If GetVar(CharPath & Name & ".chr", "FLAGS", "Ban") <> "0" Then
    '    Call SendData(ToIndex, UserIndex, 0, "||El personaje ya se encuentra baneado." & FONTTYPE_INFO)
    '    Exit Sub
    'End If
    If Baneado = 1 Then
        Call WriteVar(CharPath & name & ".chr", "FLAGS", "Ban", 1)
    Else
        Call WriteVar(CharPath & name & ".chr", "FLAGS", "Ban", 0)
    End If

End Function

Public Sub SendCharInfo(ByVal UserName As String, UserIndex As Integer)
    Dim Data As String
    '¿Existe el personaje?

    If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub


    Dim UserFile As String
    UserFile = CharPath & UCase$(UserName) & ".chr"

    If FileExist(UserFile, vbNormal) = False Then Exit Sub
    Data = "CHRINFO" & UserName
    Data = Data & "," & ListaRazas(val(GetVar(UserFile, "INIT", "Raza"))) & "," & ListaClases(val(GetVar(UserFile, "INIT", "Clase"))) & _
           "," & GeneroLetras(val(GetVar(UserFile, "INIT", "Genero"))) & ","
    Data = Data & val(GetVar(UserFile, "STATS", "ELV")) & "," & val(GetVar(UserFile, "STATS", "GLD")) & "," & val(GetVar(UserFile, "STATS", "BANCO")) & ","

    Data = Data & val(GetVar(UserFile, "Guild", "FundoClan")) & _
           "," & GetVar(UserFile, "Guild", "ClanFundado") & "," _
           & val(GetVar(UserFile, "Guild", "Solicitudes")) & "," _
           & val(GetVar(UserFile, "Guild", "SolicitudesRechazadas")) & "," _
           & val(GetVar(UserFile, "Guild", "VecesFueGuildLeader")) & "," _
           & val(GetVar(UserFile, "Guild", "ClanesParticipo")) & ","

    Data = Data & val(GetVar(UserFile, "FACCIONES", "Bando")) & "," & val(GetVar(UserFile, "FACCIONES", "Matados0")) & "," & val(GetVar(UserFile, "FACCIONES", "Matados1")) & "," & val(GetVar(UserFile, "FACCIONES", "Matados2"))
    Call SendData(ToIndex, UserIndex, 0, Data)

End Sub

Function CalcularTiempoSilenciado(UserIndex As Integer) As Integer

    If UserList(UserIndex).flags.Silenciado = 1 Then CalcularTiempoSilenciado = 1 + (UserList(UserIndex).Counters.TiempoSilenc - TiempoTranscurrido(UserList(UserIndex).Counters.PenaSilenc)) \ 60    'matute

End Function
Function CalcularTiempoCarcel(UserIndex As Integer) As Integer

    If UserList(UserIndex).flags.Encarcelado = 1 Then CalcularTiempoCarcel = 1 + (UserList(UserIndex).Counters.TiempoPena - TiempoTranscurrido(UserList(UserIndex).Counters.Pena)) \ 60

End Function

Public Function BANCheck(ByVal name As String) As Boolean
'If Inbaneable(Name) Then Exit Function

    BANCheck = (val(GetVar(App.path & "\charfile\" & name & ".chr", "FLAGS", "Ban")) = 1)    'Or _
                                                                                             (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "AdminBan")) = 1)

End Function
Function ExistePersonaje(name As String) As Boolean

    If FileExist(App.path & "\CHARFILE\" & UCase$(name) & ".chr", vbNormal) = True Then

        ExistePersonaje = True
    End If
End Function
Function AgregarAClan(ByVal name As String, ByVal Clan As String) As Boolean

    
        Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "GUILD", "GuildName", Clan)
        Valor = val(GetVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "GUILD", "ClanesParticipo"))
        Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "GUILD", "ClanesParticipo", Valor = Valor + 1)
        Valor = val(GetVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "GUILD", "GuildPts"))
        Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "GUILD", "GuildPts", Valor = Valor + 25)
        AgregarAClan = True

End Function
Sub RechazarSolicitud(ByVal name As String)


    Valor = GetVar(App.path & "\CHARFILE\" & UCase$(name), "GUILD", val("SolicitudesRechazadas"))
    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name), "GUILD", "SolicitudesRechazadas", Valor = Valor + 1)

End Sub

