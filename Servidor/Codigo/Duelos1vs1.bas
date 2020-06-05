Attribute VB_Name = "Duelos1vs1"
   Public Sub ComensarDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EstanDueleando = True
    UserList(UserIndex).flags.Oponentes = TIndex
    UserList(TIndex).flags.EstanDueleando = True
    Call WarpUserChar(TIndex, 5, 40, 39)
    UserList(TIndex).flags.Oponentes = UserIndex
    Call WarpUserChar(UserIndex, 5, 64, 56)
    Call SendData(ToAll, 0, 0, "||" & UserList(TIndex).Name & " y " & UserList(UserIndex).Name & " van a competir en un duelo." & FONTTYPE_TALK)
    End Sub
    Public Sub ResetDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EsperandoDuelos = False
    UserList(UserIndex).flags.Oponentes = 0
    UserList(UserIndex).flags.EstanDueleando = False
    Call WarpUserChar(UserIndex, 1, 50, 50)
    Call WarpUserChar(TIndex, 1, 50, 54)
    UserList(TIndex).flags.EsperandoDuelos = False
    UserList(TIndex).flags.Oponentes = 0
    UserList(TIndex).flags.EstanDueleando = False
    End Sub
Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer, Optional ByVal Apuesta As String)
    Call SendData(ToAll, Ganador, 0, "||" & UserList(Ganador).Name & " venció en duelo a " & UserList(Perdedor).Name & "." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
    Call SendData(ToIndex, Ganador, 0, "||Ganaste " & Apuesta & " monedas de oro." & FONTTYPE_INFO)
    UserList(Ganador).Stats.GLD = UserList(Ganador).Stats.GLD + Apuesta
End Sub
    Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||El duelo ha sido cancelado. El perdedor fué " & UserList(Perdedor).Name & "." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
    End Sub
