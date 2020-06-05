Attribute VB_Name = "Mod_DeathAutomatico"
Option Explicit
Private cantdeath As Integer
Private Const mapadeath As Integer = 205
Private Const posideath As Integer = 50
Private Const posideathy As Integer = 50
Public deathac As Boolean
Public deathesp As Boolean
Public Cantidad As Integer
Private Const esperadeath = 52
Private Const esperadeathy = 27
Private Death_Luchadores() As Integer


Sub death_entra(ByVal UserIndex As Integer)
    On Error GoTo errordm:
    Dim i As Long
    If deathac = False Then
        Call SendData(ToIndex, 0, 0, "||No hay ningun Deathmatch!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If deathesp = False Then
        Call SendData(ToIndex, 0, 0, "||El Deathmatch ya ha comenzado, te quedaste fuera!" & FONTTYPE_INFO)
        Exit Sub
    End If

    For i = LBound(Death_Luchadores) To UBound(Death_Luchadores)
        If (Death_Luchadores(i) = UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya estas dentro!" & FONTTYPE_WARNING)
            Exit Sub
        End If
    Next i

    For i = LBound(Death_Luchadores) To UBound(Death_Luchadores)
        If (Death_Luchadores(i) = -1) Then
            Death_Luchadores(i) = UserIndex
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            FuturePos.Map = mapadeath
            FuturePos.X = esperadeath: FuturePos.Y = esperadeathy
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Death_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
            UserList(Death_Luchadores(i)).flags.Death = True
            Call SendData(ToIndex, UserIndex, 0, "||Estas dentro del deathmatch!" & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "||DeathMatch: Entra el participante " & UserList(UserIndex).name & FONTTYPE_INFO)
            If (i = UBound(Death_Luchadores)) Then
                Call SendData(ToAll, 0, 0, "||DeathMatch: Empieza el DeathMach!!" & FONTTYPE_GUILD)
                deathesp = False
                Call Deathauto_empieza
            End If
            Exit Sub
        End If
    Next i
errordm:
End Sub

Sub death_comienza(ByVal wetas As Integer)
    On Error GoTo errordm
    If deathac = True Then
        Call SendData(ToIndex, 0, 0, "||Ya hay un deathmatch!!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If deathesp = True Then
        Call SendData(ToIndex, 0, 0, "||La deathmatch ya ha comenzado!" & FONTTYPE_INFO)
        Exit Sub
    End If

    cantdeath = wetas
    Cantidad = cantdeath
    Call SendData(ToAll, 0, 0, "||DeathMatch: Esta empezando un nuevo deathmatch para " & cantdeath & " participantes. Para participar envia /DEATH - (NO cae Inventario) " & FONTTYPE_GUILD)
    Call SendData(ToAll, 0, 0, "TW48")
    deathac = True
    deathesp = True
    frmMain.Cancelaaut.Enabled = True
    ReDim Death_Luchadores(1 To cantdeath) As Integer

    Dim i As Long
    For i = LBound(Death_Luchadores) To UBound(Death_Luchadores)
        Death_Luchadores(i) = -1
    Next i
errordm:
End Sub

Sub death_muere(ByVal UserIndex As Integer)
    On Error GoTo errord
    If UserList(UserIndex).flags.Death = True Then
        Call WarpUserChar(UserIndex, 1, 50, 50, True)
        UserList(UserIndex).flags.Death = False
        Cantidad = Cantidad - 1
        If Cantidad = 1 Or MapInfo(mapadeath).NumUsers = 1 Then
            terminodeat = True
            Call SendData(ToAll, 0, 0, "||DeathMatch: Termina la DeathMatch! El Ganador Debe escribir /GANADOR para recibir su recompensa!!!" & FONTTYPE_GUILD)
            Exit Sub
        End If
        If Cantidad = 0 Then
            terminodeat = False
            deathesp = False
            deathac = False
            Call SendData(ToAll, 0, 0, "||DeathMatch: El ganador de la deatmatch desconecto. Nadie ha ganado premios." & FONTTYPE_GUILD)
        End If
    End If
errord:
End Sub

Sub Death_Cancela()
    On Error GoTo errordm
    If deathac = False And deathesp = False Then Exit Sub
    deathesp = False
    deathac = False
    Call SendData(ToAll, 0, 0, "||DeathMatch: DeathMatch Automatico Cancelada Por un Game Master" & FONTTYPE_GUILD)
    frmMain.Torneos.Enabled = True
    frmMain.Cancelaaut.Enabled = False

    Dim i As Long
    For i = LBound(Death_Luchadores) To UBound(Death_Luchadores)
        If (Death_Luchadores(i) <> -1) Then
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            FuturePos.Map = 1
            FuturePos.X = 50: FuturePos.Y = 50
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Death_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
            UserList(Death_Luchadores(i)).flags.Death = False
        End If
    Next i
errordm:
End Sub

Sub Deathauto_Cancela()
    On Error GoTo errordmm

    If deathac = False And deathesp = False Then Exit Sub

    deathesp = False
    deathac = False
    Call SendData(ToAll, 0, 0, "||DeathMatch: DeathMatch Automatico cancelada por falta de participantes." & FONTTYPE_GUILD)
    frmMain.Torneos.Enabled = True
    Dim i As Long
    For i = LBound(Death_Luchadores) To UBound(Death_Luchadores)
        If (Death_Luchadores(i) <> -1) Then
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            FuturePos.Map = 1
            FuturePos.X = 50: FuturePos.Y = 50
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Death_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
            UserList(Death_Luchadores(i)).flags.Death = False
        End If
    Next i
errordmm:
End Sub

Sub Deathauto_empieza()
    On Error GoTo errordm

    Dim i As Long
    For i = LBound(Death_Luchadores) To UBound(Death_Luchadores)
        If (Death_Luchadores(i) <> -1) Then
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            FuturePos.Map = mapadeath
            FuturePos.X = posideath: FuturePos.Y = posideathy
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Death_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
        End If
    Next i
errordm:
End Sub

