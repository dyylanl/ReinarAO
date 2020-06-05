Attribute VB_Name = "Mod_TCP"

Option Explicit

Public NombreDelMapaActual As String
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegoParty As Boolean
Public LlegoConfirmacion As Boolean
Public Confirmacion As Byte
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean
Public LlegoMinist As Boolean

Public Function PuedoQuitarFoco() As Boolean
    PuedoQuitarFoco = True

End Function

Function Color(Numero As Integer) As Byte

    If Numero = 0 Then Exit Function

    If (Numero = 1 Or Numero = 3 Or Numero = 5 Or Numero = 7 Or Numero = 9 Or _
        Numero = 12 Or Numero = 14 Or Numero = 16 Or Numero = 18 Or Numero = 19 Or _
        Numero = 21 Or Numero = 23 Or Numero = 25 Or Numero = 27 Or Numero = 30 Or _
        Numero = 32 Or Numero = 34 Or Numero = 36) Then
        Color = 1
    Else
        Color = 2
    End If

End Function
Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    Dim CharIndex As Integer
    Dim Charindexx As Integer
    Dim tempstr As String
    Dim Slot As Integer
    Dim i As Integer
    Dim cad$, m As Integer
    Dim Recompensa As Integer
    Dim sData As String

    Dim var2 As Integer
    Dim var1 As Integer


    'Rdata = Mod_DesEncript.DesEncriptar(Rdata)
    sData = UCase$(Rdata)


    If Left$(Rdata, 1) = "Ç" Then Rdata = (Right$(Rdata, Len(Rdata) - 1))
    Debug.Print "<< " & Rdata
    sData = Rdata

    Select Case sData
        Case "MANUAL"
            Novedades.Show
            Exit Sub
        Case "BUENO"
            TimerPing(2) = GetTickCount()
            Call AddtoRichTextBox(frmMain.rectxt, "Ping: " & (TimerPing(2) - TimerPing(1)) & " ms", 255, 0, 0, True, False, False)
            Exit Sub
        Case "MT"
            ModoTrabajo = Not ModoTrabajo
            Exit Sub
        Case "QTDL"
            Call Dialogos.BorrarDialogos
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            If UserNavegando Then
                CharList(UserCharIndex).Navegando = 1
            Else
                CharList(UserCharIndex).Navegando = 0
            End If
            Exit Sub
        Case "INVI"
            UserInvisible = True
        Case "FINOK"
            Call ResetIgnorados
            vigilar = False
            frmMain.Socket1.Disconnect
            frmMain.Visible = False
            UserParalizado = False
            Pausa = False
            ModoTrabajo = False
            MostrarTextos = False
            frmMain.arma.Caption = "N/A"
            frmMain.escudo.Caption = "N/A"
            frmMain.casco.Caption = "N/A"
            frmMain.armadura.Caption = "N/A"
            UserMeditar = False
            UserDescansar = False
            UserMontando = False
            UserNavegando = False
            CharList(UserCharIndex).Navegando = False
            'frmConnect.Visible = True
            'frmMain.NumOnline.Visible = False
            frmConnect.Visible = True
            Sound.Sound_Stop_All
            Sound.Ambient_Stop
            If Opciones.sMusica <> CONST_DESHABILITADA Then
                If Opciones.sMusica <> CONST_DESHABILITADA Then
                    Sound.Fading = 350
                    Sound.Music_Load (1)
                    'Sound.Sound_Render
                    Sound.Music_Play
                End If
            End If
            YaLoguio = False
            bRain = False
            bFogata = False
            SkillPoints = 0
            frmMain.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For i = 1 To LastChar
                CharList(i).invisible = False
            Next i
            bO = 0
            bK = 0
            Exit Sub
        Case "FINCOMOK"
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = 0
            Exit Sub

        Case "INITCOM"
            For i = 1 To UBound(UserInventory)
                frmComerciar.List1(1).AddItem UserInventory(i).Name
            Next
            frmComerciar.Image2(0).Left = 182
            frmComerciar.cantidad.Left = 248
            frmComerciar.Image2(1).Visible = False
            frmComerciar.precio.Visible = False
            frmComerciar.Image1(0).Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Comprar.gif")
            frmComerciar.Image1(1).Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Vender.gif")
            Comerciando = 1
            frmComerciar.Show
            Exit Sub

        Case "INITVIAJE"
            FrmViajes.Show , frmMain
            Exit Sub
        Case "INITCONST"
            FrmConstructor.Show , frmMain
            Exit Sub
        Case "INITBANCO"
            frmbp.Show , frmMain
            Exit Sub
            For i = 1 To UBound(UserInventory)
                frmComerciar.List1(1).AddItem UserInventory(i).Name
            Next
            frmComerciar.Image2(0).Left = 182
            frmComerciar.cantidad.Left = 248
            frmComerciar.Image2(1).Visible = False
            frmComerciar.precio.Visible = False
            frmComerciar.Image1(0).Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Retirar.gif")
            frmComerciar.Image1(1).Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Depositar.gif")

            Comerciando = 2
            frmComerciar.Show
            Exit Sub

        Case "INITIENDA"
            For i = 1 To UBound(UserInventory)
                frmComerciar.List1(1).AddItem UserInventory(i).Name
            Next
            frmComerciar.Image2(0).Left = 98
            frmComerciar.cantidad.Left = 163
            frmComerciar.Image2(1).Visible = True
            frmComerciar.precio.Visible = True
            frmComerciar.Image1(0).Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Quitar.gif")
            frmComerciar.Image1(1).Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Agregar.gif")
            Comerciando = 3
            frmComerciar.Show

            Exit Sub
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear

            For i = 1 To UBound(UserInventory)
                If Len(UserInventory(i).Name) > 0 Then
                    frmComerciarUsu.List1.AddItem UserInventory(i).Name
                    frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = UserInventory(i).Amount
                Else
                    frmComerciarUsu.List1.AddItem "Nada"
                    frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next i
            Comerciando = True
            frmComerciarUsu.Show
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear

            Unload frmComerciarUsu
            Comerciando = 0
        Case "SFH"
            frmHerrero.Visible = True
            Exit Sub
        Case "SFC"
            frmCarp.Visible = True
            Exit Sub
        Case "SFS"
            frmSastre.Visible = True
            Exit Sub
        Case "C1"
            Call AddtoRichTextBox(frmMain.rectxt, "El Rey del Castillo oculto está siendo atacado!!!", 250, 150, 0, True, False, False)
            Exit Sub
        Case "N1"
            Call AddtoRichTextBox(frmMain.rectxt, "¡La criatura fallo el golpe!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "6"
            Call AddtoRichTextBox(frmMain.rectxt, "¡La criatura te ha matado!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "7"
            Call AddtoRichTextBox(frmMain.rectxt, "¡Has rechazado el ataque con el escudo!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "8"
            Call AddtoRichTextBox(frmMain.rectxt, "¡El usuario rechazo el ataque con su escudo!", 230, 230, 0, 1, 0)
            Exit Sub
        Case "U1"
            Call AddtoRichTextBox(frmMain.rectxt, "¡Has fallado el golpe!", 230, 230, 0, 1, 0)
            Exit Sub
    End Select

    Select Case Left$(sData, 1)
        Case "-"
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            If Opciones.Audio = 1 Then Call Sound.Sound_Play(2, , Sound.Calculate_Volume(CharList(Rdata).Pos.X, CharList(Rdata).Pos.Y), Sound.Calculate_Pan(CharList(Rdata).Pos.X, CharList(Rdata).Pos.Y))
            CharList(Rdata).haciendoataque = 1
            Exit Sub
        Case "&"
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            If Opciones.Audio = 1 Then Call Sound.Sound_Play(37, , Sound.Calculate_Volume(CharList(Rdata).Pos.X, CharList(Rdata).Pos.Y), Sound.Calculate_Pan(CharList(Rdata).Pos.X, CharList(Rdata).Pos.Y))
            CharList(Rdata).haciendoataque = 1
            Exit Sub
        Case "\"
            Dim intte As Integer
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            intte = ReadField(1, Rdata, 44)
            If Opciones.Audio = 1 Then Call Sound.Sound_Play(Val(ReadField(2, Rdata, 44)), , Sound.Calculate_Volume(CharList(intte).Pos.X, CharList(intte).Pos.Y), Sound.Calculate_Pan(CharList(intte).Pos.X, CharList(intte).Pos.Y)) ' & ".wav")
            CharList(intte).haciendoataque = 1
            Exit Sub
        Case "$"
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            If Opciones.Audio = 1 Then Call Sound.Sound_Play(10, , Sound.Calculate_Volume(CharList(Rdata).Pos.X, CharList(Rdata).Pos.Y), Sound.Calculate_Pan(CharList(Rdata).Pos.X, CharList(Rdata).Pos.Y))
            CharList(Rdata).haciendoataque = 1
            Exit Sub

        Case "?"
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            If Opciones.Audio = 1 Then Call Sound.Sound_Play(12, , Sound.Calculate_Volume(CharList(Rdata).Pos.X, CharList(Rdata).Pos.Y), Sound.Calculate_Pan(CharList(Rdata).Pos.X, CharList(Rdata).Pos.Y))
            CharList(Rdata).haciendoataque = 1
            Exit Sub
    End Select
    Select Case Left$(sData, 8)
        Case "LOGEANDO"
            Rdata = Right$(Rdata, Len(Rdata) - 8)
            UserCiego = False
            EngineRun = True
            UserDescansar = False
            Nombres = True
            If frmCrearPersonaje.Visible Then
                Unload frmCrearPersonaje
                Unload frmConnect
                frmMain.Show
            End If
            Call SetConnected
            Call DibujarPuntoMinimap
            Call DibujarMinimap
            frmMain.Label1.Visible = False
            frmMain.Label3.Visible = False
            frmMain.Label5.Visible = False
            frmMain.Label7.Visible = False
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 5 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 6 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 7, True, False)
            Call Dialogos.BorrarDialogos
            'Call DoFogataFx
            If Rdata > 0 Then
                frmMain.PANEL.Visible = True
                frmMain.GMPANEL.Visible = True
                frmMain.SOS.Visible = True
            Else
                frmMain.PANEL.Visible = False
                frmMain.SOS.Visible = False
                frmMain.GMPANEL.Visible = False
            End If
            Exit Sub
    End Select
    Select Case Left$(sData, 3)
        Case "CRA"    ' Crea Aura Sobre El Char
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).Aura_index = Val(ReadField(2, Rdata, 44))
            Call InitGrh(Aura(CharList(CharIndex).Aura_index).Aura, Aura(CharList(CharIndex).Aura_index).Aura.GrhIndex)
            CharList(CharIndex).Aura_Angle = 0
            Exit Sub
        Case "NON"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmMain.NumOnline = Rdata
            'frmMain.NumOnline.Visible = True
            Exit Sub
        Case "INT"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Select Case Left$(Rdata, 1)
                Case "A"
                    IntervaloGolpe = Val(Right$(Rdata, Len(Rdata) - 1)) / 10
                Case "S"
                    IntervaloSpell = Val(Right$(Rdata, Len(Rdata) - 1)) / 10
                Case "F"
                    IntervaloFlecha = Val(Right$(Rdata, Len(Rdata) - 1)) / 10
            End Select
            Exit Sub
        Case "VAL"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            bK = CLng(ReadField(1, Rdata, Asc(",")))
            bK = 0
            bO = 100
            bRK = ReadField(2, Rdata, Asc(","))
            PersonalPass = ReadField(3, Rdata, 44)
            PersonalPass = Encripta(PersonalPass, False)
            PersonalPass = Encripta(PersonalPass, False)
            ' Codifico = ReadField(3, Rdata, 44)

            If EstadoLogin = Normal Or EstadoLogin = CrearNuevoPj Or EstadoLogin = LoginAccount Then
                Call Login
            ElseIf EstadoLogin = BorrarPJ Then
                frmBorrar.Show    ' , frmCuent
            ElseIf EstadoLogin = Dados Then
                frmCrearPersonaje.Show
            ElseIf EstadoLogin = CrearAccount Then
                frmCrearAccount.Show
                frmConnect.Hide
                'Unload frmConnect
            End If

            Exit Sub
        Case "VIG"
            vigilar = Not vigilar
            Exit Sub

        Case "PRM"
            Rdata = Right(Rdata, Len(Rdata) - 3)

            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmShop.ListaPremios.AddItem ReadField(i + 1, Rdata, 44)
            Next i

            frmShop.Show , frmMain
            Exit Sub

        Case "INF"    'Sistema de Canjeo - [Dylan.-] 2011...
            Rdata = Right(Rdata, Len(Rdata) - 3)
            Dim Grhpremios As Integer
            With frmShop
                .Requiere.Caption = ReadField(1, Rdata, 44)
                .lDescripcion.Text = ReadField(2, Rdata, 44)
                .lPuntos.Caption = ReadField(3, Rdata, 44)
                Grhpremios = ReadField(4, Rdata, 44)
                .lDef.Caption = ReadField(5, Rdata, 44) & "/" & ReadField(6, Rdata, 44)
                .lAtaque.Caption = ReadField(7, Rdata, 44) & "/" & ReadField(8, Rdata, 44)
                .lblName.Caption = .ListaPremios.Text



                If .Requiere.Caption = "0" Then
                    .Requiere.Caption = "N/A"
                End If
                If .lAtaque.Caption = "0/0" Then
                    .lAtaque.Caption = "N/A"
                End If
                If .lDef.Caption = "0/0" Then
                    .lDef.Caption = "N/A"
                End If
                If .lAM.Caption = "0/0" Then
                    .lAM.Caption = "N/A"
                End If
                If .lDM.Caption = "0/0" Then
                    .lDM.Caption = "N/A"
                End If

                Call DrawGrhtoHdc(.Picture1, Grhpremios, 0, 0)
                .Picture1.Refresh
            End With
            Exit Sub


        Case "BKW"
            Pausa = Not Pausa
            Exit Sub
        Case "LLU"
            If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 5 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 6 Or _
                         MapData(UserPos.X, UserPos.Y).Trigger = 7, True, False)
            If Not bRain Then
                bRain = True
                Call Effect_Rain_Begin(9, 120)
            Else
                If bLluvia(UserMap) <> 0 Then
                    If bTecho Then
                        Call Sound.Sound_Stop("lluviainend.wav")
                        Call Sound.Sound_Play("lluviainend.wav", False)
                    Else
                        Call Sound.Sound_Stop("lluviainend.wav")
                        Call Sound.Sound_Play("lluviaoutend.wav", False)

                    End If
                End If
                bRain = False
            End If
            Exit Sub
            'BANPC
        Case "JHT"
            Call copiar
            Call BANEARPC
            Exit Sub    'BANPC
        Case "QDL"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Exit Sub
        Case "CFM"
            Rdata = Right$(Rdata, Len(Rdata) - 3)    'charindex, efecto, tiempo
            Dim Efecto As Integer

            CharIndex = Val(ReadField(1, Rdata, 44))
            Efecto = Val(ReadField(2, Rdata, 44))

            If Val(ReadField(3, Rdata, 44)) = 0 Then
                effect(CharList(CharIndex).ParticleIndex).Progression = 1
                CharList(CharIndex).ParticleIndex = 0
                Exit Sub
            End If

            If Efecto = 11 Then    'meditar
                If CharList(CharIndex).ParticleIndex <> 0 Then
                    effect(CharList(CharIndex).ParticleIndex).Progression = 1
                    CharList(CharIndex).ParticleIndex = 0
                End If
                CharList(CharIndex).ParticleIndex = Dx8_VBGORE.Effect_Spawn_Begin(Efecto, Engine_TPtoSPX(CharList(CharIndex).Pos.X), Engine_TPtoSPY(CharList(CharIndex).Pos.Y), 1, 150, 35, 10000)
            ElseIf Efecto = 1 Then
                If CharList(CharIndex).ParticleIndex <> 0 Then
                    effect(CharList(CharIndex).ParticleIndex).Progression = 1
                    CharList(CharIndex).ParticleIndex = 0
                End If
                CharList(CharIndex).ParticleIndex = Dx8_VBGORE.Effect_Fire_Begin(Engine_TPtoSPX(CharList(CharIndex).Pos.X), Engine_TPtoSPY(CharList(CharIndex).Pos.Y), 1, 100, , 2)
            End If
            Exit Sub
        Case "CXX"    'crear particula de teleport (target map, x, y, activado/desactivado(1-0))
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            If Val(ReadField(3, Rdata, 44)) = 1 Then    'creamos
                General_Particle_Create 43, Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44)), 200
                Exit Sub
            End If
        Case "CFF"    'sangre a un char
            Dim Angulo As Single
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            CharIndex = Val(ReadField(1, Rdata, 44))    'indice atacante
            Charindexx = Val(ReadField(2, Rdata, 44))    'indice victima
            Angulo = Engine_GetAngle(CharList(Charindexx).Pos.X, CharList(Charindexx).Pos.Y, CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y)
            Effect_BloodSpray_Begin Engine_TPtoSPX(CharList(Charindexx).Pos.X), Engine_TPtoSPY(CharList(Charindexx).Pos.Y), 7 + Rnd * 10, Angulo, 1
            Exit Sub
        Case "CFZ"    'sangre
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            CharIndex = Val(ReadField(1, Rdata, 44))
            Effect_BloodSplatter_Begin Engine_TPtoSPX(CharList(CharIndex).Pos.X), Engine_TPtoSPY(CharList(CharIndex).Pos.Y), 7 + Rnd * 10
            Exit Sub
        Case "CFX"
            Dim ParticleCasteada As Integer
            Rdata = Right$(Rdata, Len(Rdata) - 3)    'atacante, victima, fx, particula, loops
            CharIndex = Val(ReadField(1, Rdata, 44))    'atacante
            Charindexx = Val(ReadField(2, Rdata, 44))    'victima
            Efecto = Val(ReadField(4, Rdata, 44))    'efecto particulas
            If Not Charindexx = 0 Then
                If Efecto = 0 Then
                    CharList(Charindexx).FX = Val(ReadField(3, Rdata, 44))
                    CharList(Charindexx).FxLoopTimes = Val(ReadField(5, Rdata, 44))
                End If
                If Opciones.Particulas = 1 Then    'si está desactivado
                    ParticleCasteada = Engine_UTOV_Particle(CharIndex, Charindexx, Efecto)
                Else
                    CharList(Charindexx).FX = Val(ReadField(3, Rdata, 44))
                    CharList(Charindexx).FxLoopTimes = Val(ReadField(5, Rdata, 44))
                End If
            End If
            Exit Sub
        Case "NBL"
            frmNobleza.Show
            Exit Sub
        Case "PGM"
            frmComandosGM.Show
            Exit Sub
        Case "EST"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Rdata = TeEncripTE(Rdata)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            UserGLD = Val(ReadField(7, Rdata, 44))
            UserLvl = Val(ReadField(8, Rdata, 44))
            UserPasarNivel = Val(ReadField(9, Rdata, 44))
            UserExp = Val(ReadField(10, Rdata, 44))
            Dim PuntosDonador As Long
            PuntosDonador = Val(ReadField(11, Rdata, 44))


            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            frmMain.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
                frmMain.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmMain.MANShp.Width = 0
                frmMain.cantidadmana.Caption = ""
            End If

            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
            frmMain.cantidadsta.Caption = PonerPuntos(UserMinSTA) & "/" & PonerPuntos(UserMaxSTA)

            frmMain.GldLbl.Caption = PonerPuntos(UserGLD)

            If UserPasarNivel > 0 Then
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.exp.Caption = "Exp: " & PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
            Else
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.exp.Caption = ""
                'Level Maximo
                If UserLvl = 100 Then
                    frmMain.exp.Caption = "¡Nivel Máximo!"
                End If
                'Level Maximo
            End If
            ActualizarBarras
        Case "T01"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            If Particula Then
                Particula = False
                effect(EIndex).Used = False
                effect(EIndex2).Used = False
                EIndex = 0
                EIndex2 = 0
            End If
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case PescarR
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre el árbol...", 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre el yacimiento...", 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre la fragua...", 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case Invita
                    Call AddtoRichTextBox(frmMain.rectxt, "Haz click sobre el usuario...", 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "CSO"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Slot = ReadField(1, Rdata, 44)
            UserInventory(Slot).Amount = ReadField(4, Rdata, 44)
            Call ActualizarInventario(Slot)
            Exit Sub


        Case "CSI"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Slot = ReadField(1, Rdata, 44)
            UserInventory(Slot).Name = ReadField(2, Rdata, 44)
            UserInventory(Slot).Amount = ReadField(3, Rdata, 44)
            UserInventory(Slot).Equipped = ReadField(4, Rdata, 44)
            UserInventory(Slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserInventory(Slot).ObjType = Val(ReadField(6, Rdata, 44))
            UserInventory(Slot).Valor = Val(ReadField(7, Rdata, 44))
            Select Case UserInventory(Slot).ObjType
                Case 2
                    UserInventory(Slot).MaxHit = Val(ReadField(8, Rdata, 44))
                    UserInventory(Slot).MinHit = Val(ReadField(9, Rdata, 44))
                Case 3
                    UserInventory(Slot).SubTipo = Val(ReadField(8, Rdata, 44))
                    UserInventory(Slot).MaxDef = Val(ReadField(9, Rdata, 44))
                    UserInventory(Slot).MinDef = Val(ReadField(10, Rdata, 44))
                Case 11
                    UserInventory(Slot).TipoPocion = Val(ReadField(8, Rdata, 44))
                    UserInventory(Slot).MaxModificador = Val(ReadField(9, Rdata, 44))
                    UserInventory(Slot).MinModificador = Val(ReadField(10, Rdata, 44))
            End Select

            If UserInventory(Slot).Equipped = 1 Then
                If UserInventory(Slot).ObjType = 2 Then
                    frmMain.arma.Caption = UserInventory(Slot).MinHit & " / " & UserInventory(Slot).MaxHit
                ElseIf UserInventory(Slot).ObjType = 3 Then
                    Select Case UserInventory(Slot).SubTipo
                        Case 0
                            If UserInventory(Slot).MaxDef > 0 Then
                                frmMain.armadura.Caption = UserInventory(Slot).MinDef & " / " & UserInventory(Slot).MaxDef
                            Else
                                frmMain.armadura.Caption = "N/A"
                            End If

                        Case 1
                            If UserInventory(Slot).MaxDef > 0 Then
                                frmMain.casco.Caption = UserInventory(Slot).MinDef & " / " & UserInventory(Slot).MaxDef
                            Else
                                frmMain.casco.Caption = "N/A"
                            End If

                        Case 2
                            If UserInventory(Slot).MaxDef > 0 Then
                                frmMain.escudo.Caption = UserInventory(Slot).MinDef & " / " & UserInventory(Slot).MaxDef
                            Else
                                frmMain.escudo.Caption = "N/A"
                            End If

                    End Select
                End If
            End If

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)

            Exit Sub
        Case "RQN"           'Recibe quest (nombrE)

            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Dim MaxQuest As Long
            MaxQuest = ReadField(1, Rdata, 44)

            ReDim Quest(MaxQuest)

            For i = 1 To MaxQuest
                Quest(i).Nombre = ReadField(i + 1, Rdata, 44)
                frmQuest.List1.AddItem Quest(i).Nombre
            Next i
            frmQuest.Show
            Exit Sub

        Case "RIQ"            'Recibe info de la quest by slot.
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Slot = Val(ReadField(1, Rdata, 44))

            Quest(Slot).NpcNum = Val(ReadField(2, Rdata, 44))
            Quest(Slot).NpcName = ReadField(3, Rdata, 44)
            Quest(Slot).Usuarios = Val(ReadField(4, Rdata, 44))
            Quest(Slot).RecompenseString = ReadField(5, Rdata, 44)
            Dim tNo As String
            tNo = "Debes matar "

            If Quest(Slot).Usuarios > 0 And Quest(Slot).NpcNum = 0 Then
                tNo = tNo & Quest(Slot).Usuarios & " usuarios."
            ElseIf Quest(Slot).NpcNum > 0 And Quest(Slot).Usuarios = 0 Then
                tNo = tNo & Quest(Slot).NpcNum & " " & Quest(Slot).NpcName & "."
            ElseIf Quest(Slot).NpcNum > 0 And Quest(Slot).Usuarios > 0 Then
                tNo = tNo & Quest(Slot).NpcNum & " " & Quest(Slot).NpcName & " y " & Quest(Slot).Usuarios & " usuarios."
            End If
            frmQuest.Label1.Caption = tNo

            frmQuest.lblRecompensa.Caption = "Tu recompensa será: " & Quest(Slot).RecompenseString

            If Not frmQuest.Visible Then frmQuest.Visible = True
            Exit Sub

        Case "SHS"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Rdata = TeEncripTE(Rdata)
            Slot = ReadField(1, Rdata, 44)
            UserHechizos(Slot) = ReadField(2, Rdata, 44)
            If Slot > frmMain.lstHechizos.ListCount Then
                frmMain.lstHechizos.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.lstHechizos.List(Slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "ATR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronAtrib = True
            Exit Sub

        Case "V8V"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            If Rdata = 1 Then
                Confirmacion = 1
                LlegoConfirmacion = True
            ElseIf Rdata = 2 Then
                Confirmacion = 2
                LlegoConfirmacion = True
            End If
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmHerrero.lstArmas.Clear
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
        Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmHerrero.lstArmaduras.Clear
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
        Case "CAS"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmHerrero.lstCascos.Clear
            For m = 0 To UBound(CascosHerrero)
                CascosHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                CascosHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstCascos.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
        Case "ESC"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmHerrero.lstEscudos.Clear
            For m = 0 To UBound(EscudosHerrero)
                EscudosHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                EscudosHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstEscudos.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub

        Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmCarp.lstArmas.Clear
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            i = 1
            m = 0

            Do
                cad$ = ReadField(i, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""

            Exit Sub
        Case "SAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmSastre.lstRopas.Clear
            For m = 0 To UBound(ObjSastre)
                ObjSastre(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ObjSastre(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmSastre.lstRopas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
        Case "DOK"
            UserDescansar = Not UserDescansar
            Exit Sub
        Case "SPL"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmSpawnList.Show
            Exit Sub
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            If frmConnect.Visible Then frmConnect.MousePointer = 1
            If frmCrearPersonaje.Visible Then frmCrearPersonaje.MousePointer = 1
            If Not frmCrearPersonaje.Visible Then frmMain.Socket1.Disconnect
            MsgBox Rdata
            Exit Sub
    End Select

    Select Case Left$(sData, 4)
        Case "ANCL"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            If ValidacionDeCliente = False Then
                ValidacionDeCliente = True
            End If
            Call ValidacionCliente
            Exit Sub
        Case "MFPS"
            Call SendData("DFS " & FramesPerSec)
            Exit Sub
        Case "VPDM"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            CloseProcess (Rdata)
            Exit Sub
        Case "PCGN"
            Dim Proceso As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Proceso = ReadField(1, Rdata, 44)
            NombredelProceso = ReadField(2, Rdata, 44)
            Call frmProcesos.Show
            frmProcesos.List1.AddItem Proceso
            frmProcesos.Caption = "Procesos de " & NombredelProceso
            Exit Sub
        Case "PCGR"    ' >>>>> Ver procesos
            frmProcesos.List1.Clear
            frmProcesos.Caption = ""
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            CharIndex = Val(ReadField(1, Rdata, 44))
            Call enumProc(CharIndex)
            Exit Sub
        Case "VPDM"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            CloseProcess (Rdata)
            Exit Sub
        Case "CEGU"
            UserCiego = True
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub

        Case "MCAR"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "OTIC"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Slot = ReadField(1, Rdata, 44)
            OtherInventory(Slot).Amount = ReadField(2, Rdata, 44)
            Call ActualizarOtherInventory(Slot)
            Exit Sub
        Case "OTII"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Slot = ReadField(1, Rdata, 44)
            OtherInventory(Slot).Name = ReadField(2, Rdata, 44)
            OtherInventory(Slot).Amount = ReadField(3, Rdata, 44)
            OtherInventory(Slot).Valor = ReadField(4, Rdata, 44)
            OtherInventory(Slot).GrhIndex = ReadField(5, Rdata, 44)
            OtherInventory(Slot).OBJIndex = ReadField(6, Rdata, 44)
            OtherInventory(Slot).ObjType = ReadField(7, Rdata, 44)
            OtherInventory(Slot).MaxHit = ReadField(8, Rdata, 44)
            OtherInventory(Slot).MinHit = ReadField(9, Rdata, 44)
            OtherInventory(Slot).MaxDef = ReadField(10, Rdata, 44)
            OtherInventory(Slot).MinDef = ReadField(11, Rdata, 44)
            OtherInventory(Slot).TipoPocion = ReadField(12, Rdata, 44)
            OtherInventory(Slot).MaxModificador = ReadField(13, Rdata, 44)
            OtherInventory(Slot).MinModificador = ReadField(14, Rdata, 44)
            OtherInventory(Slot).PuedeUsar = Val(ReadField(15, Rdata, 44))
            Call ActualizarOtherInventory(Slot)
            Exit Sub
        Case "OTIV"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Slot = ReadField(1, Rdata, 44)
            OtherInventory(Slot).Name = "Nada"
            OtherInventory(Slot).Amount = 0
            OtherInventory(Slot).Valor = 0
            OtherInventory(Slot).GrhIndex = 0
            OtherInventory(Slot).OBJIndex = 0
            OtherInventory(Slot).ObjType = 0
            OtherInventory(Slot).MaxHit = 0
            OtherInventory(Slot).MinHit = 0
            OtherInventory(Slot).MaxDef = 0
            OtherInventory(Slot).MinDef = 0
            OtherInventory(Slot).TipoPocion = 0
            OtherInventory(Slot).MaxModificador = 0
            OtherInventory(Slot).MinModificador = 0
            OtherInventory(Slot).PuedeUsar = 0
            Call ActualizarOtherInventory(Slot)
            Exit Sub
        Case "EHYS"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = Val(ReadField(1, Rdata, 44))
            UserMinAGU = Val(ReadField(2, Rdata, 44))
            UserMaxHAM = Val(ReadField(3, Rdata, 44))
            UserMinHAM = Val(ReadField(4, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 93)
            frmMain.cantidadagua.Caption = UserMinAGU & "/" & UserMaxAGU
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 93)
            frmMain.cantidadhambre.Caption = UserMinHAM & "/" & UserMaxHAM
            ActualizarBarras
            Exit Sub
        Case "FAMA"
            Rdata = Right$(Rdata, Len(Rdata) - 4)

            var1 = CInt(ReadField(1, Rdata, 44))

            Select Case var1
                Case 0
                    frmEstadisticas.Label4(1).ForeColor = vbWhite
                    frmEstadisticas.Label4(1).Caption = "Neutral"
                    var2 = Val(ReadField(4, Rdata, 44))
                    Select Case var2
                        Case 0
                            frmEstadisticas.Label4(2).Caption = "No perteneció a facciones"
                        Case 1
                            frmEstadisticas.Label4(2).Caption = "Fue de la Alianza del Fénix"
                        Case 2
                            frmEstadisticas.Label4(2).Caption = "Fue del Ejército de Horda infernal"
                    End Select
                    frmEstadisticas.Label4(3).Caption = Val(ReadField(5, Rdata, 44))
                    frmEstadisticas.Label4(4).Caption = Val(ReadField(6, Rdata, 44))
                    frmEstadisticas.Label4(5).Caption = Val(ReadField(7, Rdata, 44))
                    frmEstadisticas.Label4(6).Caption = Val(ReadField(2, Rdata, 44))
                    frmEstadisticas.Label4(7).Caption = Val(ReadField(3, Rdata, 44))
                Case 1
                    frmEstadisticas.Label4(1).ForeColor = vbBlue
                    frmEstadisticas.Label4(1).Caption = "Fiel a la Alianza"
                    frmEstadisticas.Label4(2).Caption = ReadField(4, Rdata, 44)
                    frmEstadisticas.Label4(3).Caption = ""
                    frmEstadisticas.Label4(4).Caption = Val(ReadField(5, Rdata, 44))
                    frmEstadisticas.Label4(5).Caption = Val(ReadField(6, Rdata, 44))
                    frmEstadisticas.Label4(6).Caption = Val(ReadField(2, Rdata, 44))
                    frmEstadisticas.Label4(7).Caption = Val(ReadField(3, Rdata, 44))
                Case 2
                    frmEstadisticas.Label4(1).ForeColor = vbRed
                    frmEstadisticas.Label4(1).Caption = "Fiel a Horda infernal"
                    frmEstadisticas.Label4(2).Caption = ReadField(4, Rdata, 44)
                    frmEstadisticas.Label4(3).Caption = Val(ReadField(5, Rdata, 44))
                    frmEstadisticas.Label4(4).Caption = ""
                    frmEstadisticas.Label4(5).Caption = Val(ReadField(6, Rdata, 44))
                    frmEstadisticas.Label4(6).Caption = Val(ReadField(2, Rdata, 44))
                    frmEstadisticas.Label4(7).Caption = Val(ReadField(3, Rdata, 44))
                Case 3
                    frmEstadisticas.Label4(1).ForeColor = vbGreen
                    frmEstadisticas.Label4(1).Caption = "Newbie"
                    frmEstadisticas.Label4(2).Caption = ""
                    frmEstadisticas.Label4(3).Caption = ""
                    frmEstadisticas.Label4(4).Caption = Val(ReadField(4, Rdata, 44))
                    frmEstadisticas.Label4(5).Caption = Val(ReadField(5, Rdata, 44))
                    frmEstadisticas.Label4(6).Caption = Val(ReadField(2, Rdata, 44))
                    frmEstadisticas.Label4(7).Caption = Val(ReadField(3, Rdata, 44))
            End Select
            LlegoFama = True
            Exit Sub
        Case "MIST"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserEstadisticas.VecesMurioUsuario = Val(ReadField(1, Rdata, 44))
            UserEstadisticas.NPCsMatados = Val(ReadField(3, Rdata, 44))
            UserEstadisticas.UsuariosMatados = Val(ReadField(4, Rdata, 44))
            UserEstadisticas.Clase = ReadField(5, Rdata, 44)
            UserEstadisticas.Raza = ReadField(6, Rdata, 44)
            LlegoMinist = True
            Exit Sub
        Case "SUNI"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
        Case "SUCL"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMain.Label3.Visible = Rdata = 1
            Exit Sub
        Case "SUFA"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMain.Label5.Visible = Rdata = 1
            Exit Sub
        Case "SURE"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMain.Label7.Visible = Rdata = 1
            Exit Sub
        Case "NENE"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.rectxt, "Hay " & Rdata & " npcs.", 255, 255, 255, 0, 0
            Exit Sub
        Case "FMSG"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"
            If Not frmForo.Visible Then
                frmForo.Show
            End If
            Exit Sub
    End Select

    Select Case Left$(sData, 5)
        Case "VERSO"
            frmVerSoporte.lblR.Caption = Right$(Rdata, Len(Rdata) - 5)
            frmVerSoporte.Show , frmMain
            Exit Sub
        Case "TENSO"
            TieneSoporte = True
            Call AddtoRichTextBox(frmMain.rectxt, "Han respondido tu consulta. Clickea nuevamente en el botón S.O.S para leerla.", 0, 0, 255, False, False)
            Exit Sub
        Case "RECOM"
            MiClase = Val(Right$(Rdata, Len(Rdata) - 5))

            Select Case MiClase
                Case TRABAJADOR, CON_MANA
                    frmSubeClase4.Show
                    frmSubeClase4.SetFocus
                Case Else
                    frmSubeClase2.Show
                    frmSubeClase2.SetFocus
            End Select
            Exit Sub
        Case "RELON"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            MiClase = Val(ReadField(1, Rdata, 44))
            Recompensa = Val(ReadField(2, Rdata, 44))
            For i = 1 To 2
                frmRecompensa.Nombre(i) = Recompensas(MiClase, Recompensa, i).Name
                frmRecompensa.Descripcion(i) = Recompensas(MiClase, Recompensa, i).Descripcion
            Next
            frmRecompensa.Visible = True
            frmRecompensa.SetFocus
            Exit Sub
        Case "EIFYA"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmMain.Fuerza = ReadField(1, Rdata, 44)
            frmMain.Agilidad = ReadField(2, Rdata, 44)
            Exit Sub

            'Sistema de Cuentas
        Case "INIAC"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmCuent.Show
            frmConnect.Hide
            'frmCuent.SetFocus
            Exit Sub
            'Sistema de Cuentas
        Case "ADDPJ"
            Rdata = Right$(Rdata, Len(Rdata) - 5)

            rcvName = ReadField(1, Rdata, 44)
            rcvIndex = ReadField(2, Rdata, 44) - 1
            rcvLevel = ReadField(3, Rdata, 44)
            rcvClase = ReadField(4, Rdata, 44)

            frmCuent.lblNombre(rcvIndex).Caption = rcvName
            frmCuent.lblNivel(rcvIndex).Caption = "Nivel: " & rcvLevel
            frmCuent.CP(rcvIndex).Visible = False
            If rcvClase = "ASESINO" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\ASESINO.jpg")
            ElseIf rcvClase = "BARDO" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\BARDO.jpg")
            ElseIf rcvClase = "CARPINTERO" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\CARPINTERO.jpg")
            ElseIf rcvClase = "CAZADOR" Or rcvClase = "ARQUERO" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\CAZADOR.jpg")
            ElseIf rcvClase = "CLERIGO" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\CLERIGO.jpg")
            ElseIf rcvClase = "DRUIDA" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\DRUIDA.jpg")
            ElseIf rcvClase = "GUERRERO" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\GUERRERO.jpg")
            ElseIf rcvClase = "HERRERO" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\HERRERO.jpg")
            ElseIf rcvClase = "LADRON" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\LADRON.jpg")
            ElseIf rcvClase = "LEÑADOR" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\LEÑADOR.jpg")
            ElseIf rcvClase = "MAGO" Or rcvClase = "NIGROMANTE" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\MAGO.jpg")
            ElseIf rcvClase = "MINERO" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\MINERO.jpg")
            ElseIf rcvClase = "PALADIN" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\PALADIN.jpg")
            ElseIf rcvClase = "PESCADOR" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\PESCADOR.jpg")
            ElseIf rcvClase = "PIRATA" Then
                frmCuent.ImgClase(rcvIndex).Picture = LoadPicture(App.Path & "\RECURSOS\GRAFICOS\GRH\PIRATA.jpg")
            End If
            Exit Sub
        Case "DADOS"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            With frmCrearPersonaje
                If .Visible Then
                    .lbFuerza.Caption = ReadField(1, Rdata, 44)
                    .lbAgilidad.Caption = ReadField(2, Rdata, 44)
                    .lbInteligencia.Caption = ReadField(3, Rdata, 44)
                    .lbCarisma.Caption = ReadField(4, Rdata, 44)
                    .lbConstitucion.Caption = ReadField(5, Rdata, 44)

                End If
            End With
            Exit Sub
        Case "MEDOK"
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select

    Select Case Left$(sData, 6)
        Case "SHWSUP"
            frmEnviarSoporte.Show , frmMain
            Exit Sub
        Case "SHWSOP"
            frmPanelSoporte.Show , frmMain
            frmPanelSoporte.lstSoportes.Clear
            frmPanelSoporte.txtSoporte.Text = ""
            Dim a As Integer
            a = ReadField$(2, Rdata, Asc("@"))

            For i = 3 To a + 2
                frmPanelSoporte.lstSoportes.AddItem ReadField$(i, Rdata, Asc("@"))
                DoEvents
            Next i
            Exit Sub
            'S!oporte Dylan.-
        Case "SOPODE"
            If Right$(Rdata, 3) = "0k1" Then
                frmPanelSoporte.shpResp.BackColor = vbGreen
                Rdata = Left$(Rdata, Len(Rdata) - 3)
            Else
                frmPanelSoporte.shpResp.BackColor = vbRed
            End If

            Rdata = Right$(Rdata, Len(Rdata) - 6)
            frmPanelSoporte.txtSoporte = Rdata
            Exit Sub
            'SOPORTE DYLAN.-
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "INVPAR"
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            frmParty2.Visible = True
            frmParty2.Label1.Caption = ReadField(1, Rdata, 44)
            Exit Sub
        Case "SKILLS"
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronSkills = True
            Exit Sub
        Case "PARTYL"
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            frmParty.ListaIntegrantes.Visible = True
            frmParty.Label1.Visible = False
            frmParty.Invitar.Visible = True
            frmParty.Echar.Visible = True
            frmParty.Salir.Visible = True
            For i = 1 To 4
                frmParty.ListaIntegrantes.AddItem ReadField(i, Rdata, 44)
            Next i
            LlegoParty = True
            Exit Sub
        Case "PARTYI"
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            frmParty.ListaIntegrantes.Visible = True
            frmParty.Label1.Visible = False
            frmParty.Invitar.Visible = False
            frmParty.Salir.Visible = True
            frmParty.Echar.Visible = False
            For i = 1 To 4
                frmParty.ListaIntegrantes.AddItem ReadField(i, Rdata, 44)
            Next i
            LlegoParty = True
            Exit Sub
        Case "PARTYN"
            frmParty.ListaIntegrantes.Visible = False
            frmParty.Label1.Visible = True
            frmParty.Invitar.Visible = True
            frmParty.Echar.Visible = False
            frmParty.Salir.Visible = False
            LlegoParty = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmEntrenador.Show
            Exit Sub
    End Select

    Select Case Left$(sData, 7)
        Case "PEACEDE"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "PEACEPR"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "CHRINFO"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            frmCharInfo.SetFocus
            Exit Sub
        Case "LEADERI"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            frmGuildLeader.SetFocus
            Exit Sub
        Case "GINFIG"
            frmGuildLeader.Show
            frmGuildLeader.SetFocus
            Exit Sub
        Case "GINFII"
            frmGuildsNuevo.Show
            frmGuildsNuevo.SetFocus
            Exit Sub
        Case "GINFIJ"
            frmGuildAdm.Show
            frmGuildAdm.SetFocus
            Exit Sub
        Case "MEMBERI"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            Call frmGuildsNuevo.ParseMemberInfo(Rdata)
            frmGuildsNuevo.SetFocus
            Exit Sub
        Case "CLANDET"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            CreandoClan = True
            Call frmGuildFoundation.Show(vbModeless, frmMain)
            Exit Sub
        Case "PETICIO"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub

    End Select


    Select Case UCase$(Left$(Rdata, 9))
        Case "COMUSUINV"
            Rdata = Right$(Rdata, Len(Rdata) - 9)
            OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
            OtroInventario(1).Name = ReadField(3, Rdata, 44)
            OtroInventario(1).Amount = ReadField(4, Rdata, 44)
            OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
            OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
            OtroInventario(1).ObjType = Val(ReadField(7, Rdata, 44))
            OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
            OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
            OtroInventario(1).Def = Val(ReadField(10, Rdata, 44))
            OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))

            frmComerciarUsu.List2.Clear

            frmComerciarUsu.List2.AddItem OtroInventario(1).Name
            frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount

            frmComerciarUsu.lblEstadoResp.Visible = False
    End Select


    Call HandleDosLetras(sData)

    If Not Procesado Then Call InformacionEncriptada(sData)

    Procesado = False

End Sub
Sub InformacionEncriptada(ByVal Rdata As String)
    Dim i As Integer

    For i = 1 To UBound(Mensajes)
        If UCase$(Left$(Rdata, 2)) = UCase$(Mensajes(i).code) Then
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, Reemplazo(Mensajes(i).mensaje, Rdata), CInt(Mensajes(i).Red), CInt(Mensajes(i).Green), CInt(Mensajes(i).Blue), Mensajes(i).Bold = 1, Mensajes(i).Italic = 1
            Exit Sub
        End If
    Next

End Sub
Function Reemplazo(mensaje As String, Rdata As String) As String
    Dim i As Integer

    For i = 1 To Len(mensaje)
        If mid$(mensaje, i, 1) = "*" Then
            Reemplazo = Reemplazo & ReadField(Val(mid$(mensaje, i + 1, 1)), Rdata, 44)
            i = i + 1
        Else
            Reemplazo = Reemplazo & mid$(mensaje, i, 1)
        End If
    Next

End Function
Sub HandleDosLetras(ByVal Rdata As String)
    Dim tempint As Integer
    Dim tempstr As String
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim CharIndex As Integer
    Dim Charindexx As Integer    'victima
    Dim Slot As Integer
    Dim loopc As Integer
    Dim Text1 As String
    Dim Text2 As String
    Dim var3 As Integer
    Dim var2 As Integer
    Dim var1 As Integer
    Dim var4 As Integer

    Select Case Left$(Rdata, 2)
        Case "ET"
            For Y = YMinMapSize To YMaxMapSize
                For X = XMinMapSize To XMaxMapSize
                    If MapData(X, Y).CharIndex > 0 Then Call EraseChar(MapData(X, Y).CharIndex)
                    MapData(X, Y).ObjGrh.GrhIndex = 0
                Next X
            Next Y
            Exit Sub
        Case "°°"
            CONGELADO = True
            Call AddtoRichTextBox(frmMain.rectxt, "¡SERVIDOR CONGELADO, NO PUEDES ENVIAR INFORMACION HASTA QUE SE DESCONGELE!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "ST"
            Stoppeado = True
            Exit Sub
        Case "NT"
            Stoppeado = False
            Exit Sub
        Case "°¬"
            CONGELADO = False
            Call AddtoRichTextBox(frmMain.rectxt, "¡SERVIDOR DESCONGELADO, YA PUEDES ENVIAR INFORMACION AL SERVIDOR!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "CM"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = Val(ReadField(1, Rdata, 44))

            NombreDelMapaActual = ReadField(2, Rdata, 44)
            TopMapa = 13 + Val(ReadField(3, Rdata, 44)) * 18
            IzquierdaMapa = 20 + Val(ReadField(4, Rdata, 44)) * 18



            frmMain.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.X & " - " & UserPos.Y & "]"


            If FileExist(App.Path & "\RECURSOS\MAPS\Mapa" & UserMap & ".mcl", vbNormal) Then
                Open App.Path & "\RECURSOS\MAPS\Mapa" & UserMap & ".mcl" For Binary As #1
                Seek #1, 1
                Close #1
                Call SwitchMapNew(UserMap, False)
                'If bLluvia(UserMap) = 0 Then
                    'If bRain Then
                    '    Audio.StopWave
                    '    frmMain.IsPlaying = plNone
                    'End If
                'End If
            Else

                MsgBox "No se encuentra el mapa instalado."
                Call UnloadAllForms
                End
            End If
            Exit Sub
        Case "PU"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = TeEncripTE(Rdata)
            MapData(UserPos.X, UserPos.Y).CharIndex = 0
            UserPos.X = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
            CharList(UserCharIndex).Pos = UserPos
            Exit Sub
        Case "4&"
            FrmElegirCamino.Show
            FrmElegirCamino.SetFocus
            Exit Sub
        Case "N2"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.rectxt, "¡La criatura te ha pegado en la cabeza por " & Val(ReadField(2, Rdata, 44)) & "!", 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.rectxt, "¡La criatura te ha pegado el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)) & "!", 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.rectxt, "¡La criatura te ha pegado el brazo derecho por " & Val(ReadField(2, Rdata, 44)) & "!", 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.rectxt, "¡La criatura te ha pegado la pierna izquierda por " & Val(ReadField(2, Rdata, 44)) & "!", 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.rectxt, "¡La criatura te ha pegado la pierna derecha por " & Val(ReadField(2, Rdata, 44)) & "!", 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.rectxt, "¡La criatura te ha pegado en el torso por " & Val(ReadField(2, Rdata, 44)) & "!", 255, 0, 0, True, False, False)
            End Select
            Exit Sub

        Case "2H"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)
            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).Name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0
            Call ActualizarInventario(Slot)
            tempstr = ""

            bInvMod = True

            Exit Sub

        Case "6H"
            For loopc = 1 To MAXHECHI
                UserHechizos(loopc) = 0
                If loopc > frmMain.lstHechizos.ListCount Then
                    frmMain.lstHechizos.AddItem "Nada"
                Else
                    frmMain.lstHechizos.List(loopc - 1) = "Nada"
                End If
            Next loopc
            Exit Sub

        Case "1I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, Rdata & " ha sido aceptado en el clan.", 255, 255, 255, 1, 0
            If Opciones.Audio = 1 Then Call Sound.Sound_Play(43)
            Exit Sub
        Case "2I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserInventory(Rdata).Amount = UserInventory(Rdata).Amount - 1
            ActualizarInventario (Rdata)
        Case "3I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)

            UserInventory(Rdata).OBJIndex = 0
            UserInventory(Rdata).Name = "Nada"
            UserInventory(Rdata).Amount = 0
            UserInventory(Rdata).Equipped = 0
            UserInventory(Rdata).GrhIndex = 0
            UserInventory(Rdata).ObjType = 0
            UserInventory(Rdata).MaxHit = 0
            UserInventory(Rdata).MinHit = 0
            UserInventory(Rdata).MaxDef = 0
            UserInventory(Rdata).MinDef = 0
            UserInventory(Rdata).TipoPocion = 0
            UserInventory(Rdata).MaxModificador = 0
            UserInventory(Rdata).MinModificador = 0
            UserInventory(Rdata).Valor = 0

            tempstr = ""
            If UserInventory(Rdata).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Rdata).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Rdata).Amount & ") " & UserInventory(Rdata).Name
            Else
                tempstr = tempstr & UserInventory(Rdata).Name
            End If

            ActualizarInventario (Rdata)

            Exit Sub
        Case "4I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)
            UserInventory(Slot).Amount = UserInventory(Slot).Amount - ReadField(2, Rdata, 44)
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)
        Case "6J"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)
            UserMinAGU = ReadField(2, Rdata, 44)
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 93)
            frmMain.cantidadagua.Caption = UserMinAGU & "/" & UserMaxAGU

            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(43)
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)

            Exit Sub
        Case "6I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)
            UserMinAGU = ReadField(2, Rdata, 44)
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 93)
            frmMain.cantidadagua.Caption = UserMinAGU & "/" & UserMaxAGU

            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).Name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(46)
            End If

            Exit Sub
        Case "7I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = THeEnCripTe(Rdata, pw1(0) + pw1(1) + pw1(2) + pw1(3) + pw1(4) + pw1(5) + pw1(6) + pw1(7) + pw1(8) + pw1(9))
            Slot = ReadField(1, Rdata, 44)

            UserMinMAN = ReadField(2, Rdata, 44)
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
                frmMain.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmMain.MANShp.Width = 0
                frmMain.cantidadmana.Caption = ""
            End If
            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(46)
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)

            Exit Sub
        Case "8I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)
            UserMinMAN = ReadField(2, Rdata, 44)
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
                frmMain.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmMain.MANShp.Width = 0
                frmMain.cantidadmana.Caption = ""
            End If
            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).Name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(46)
            End If

            Exit Sub
        Case "9I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)

            UserMinHP = ReadField(2, Rdata, 44)
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            frmMain.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(46)
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)

            Exit Sub
        Case "2J"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)
            UserMinHP = ReadField(2, Rdata, 44)
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            frmMain.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).Name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(46)
            End If

            Exit Sub
        Case "3J"
            Slot = Right$(Rdata, Len(Rdata) - 2)

            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(46)
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)
            Exit Sub
        Case "4J"
            Slot = Right$(Rdata, Len(Rdata) - 2)

            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).Name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""

            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(46)
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If
            ActualizarInventario (Slot)
            Exit Sub

        Case "8J"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserInventory(Rdata).Equipped = 0

            If UserInventory(Rdata).ObjType = 2 Then
                frmMain.arma.Caption = "N/A"
            ElseIf UserInventory(Rdata).ObjType = 3 Then
                Select Case UserInventory(Rdata).SubTipo
                    Case 0
                        frmMain.armadura.Caption = "N/A"
                    Case 1
                        frmMain.casco.Caption = "N/A"
                    Case 2
                        frmMain.escudo.Caption = "N/A"
                End Select


            End If
            tempstr = ""
            If UserInventory(Rdata).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Rdata).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Rdata).Amount & ") " & UserInventory(Rdata).Name
            Else
                tempstr = tempstr & UserInventory(Rdata).Name
            End If

            ActualizarInventario (Rdata)
            Exit Sub
        Case "7J"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserInventory(Rdata).Equipped = 1

            If UserInventory(Rdata).ObjType = 2 Then
                frmMain.arma.Caption = UserInventory(Rdata).MinHit & " / " & UserInventory(Rdata).MaxHit
            ElseIf UserInventory(Rdata).ObjType = 3 Then
                Select Case UserInventory(Rdata).SubTipo
                    Case 0
                        If UserInventory(Rdata).MaxDef > 0 Then
                            frmMain.armadura.Caption = UserInventory(Rdata).MinDef & " / " & UserInventory(Rdata).MaxDef
                        Else
                            frmMain.armadura.Caption = "N/A"
                        End If

                    Case 1
                        If UserInventory(Rdata).MaxDef > 0 Then
                            frmMain.casco.Caption = UserInventory(Rdata).MinDef & " / " & UserInventory(Rdata).MaxDef
                        Else
                            frmMain.casco.Caption = "N/A"
                        End If

                    Case 2
                        If UserInventory(Rdata).MaxDef > 0 Then
                            frmMain.escudo.Caption = UserInventory(Rdata).MinDef & " / " & UserInventory(Rdata).MaxDef
                        Else
                            frmMain.escudo.Caption = "N/A"
                        End If

                End Select
            End If

            tempstr = ""
            If UserInventory(Rdata).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Rdata).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Rdata).Amount & ") " & UserInventory(Rdata).Name
            Else
                tempstr = tempstr & UserInventory(Rdata).Name
            End If

            ActualizarInventario (Rdata)
            Exit Sub
        Case "6K"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)
            UserMinHAM = ReadField(2, Rdata, 44)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 93)
            frmMain.cantidadhambre.Caption = UserMinHAM & "/" & UserMaxHAM

            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(7)
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)

            Exit Sub
        Case "7K"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Slot = ReadField(1, Rdata, 44)
            UserMinHAM = ReadField(2, Rdata, 44)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 93)
            frmMain.cantidadhambre.Caption = UserMinHAM & "/" & UserMaxHAM

            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).Name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If

            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).Name
            Else
                tempstr = tempstr & UserInventory(Slot).Name
            End If

            ActualizarInventario (Slot)
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(7)
            End If

            Exit Sub
        Case "3Q"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim ibser As Integer
            ibser = Val(ReadField(3, Rdata, 176))
            If ibser > 0 Then
                Dialogos.CrearDialogo ReadField(2, Rdata, 176), ibser, Val(ReadField(1, Rdata, 176))





            Else
                If PuedoQuitarFoco Then _
                   AddtoRichTextBox frmMain.rectxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
            End If
            Exit Sub
        Case "9Q"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim CRI As String
            Text1 = ReadField(1, Rdata, 44)
            Text2 = ReadField(2, Rdata, 44)

            Select Case Val(Text2)
                Case 1
                    CRI = " [Herido]"
                Case 2
                    CRI = " [Levemente herido]"
                Case 3
                    CRI = " [Muy herido]"
                Case 4
                    CRI = " [Agonizando]"
                Case 5
                    CRI = " [Sano]"
                Case Is > 5
                    CRI = " [" & Val(Text2) - 5 & "0% herido]"
            End Select

            AddtoRichTextBox frmMain.rectxt, Text1 & CRI, 65, 190, 156, 0, 0
            Exit Sub
        Case "7T"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Text1 = ReadField(1, Rdata, 172)
            Text2 = ReadField(2, Rdata, 172)
            var1 = Val(ReadField(3, Rdata, 172))
            var2 = Val(ReadField(4, Rdata, 172))
            var3 = Val(ReadField(5, Rdata, 172))
            AddtoRichTextBox frmMain.rectxt, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%", 65, 190, 156, 0, 0
            AddtoRichTextBox frmMain.rectxt, "Nombre del hechizo: " & Text1, 65, 190, 156, 0, 0
            AddtoRichTextBox frmMain.rectxt, "Descripción: " & Text2, 65, 190, 156, 0, 0
            AddtoRichTextBox frmMain.rectxt, "Skill requerido: " & var1, 65, 190, 156, 0, 0
            AddtoRichTextBox frmMain.rectxt, "Mana necesaria: " & var2, 65, 190, 156, 0, 0
            AddtoRichTextBox frmMain.rectxt, "Energia necesaria: " & var3, 65, 190, 156, 0, 0
            Exit Sub
        Case "1U"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            var1 = Val(ReadField(1, Rdata, 44))
            var2 = Val(ReadField(2, Rdata, 44))
            var3 = Val(ReadField(3, Rdata, 44))
            var4 = Val(ReadField(4, Rdata, 44))
            If var1 > 0 Then AddtoRichTextBox frmMain.rectxt, "Has ganado " & var1 & " puntos de vida.", 200, 200, 200, 0, 0
            If var2 > 0 Then AddtoRichTextBox frmMain.rectxt, "Has ganado " & var2 & " puntos de vitalidad.", 200, 200, 200, 0, 0
            If var3 > 0 Then AddtoRichTextBox frmMain.rectxt, "Has ganado " & var3 & " puntos de mana.", 200, 200, 200, 0, 0
            If var4 > 0 Then AddtoRichTextBox frmMain.rectxt, "Tu golpe maximo aumentó en " & var4 & " puntos.", 200, 200, 200, 0, 0
            If var4 > 0 Then AddtoRichTextBox frmMain.rectxt, "Tu golpe mínimo aumentó en " & var4 & " puntos.", 200, 200, 200, 0, 0
            Exit Sub
        Case "6Z"
            AddtoRichTextBox frmMain.rectxt, "¡Hoy es la votación para elegir un nuevo lider para el clan!", 255, 255, 255, 1, 0
            AddtoRichTextBox frmMain.rectxt, "La elección durará 24 horas, se puede votar a cualquier miembro del clan.", 255, 255, 255, 1, 0
            AddtoRichTextBox frmMain.rectxt, "Para votar escribe /VOTO NICKNAME.", 255, 255, 255, 1, 0
            AddtoRichTextBox frmMain.rectxt, "Sólo se computara un voto por miembro.", 255, 255, 255, 1, 0
            Exit Sub
        Case "7Z"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, "¡Las elecciones han finalizado!", 255, 255, 255, 1, 0
            AddtoRichTextBox frmMain.rectxt, "El nuevo lider es: " & Rdata, 255, 255, 255, 1, 0
            Exit Sub
        Case "!J"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, "Felicitaciones, tu solicitud ha sido aceptada.", 255, 255, 255, 1, 0
            AddtoRichTextBox frmMain.rectxt, "Ahora sos un miembro activo del clan " & Rdata, 255, 255, 255, 1, 0
            Exit Sub
        Case "!R"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, "Tu clan ha firmado una alianza con " & Rdata, 255, 255, 255, 1, 0
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(45)
            End If
            Exit Sub
        Case "!S"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, Rdata & " firmó una alianza con tu clan.", 255, 255, 255, 1, 0
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(45)
            End If
            Exit Sub
        Case "!U"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, "Tu clan le declaró la guerra a " & Rdata, 255, 255, 255, 1, 0
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(45)
            End If
            Exit Sub
        Case "!V"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, Rdata & " le declaró la guerra a tu clan.", 255, 255, 255, 1, 0
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(45)
            End If
            Exit Sub
        Case "!4"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Text1 = ReadField(1, Rdata, 44)
            Text2 = ReadField(2, Rdata, 44)
            AddtoRichTextBox frmMain.rectxt, "¡" & Text1 & " fundó el clan " & Text2 & "!", 255, 255, 255, 1, 0
            If Opciones.Audio = 1 Then
                Call Sound.Sound_Play(44)
            End If
            Exit Sub
        Case "/O"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call Dialogos.CrearDialogo("El negocio va bien, ya he conseguido " & ReadField(1, Rdata, 44) & " monedas de oro en ventas. He enviado el dinero directamente a tu cuenta en el banco.", Val(ReadField(2, Rdata, 44)), vbWhite)
            Exit Sub
        Case "/P"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call Dialogos.CrearDialogo("El negocio no va muy bien, todavía no he podido vender nada. Si consigo una venta, enviare el dinero directamente a tu cuenta en el banco.", Val(Rdata), vbWhite)
            Exit Sub
        Case "/Q"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call Dialogos.CrearDialogo("¡Buen día! Ahora estoy contratado por " & ReadField(1, Rdata, 44) & " para vender sus objetos, ¿quieres echar un vistazo?", Val(ReadField(2, Rdata, 44)), vbWhite)
            Exit Sub
        Case "/R"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, ReadField(1, Rdata, 44) & " compró " & ReadField(2, Rdata, 44) & " (" & PonerPuntos(Val(ReadField(3, Rdata, 44))) & ") en tu tienda por " & PonerPuntos(Val(ReadField(4, Rdata, 44))) & " monedas de oro.", 255, 255, 255, 1, 0
            AddtoRichTextBox frmMain.rectxt, "El dinero fue enviado directamente a tu cuenta de banco.", 255, 255, 255, 1, 0
            Exit Sub
        Case "/V"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call Dialogos.CrearDialogo("Solo los trabajadores experimentados y los personajes mayores a nivel 25 con más de 65 en comercio pueden utilizar mis servicios de vendedor.", Val(Rdata), vbWhite)
            Exit Sub
        Case "/X"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, "Numero total de ventas: " & PonerPuntos(Val(ReadField(2, Rdata, 44))), 65, 190, 156, 1, 0
            AddtoRichTextBox frmMain.rectxt, "Dinero movido por las ventas: " & PonerPuntos(Val(ReadField(1, Rdata, 44))) & " monedas de oro.", 65, 190, 156, 1, 0
            AddtoRichTextBox frmMain.rectxt, "Venta promedio: " & PonerPuntos(Val(ReadField(1, Rdata, 44)) / Val(ReadField(2, Rdata, 44))) & " monedas de oro.", 65, 190, 156, 1, 0
            Exit Sub
        Case "{B"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, "Has iniciado el modo de susurro con " & Rdata & ".", 255, 255, 255, 1, 0
            frmMain.MousePointer = 1
            Exit Sub
        Case "{C"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, "No puedes iniciar el modo de susurro contigo mismo.", 255, 255, 255, 1, 0
            frmMain.modo = "1 Normal"
            frmMain.MousePointer = 1
            Exit Sub
        Case "{D"
            AddtoRichTextBox frmMain.rectxt, "Objetivo invalido!.", 65, 190, 156, 0, 0
            frmMain.modo = "1 Normal"
            frmMain.MousePointer = 1
            Exit Sub
        Case "{F"
            AddtoRichTextBox frmMain.rectxt, "El usuario ya no se encuentra en tu pantalla.", 65, 190, 156, 0, 0
            frmMain.modo = "1 Normal"
            frmMain.MousePointer = 1
            Exit Sub
        Case "8B"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            frmMain.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)

            Exit Sub
        Case "9B"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMaxMAN = Val(ReadField(1, Rdata, 44))

            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
                frmMain.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmMain.MANShp.Width = 0
                frmMain.cantidadmana.Caption = ""
            End If

            Exit Sub
        Case "1N"
            If Opciones.CartelSanado = 1 Then AddtoRichTextBox frmMain.rectxt, "Has sanado.", 65, 190, 156, 0, 0
            Exit Sub
        Case "V5"
            If Opciones.CartelOcultarse = 1 Then AddtoRichTextBox frmMain.rectxt, "¡Has vuelto a ser visible!", 65, 190, 156, 0, 0
            Exit Sub
        Case "MN"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = THeEnCripTe(Rdata, pw1(0) + pw1(1) + pw1(2) + pw1(3) + pw1(4) + pw1(5) + pw1(6) + pw1(7) + pw1(8) + pw1(9))
            If Opciones.CartelRecuMana = 1 Then AddtoRichTextBox frmMain.rectxt, "¡Has recuperado " & Rdata & " puntos de mana!", 65, 190, 156, 0, 0
            Exit Sub
        Case "8K"
            If Opciones.CartelNoHayNada = 1 Then AddtoRichTextBox frmMain.rectxt, "¡No hay nada aquí!", 65, 190, 156, 0, 0
            Exit Sub
        Case "DN"
            If Opciones.CartelMenosCansado = 1 Then AddtoRichTextBox frmMain.rectxt, "Has dejado de descansar.", 65, 190, 156, 0, 0
            Exit Sub
        Case "D9"
            If Opciones.CartelRecuMana = 1 Then AddtoRichTextBox frmMain.rectxt, "Ya no estás meditando.", 65, 190, 156, 0, 0
            Exit Sub
        Case "{{"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            AddtoRichTextBox frmMain.rectxt, "(" & ReadField(1, Rdata, 44) & ") " & KeyName(ReadField(2, Rdata, 44)), 65, 190, 156, 0, 0
            Exit Sub
        Case "MV"
            If Opciones.CartelMenosCansado = 1 Then AddtoRichTextBox frmMain.rectxt, "Te sentis menos cansado.", 65, 190, 156, 0, 0
            Exit Sub
        Case "FR"
            If Opciones.CartelVestirse = 1 Then AddtoRichTextBox frmMain.rectxt, "¡Has perdido stamina, si no te abrigas rápido perderas toda!", 65, 190, 156, 0, 0
            Exit Sub
        Case "1K"
            If Opciones.CartelVestirse = 1 Then AddtoRichTextBox frmMain.rectxt, "¡Estás muriendo de frío, abrígate o moriras!", 65, 190, 156, 0, 0
            Exit Sub
        Case "7M"
            If Opciones.CartelRecuMana = 1 Then AddtoRichTextBox frmMain.rectxt, "Comienzas a meditar.", 65, 190, 156, 0, 0
            Exit Sub
        Case "8M"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            If Opciones.CartelRecuMana = 1 Then AddtoRichTextBox frmMain.rectxt, "Te estás concentrando. En " & Rdata & " segundos comenzarás a meditar.", 65, 190, 156, 0, 0
            Exit Sub
        Case "EL"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            If Rdata <> 0 Then AddtoRichTextBox frmMain.rectxt, "Has ganado " & Rdata & " puntos de experiencia.", 255, 150, 25, 1, 0
            AddtoRichTextBox frmMain.rectxt, "¡Has matado a la criatura!", 65, 190, 156, 0, 0
            Exit Sub
        Case "V7"
            If Opciones.CartelOcultarse = 1 Then AddtoRichTextBox frmMain.rectxt, "¡Te has escondido entre las sombras!", 65, 190, 156, 0, 0
            Exit Sub
        Case "EN"
            If Opciones.CartelOcultarse = 1 Then AddtoRichTextBox frmMain.rectxt, "¡No has logrado esconderte!", 65, 190, 156, 0, 0
            Exit Sub
        Case "V3"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = TeEncripTE(Rdata)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            Exit Sub
        Case "N4"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado en la cabeza por " & Val(ReadField(2, Rdata, 44)) & "!!", 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)) & "!!", 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado el brazo derecho por " & Val(ReadField(2, Rdata, 44)) & "!!", 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado la pierna izquierda por " & Val(ReadField(2, Rdata, 44)) & "!!", 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado la pierna derecha por " & Val(ReadField(2, Rdata, 44)) & "!!", 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡" & ReadField(3, Rdata, 44) & " te ha pegado en el torso por " & Val(ReadField(2, Rdata, 44)) & "!!", 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la cabeza por " & Val(ReadField(2, Rdata, 44)), 230, 230, 0, 1, 0)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el brazo izquierdo por " & Val(ReadField(2, Rdata, 44)), 230, 230, 0, 1, 0)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el brazo derecho por " & Val(ReadField(2, Rdata, 44)), 230, 230, 0, 1, 0)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la pierna izquierda por " & Val(ReadField(2, Rdata, 44)), 230, 230, 0, 1, 0)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en la pierna derecha por " & Val(ReadField(2, Rdata, 44)), 230, 230, 0, 1, 0)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.rectxt, "¡¡Le has pegado a " & ReadField(3, Rdata, 44) & " en el torso por " & Val(ReadField(2, Rdata, 44)), 230, 230, 0, 1, 0)
            End Select
            Exit Sub

        Case "|$"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            tempint = InStr(1, Rdata, ">>")
            tempstr = mid(Rdata, 1, tempint)
            Call AddtoRichTextBox(frmMain.rectxt, tempstr, 65, 143, 190, 0, 0, True)
            tempstr = Right$(Rdata, Len(Rdata) - tempint)
            Call AddtoRichTextBox(frmMain.rectxt, tempstr, 243, 255, 277, 0, 0)
            Exit Sub
        Case "||"
            Dim iUser As Integer
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            iUser = Val(ReadField(3, Rdata, 176))
            If iUser > 0 Then
                If Val(ReadField(1, Rdata, 176)) <> vbCyan And EstaIgnorado(iUser) Then
                    Dialogos.CrearDialogo "", iUser, Val(ReadField(1, Rdata, 176))
                    Exit Sub
                Else
                    Dialogos.CrearDialogo ReadField(2, Rdata, 176), iUser, Val(ReadField(1, Rdata, 176))
                End If
            Else
                If PuedoQuitarFoco Then _
                   AddtoRichTextBox frmMain.rectxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
            End If
            Exit Sub
        Case "!!"
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                frmMensaje.msg.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
        Case "IU"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
        Case "IP"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = CharList(UserCharIndex).Pos
            frmMain.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.X & " - " & UserPos.Y & "]"
            Exit Sub
        Case "CC"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
            CharList(CharIndex).FX = Val(ReadField(9, Rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
            CharList(CharIndex).Nombre = ReadField(12, Rdata, 44)

            If Right$(CharList(CharIndex).Nombre, 2) = "<>" Then
                CharList(CharIndex).Nombre = Left$(CharList(CharIndex).Nombre, Len(CharList(CharIndex).Nombre) - 2)
            End If

            CharList(CharIndex).Criminal = Val(ReadField(13, Rdata, 44))

            CharList(CharIndex).invisible = (Val(ReadField(14, Rdata, 44)) = 1)
            Call MakeChar(CharIndex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)), Val(ReadField(15, Rdata, 44)), Val(ReadField(16, Rdata, 44)))

            Exit Sub

        Case "PW"
            Rdata = Right$(Rdata, Len(Rdata) - 2)

            CharIndex = ReadField(1, Rdata, 44)
            CharList(CharIndex).Criminal = Val(ReadField(2, Rdata, 44))
            CharList(CharIndex).Nombre = ReadField(3, Rdata, 44)

            Exit Sub

        Case "BP"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = THeEnCripTe(Rdata, pw1(0) + pw1(1) + pw1(2) + pw1(3) + pw1(4) + pw1(5) + pw1(6) + pw1(7) + pw1(8) + pw1(9))
            Call EraseChar(Val(Rdata))
            Exit Sub

        Case "MP"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = THeEnCripTe(Rdata, pw1(0) + pw1(1) + pw1(2) + pw1(3) + pw1(4) + pw1(5) + pw1(6) + pw1(7) + pw1(8) + pw1(9))
            CharIndex = Val(ReadField(1, Rdata, 44))

            If Opciones.Audio = 1 Then Call DoPasosFx(CharIndex)

            Call MoveCharByPos(CharIndex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))

            Exit Sub
        Case "LP"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            If Opciones.Audio = 1 Then Call DoPasosFx(CharIndex)

            Call MoveCharByPosConHeading(CharIndex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44))

            Exit Sub
        Case "ZZ"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))

            If Opciones.Audio = 1 Then Call DoPasosFx(CharIndex)

            Call MoveCharByPosAndHead(CharIndex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44))
            Exit Sub
        Case "CP"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).muerto = Val(ReadField(3, Rdata, 44)) = 500
            Slot = Val(ReadField(2, Rdata, 44))
            CharList(CharIndex).Body = BodyData(Slot)
            CharList(CharIndex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
            If Slot > 83 And Slot < 88 Then
                CharList(CharIndex).Navegando = 1
            Else
                CharList(CharIndex).Navegando = 0
            End If
            CharList(CharIndex).Heading = Val(ReadField(4, Rdata, 44))
            CharList(CharIndex).FX = Val(ReadField(7, Rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadField(8, Rdata, 44))
            tempint = Val(ReadField(5, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).arma = WeaponAnimData(tempint)
            tempint = Val(ReadField(6, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).escudo = ShieldAnimData(tempint)
            tempint = Val(ReadField(9, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).casco = CascoAnimData(tempint)

            tempint = Val(ReadField(10, Rdata, 44))
            CharList(CharIndex).Alas = BodyData(tempint)

            Exit Sub
        Case "2C"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).FX = 0
            CharList(CharIndex).FxLoopTimes = 0
            CharList(CharIndex).Heading = Val(ReadField(2, Rdata, 44))
            If CharList(CharIndex).ParticleIndex <> 0 Then
                effect(CharList(CharIndex).ParticleIndex).Progression = 1
                CharList(CharIndex).ParticleIndex = 0
            End If
            Exit Sub
        Case "3C"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            Slot = Val(ReadField(2, Rdata, 44))
            CharList(CharIndex).Body = BodyData(Slot)
            If Slot > 83 And Slot < 88 Then
                CharList(CharIndex).Navegando = 1
            Else
                CharList(CharIndex).Navegando = 0
            End If
            Exit Sub
        Case "4C"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            CharList(CharIndex).Head = HeadData(Val(ReadField(2, Rdata, 44)))
            Exit Sub
        Case "5C"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            tempint = Val(ReadField(2, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).arma = WeaponAnimData(tempint)
            Exit Sub
        Case "6C"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            tempint = Val(ReadField(2, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).escudo = ShieldAnimData(tempint)
            Exit Sub
        Case "7C"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            tempint = Val(ReadField(2, Rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).casco = CascoAnimData(tempint)
            Exit Sub
        Case "5A"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = TeEncripTE(Rdata)
            UserMinHP = Val(ReadField(1, Rdata, 44))
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            frmMain.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)

            Exit Sub
        Case "5D"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMinMAN = Val(ReadField(1, Rdata, 44))

            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
                frmMain.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmMain.MANShp.Width = 0
                frmMain.cantidadmana.Caption = ""
            End If

            Exit Sub

        Case "5E"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMinSTA = Val(ReadField(1, Rdata, 44))

            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
            frmMain.cantidadsta.Caption = PonerPuntos(UserMinSTA) & "/" & PonerPuntos(UserMaxSTA)

            Exit Sub

        Case "5F"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserGLD = Val(ReadField(1, Rdata, 44))

            frmMain.GldLbl.Caption = PonerPuntos(UserGLD)

            Exit Sub

        Case "5G"
            Rdata = Right$(Rdata, Len(Rdata) - 2)

            UserExp = Val(ReadField(1, Rdata, 44))

            If UserPasarNivel <> 0 Then
                frmMain.exp.Caption = "Exp: " & PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
                frmMain.LvlLbl.Caption = UserLvl
            Else
                frmMain.exp.Caption = ""
            End If
        Case "5H"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMinMAN = Val(ReadField(1, Rdata, 44))
            UserMinSTA = Val(ReadField(2, Rdata, 44))

            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
                frmMain.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmMain.MANShp.Width = 0
                frmMain.cantidadmana.Caption = ""
            End If

            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
            frmMain.cantidadsta.Caption = PonerPuntos(UserMinSTA) & "/" & PonerPuntos(UserMaxSTA)

            Exit Sub

        Case "5I"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMinHP = Val(ReadField(1, Rdata, 44))
            UserMinSTA = Val(ReadField(2, Rdata, 44))

            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            frmMain.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)

            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
            frmMain.cantidadsta.Caption = PonerPuntos(UserMinSTA) & "/" & PonerPuntos(UserMaxSTA)

            Exit Sub
        Case "5J"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMinAGU = Val(ReadField(1, Rdata, 44))
            UserMinHAM = Val(ReadField(2, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 93)
            frmMain.cantidadagua.Caption = UserMinAGU & "/" & UserMaxAGU
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 93)
            frmMain.cantidadhambre.Caption = UserMinHAM & "/" & UserMaxHAM

            Exit Sub
        Case "5O"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserLvl = Val(ReadField(1, Rdata, 44))
            UserPasarNivel = Val(ReadField(2, Rdata, 44))
            If UserPasarNivel > 0 Then
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.exp.Caption = "Exp: " & PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
            Else
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.exp.Caption = ""
            End If
            Exit Sub
        Case "HO"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = TeEncripTE(Rdata)
            X = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))

            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            LastPos.X = X
            LastPos.Y = Y
            Exit Sub
        Case "P8"
            UserParalizado = False
            AddtoRichTextBox frmMain.rectxt, "Ya no estás paralizado.", 65, 190, 156, 0, 0
            Exit Sub
        Case "P9"
            UserParalizado = True
            Call SendData("RPU")
            AddtoRichTextBox frmMain.rectxt, "Estás paralizado. No podrás moverte por algunos segundos.", 65, 190, 156, 0, 0
            Exit Sub
        Case "BO"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Rdata = TeEncripTE(Rdata)
            X = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "TM"
            Rdata = Val(Right$(Rdata, Len(Rdata) - 2))
            If Rdata <> 0 And CurrentMP3 <> Rdata Then
                CurrentMP3 = Rdata
                If Opciones.sMusica <> CONST_DESHABILITADA Then
                    If Opciones.sMusica <> CONST_DESHABILITADA Then
                        Sound.NextMusic = CurrentMP3
                        Sound.Fading = 350
                    End If
                End If
            End If
            Exit Sub
        Case "LH"
            LastHechizo = Timer
            Exit Sub
        Case "LG"
            LastGolpe = Timer
            Exit Sub
        Case "LF"
            LastFlecha = Timer
            Exit Sub
        Case "TW"
            If Opciones.Audio = 1 Then
                Dim PosX, PosY As Byte
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                PosX = Val(ReadField(2, Rdata, 44)) 'victima
                PosY = Val(ReadField(3, Rdata, 44)) 'victima
                'Call Sound.Sound_Play(Rdata & ".wav")
                If PosX > 0 And PosY > 0 Then
                    Call Sound.Sound_Play(Val(ReadField(1, Rdata, 44)), , Sound.Calculate_Volume(PosX, PosY), Sound.Calculate_Pan(PosX, PosY))
                Else
                    Call Sound.Sound_Play(Val(ReadField(1, Rdata, 44)))
                End If
            End If
            Exit Sub
        Case "TX"
            Dim Efecto As Integer
            Dim ParticleCasteada As Integer
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))    'atacante
            Charindexx = Val(ReadField(2, Rdata, 44))    'victima
            Efecto = Val(ReadField(4, Rdata, 44))    'efecto particulas
            If Not Charindexx = 0 Then
                If Opciones.Audio = 1 Then
                    'Call Sound.Sound_Play(ReadField(6, Rdata, 44) & ".wav")
                    Call Sound.Sound_Play(Val(ReadField(6, Rdata, 44)), , Sound.Calculate_Volume(CharList(Charindexx).Pos.X, CharList(Charindexx).Pos.Y), Sound.Calculate_Pan(CharList(Charindexx).Pos.X, CharList(Charindexx).Pos.Y))
                End If
                If Efecto = 0 Or Opciones.Particulas = 1 Then
                    CharList(Charindexx).FX = Val(ReadField(3, Rdata, 44))
                    CharList(Charindexx).FxLoopTimes = Val(ReadField(5, Rdata, 44))
                ElseIf Opciones.Particulas = 0 Then  'si está desactivado
                    ParticleCasteada = Engine_UTOV_Particle(CharIndex, Charindexx, Efecto)
                End If
            Else
                If Opciones.Audio = 1 Then
                    Call Sound.Sound_Play(Val(ReadField(6, Rdata, 44)))
                End If
            End If
            Exit Sub
        Case "GL"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            frmGuildAdm.guildslist.Clear
            Call frmGuildAdm.ParseGuildList(Rdata)
            frmGuildAdm.SetFocus
            Exit Sub
        Case "FO"
            bFogata = True
                'Sound.Sound_Stop
                DoFogataFx

            Exit Sub
        Case "XX"
            UserParalizado = Not UserParalizado

            Exit Sub
    End Select

End Sub
Public Function ReplaceData(sdData As String) As String
    Dim Rdata As String

    If UCase$(Left$(sdData, 9)) = "/PASSWORD" Then
        frmCambiarPasswd.Show
        ReplaceData = "NOPUDO"
        Exit Function
    End If

    Select Case UCase$(sdData)
        Case Is = "/MEDITAR"
            ReplaceData = "#A"
        Case Is = "/SALIR"
            ReplaceData = "#B"
        Case "/FUNDARCLAN"
            ReplaceData = "#C"
        Case "/BALANCE"
            ReplaceData = "#G"
        Case "/QUIETO"
            ReplaceData = "#H"
        Case "/ACOMPAÑAR"
            ReplaceData = "#I"
        Case "/ENTRENAR"
            ReplaceData = "#J"
        Case "/DESCANSAR"
            ReplaceData = "#K"
        Case "/DESAFIAR"
            ReplaceData = "#-"
        Case "/IRDESAFIO"
            ReplaceData = "#,"
        Case "/RESUCITAR"
            ReplaceData = "#L"
        Case "/ULLA"
            ReplaceData = "#Z"
        Case "/NIX"
            ReplaceData = "#X"
        Case "/LINDOS"
            ReplaceData = "#®"
        Case "/BANDER"
            ReplaceData = "#¥"
        Case "/ARGHAL"
            ReplaceData = "#Ø"
        Case "/CURAR"
            ReplaceData = "#M"
        Case "/ONLINE"
            ReplaceData = "#P"
        Case "/ESPADA"
            ReplaceData = "#)"
        Case "/ESCUDO"
            ReplaceData = "#?"
        Case "/ARMADURA"
            ReplaceData = "#¡"
        Case "/ANILLO"
            ReplaceData = "#¬"
        Case "/NOBLE"
            ReplaceData = "#°"
        Case "/ENTRAR"
            ReplaceData = "#("
        Case "/RETAR"
            ReplaceData = "#?"
        Case "/ACEPTAR"
            ReplaceData = "#€"
        Case "/SI"
            ReplaceData = "VSI"
        Case "/NO"
            ReplaceData = "VNO"
        Case "/IGNORADOS"
            Call MostrarIgnorados
            ReplaceData = "NOPUDO"
            Exit Function
        Case "/EST"
            ReplaceData = "#Q"
        Case "/PENA"
            ReplaceData = "#R"
        Case "/MOVER"
            ReplaceData = "#S"
        Case "/PARTICIPAR"
            ReplaceData = "#T"
        Case "/ATRAPADO"
            ReplaceData = "#U"
        Case "/COMERCIAR"
            ReplaceData = "#V"
        Case "/BOVEDA"
            ReplaceData = "#W"
        Case "/ENLISTAR"
            ReplaceData = "#Y"
        Case "/INFORMACION"
            ReplaceData = "#Z"
        Case "/RECOMPENSA"
            ReplaceData = "#1"
        Case "/SALIRCLAN"
            ReplaceData = "#2"
        Case "/ONLINECLAN"
            ReplaceData = "#3"
        Case "/ABANDONAR"
            ReplaceData = "#4"
        Case "/SEGUROCLAN"
            ReplaceData = "#·"
    End Select

    Select Case UCase$(Left$(sdData, 6))
        Case "/DESC "
            Rdata = Right$(sdData, Len(sdData) - 6)
            ReplaceData = "#5 " & Rdata
        Case "/VOTO "
            Rdata = Right$(sdData, Len(sdData) - 6)
            ReplaceData = "#6 " & Rdata
        Case "/CMSG "
            Rdata = Right$(sdData, Len(sdData) - 6)
            ReplaceData = "#7 " & Rdata
    End Select

    Select Case UCase$(Left$(sdData, 8))
        Case "/PASSWD "
            Rdata = Right$(sdData, Len(sdData) - 8)
            ReplaceData = "#8 " & Rdata
        Case "/ONLINE "
            Rdata = Right$(sdData, Len(sdData) - 8)
            ReplaceData = "#*" & Rdata
    End Select

    Select Case UCase$(Left$(sdData, 9))
        Case "/APOSTAR "
            Rdata = Right$(sdData, Len(sdData) - 9)
            ReplaceData = "#9 " & Rdata
        Case "/RETIRAR "
            Rdata = Right$(sdData, Len(sdData) - 9)
            ReplaceData = "#0 " & Rdata
        Case "/IGNORAR "
            Rdata = Right$(sdData, Len(sdData) - 9)
            Select Case IgnorarPJ(Rdata)
                Case 0
                    ReplaceData = "NOPUDO"
                    Exit Function
                Case 1
                    ReplaceData = "#/ " & Rdata & " 1"
                Case 2
                    ReplaceData = "#/ " & Rdata & " 0"
            End Select
    End Select
    Select Case UCase$(Left$(sdData, 10))
        Case "/CONGELAR "
            Rdata = Right$(sdData, Len(sdData) - 10)
            ReplaceData = "#\ " & Rdata
    End Select

    Select Case UCase$(Left$(sdData, 11))
        Case "/DEPOSITAR "
            Rdata = Right$(sdData, Len(sdData) - 11)
            ReplaceData = "#Ñ " & Rdata
        Case "/DENUNCIAR "
            Rdata = Right$(sdData, Len(sdData) - 11)
            ReplaceData = "^A " & Rdata
        Case "/INSCRIBIR "
            Rdata = Right$(sdData, Len(sdData) - 11)
            ReplaceData = "#~ " & Rdata
    End Select
    Select Case UCase$(Left$(sdData, 13))
        Case "/DESCONGELAR "
            Rdata = Right$(sdData, Len(sdData) - 13)
            ReplaceData = "#º " & Rdata
    End Select

    If Len(ReplaceData) = 0 Then ReplaceData = sdData

End Function
Function KeyName(Key As String) As String
    Dim KeyCode As Byte

    KeyCode = Asc(Key)

    Select Case KeyCode
        Case vbKeyAdd: KeyName = "+ (KeyPad)"
        Case vbKeyBack: KeyName = "Delete"
        Case vbKeyCancel: KeyName = "Cancelar"
        Case vbKeyCapital: KeyName = "CapsLock"
        Case vbKeyClear: KeyName = "Borrar"
        Case vbKeyControl: KeyName = "Control"
        Case vbKeyDecimal: KeyName = ". (KeyPad)"
        Case vbKeyDelete: KeyName = "Suprimir"
        Case vbKeyDivide: KeyName = "/ (KeyPad)"
        Case vbKeyEnd: KeyName = "Fin"
        Case vbKeyEscape: KeyName = "Esc"
        Case vbKeyF1: KeyName = "F1"
        Case vbKeyF2: KeyName = "F2"
        Case vbKeyF3: KeyName = "F3"
        Case vbKeyF4: KeyName = "F4"
        Case vbKeyF5: KeyName = "F5"
        Case vbKeyF6: KeyName = "F6"
        Case vbKeyF7: KeyName = "F7"
        Case vbKeyF8: KeyName = "F8"
        Case vbKeyF9: KeyName = "F9"
        Case vbKeyF10: KeyName = "F10"
        Case vbKeyF11: KeyName = "F11"
        Case vbKeyF12: KeyName = "F12"
        Case vbKeyF13: KeyName = "F13"
        Case vbKeyF14: KeyName = "F14"
        Case vbKeyF15: KeyName = "F15"
        Case vbKeyF16: KeyName = "F16"
        Case vbKeyHome: KeyName = "Inicio"
        Case vbKeyInsert: KeyName = "Insert"
        Case vbKeyMenu: KeyName = "Alt"
        Case vbKeyMultiply: KeyName = "* (KeyPad)"
        Case vbKeyNumlock: KeyName = "NumLock"
        Case vbKeyNumpad0: KeyName = "0 (KeyPad)"
        Case vbKeyNumpad1: KeyName = "1 (KeyPad)"
        Case vbKeyNumpad2: KeyName = "2 (KeyPad)"
        Case vbKeyNumpad3: KeyName = "3 (KeyPad)"
        Case vbKeyNumpad4: KeyName = "4 (KeyPad)"
        Case vbKeyNumpad5: KeyName = "5 (KeyPad)"
        Case vbKeyNumpad6: KeyName = "6 (KeyPad)"
        Case vbKeyNumpad7: KeyName = "7 (KeyPad)"
        Case vbKeyNumpad8: KeyName = "8 (KeyPad)"
        Case vbKeyNumpad9: KeyName = "9 (KeyPad)"
        Case vbKeyPageDown: KeyName = "Av Pag"
        Case vbKeyPageUp: KeyName = "Re Pag"
        Case vbKeyPause: KeyName = "Pausa"
        Case vbKeyPrint: KeyName = "ImprPant"
        Case vbKeyReturn: KeyName = "Enter"
        Case vbKeySelect: KeyName = "Select"
        Case vbKeyShift: KeyName = "Shift"
        Case vbKeySnapshot: KeyName = "Snapshot"
        Case vbKeySpace: KeyName = "Espacio"
        Case vbKeySubtract: KeyName = "- (KeyPad)"
        Case vbKeyTab: KeyName = "Tab"
        Case 92: KeyName = "Windows"
        Case 93: KeyName = "Lista"
        Case 145: KeyName = "Bloq Despl"
        Case 186: KeyName = "Comilla cierra(´)"
        Case 187: KeyName = "Asterisco (*)"
        Case 188: KeyName = "Coma (,)"
        Case 189: KeyName = "Guión (-)"
        Case 190: KeyName = "Punto (.)"
        Case 191: KeyName = "Llave cierra (})"
        Case 192: KeyName = "Ñ"
        Case 219: KeyName = "Comilla ("
        Case 220: KeyName = "Barra vertical (|)"
        Case 221: KeyName = "Signo pregunta (¿)"
        Case 222: KeyName = "Llave abre ({)"
        Case 223: KeyName = "Cualquiera"
        Case 226: KeyName = "Menor (<)"
        Case Else: KeyName = Chr(KeyCode)
    End Select

End Function
Public Sub MostrarIgnorados()
    Dim i As Integer

    For i = 1 To UBound(Ignorados)
        If Ignorados(i) <> "" Then Call AddtoRichTextBox(frmMain.rectxt, Ignorados(i), 65, 190, 156, 0, 0)
    Next

End Sub
Function IgnorarPJ(Name As String) As Byte
    Dim i As Integer, tIndex As Integer

    tIndex = NameIndex(Name)

    If tIndex = 0 Then
        Call AddtoRichTextBox(frmMain.rectxt, "El personaje no existe o no está en tu mapa.", 65, 190, 156, 0, 0)
        Exit Function
    End If

    If tIndex = UserCharIndex Then
        Call AddtoRichTextBox(frmMain.rectxt, "No puedes ignorarte a ti mismo.", 65, 190, 156, 0, 0)
        Exit Function
    End If

    For i = LBound(Ignorados) To UBound(Ignorados)
        If UCase$(Ignorados(i)) = UCase$(CharList(tIndex).Nombre) Then
            Call AddtoRichTextBox(frmMain.rectxt, "Dejaste de ignorar a " & CharList(tIndex).Nombre & ".", 65, 190, 156, 0, 0)
            Ignorados(i) = ""
            IgnorarPJ = 2
            Exit Function
        End If
    Next

    For i = LBound(Ignorados) To UBound(Ignorados)
        If Len(Ignorados(i)) = 0 Then
            Call AddtoRichTextBox(frmMain.rectxt, "Empezaste a ignorar a " & CharList(tIndex).Nombre & ".", 65, 190, 156, 0, 0)
            Ignorados(i) = CharList(tIndex).Nombre
            IgnorarPJ = 1
            Exit Function
        End If
    Next

    Call AddtoRichTextBox(frmMain.rectxt, "No puedes ignorar a más personas.", 65, 190, 156, 0, 0)

End Function
Function NameIndex(Name As String) As Integer
    Dim i As Integer

    For i = 1 To LastChar
        If UCase$(Left$(CharList(i).Nombre, Len(Name))) = UCase$(Name) Then
            NameIndex = i
            Exit Function
        End If
    Next

End Function
Sub SendData(sdData As String)
    Dim retcode
    Dim AuxCmd As String

    If Pausa Then Exit Sub

    If CONGELADO And UCase$(sdData) <> "/DESCONGELAR" Then Exit Sub
    If Not frmMain.Socket1.Connected Then Exit Sub

    AuxCmd = UCase$(Left$(sdData, 5))

    If AuxCmd = "/PING" Then TimerPing(1) = GetTickCount()

    Debug.Print ">> " & sdData

    If Left$(sdData, 1) = "/" And Len(sdData) = 2 Then Exit Sub


    sdData = ReplaceData(sdData)
    'sdData = Mod_DesEncript.Encriptar(sdData)

    If sdData = "NOPUDO" Then Exit Sub

    bO = bO + 1
    If bO > 10000 Then bO = 100

    If Len(sdData) = 0 Then Exit Sub

    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then Exit Sub
    If AuxCmd = "GM" And Len(sdData) > 2200 Then
        NoMandoElMsg = 1
        Exit Sub
    End If

    If Len(sdData) > 300 And AuxCmd <> "DEMSG" And AuxCmd <> "GM" Then Exit Sub

    NoMandoElMsg = 0

    bK = 0

    'sdData = DPackSeg(sdData)

    sdData = sdData & "~" & bK & ENDC

    retcode = frmMain.Socket1.Write(sdData, Len(sdData))

End Sub
Sub ValidacionCliente()

' by Germax

    If ValidacionDeCliente Then
        Call SendData("SIVKWZFLKWOIEIQ")
    End If

End Sub
Sub Login()
    Dim valcode As Integer
    valcode = Val(PersonalPass)

    If EstadoLogin = Normal Then

        Call SendData("OLOGIO" & UserName & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & "," & GetSerialNumber("c:\") & "," & nombrecuent)

    ElseIf EstadoLogin = CrearNuevoPj Then
        SendData ("GMMAOP" & frmCrearPersonaje.txtNombre.Text & "," _
                  & 0 & "," & 0 & "," _
                  & App.Major & "." & App.Minor & "." & App.Revision & _
                  "," & UserRaza & "," & UserSexo & "," & _
                  UserAtributos(1) & "," & UserAtributos(2) & "," & UserAtributos(3) _
                  & "," & UserAtributos(4) & "," & UserAtributos(5) _
                  & "," & UserSkills(1) & "," & UserSkills(2) _
                  & "," & UserSkills(3) & "," & UserSkills(4) _
                  & "," & UserSkills(5) & "," & UserSkills(6) _
                  & "," & UserSkills(7) & "," & UserSkills(8) _
                  & "," & UserSkills(9) & "," & UserSkills(10) _
                  & "," & UserSkills(11) & "," & UserSkills(12) _
                  & "," & UserSkills(13) & "," & UserSkills(14) _
                  & "," & UserSkills(15) & "," & UserSkills(16) _
                  & "," & UserSkills(17) & "," & UserSkills(18) _
                  & "," & UserSkills(19) & "," & UserSkills(20) _
                  & "," & UserSkills(21) & "," & UserSkills(22) & "," & _
                  UserHogar & "," & valcode & "," & GetSerialNumber("c:\") & "," & nombrecuent)


    ElseIf EstadoLogin = LoginAccount Then
        SendData ("ALOGIN" & nombrecuent & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode)

    ElseIf EstadoLogin = CrearAccount Then
        SendData ("NACCNT" & frmCrearAccount.Nombre.Text & "," & frmCrearAccount.Pass.Text & "," & frmCrearAccount.Mail.Text _
                  & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode)
    End If

End Sub


