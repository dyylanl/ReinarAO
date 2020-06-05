Attribute VB_Name = "Mod_General"
'FénixAO en DX8 by ·Parra, Thusing y DarkTester

Option Explicit

Public Config_Particles As Boolean
'Public Audio As New clsAudio
Public Sound As clsSoundEngine

Public pw1(10) As String

'Sound constants
Public Const MUS_Inicio As String = "1"
Public Const MUS_CrearPersonaje As String = "7"
Public Const MUS_VolverInicio As String = "1"
Public Const SND_CLICK  As String = "119"
Public Const SND_MONTANDO As String = "23"
Public Const SND_PASOS1 As String = "23"
Public Const SND_PASOS2 As String = "24"
Public Const SND_NAVEGANDO As String = "50"
Public Const SND_OVER As String = "1" ' "click2.Wav"
Public Const SND_DICE As String = "1" ' "cupdice.Wav"
Public Const SND_FUEGO As Integer = 79
Public Const GRH_FOGATA As Integer = 1521



Public atacar As Integer
Public IsClan As Byte

Public Desplazar As Boolean
Public vigilar As Boolean


Public rG(1 To 9, 1 To 3) As Byte

Public bO As Integer
Public bK As Long
Public bRK As Long
Public banners As String

Public bInvMod As Boolean

Public bFogata As Boolean

Public bLluvia() As Byte

Type Recompensa
    Name As String
    Descripcion As String
End Type

Public Recompensas(1 To 60, 1 To 3, 1 To 2) As Recompensa

Global i&, j&, k&
Global msg$, MsgErr$, NumErr&
Global Cont%, Opc%

' Fin de ModPrincipal.bas
Public Sub EstablecerRecompensas()

    Recompensas(MINERO, 1, 1).Name = "Fortaleza del Trabajador"
    Recompensas(MINERO, 1, 1).Descripcion = "Aumenta la vida en 120 puntos."

    Recompensas(MINERO, 1, 2).Name = "Suerte de Novato"
    Recompensas(MINERO, 1, 2).Descripcion = "Al morir hay 20% de probabilidad de no perder los minerales."

    Recompensas(MINERO, 2, 1).Name = "Destrucción Mágica"
    Recompensas(MINERO, 2, 1).Descripcion = "Inmunidad al paralisis lanzado por otros usuarios."

    Recompensas(MINERO, 2, 2).Name = "Pica Fuerte"
    Recompensas(MINERO, 2, 2).Descripcion = "Permite minar 20% más cantidad de hierro y la plata."

    Recompensas(MINERO, 3, 1).Name = "Gremio del Trabajador"
    Recompensas(MINERO, 3, 1).Descripcion = "Permite minar 20% más cantidad de oro."

    Recompensas(MINERO, 3, 2).Name = "Pico de la Suerte"
    Recompensas(MINERO, 3, 2).Descripcion = "Al morir hay 30% de probabilidad de que no perder los minerales (acumulativo con Suerte de Novato.)"


    Recompensas(HERRERO, 1, 1).Name = "Yunque Rojizo"
    Recompensas(HERRERO, 1, 1).Descripcion = "25% de probabilidad de gastar la mitad de lingotes en la creación de objetos (Solo aplicable a armas y armaduras)."

    Recompensas(HERRERO, 1, 2).Name = "Maestro de la Forja"
    Recompensas(HERRERO, 1, 2).Descripcion = "Reduce los costos de cascos y escudos a un 50%."

    Recompensas(HERRERO, 2, 1).Name = "Experto en Filos"
    Recompensas(HERRERO, 2, 1).Descripcion = "Permite crear las mejores armas (Espada Neithan, Espada Neithan + 1, Espada de Plata + 1 y Daga Infernal)."

    Recompensas(HERRERO, 2, 2).Name = "Experto en Corazas"
    Recompensas(HERRERO, 2, 2).Descripcion = "Permite crear las mejores armaduras (Armaduras de las Tinieblas, Armadura Legendaria y Armaduras del Dragón)."

    Recompensas(HERRERO, 3, 1).Name = "Fundir Metal"
    Recompensas(HERRERO, 3, 1).Descripcion = "Reduce a un 50% la cantidad de lingotes utilizados en fabricación de Armas y Armaduras (acumulable con Yunque Rojizo)."

    Recompensas(HERRERO, 3, 2).Name = "Trabajo en Serie"
    Recompensas(HERRERO, 3, 2).Descripcion = "10% de probabilidad de crear el doble de objetos de los asignados con la misma cantidad de lingotes."


    Recompensas(TALADOR, 1, 1).Name = "Músculos Fornidos"
    Recompensas(TALADOR, 1, 1).Descripcion = "Permite talar 20% más cantidad de madera."

    Recompensas(TALADOR, 1, 2).Name = "Tiempos de Calma"
    Recompensas(TALADOR, 1, 2).Descripcion = "Evita tener hambre y sed."


    Recompensas(CARPINTERO, 1, 1).Name = "Experto en Arcos"
    Recompensas(CARPINTERO, 1, 1).Descripcion = "Permite la creación de los mejores arcos (Élfico y de las Tinieblas)."

    Recompensas(CARPINTERO, 1, 2).Name = "Experto de Varas"
    Recompensas(CARPINTERO, 1, 2).Descripcion = "Permite la creación de las mejores varas (Engarzadas)."

    Recompensas(CARPINTERO, 2, 1).Name = "Fila de Leña"
    Recompensas(CARPINTERO, 2, 1).Descripcion = "Aumenta la creación de flechas a 20 por vez."

    Recompensas(CARPINTERO, 2, 2).Name = "Espíritu de Navegante"
    Recompensas(CARPINTERO, 2, 2).Descripcion = "Reduce en un 20% el coste de madera de las barcas."


    Recompensas(PESCADOR, 1, 1).Name = "Favor de los Dioses"
    Recompensas(PESCADOR, 1, 1).Descripcion = "Pescar 20% más cantidad de pescados."

    Recompensas(PESCADOR, 1, 2).Name = "Pesca en Alta Mar"
    Recompensas(PESCADOR, 1, 2).Descripcion = "Al pescar en barca hay 10% de probabilidad de obtener pescados más caros."


    Recompensas(MAGO, 1, 1).Name = "Pociones de Espíritu"
    Recompensas(MAGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

    Recompensas(MAGO, 1, 2).Name = "Pociones de Vida"
    Recompensas(MAGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(MAGO, 2, 1).Name = "Vitalidad"
    Recompensas(MAGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

    Recompensas(MAGO, 2, 2).Name = "Fortaleza Mental"
    Recompensas(MAGO, 2, 2).Descripcion = "Libera el limite de mana máximo."

    Recompensas(MAGO, 3, 1).Name = "Furia del Relámpago"
    Recompensas(MAGO, 3, 1).Descripcion = "Aumenta el daño base máximo de la Descarga Eléctrica en 10 puntos."

    Recompensas(MAGO, 3, 2).Name = "Destrucción"
    Recompensas(MAGO, 3, 2).Descripcion = "Aumenta el daño base mínimo del Apocalipsis en 10 puntos."


    Recompensas(NIGROMANTE, 1, 1).Name = "Pociones de Espíritu"
    Recompensas(NIGROMANTE, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

    Recompensas(NIGROMANTE, 1, 2).Name = "Pociones de Vida"
    Recompensas(NIGROMANTE, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(NIGROMANTE, 2, 1).Name = "Vida del Invocador"
    Recompensas(NIGROMANTE, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

    Recompensas(NIGROMANTE, 2, 2).Name = "Alma del Invocador"
    Recompensas(NIGROMANTE, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

    Recompensas(NIGROMANTE, 3, 1).Name = "Semillas de las Almas"
    Recompensas(NIGROMANTE, 3, 1).Descripcion = "Aumenta el daño base mínimo de la magia en 10 puntos."

    Recompensas(NIGROMANTE, 3, 2).Name = "Bloqueo de las Almas"
    Recompensas(NIGROMANTE, 3, 2).Descripcion = "Aumenta la evasión en un 5%."


    Recompensas(PALADIN, 1, 1).Name = "Pociones de Espíritu"
    Recompensas(PALADIN, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

    Recompensas(PALADIN, 1, 2).Name = "Pociones de Vida"
    Recompensas(PALADIN, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(PALADIN, 2, 1).Name = "Aura de Vitalidad"
    Recompensas(PALADIN, 2, 1).Descripcion = "Aumenta la vida en 5 puntos y el mana en 10 puntos."

    Recompensas(PALADIN, 2, 2).Name = "Aura de Espíritu"
    Recompensas(PALADIN, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

    Recompensas(PALADIN, 3, 1).Name = "Gracia Divina"
    Recompensas(PALADIN, 3, 1).Descripcion = "Reduce el coste de mana de Remover Paralisis a 250 puntos."

    Recompensas(PALADIN, 3, 2).Name = "Favor de los Enanos"
    Recompensas(PALADIN, 3, 2).Descripcion = "Aumenta en 5% la posibilidad de golpear al enemigo con armas cuerpo a cuerpo."


    Recompensas(CLERIGO, 1, 1).Name = "Pociones de Espíritu"
    Recompensas(CLERIGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

    Recompensas(CLERIGO, 1, 2).Name = "Pociones de Vida"
    Recompensas(CLERIGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(CLERIGO, 2, 1).Name = "Signo Vital"
    Recompensas(CLERIGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

    Recompensas(CLERIGO, 2, 2).Name = "Espíritu de Sacerdote"
    Recompensas(CLERIGO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

    Recompensas(CLERIGO, 3, 1).Name = "Sacerdote Experto"
    Recompensas(CLERIGO, 3, 1).Descripcion = "Aumenta la cura base de Curar Heridas Graves en 20 puntos."

    Recompensas(CLERIGO, 3, 2).Name = "Alzamientos de Almas"
    Recompensas(CLERIGO, 3, 2).Descripcion = "El hechizo de Resucitar cura a las personas con su mana, energía, hambre y sed llenas y cuesta 1.100 de mana."


    Recompensas(BARDO, 1, 1).Name = "Pociones de Espíritu"
    Recompensas(BARDO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

    Recompensas(BARDO, 1, 2).Name = "Pociones de Vida"
    Recompensas(BARDO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(BARDO, 2, 1).Name = "Melodía Vital"
    Recompensas(BARDO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

    Recompensas(BARDO, 2, 2).Name = "Melodía de la Meditación"
    Recompensas(BARDO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

    Recompensas(BARDO, 3, 1).Name = "Concentración"
    Recompensas(BARDO, 3, 1).Descripcion = "Aumenta la probabilidad de Apuñalar a un 20% (con 100 skill)."

    Recompensas(BARDO, 3, 2).Name = "Melodía Caótica"
    Recompensas(BARDO, 3, 2).Descripcion = "Aumenta el daño base del Apocalipsis y la Descarga Electrica en 5 puntos."


    Recompensas(DRUIDA, 1, 1).Name = "Pociones de Espíritu"
    Recompensas(DRUIDA, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

    Recompensas(DRUIDA, 1, 2).Name = "Pociones de Vida"
    Recompensas(DRUIDA, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(DRUIDA, 2, 1).Name = "Grifo de la Vida"
    Recompensas(DRUIDA, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

    Recompensas(DRUIDA, 2, 2).Name = "Poder del Alma"
    Recompensas(DRUIDA, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

    Recompensas(DRUIDA, 3, 1).Name = "Raíces de la Naturaleza"
    Recompensas(DRUIDA, 3, 1).Descripcion = "Reduce el coste de mana de Inmovilizar a 250 puntos."

    Recompensas(DRUIDA, 3, 2).Name = "Fortaleza Natural"
    Recompensas(DRUIDA, 3, 2).Descripcion = "Aumenta la vida de los elementales invocados en 75 puntos."


    Recompensas(ASESINO, 1, 1).Name = "Pociones de Espíritu"
    Recompensas(ASESINO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

    Recompensas(ASESINO, 1, 2).Name = "Pociones de Vida"
    Recompensas(ASESINO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(ASESINO, 2, 1).Name = "Sombra de Vida"
    Recompensas(ASESINO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

    Recompensas(ASESINO, 2, 2).Name = "Sombra Mágica"
    Recompensas(ASESINO, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

    Recompensas(ASESINO, 3, 1).Name = "Daga Mortal"
    Recompensas(ASESINO, 3, 1).Descripcion = "Aumenta el daño de Apuñalar a un 70% más que el golpe."

    Recompensas(ASESINO, 3, 2).Name = "Punteria mortal"
    Recompensas(ASESINO, 3, 2).Descripcion = "Las chances de apuñalar suben a 25% (Con 100 skills)."


    Recompensas(CAZADOR, 1, 1).Name = "Pociones de Espíritu"
    Recompensas(CAZADOR, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

    Recompensas(CAZADOR, 1, 2).Name = "Pociones de Vida"
    Recompensas(CAZADOR, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(CAZADOR, 2, 1).Name = "Fortaleza del Oso"
    Recompensas(CAZADOR, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

    Recompensas(CAZADOR, 2, 2).Name = "Fortaleza del Leviatán"
    Recompensas(CAZADOR, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

    Recompensas(CAZADOR, 3, 1).Name = "Precisión"
    Recompensas(CAZADOR, 3, 1).Descripcion = "Aumenta la puntería con arco en un 10%."

    Recompensas(CAZADOR, 3, 2).Name = "Tiro Preciso"
    Recompensas(CAZADOR, 3, 2).Descripcion = "Las flechas que golpeen la cabeza ignoran la defensa del casco."


    Recompensas(ARQUERO, 1, 1).Name = "Flechas Mortales"
    Recompensas(ARQUERO, 1, 1).Descripcion = "1.500 flechas que caen al morir."

    Recompensas(ARQUERO, 1, 2).Name = "Pociones de Vida"
    Recompensas(ARQUERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(ARQUERO, 2, 1).Name = "Vitalidad Élfica"
    Recompensas(ARQUERO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

    Recompensas(ARQUERO, 2, 2).Name = "Paso Élfico"
    Recompensas(ARQUERO, 2, 2).Descripcion = "Aumenta la evasión en un 5%."

    Recompensas(ARQUERO, 3, 1).Name = "Ojo del Águila"
    Recompensas(ARQUERO, 3, 1).Descripcion = "Aumenta la puntería con arco en un 5%."

    Recompensas(ARQUERO, 3, 2).Name = "Disparo Élfico"
    Recompensas(ARQUERO, 3, 2).Descripcion = "Aumenta el daño base mínimo de las flechas en 5 puntos y el máximo en 3 puntos."


    Recompensas(GUERRERO, 1, 1).Name = "Pociones de Poder"
    Recompensas(GUERRERO, 1, 1).Descripcion = "80 pociones verdes y 100 amarillas que no caen al morir."

    Recompensas(GUERRERO, 1, 2).Name = "Pociones de Vida"
    Recompensas(GUERRERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

    Recompensas(GUERRERO, 2, 1).Name = "Vida del Mamut"
    Recompensas(GUERRERO, 2, 1).Descripcion = "Aumenta la vida en 5 puntos."

    Recompensas(GUERRERO, 2, 2).Name = "Piel de Piedra"
    Recompensas(GUERRERO, 2, 2).Descripcion = "Aumenta la defensa permanentemente en 2 puntos."

    Recompensas(GUERRERO, 3, 1).Name = "Cuerda Tensa"
    Recompensas(GUERRERO, 3, 1).Descripcion = "Aumenta la puntería con arco en un 10%."

    Recompensas(GUERRERO, 3, 2).Name = "Resistencia Mágica"
    Recompensas(GUERRERO, 3, 2).Descripcion = "Reduce la duración de la parálisis de un minuto a 45 segundos."


    Recompensas(PIRATA, 1, 1).Name = "Marejada Vital"
    Recompensas(PIRATA, 1, 1).Descripcion = "Aumenta la vida en 20 puntos."

    Recompensas(PIRATA, 1, 2).Name = "Aventurero Arriesgado"
    Recompensas(PIRATA, 1, 2).Descripcion = "Permite entrar a los dungeons independientemente del nivel."

    Recompensas(PIRATA, 2, 1).Name = "Riqueza"
    Recompensas(PIRATA, 2, 1).Descripcion = "10% de probabilidad de no perder los objetos al morir."

    Recompensas(PIRATA, 2, 2).Name = "Escamas del Dragón"
    Recompensas(PIRATA, 2, 2).Descripcion = "Aumenta la vida en 40 puntos."

    Recompensas(PIRATA, 3, 1).Name = "Magia Tabú"
    Recompensas(PIRATA, 3, 1).Descripcion = "Inmunidad a la paralisis."

    Recompensas(PIRATA, 3, 2).Name = "Cuerda de Escape"
    Recompensas(PIRATA, 3, 2).Descripcion = "Permite salir del juego en solo dos segundos."


    Recompensas(LADRON, 1, 1).Name = "Codicia"
    Recompensas(LADRON, 1, 1).Descripcion = "Aumenta en 10% la cantidad de oro robado."

    Recompensas(LADRON, 1, 2).Name = "Manos Sigilosas"
    Recompensas(LADRON, 1, 2).Descripcion = "Aumenta en 5% la probabilidad de robar exitosamente."

    Recompensas(LADRON, 2, 1).Name = "Pies sigilosos"
    Recompensas(LADRON, 2, 1).Descripcion = "Permite moverse mientrás se está oculto."

    Recompensas(LADRON, 2, 2).Name = "Ladrón Experto"
    Recompensas(LADRON, 2, 2).Descripcion = "Permite el robo de objetos (10% de probabilidad)."

    Recompensas(LADRON, 3, 1).Name = "Robo Lejano"
    Recompensas(LADRON, 3, 1).Descripcion = "Permite robar a una distancia de hasta 4 tiles."

    Recompensas(LADRON, 3, 2).Name = "Fundido de Sombra"
    Recompensas(LADRON, 3, 2).Descripcion = "Aumenta en 10% la probabilidad de robar objetos."

End Sub

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\RECURSOS\Graficos\"
End Function

Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, Optional Red As Integer = -1, Optional Green As Integer, Optional Blue As Integer, Optional Bold As Boolean, Optional Italic As Boolean, Optional bCrLf As Boolean)

    If Opciones.ConsolaActivada = False Then Exit Sub
    With RichTextBox
        If (Len(.Text)) > 4000 Then .Text = ""
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0

        .SelBold = IIf(Bold, True, False)
        .SelItalic = IIf(Italic, True, False)

        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)

        .SelText = IIf(bCrLf, Text, Text & vbCrLf)

        RichTextBox.Refresh
    End With

End Sub
Sub AddtoTextBox(TextBox As TextBox, Text As String)

    TextBox.SelStart = Len(TextBox.Text)
    TextBox.SelLength = 0

    TextBox.SelText = Chr(13) & Chr(10) & Text

End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))

        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False
            Exit Function
        End If

    Next i

    AsciiValidos = True

End Function



Function CheckUserData(checkemail As Boolean) As Boolean

    Dim loopc As Integer
    Dim CharAscii As Integer

    If checkemail Then
        If UserEmail = "" Then
            MsgBox ("Direccion de email invalida")
            Exit Function
        End If
    End If

    If UserPassword = "" Then
        MsgBox "Ingrese la contraseña de su personaje.", vbInformation, "Password"
        Exit Function
    End If

    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MsgBox "El password es inválido." & vbCrLf & vbCrLf & "Volvé a intentarlo otra vez." & vbCrLf & "Si el password es ese, verifica el estado del BloqMayús.", vbExclamation, "Password inválido"
            Exit Function
        End If
    Next loopc

    If UserName = "" Then
        MsgBox "Tenés que ingresar el Nombre de tu Personaje para poder Jugar.", vbExclamation, "Nombre inválido"
        Exit Function
    End If

    If Len(UserName) > 20 Then
        MsgBox ("El Nombre de tu Personaje debe tener menos de 20 letras.")
        Exit Function
    End If

    For loopc = 1 To Len(UserName)

        CharAscii = Asc(mid$(UserName, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MsgBox "El Nombre del Personaje ingresado es inválido." & vbCrLf & vbCrLf & "Verifica que no halla errores en el tipeo del Nombre de tu Personaje.", vbExclamation, "Carácteres inválidos"
            Exit Function
        End If

    Next loopc


    CheckUserData = True

End Function
Sub UnloadAllForms()
    On Error Resume Next
    Dim mifrm As Form

    For Each mifrm In Forms
        Unload mifrm
    Next

End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean

    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If


    If KeyAscii < 32 Or KeyAscii = 44 Then
        LegalCharacter = False
        Exit Function
    End If

    If KeyAscii > 126 Then
        LegalCharacter = False
        Exit Function
    End If


    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        LegalCharacter = False
        Exit Function
    End If


    LegalCharacter = True

End Function
Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

    RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function
Public Function TiempoTranscurrido(ByVal Desde As Single) As Single

    TiempoTranscurrido = Timer - Desde

    If TiempoTranscurrido < -5 Then
        TiempoTranscurrido = TiempoTranscurrido + 86400
    ElseIf TiempoTranscurrido < 0 Then
        TiempoTranscurrido = 0
    End If

End Function
Public Sub ProcesaEntradaCmd(ByVal Datos As String)

'If Len(Datos) = 0 Then Exit Sub
'
'If UCase$(Left$(Datos, 3)) = "/GM" Then
'    frmMSG.Show
'    Exit Sub
'End If

    Select Case Left$(Datos, 1)
        Case "\", "/"

        Case Else
            Datos = ";" & Left$(frmMain.modo, 1) & Datos

    End Select

    Call SendData(Datos)

End Sub
Public Sub ResetIgnorados()
    Dim i As Integer

    For i = 1 To UBound(Ignorados)
        Ignorados(i) = ""
    Next

End Sub
Public Function EstaIgnorado(CharIndex As Integer) As Boolean
    Dim i As Integer

    For i = 1 To UBound(Ignorados)
        If Len(Ignorados(i)) > 0 And Ignorados(i) = CharList(CharIndex).Nombre Then
            EstaIgnorado = True
            Exit Function
        End If
    Next

End Function
Sub CheckKeys()
    On Error Resume Next

    Static KeyTimer As Integer

    If KeyTimer > 0 Then
        KeyTimer = KeyTimer - 1
        Exit Sub
    End If

    If Comerciando > 0 Then Exit Sub

    If UserMoving = 0 Then
        If Not UserEstupido Then
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveMe(NORTH)
                Exit Sub
            End If

            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call MoveMe(EAST)
                Exit Sub
            End If

            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call MoveMe(SOUTH)
                Exit Sub
            End If

            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call MoveMe(WEST)
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                 GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                 GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                 GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            If kp Then Call MoveMe(Int(RandomNumber(1, 4)))
        End If
    End If

End Sub
Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
    Dim i As Integer, LastPos As Integer, FieldNum As Integer

    For i = 1 To Len(Text)
        If mid(Text, i, 1) = Chr(SepASCII) Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr(SepASCII), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next

    If FieldNum + 1 = Pos Then ReadField = mid(Text, LastPos + 1)

End Function
Public Function PonerPuntos(Numero As Long) As String
    Dim i As Integer
    Dim Cifra As String

    Cifra = Str(Numero)
    Cifra = Right$(Cifra, Len(Cifra) - 1)
    For i = 0 To 4
        If Len(Cifra) - 3 * i >= 3 Then
            If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
                PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
            End If
        Else
            If Len(Cifra) - 3 * i > 0 Then
                PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
            End If
            Exit For
        End If
    Next

    PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)

End Function
Function FileExist(File As String, FileType As VbFileAttribute) As Boolean

    FileExist = Len(Dir$(File, FileType)) > 0

End Function
Function Traduccion(Original As String) As String
    Dim i As Integer, Char As Integer

    For i = 1 To Len(Original)
        Char = Asc(mid$(Original, i, 1)) - 232 - i ^ 2
        Do Until Char > 0
            Char = Char + 255
        Loop
        Traduccion = Traduccion & Chr$(Char)
    Next

End Function
Sub CargarMensajes()
    Dim i As Integer, NumMensajes As Integer, Leng As Byte

    Open App.Path & "\RECURSOS\Init\Mensajes.dat" For Binary As #1
    Seek #1, 1

    Get #1, , NumMensajes

    ReDim Mensajes(1 To NumMensajes) As Mensajito

    For i = 1 To NumMensajes
        Mensajes(i).code = Space$(2)
        Get #1, , Mensajes(i).code
        Mensajes(i).code = Traduccion(Mensajes(i).code)

        Get #1, , Leng
        Mensajes(i).mensaje = Space$(Leng)
        Get #1, , Mensajes(i).mensaje
        Mensajes(i).mensaje = Traduccion(Mensajes(i).mensaje)

        Get #1, , Mensajes(i).Red
        Get #1, , Mensajes(i).Green
        Get #1, , Mensajes(i).Blue
        Get #1, , Mensajes(i).Bold
        Get #1, , Mensajes(i).Italic
    Next

    Close #1

End Sub
Public Sub ActualizarInformacionComercio(Index As Integer)

    Select Case Index
        Case 0
            frmComerciar.Label1(0).Caption = PonerPuntos(OtherInventory(frmComerciar.List1(0).ListIndex + 1).Valor)
            If OtherInventory(frmComerciar.List1(0).ListIndex + 1).Amount <> 0 Then
                frmComerciar.Label1(1).Caption = PonerPuntos(CLng(OtherInventory(frmComerciar.List1(0).ListIndex + 1).Amount))
            ElseIf OtherInventory(frmComerciar.List1(0).ListIndex + 1).Name <> "Nada" Then
                frmComerciar.Label1(1).Caption = "Ilimitado"
            Else
                frmComerciar.Label1(1).Caption = 0
            End If

            frmComerciar.Label1(5).Caption = OtherInventory(frmComerciar.List1(0).ListIndex + 1).Name
            frmComerciar.List1(0).ToolTipText = OtherInventory(frmComerciar.List1(0).ListIndex + 1).Name

            Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).ObjType
                Case 2
                    frmComerciar.Label1(3).Caption = "Max Golpe:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxHit
                    frmComerciar.Label1(4).Caption = "Min Golpe:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinHit
                    frmComerciar.Label1(3).Visible = True
                    frmComerciar.Label1(4).Visible = True
                    frmComerciar.Label1(2).Caption = "Arma:"
                    frmComerciar.Label1(2).Visible = True
                Case 3
                    frmComerciar.Label1(3).Visible = True
                    frmComerciar.Label1(3).Caption = "Defensa máxima: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef
                    frmComerciar.Label1(4).Caption = "Defensa mínima: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinDef
                    frmComerciar.Label1(4).Visible = True
                    frmComerciar.Label1(2).Visible = True
                    frmComerciar.Label1(2).Caption = "Casco/Escudo/Armadura"
                    If OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef = 0 Then
                        frmComerciar.Label1(3).Visible = False
                        frmComerciar.Label1(4).Caption = "Esta ropa no tiene defensa."
                    End If
                    If OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef > 0 Then
                        frmComerciar.Label1(3).Visible = False
                        frmComerciar.Label1(4).Caption = "Defensa: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinDef & "/" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef
                    End If
                Case 11
                    frmComerciar.Label1(3).Caption = "Max Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxModificador
                    frmComerciar.Label1(4).Caption = "Min Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinModificador

                    frmComerciar.Label1(3).Visible = True
                    frmComerciar.Label1(4).Visible = True
                    frmComerciar.Label1(2).Visible = True
                    frmComerciar.Label1(2).Caption = "Min Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).TipoPocion
                    Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).TipoPocion
                        Case 1
                            frmComerciar.Label1(2).Caption = "Modifica Agilidad:"
                        Case 2
                            frmComerciar.Label1(2).Caption = "Modifica Fuerza:"
                        Case 3
                            frmComerciar.Label1(2).Caption = "Repone Vida:"
                        Case 4
                            frmComerciar.Label1(2).Caption = "Repone Mana:"
                        Case 5
                            frmComerciar.Label1(2).Caption = "- Cura Envenenamiento -"
                            frmComerciar.Label1(3).Visible = False
                            frmComerciar.Label1(4).Visible = False
                    End Select
                Case 24
                    frmComerciar.Label1(3).Visible = False
                    frmComerciar.Label1(4).Visible = False
                    frmComerciar.Label1(2).Visible = True
                    frmComerciar.Label1(2).Caption = "- Hechizo -"
                Case 31
                    frmComerciar.Label1(3).Visible = True
                    frmComerciar.Label1(4).Visible = True
                    frmComerciar.Label1(2).Visible = True
                    frmComerciar.Label1(2).Caption = "- Fragata -"
                    frmComerciar.Label1(4).Caption = "Min/Max Golpe: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinHit & "/" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxHit
                    frmComerciar.Label1(3).Caption = "Defensa:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).Def
                    frmComerciar.Label1(4).Visible = True
                Case Else
                    frmComerciar.Label1(2).Visible = False
                    frmComerciar.Label1(3).Visible = False
                    frmComerciar.Label1(4).Visible = False
            End Select

            If OtherInventory(frmComerciar.List1(0).ListIndex + 1).PuedeUsar > 0 Then
                frmComerciar.Label1(6).Caption = "No podés usarlo ("
                Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).PuedeUsar
                    Case 1
                        frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Genero)"
                    Case 2
                        frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Clase)"
                    Case 3
                        frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Facción)"
                    Case 4
                        frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Skill)"
                    Case 5
                        frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Raza)"
                End Select
            Else
                frmComerciar.Label1(6).Caption = ""
            End If

            If OtherInventory(frmComerciar.List1(0).ListIndex + 1).GrhIndex > 0 Then
                Call DrawGrhtoHdc(frmComerciar.Picture1, OtherInventory(frmComerciar.List1(0).ListIndex + 1).GrhIndex, 0, 0)
            Else
                frmComerciar.Picture1.Picture = LoadPicture()
            End If

        Case 1
            frmComerciar.Label1(0).Caption = PonerPuntos(UserInventory(frmComerciar.List1(1).ListIndex + 1).Valor)
            frmComerciar.Label1(1).Caption = PonerPuntos(UserInventory(frmComerciar.List1(1).ListIndex + 1).Amount)
            frmComerciar.Label1(5).Caption = UserInventory(frmComerciar.List1(1).ListIndex + 1).Name

            frmComerciar.List1(1).ToolTipText = UserInventory(frmComerciar.List1(1).ListIndex + 1).Name
            Select Case UserInventory(frmComerciar.List1(1).ListIndex + 1).ObjType
                Case 2
                    frmComerciar.Label1(2).Caption = "Arma:"
                    frmComerciar.Label1(3).Caption = "Max Golpe:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxHit
                    frmComerciar.Label1(4).Caption = "Min Golpe:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinHit
                    frmComerciar.Label1(3).Visible = True
                    frmComerciar.Label1(2).Visible = True
                    frmComerciar.Label1(4).Visible = True
                Case 3
                    frmComerciar.Label1(3).Visible = True
                    frmComerciar.Label1(3).Caption = "Defensa máxima: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef
                    frmComerciar.Label1(4).Caption = "Defensa mínima: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinDef
                    frmComerciar.Label1(4).Visible = True
                    frmComerciar.Label1(2).Visible = True
                    frmComerciar.Label1(2).Caption = "Casco/Escudo/Armadura"
                    If UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef = 0 Then
                        frmComerciar.Label1(3).Visible = False
                        frmComerciar.Label1(4).Caption = "Esta ropa no tiene defensa."
                    End If
                    If UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef > 0 Then
                        frmComerciar.Label1(3).Visible = False
                        frmComerciar.Label1(4).Caption = "Defensa " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinDef & "/" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef
                    End If
                Case 11
                    frmComerciar.Label1(3).Caption = "Max Efecto:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxModificador
                    frmComerciar.Label1(4).Caption = "Min Efecto:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinModificador

                    frmComerciar.Label1(3).Visible = True
                    frmComerciar.Label1(4).Visible = True
                    frmComerciar.Label1(2).Visible = True

                    Select Case UserInventory(frmComerciar.List1(1).ListIndex + 1).TipoPocion
                        Case 1
                            frmComerciar.Label1(2).Caption = "Aumenta Agilidad"
                        Case 2
                            frmComerciar.Label1(2).Caption = "Aumenta Fuerza"
                        Case 3
                            frmComerciar.Label1(2).Caption = "Repone Vida"
                        Case 4
                            frmComerciar.Label1(2).Caption = "Repone Mana"
                        Case 5
                            frmComerciar.Label1(2).Caption = "- Cura Envenenamiento -"
                            frmComerciar.Label1(3).Visible = False
                            frmComerciar.Label1(4).Visible = False
                    End Select
                Case 24
                    frmComerciar.Label1(3).Visible = False
                    frmComerciar.Label1(4).Visible = False
                    frmComerciar.Label1(2).Caption = "- Hechizo -"
                    frmComerciar.Label1(2).Visible = True
                Case 31
                    frmComerciar.Label1(3).Visible = True
                    frmComerciar.Label1(4).Visible = True
                    frmComerciar.Label1(2).Caption = "- Fragata -"
                    frmComerciar.Label1(4).Caption = "Min/Max Golpe: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinHit & "/" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxHit
                    frmComerciar.Label1(3).Caption = "Defensa:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).Def
                    frmComerciar.Label1(4).Visible = True
                    frmComerciar.Label1(2).Visible = True
                Case Else
                    frmComerciar.Label1(2).Visible = False
                    frmComerciar.Label1(3).Visible = False
                    frmComerciar.Label1(4).Visible = False
            End Select

            If UserInventory(frmComerciar.List1(1).ListIndex + 1).GrhIndex > 0 Then
                Call DrawGrhtoHdc(frmComerciar.Picture1, UserInventory(frmComerciar.List1(1).ListIndex + 1).GrhIndex, 0, 0)
            Else
                frmComerciar.Picture1.Picture = LoadPicture()
            End If

    End Select

    frmComerciar.Picture1.Refresh

End Sub


Sub Main()
    Call InitCommonControlsVB
    'frmMain.seguridad.Interval = 8000
    'frmMain.seguridad.Enabled = True
    'frmMain.detectarclick.Interval = 250    ' va y aca para desactivar la seguridad hacela mas facil y cambia los true por false si de una
    'frmMain.detectarclick.Enabled = True

    pw1(1) = "H"

    frmCargando.Show
    AddtoRichTextBox frmCargando.Status, "Cargando Opciones...", 255, 150, 50, 1, , False
    
    CargarOpciones
    
    If Opciones.RegistroLibrerias = 0 Then
        MsgBox "Aún no se ha ejecutado el 'Registrador de Librerias' ubicado en " & App.Path & "\Registrador de Librerias.exe para un correcto funcionamiento del juego. Recomendamos ejecutarlo, de lo contrario podrían surgir errores al intentar jugar.", vbCritical, "Atención antes de jugar"
        End
    End If

    'Mod_T0.COMPROBARBANPC
    'Mod_T0.COMPROBARBANPC1
    'Mod_T0.COMPROBARBANPC2
    'Mod_T0.COMPROBARBANPC3
    
    If App.PrevInstance Then
        Call MsgBox("¡Lhirius AO ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

    frmCargando.Refresh

    UserParalizado = False
    AddtoRichTextBox frmCargando.Status, "Resolución...", 255, 150, 50, , , True
    pw1(4) = "s"
    
    If MsgBox("¿Desea Reproducir el Juego en Pantalla Completa?", vbQuestion + vbYesNo, "Resolución") = vbYes Then
        SetResolucion
        frmMain.WindowState = vbMaximized
        frmConnect.WindowState = vbMaximized
        frmCuent.WindowState = vbMaximized
    End If
    
    AddtoRichTextBox frmCargando.Status, "Listo.", 255, 150, 50, 1, , False
    
    AddtoRichTextBox frmCargando.Status, "Iniciando constantes...", 255, 150, 50, 0, , True
    
    Call IniciarConstantes

    AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, , False
    pw1(3) = "z"
    AddtoRichTextBox frmCargando.Status, "Cargando Sonidos....", 255, 150, 50, , , True
    pw1(2) = "l"
    pw1(0) = "m"
    

    'Inicializamos el sonido
    Call CargarPasos
    Set Sound = New clsSoundEngine
    If Sound.Initialize_Engine(frmMain.hWnd, App.Path & "\Recursos\WAV\", App.Path & "\Recursos\MP3\", App.Path & "\Recursos\MIDI", False, Opciones.Audio, Opciones.sMusica, Opciones.FXVolume, Opciones.MusicVolume, Opciones.InvertirSonido) = False Then
        MsgBox "¡No se ha logrado iniciar el engine de DirectSound! Reinstale los últimos controladores de DirectX. No habrá soporte de audio en el juego.", vbCritical, "Advertencia"
        End
    End If
 
    If Opciones.sMusica <> CONST_DESHABILITADA Then
        'Sound.NextMusic = 1
        
        Sound.Fading = 350
        Sound.Music_Load (1)
        'Sound.Sound_Render
        Sound.MusicActual = 1
        Sound.Music_Play
    End If

    AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, , False
    pw1(5) = "J"
    

    pw1(8) = "Q"
    pw1(6) = "x"
    pw1(7) = "I"
    pw1(9) = "i"

    ENDC = Chr(1)

    'UserMap = 1

    Call CargarAnimsExtra
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarMensajes
    Call EstablecerRecompensas
    Call InitTileEngine(frmMain.renderer.hWnd, 32, 32, 13, 17)
    InitGrh estrella, 31988
    Unload frmCargando
    'YaPrendioLuces = True
    'SwitchMapNew UserMap, False
    
    frmConnect.Visible = True

    prgRun = True
    Pausa = False
    ' Empieza el bucle
    Call ShowNextFrame

    EngineRun = False
    Call UnloadAllForms
    Call DeInitTileEngine

    End

    'ManejadorErrores:
    '    End

End Sub



Sub WriteVar(File As String, Main As String, Var As String, value As String)


    writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(File As String, Main As String, Var As String) As String

    Dim sSpaces As String
    Dim szReturn As String

    szReturn = ""

    sSpaces = Space(5000)


    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

    GetVar = RTrim(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function
Public Function CheckMailString(ByRef sString As String) As Boolean
    On Error GoTo errHnd:
    Dim lPos As Long, lX As Long

    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        If Not InStr(lPos, sString, ".", vbBinaryCompare) > (lPos + 1) Then Exit Function

        For lX = 0 To Len(sString) - 1
            If Not lX = (lPos - 1) And Not CMSValidateChar_(Asc(mid$(sString, (lX + 1), 1))) Then Exit Function
        Next lX

        CheckMailString = True
    End If

errHnd:

End Function
Private Function CMSValidateChar_(ByRef iAsc As Integer) As Boolean

    CMSValidateChar_ = iAsc = 46 Or (iAsc >= 48 And iAsc <= 57) Or _
                       (iAsc >= 65 And iAsc <= 90) Or _
                       (iAsc >= 97 And iAsc <= 122) Or _
                       (iAsc = 95) Or (iAsc = 45)

End Function

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.renderer.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmMain.renderer.ScaleHeight \ 64
End Sub


Sub SetConnected()

    Unload frmConnect
    
    txtNombre = vbNullString
    txtPasswdAsteriscos = vbNullString
    txtPasswd = vbNullString
    FocoPasswd = False
    
    frmMain.Label8.Caption = PJClickeado

    frmMain.Visible = True

    Colorinicial = 1
    YaPrendioLuces = False
    ColorMuerto = 1
    'If Opciones.sMusica = 1 Then Sound.Music_Stop
    OroFalso = 0
    Vidafalsa = 0
    Manafalsa = 0
    Aguafalsa = 0
    HambreFalsa = 0
    If frmMain.Visible = True And YaLoguio = False Then
        YaLoguio = True
        Light.Light_Remove_All
    End If
End Sub
Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long

    Dim StreamFile As String
    StreamFile = App.Path & "\RECURSOS\init\" & "Particulas.ini"

    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))

    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As stream

    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).X1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).Y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).X2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).Y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = 1    'Val(General_Var_Get(StreamFile, Val(LoopC), "AlphaBlend"))
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))

        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")

        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")

        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = Val(General_Field_Read(i, GrhListing, Asc(",")))
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).R = Val(General_Field_Read(1, TempSet, Asc(",")))
            StreamData(loopc).colortint(ColorSet - 1).G = Val(General_Field_Read(2, TempSet, Asc(",")))
            StreamData(loopc).colortint(ColorSet - 1).b = Val(General_Field_Read(3, TempSet, Asc(",")))
        Next ColorSet

    Next loopc
End Sub
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long


    Dim Rgb_List(0 To 3) As Long
    Rgb_List(0) = RGB(StreamData(ParticulaInd).colortint(0).R, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).b)
    Rgb_List(1) = RGB(StreamData(ParticulaInd).colortint(1).R, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).b)
    Rgb_List(2) = RGB(StreamData(ParticulaInd).colortint(2).R, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).b)
    Rgb_List(3) = RGB(StreamData(ParticulaInd).colortint(3).R, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).b)


    General_Particle_Create = Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, Rgb_List(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
                                                    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).X1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).angle, _
                                                    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
                                                    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
                                                    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).X2, _
                                                    StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
                                                    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Sub DibujarPuntoMinimap()
    frmMain.UserPosicion.Left = UserPos.X - 13
    frmMain.UserPosicion.Top = UserPos.Y - 13
End Sub

Public Sub DibujarMinimap()
    On Error Resume Next
    frmMain.Minimap.Picture = LoadPicture(App.Path & "\RECURSOS\Graficos\Minimapa\" & UserMap & ".jpg")

End Sub

Public Function Encriptar(sTexto As String) As String
    Dim i As Integer
    Dim CodeAscii As Integer    'Almacena el codigo Ascii de la letra
    Dim sLetra As String    'Almacena una letra
    'Bucle que recorre cada letra del sTexto
    For i = 1 To Len(sTexto)
        sLetra = mid(sTexto, i, 1)    'Almacena la letra
        CodeAscii = ((Asc(sLetra) + 123) - 123)    'Obtiene el Ascii del sLetra
        If CodeAscii < 100 Then  'Si es menor que 100
            Encriptar = Encriptar & "0" & CodeAscii    'Imprime un 0 delante para que tenga 3 caracteres
        Else
            Encriptar = Encriptar & CodeAscii    'Lo deja talcual
        End If
        DoEvents    'Realiza cada evento
    Next i
End Function    'Fin de la funcion

'Public Sub TomarFoto()
'    Dim i As Integer
'    frmMain.Captura1.Area = Ventana
'    frmMain.Captura1.Captura
'    For i = 1 To 1000
'        If Not FileExist(App.Path & "\screenshot\LhiriusAO_" & i & ".bmp", vbNormal) Then Exit For
'    Next
'    Call SavePicture(frmMain.Captura1.Imagen, App.Path & "/screenshot/LhiriusAO_" & i & ".bmp")
'    Call AddtoRichTextBox(frmMain.rectxt, "Una imagen fue guardada en la carpeta de SCRENSHOT bajo el nombre de LhiriusAO_" & i & ".bmp", 255, 150, 50, False, False, False)

'End Sub
Public Function DesEncriptar(ByVal sTexto As String) As String
    On Error Resume Next   'En caso de error continua
    Dim i, T As Integer
    Dim sCodeAscii As String
    Dim lnCodeAscii As Long
    T = 1
    'Bucle que recorre el sTexto y toma de a 3 caracteres
    For i = 1 To Len(sTexto) / 3
        sCodeAscii = mid(sTexto, T, 3)    'Toma 3 caracteres y los almacena en sCodeAscii
        lnCodeAscii = ((Val(sCodeAscii) - 123) + 123)    'Tranforma sCodeAscii en numero y lo almacena en lnCodeAscii
        T = T + 3    'Aumenta en 3 la variable t
        DesEncriptar = DesEncriptar & Chr(lnCodeAscii)    'Transforma el Ascii al caracter correspondient e
        DoEvents    'Realiza eventos
    Next i
End Function    'Fin de la funcion

Public Sub CargarAuras()
    Dim TotalAuras As Long, i As Long

    TotalAuras = Val(GetVar(App.Path & "\Recursos\init\auras.dat", "INIT", "Total"))

    For i = 1 To TotalAuras
        Aura(i).Aura.GrhIndex = GetVar(App.Path & "\Recursos\init\auras.dat", "Aura" & i, "GrhIndex")
        Aura(i).R = Val(GetVar(App.Path & "\Recursos\init\auras.dat", "Aura" & i, "Rojo"))
        Aura(i).G = Val(GetVar(App.Path & "\Recursos\init\auras.dat", "Aura" & i, "Verde"))
        Aura(i).b = Val(GetVar(App.Path & "\Recursos\init\auras.dat", "Aura" & i, "Azul"))
        Aura(i).Giratoria = Val(GetVar(App.Path & "\Recursos\init\auras.dat", "Aura" & i, "Giratoria"))
        Aura(i).offset = Val(GetVar(App.Path & "\Recursos\init\auras.dat", "Aura" & i, "OffSet"))
    Next i
End Sub

Sub IniciarConstantes()
rG(1, 1) = 255
    rG(1, 2) = 255
    rG(1, 3) = 255

    rG(2, 1) = 0
    rG(2, 2) = 128
    rG(2, 3) = 255

    rG(3, 1) = 255
    rG(3, 2) = 0
    rG(3, 3) = 0

    rG(4, 1) = 255
    rG(4, 2) = 255
    rG(4, 3) = 0

    rG(5, 1) = 130
    rG(5, 2) = 130
    rG(5, 3) = 130

    rG(6, 1) = 210    'Consilio de Arghal.
    rG(6, 2) = 50
    rG(6, 3) = 0

    rG(7, 1) = 0    'Consejo de Banderbill.
    rG(7, 2) = 215
    rG(7, 3) = 215

    rG(8, 1) = 7
    rG(8, 2) = 155
    rG(8, 3) = 0    'Semidioses, un verde oscurito

    rG(9, 1) = 0
    rG(9, 2) = 255
    rG(9, 3) = 0    'Noble, Verde llamativo

    ReDim Ciudades(1 To NUMCIUDADES) As String
    Ciudades(1) = "Ullathorpe"
    Ciudades(2) = "Nix"
    Ciudades(3) = "Banderbill"

    ReDim CityDesc(1 To NUMCIUDADES) As String
    CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
    CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
    CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

    ReDim ListaRazas(1 To NUMRAZAS) As String
    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ReDim ListaClases(1 To NUMCLASES) As String
    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Ladron"
    ListaClases(6) = "Bardo"
    ListaClases(7) = "Druida"
    ListaClases(8) = "Bandido"
    ListaClases(9) = "Paladin"
    ListaClases(10) = "Arquero"
    ListaClases(11) = "Pescador"
    ListaClases(12) = "Herrero"
    ListaClases(13) = "Leñador"
    ListaClases(14) = "Minero"
    ListaClases(15) = "Carpintero"
    ListaClases(16) = "Pirata"

    ReDim SkillsNames(1 To NUMSKILLS) As String
    SkillsNames(1) = "Magia"
    SkillsNames(2) = "Robar"
    SkillsNames(3) = "Tacticas de combate"
    SkillsNames(4) = "Combate con armas"
    SkillsNames(5) = "Meditar"
    SkillsNames(6) = "Apuñalar"
    SkillsNames(7) = "Ocultarse"
    SkillsNames(8) = "Supervivencia"
    SkillsNames(9) = "Talar árboles"
    SkillsNames(10) = "Defensa con escudos"
    SkillsNames(11) = "Pesca"
    SkillsNames(12) = "Mineria"
    SkillsNames(13) = "Carpinteria"
    SkillsNames(14) = "Herreria"
    SkillsNames(15) = "Liderazgo"
    SkillsNames(16) = "Domar animales"
    SkillsNames(17) = "Armas de proyectiles"
    SkillsNames(18) = "Wresterling"
    SkillsNames(19) = "Navegacion"
    SkillsNames(20) = "Sastrería"
    SkillsNames(21) = "Comercio"
    SkillsNames(22) = "Resistencia Mágica"

    ReDim UserSkills(1 To NUMSKILLS) As Integer
    ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
    ReDim AtributosNames(1 To NUMATRIBUTOS) As String
    
    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub
Public Function General_Distance_Get(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
'**************************************************************
'Author: Augusto José Rando
'Co-AUthor: Lorwik
'Last Modify Date: Unknown
'
'**************************************************************
 
General_Distance_Get = Abs(X1 - X2) + Abs(Y1 - Y2)
 
End Function
Public Sub CargarPasos()
 
    ReDim Pasos(1 To NUM_PASOS) As tPaso
 
    Pasos(CONST_BOSQUE).CantPasos = 2
    ReDim Pasos(CONST_BOSQUE).Wav(1 To Pasos(CONST_BOSQUE).CantPasos) As Integer
    Pasos(CONST_BOSQUE).Wav(1) = 201
    Pasos(CONST_BOSQUE).Wav(2) = 202
 
    Pasos(CONST_NIEVE).CantPasos = 2
    ReDim Pasos(CONST_NIEVE).Wav(1 To Pasos(CONST_NIEVE).CantPasos) As Integer
    Pasos(CONST_NIEVE).Wav(1) = 199
    Pasos(CONST_NIEVE).Wav(2) = 200
 
    Pasos(CONST_CABALLO).CantPasos = 2
    ReDim Pasos(CONST_CABALLO).Wav(1 To Pasos(CONST_CABALLO).CantPasos) As Integer
    Pasos(CONST_CABALLO).Wav(1) = 23
    Pasos(CONST_CABALLO).Wav(2) = 24
 
    Pasos(CONST_DUNGEON).CantPasos = 2
    ReDim Pasos(CONST_DUNGEON).Wav(1 To Pasos(CONST_DUNGEON).CantPasos) As Integer
    Pasos(CONST_DUNGEON).Wav(1) = 23
    Pasos(CONST_DUNGEON).Wav(2) = 24
 
    Pasos(CONST_DESIERTO).CantPasos = 2
    ReDim Pasos(CONST_DESIERTO).Wav(1 To Pasos(CONST_DESIERTO).CantPasos) As Integer
    Pasos(CONST_DESIERTO).Wav(1) = 197
    Pasos(CONST_DESIERTO).Wav(2) = 198
 
    Pasos(CONST_PISO).CantPasos = 2
    ReDim Pasos(CONST_PISO).Wav(1 To Pasos(CONST_PISO).CantPasos) As Integer
    Pasos(CONST_PISO).Wav(1) = 23
    Pasos(CONST_PISO).Wav(2) = 24
 
    Pasos(CONST_PESADO).CantPasos = 3
    ReDim Pasos(CONST_PESADO).Wav(1 To Pasos(CONST_PESADO).CantPasos) As Integer
    Pasos(CONST_PESADO).Wav(1) = 220
    Pasos(CONST_PESADO).Wav(2) = 221
    Pasos(CONST_PESADO).Wav(3) = 222
 
End Sub
Private Sub CargarOpciones()

    Opciones.RegistroLibrerias = Val(GetVar(App.Path & "\RECURSOS\init\Opciones.opc", "Librerias", "Registro"))
    Opciones.CartelOcultarse = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "Ocultarse"))
    Opciones.CartelMenosCansado = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "MenosCansado"))
    Opciones.CartelVestirse = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "Vestirse"))
    Opciones.CartelNoHayNada = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "NoHayNada"))
    Opciones.CartelRecuMana = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "RecuMana"))
    Opciones.CartelSanado = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "Sanado"))
    Opciones.FPSConfig = Val(GetVar(App.Path & "\RECURSOS\init\opciones.opc", "CONFIG VIDEO", "FPS"))
    Opciones.bGraphics = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG VIDEO", "Densidad"))
    Opciones.Particulas = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG VIDEO", "Particulas"))
    Opciones.ConsolaActivada = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "Consola_Activada"))
    Opciones.InvertirSonido = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "Invertir"))
    Opciones.sMusica = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "Musica"))
    Opciones.Ambient = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "Ambient"))
    Opciones.AmbientVol = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "VolAmbient"))
    Opciones.MusicVolume = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "VolMusic"))
    Opciones.FXVolume = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "VolAudio"))
    Opciones.Audio = Val(GetVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "FXSound"))
    
End Sub
