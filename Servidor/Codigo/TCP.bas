Attribute VB_Name = "TCP"
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

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                      (ByVal hwnd As Long, ByVal lpOperation As String, _
                                       ByVal lpFile As String, ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public usercorreo As String

Public Const SOCKET_BUFFER_SIZE = 3072
Public Enpausa As Boolean

Public Const COMMAND_BUFFER_SIZE = 1000
Public entorneo As Byte

Public Const NingunArma = 2
Public Const NingunAlas As Integer = 0
Dim Response As String
Dim Start As Single, Tmr As Single
Attribute Tmr.VB_VarUserMemId = 1073741828


Public Const ToIndex = 0
Public Const ToAll = 1
Public Const ToMap = 2
Public Const ToPCArea = 3
Public Const ToNone = 4
Public Const ToAllButIndex = 5
Public Const ToMapButIndex = 6
Public Const ToGM = 7
Public Const ToNPCArea = 8
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10
Public Const ToPCAreaButIndex = 11
Public Const ToMuertos = 12
Public Const ToPCAreaVivos = 13
Public Const ToNPCAreaG = 14
Public Const ToPCAreaButIndexG = 15
Public Const ToGMArea = 16
Public Const ToPCAreaG = 17
Public Const ToAlianza = 18
Public Const ToCaos = 19
Public Const ToParty = 20
Public Const ToMoreAdmins = 21
Public Const ToActGlobal = 22
Public Const ToConse = 3

#If UsarQueSocket = 0 Then
    Public Const INVALID_HANDLE = -1
    Public Const CONTROL_ERRIGNORE = 0
    Public Const CONTROL_ERRDISPLAY = 1



    Public Const SOCKET_OPEN = 1
    Public Const SOCKET_CONNECT = 2
    Public Const SOCKET_LISTEN = 3
    Public Const SOCKET_ACCEPT = 4
    Public Const SOCKET_CANCEL = 5
    Public Const SOCKET_FLUSH = 6
    Public Const SOCKET_CLOSE = 7
    Public Const SOCKET_DISCONNECT = 7
    Public Const SOCKET_ABORT = 8


    Public Const SOCKET_NONE = 0
    Public Const SOCKET_IDLE = 1
    Public Const SOCKET_LISTENING = 2
    Public Const SOCKET_CONNECTING = 3
    Public Const SOCKET_ACCEPTING = 4
    Public Const SOCKET_RECEIVING = 5
    Public Const SOCKET_SENDING = 6
    Public Const SOCKET_CLOSING = 7


    Public Const AF_UNSPEC = 0
    Public Const AF_UNIX = 1
    Public Const AF_INET = 2


    Public Const SOCK_STREAM = 1
    Public Const SOCK_DGRAM = 2
    Public Const SOCK_RAW = 3
    Public Const SOCK_RDM = 4
    Public Const SOCK_SEQPACKET = 5


    Public Const IPPROTO_IP = 0
    Public Const IPPROTO_ICMP = 1
    Public Const IPPROTO_GGP = 2
    Public Const IPPROTO_TCP = 6
    Public Const IPPROTO_PUP = 12
    Public Const IPPROTO_UDP = 17
    Public Const IPPROTO_IDP = 22
    Public Const IPPROTO_ND = 77
    Public Const IPPROTO_RAW = 255
    Public Const IPPROTO_MAX = 256



    Public Const INADDR_ANY = "0.0.0.0"
    Public Const INADDR_LOOPBACK = "127.0.0.1"
    Public Const INADDR_NONE = "255.055.255.255"


    Public Const SOCKET_READ = 0
    Public Const SOCKET_WRITE = 1
    Public Const SOCKET_READWRITE = 2


    Public Const SOCKET_ERRIGNORE = 0
    Public Const SOCKET_ERRDISPLAY = 1


    Public Const WSABASEERR = 24000
    Public Const WSAEINTR = 24004
    Public Const WSAEBADF = 24009
    Public Const WSAEACCES = 24013
    Public Const WSAEFAULT = 24014
    Public Const WSAEINVAL = 24022
    Public Const WSAEMFILE = 24024
    Public Const WSAEWOULDBLOCK = 24035
    Public Const WSAEINPROGRESS = 24036
    Public Const WSAEALREADY = 24037
    Public Const WSAENOTSOCK = 24038
    Public Const WSAEDESTADDRREQ = 24039
    Public Const WSAEMSGSIZE = 24040
    Public Const WSAEPROTOTYPE = 24041
    Public Const WSAENOPROTOOPT = 24042
    Public Const WSAEPROTONOSUPPORT = 24043
    Public Const WSAESOCKTNOSUPPORT = 24044
    Public Const WSAEOPNOTSUPP = 24045
    Public Const WSAEPFNOSUPPORT = 24046
    Public Const WSAEAFNOSUPPORT = 24047
    Public Const WSAEADDRINUSE = 24048
    Public Const WSAEADDRNOTAVAIL = 24049
    Public Const WSAENETDOWN = 24050
    Public Const WSAENETUNREACH = 24051
    Public Const WSAENETRESET = 24052
    Public Const WSAECONNABORTED = 24053
    Public Const WSAECONNRESET = 24054
    Public Const WSAENOBUFS = 24055
    Public Const WSAEISCONN = 24056
    Public Const WSAENOTCONN = 24057
    Public Const WSAESHUTDOWN = 24058
    Public Const WSAETOOMANYREFS = 24059
    Public Const WSAETIMEDOUT = 24060
    Public Const WSAECONNREFUSED = 24061
    Public Const WSAELOOP = 24062
    Public Const WSAENAMETOOLONG = 24063
    Public Const WSAEHOSTDOWN = 24064
    Public Const WSAEHOSTUNREACH = 24065
    Public Const WSAENOTEMPTY = 24066
    Public Const WSAEPROCLIM = 24067
    Public Const WSAEUSERS = 24068
    Public Const WSAEDQUOT = 24069
    Public Const WSAESTALE = 24070
    Public Const WSAEREMOTE = 24071
    Public Const WSASYSNOTREADY = 24091
    Public Const WSAVERNOTSUPPORTED = 24092
    Public Const WSANOTINITIALISED = 24093
    Public Const WSAHOST_NOT_FOUND = 25001
    Public Const WSATRY_AGAIN = 25002
    Public Const WSANO_RECOVERY = 25003
    Public Const WSANO_DATA = 25004
    Public Const WSANO_ADDRESS = 2500
#End If

Public Data(1 To 3, 1 To 2, 1 To 2, 1 To 2) As Double
Attribute Data.VB_VarUserMemId = 1073741830
Public Onlines(1 To 3) As Long
Attribute Onlines.VB_VarUserMemId = 1073741831

Public Const Minuto = 1
Public Const Hora = 2
Public Const Dia = 3

Public Const Actual = 1
Public Const Last = 2

Public Const Enviada = 1
Public Const Recibida = 2

Public Const Mensages = 1
Public Const Letras = 2
Sub DarCuerpoYCabeza(UserBody As Integer, UserHead As Integer, Raza As Byte, Gen As Byte)

    Select Case Gen
        Case HOMBRE
            Select Case Raza

                Case HUMANO
                    UserHead = CInt(RandomNumber(1, 24))
                    If UserHead > 24 Then UserHead = 24
                    UserBody = 1
                Case ELFO
                    UserHead = CInt(RandomNumber(1, 7)) + 100
                    If UserHead > 107 Then UserHead = 107
                    UserBody = 2
                Case ELFO_OSCURO
                    UserHead = CInt(RandomNumber(1, 4)) + 200
                    If UserHead > 204 Then UserHead = 204
                    UserBody = 3
                Case ENANO
                    UserHead = RandomNumber(1, 4) + 300
                    If UserHead > 304 Then UserHead = 304
                    UserBody = 52
                Case GNOMO
                    UserHead = RandomNumber(1, 3) + 400
                    If UserHead > 403 Then UserHead = 403
                    UserBody = 52
                Case Else
                    UserHead = 1
                    UserBody = 1

            End Select
        Case MUJER
            Select Case Raza
                Case HUMANO
                    UserHead = CInt(RandomNumber(1, 4)) + 69
                    If UserHead > 73 Then UserHead = 73
                    UserBody = 1
                Case ELFO
                    UserHead = CInt(RandomNumber(1, 5)) + 169
                    If UserHead > 174 Then UserHead = 174
                    UserBody = 2
                Case ELFO_OSCURO
                    UserHead = CInt(RandomNumber(1, 5)) + 269
                    If UserHead > 274 Then UserHead = 274
                    UserBody = 3
                Case GNOMO
                    UserHead = RandomNumber(1, 4) + 469
                    If UserHead > 473 Then UserHead = 473
                    UserBody = 52
                Case ENANO
                    UserHead = RandomNumber(1, 3) + 369
                    If UserHead > 372 Then UserHead = 372
                    UserBody = 52
                Case Else
                    UserHead = 70
                    UserBody = 1
            End Select
    End Select


End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False
            Exit Function
        End If
    Next

    AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))

        If (car < 48 Or car > 57) Then
            Numeric = False
            Exit Function
        End If

    Next

    Numeric = True

End Function
Function NombrePermitido(ByVal Nombre As String) As Boolean
    Dim i As Integer

    For i = 1 To UBound(ForbidenNames)
        If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
        End If
    Next

    NombrePermitido = True

End Function

Function ValidateAtrib(UserIndex As Integer) As Boolean
    Dim LoopC As Integer

    For LoopC = 1 To NUMATRIBUTOS
        If UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) > 23 Or UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) < 1 Then Exit Function
    Next

    ValidateAtrib = True

End Function

Function ValidateAtrib2(UserIndex As Integer) As Boolean
    Dim LoopC As Integer

    For LoopC = 1 To NUMATRIBUTOS
        If UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) > 18 Or UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) < 1 Then
            ValidateAtrib2 = False
            Exit Function
        End If
    Next

    ValidateAtrib2 = True

End Function
Function ValidateSkills(UserIndex As Integer) As Boolean
    Dim LoopC As Integer

    For LoopC = 1 To NUMSKILLS
        If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    Next

    ValidateSkills = True

End Function
Public Function IsYourChr(ByVal Account As String, ByVal PJ As String)

    Dim i As Integer
    Dim NumPjs As Integer
    Dim ChrToView As String



    NumPjs = GetVar(App.path & "\Accounts\" & UCase$(Account) & ".act", "PJS", "NumPjs")

    IsYourChr = False

    For i = 1 To NumPjs
        ChrToView = GetVar(App.path & "\Accounts\" & UCase$(Account) & ".act", "PJS", "PJ" & i)
        If ChrToView = PJ Then IsYourChr = True
    Next i

End Function

Sub ConnectAccount(ByVal UserIndex As Integer, name As String, Password As String)

    Dim i As Integer
    Dim Pjjj As String
    Dim NumPjs As Integer
    Dim ArchivodeUser As String
    Dim pos() As String
    Dim Oro() As Long
    Dim Nivel() As String
    Dim PuntosdeCanje() As Integer
    Dim OroBanco() As Byte
    Dim cosa As Integer


    If Password <> GetVar(App.path & "\Accounts\" & UCase$(name) & ".act", name, "password") Then
        Call SendData(ToIndex, UserIndex, 0, "ERRPassword incorrecto.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    UserList(UserIndex).Accounted = UCase$(name)
    UserList(UserIndex).AccountedPass = Password

    NumPjs = GetVar(App.path & "\Accounts\" & UCase$(name) & ".act", "PJS", "NumPjs")

    If TienePjs(name) = True Then
        Call SendData(ToIndex, UserIndex, 0, "INIAC" & UCase$(name) & "," & NumPjs + 1)
    Else
        Call SendData(ToIndex, UserIndex, 0, "INIAC0")
    End If
    ArchivodeUser = App.path & "\charfile\"
    For i = 1 To NumPjs
        Pjjj = GetVar(App.path & "\Accounts\" & UCase$(name) & ".act", "PJS", "PJ" & i)
        If Pjjj = "" Then Exit Sub
        Call LoadUserAccount(UCase$(Pjjj) & ".chr")
        Call SendData(ToIndex, UserIndex, 0, "ADDPJ" & Pjjj & "," & i & "," & PJEnCuenta)
    Next i
End Sub
Sub ChrToAccount(ByVal Accounted As String, tName As String)

    Dim NumPjs As Integer
    Dim N As Integer

    NumPjs = GetVar(App.path & "\Accounts\" & UCase$(Accounted) & ".act", "PJS", "NumPjs")

    If NumPjs = 1 And GetVar(App.path & "\Accounts\" & Accounted & ".act", "PJS", "PJ" & NumPjs) = "" Then
        Call WriteVar(App.path & "\Accounts\" & UCase$(Accounted) & ".act", "PJS", "NumPjs", val(NumPjs))
        Call WriteVar(App.path & "\Accounts\" & UCase$(Accounted) & ".act", "PJS", "PJ" & NumPjs, tName)
        Exit Sub
    End If

    NumPjs = NumPjs + 1

    Call WriteVar(App.path & "\Accounts\" & UCase$(Accounted) & ".act", "PJS", "NumPjs", val(NumPjs))
    Call WriteVar(App.path & "\Accounts\" & UCase$(Accounted) & ".act", "PJS", "PJ" & NumPjs, tName)


End Sub
Sub CreateAccount(ByVal Account As String, Password As String, Mail As String, UserIndex As Integer)

    On Error GoTo errhandler

    If FileExist(App.path & "\Accounts\" & UCase$(Account) & ".act", vbNormal) = True Then
        Call SendData(ToIndex, UserIndex, 0, "ERREl nombre de la cuenta ya está siendo utilizado por otro usuario.")
        Exit Sub
    End If

    Dim N As Integer
    Dim i As Integer


    N = FreeFile()

    Open App.path & "\Accounts\" & UCase$(Account) & ".act" For Output As N
    Print #N, "[" & UCase$(Account) & "]"
    Print #N, "password=" & Password
    Print #N, "mail=" & Mail
    Print #N, "ban=0"
    Print #N, "[PJS]"
    Print #N, "NumPjs=0"
    Print #N, "PJ1="
    Print #N, "PJ2="
    Print #N, "PJ3="
    Print #N, "PJ4="
    Print #N, "PJ5="
    Print #N, "PJ6="
    Print #N, "PJ7="
    Print #N, "PJ8="
    Print #N, "[Stats]"
    Print #N, "Banco=0"
    Print #N, "[BancoInventory]"
    Print #N, "CantidadItems=0"
    Print #N, "Obj1=0-0"
    Print #N, "Obj2=0-0"
    Print #N, "Obj3=0-0"
    Print #N, "Obj4=0-0"
    Print #N, "Obj5=0-0"
    Print #N, "Obj6=0-0"
    Print #N, "Obj7=0-0"
    Print #N, "Obj8=0-0"
    Print #N, "Obj9=0-0"
    Print #N, "Obj10=0-0"
    Print #N, "Obj11=0-0"
    Print #N, "Obj12=0-0"
    Print #N, "Obj13=0-0"
    Print #N, "Obj14=0-0"
    Print #N, "Obj15=0-0"
    Print #N, "Obj16=0-0"
    Print #N, "Obj17=0-0"
    Print #N, "Obj18=0-0"
    Print #N, "Obj19=0-0"
    Print #N, "Obj20=0-0"
    Print #N, "Obj21=0-0"
    Print #N, "Obj22=0-0"
    Print #N, "Obj23=0-0"
    Print #N, "Obj24=0-0"
    Print #N, "Obj25=0-0"
    Print #N, "Obj26=0-0"
    Print #N, "Obj27=0-0"
    Print #N, "Obj28=0-0"
    Print #N, "Obj29=0-0"
    Print #N, "Obj30=0-0"
    Print #N, "Obj31=0-0"
    Print #N, "Obj32=0-0"
    Print #N, "Obj33=0-0"
    Print #N, "Obj34=0-0"
    Print #N, "Obj35=0-0"
    Print #N, "Obj36=0-0"
    Print #N, "Obj37=0-0"
    Print #N, "Obj38=0-0"
    Print #N, "Obj39=0-0"
    Print #N, "Obj40=0-0"


    Close N

    DoEvents

    Call SendData(ToIndex, UserIndex, 0, "ERRLa cuenta fué creada con éxito.")
    Call CloseSocket(UserIndex)


    Exit Sub

errhandler:

    Call LogError("NewAccount - Error = " & Err.Number & " - Descripción = " & Err.Description)

End Sub

Public Function TienePjs(ByVal Account As String) As Boolean

    Dim frstPj As String

    frstPj = GetVar(App.path & "\Accounts\" & UCase$(Account) & ".act", "PJS", "PJ0")

    If frstPj <> "" Then
        TienePjs = True
    Else
        TienePjs = False
    End If

End Function

Sub ConnectNewUser(UserIndex As Integer, name As String, Body As Integer, Head As Integer, UserRaza As String, UserSexo As String, _
                   UA1 As String, UA2 As String, UA3 As String, UA4 As String, UA5 As String, _
                   US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
                   US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
                   US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
                   US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
                   US21 As String, US22 As String, Hogar As String, _
                   Disco As String, Cuenta As String)
    Dim i As Integer


    UserList(UserIndex).Ranking.DuelosGanados = 0
    UserList(UserIndex).Ranking.DuelosParejaGanados = 0
    UserList(UserIndex).Ranking.MaxRondasDesafio = 0
    UserList(UserIndex).Ranking.TorneosGanados = 0
    UserList(UserIndex).flags.RecibioDonacion = 0
    UserList(UserIndex).Stats.PuntosDonador = 0
    UserList(UserIndex).Char.Alas = NingunAlas

    If Restringido Then
        Call SendData(ToIndex, UserIndex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
        Exit Sub
    End If

    'If Not NombrePermitido(Name) Then
    '    Call SendData(ToIndex, UserIndex, 0, "ERRLos nombres de los personajes deben pertencer a la fantasia, el nombre indicado es invalido.")
    '    Call SendData(ToIndex, UserIndex, 0, "V8V" & 2)
    '    Exit Sub
    'End If

    If Not AsciiValidos(name) Then
        Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
        Call SendData(ToIndex, UserIndex, 0, "V8V" & 2)
        Exit Sub
    End If

    Dim LoopC As Integer
    Dim totalskpts As Long


    If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = True Then
        Call SendData(ToIndex, UserIndex, 0, "ERRYa existe el personaje.")
        Exit Sub
    End If

    UserList(UserIndex).flags.Muerto = 0
    UserList(UserIndex).flags.Escondido = 0
    UserList(UserIndex).flags.Guerra = False

    UserList(UserIndex).name = name
    UserList(UserIndex).Clase = CIUDADANO
    UserList(UserIndex).Raza = UserRaza
    UserList(UserIndex).Genero = UserSexo
    UserList(UserIndex).Hogar = Hogar

    Select Case UserList(UserIndex).Raza
        Case HUMANO
            UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) + 1
            UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 1
            UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) = UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) + 2
        Case ELFO
            UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 3
            UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) = UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) + 1
            UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) + 1
            UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) = UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) + 2
        Case ELFO_OSCURO
            UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) + 1
            UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 1
            UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) = UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) - 3
            UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) + 2
        Case ENANO
            UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) + 3
            UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) - 1
            UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) = UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) + 3
            UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) - 6
            UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) = UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) - 3
        Case GNOMO
            UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(fuerza) - 5
            UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) + 4
            UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) + 3
            UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) = UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) + 1
    End Select

    If Not ValidateAtrib(UserIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "ERRAtributos invalidos.")
        Call SendData(ToIndex, UserIndex, 0, "V8V" & 2)
        Exit Sub
    End If

    UserList(UserIndex).Stats.UserSkills(1) = val(US1)
    UserList(UserIndex).Stats.UserSkills(2) = val(US2)
    UserList(UserIndex).Stats.UserSkills(3) = val(US3)
    UserList(UserIndex).Stats.UserSkills(4) = val(US4)
    UserList(UserIndex).Stats.UserSkills(5) = val(US5)
    UserList(UserIndex).Stats.UserSkills(6) = val(US6)
    UserList(UserIndex).Stats.UserSkills(7) = val(US7)
    UserList(UserIndex).Stats.UserSkills(8) = val(US8)
    UserList(UserIndex).Stats.UserSkills(9) = val(US9)
    UserList(UserIndex).Stats.UserSkills(10) = val(US10)
    UserList(UserIndex).Stats.UserSkills(11) = val(US11)
    UserList(UserIndex).Stats.UserSkills(12) = val(US12)
    UserList(UserIndex).Stats.UserSkills(13) = val(US13)
    UserList(UserIndex).Stats.UserSkills(14) = val(US14)
    UserList(UserIndex).Stats.UserSkills(15) = val(US15)
    UserList(UserIndex).Stats.UserSkills(16) = val(US16)
    UserList(UserIndex).Stats.UserSkills(17) = val(US17)
    UserList(UserIndex).Stats.UserSkills(18) = val(US18)
    UserList(UserIndex).Stats.UserSkills(19) = val(US19)
    UserList(UserIndex).Stats.UserSkills(20) = val(US20)
    UserList(UserIndex).Stats.UserSkills(21) = val(US21)
    UserList(UserIndex).Stats.UserSkills(22) = val(US22)

    totalskpts = 0


    For LoopC = 1 To NUMSKILLS
        totalskpts = totalskpts + Abs(UserList(UserIndex).Stats.UserSkills(LoopC))
    Next

    If totalskpts > 10 Then
        Call LogHackAttemp(UserList(UserIndex).name & " intento hackear los skills.")

        Call CloseSocket(UserIndex)
        Exit Sub
    End If


    UserList(UserIndex).Password = UserList(UserIndex).AccountedPass
    UserList(UserIndex).Char.Account = UCase$(Cuenta)


    UserList(UserIndex).Char.Heading = SOUTH

    Call DarCuerpoYCabeza(UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Raza, UserList(UserIndex).Genero)
    UserList(UserIndex).OrigChar = UserList(UserIndex).Char

    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco

    UserList(UserIndex).Stats.MET = 1
    Dim MiInt
    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) \ 3)

    UserList(UserIndex).Stats.MaxHP = 15 + MiInt
    UserList(UserIndex).Stats.MinHP = 15 + MiInt

    UserList(UserIndex).Stats.FIT = 1


    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) \ 6)
    If MiInt = 1 Then MiInt = 2

    UserList(UserIndex).Stats.MaxSta = 20 * MiInt
    UserList(UserIndex).Stats.MinSta = 20 * MiInt

    UserList(UserIndex).Stats.MaxAGU = 100
    UserList(UserIndex).Stats.MinAGU = 100

    UserList(UserIndex).Stats.MaxHam = 100
    UserList(UserIndex).Stats.MinHam = 100




    UserList(UserIndex).Stats.MaxMAN = 0
    UserList(UserIndex).Stats.MinMAN = 0


    UserList(UserIndex).Stats.MaxHit = 2
    UserList(UserIndex).Stats.MinHit = 1

    UserList(UserIndex).Stats.GLD = 0




    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = ELUs(1)
    UserList(UserIndex).Stats.ELV = 1

    UserList(UserIndex).Invent.NroItems = 6

    UserList(UserIndex).Invent.Object(1).OBJIndex = ManzanaNewbie
    UserList(UserIndex).Invent.Object(1).Amount = 100

    UserList(UserIndex).Invent.Object(2).OBJIndex = AguaNewbie
    UserList(UserIndex).Invent.Object(2).Amount = 100

    UserList(UserIndex).Invent.Object(3).OBJIndex = DagaNewbie
    UserList(UserIndex).Invent.Object(3).Amount = 1
    UserList(UserIndex).Invent.Object(3).Equipped = 1

    Select Case UserList(UserIndex).Raza
        Case HUMANO
            UserList(UserIndex).Invent.Object(4).OBJIndex = RopaNewbieHumano
        Case ELFO
            UserList(UserIndex).Invent.Object(4).OBJIndex = RopaNewbieElfo
        Case ELFO_OSCURO
            UserList(UserIndex).Invent.Object(4).OBJIndex = RopaNewbieElfoOscuro
        Case Else
            UserList(UserIndex).Invent.Object(4).OBJIndex = RopaNewbieEnano
    End Select

    UserList(UserIndex).Invent.Object(4).Amount = 1
    UserList(UserIndex).Invent.Object(4).Equipped = 1

    UserList(UserIndex).Invent.Object(5).OBJIndex = PocionRojaNewbie
    UserList(UserIndex).Invent.Object(5).Amount = 1000

    UserList(UserIndex).Invent.Object(6).OBJIndex = 41
    UserList(UserIndex).Invent.Object(6).Amount = 1

    UserList(UserIndex).Invent.ArmourEqpSlot = 4
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(4).OBJIndex

    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(3).OBJIndex
    UserList(UserIndex).Invent.WeaponEqpSlot = 3

    UserList(UserIndex).Stats.Banco = val(GetVar(App.path & "\Accounts\" & UCase$(UserList(UserIndex).Accounted) & ".act", "STATS", "Banco"))


    Dim loopd As Integer, ln2 As String
    UserList(UserIndex).BancoInvent.NroItems = GetVar(App.path & "\Accounts\" & UCase$(UserList(UserIndex).Accounted) & ".act", "BancoInventory", "CantidadItems")
    'Lista de objetos del banco
    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        ln2 = GetVar(App.path & "\Accounts\" & UCase$(UserList(UserIndex).Accounted) & ".act", "BancoInventory", "Obj" & loopd)
        UserList(UserIndex).BancoInvent.Object(loopd).OBJIndex = val(ReadField(1, ln2, 45))
        UserList(UserIndex).BancoInvent.Object(loopd).Amount = val(ReadField(2, ln2, 45))
    Next loopd



    Call SaveUser(UserIndex, CharPath & UCase$(name) & ".chr")
    Call ChrToAccount(Cuenta, name)
    Call ConnectUser(UserIndex, name, Disco, UserList(UserIndex).AccountedPass, Cuenta)

End Sub
Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)
    On Error GoTo errhandler
    Dim LoopC As Integer

    If UserList(UserIndex).flags.Stopped = True Then Exit Sub

    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Alas)

    If UserList(UserIndex).flags.DueLeanDo = True Then
        Call SendData(ToAll, UserIndex, 0, "||El Usuario " & UserList(UserIndex).name & " ha Desconectado del Duelo" & FONTTYPE_WARNING)
        Call WarpUserChar(UserIndex, 1, 50, 50, True)
    End If

    If UserList(UserIndex).flags.Death = True Then
        Call death_muere(UserIndex)
    End If

    If UserList(UserIndex).flags.automatico = True Then
        Call Rondas_UsuarioDesconecta(UserIndex)
    End If

    Call aDos.RestarConexion(UserList(UserIndex).ip)


    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        If UserList(UserIndex).flags.Privilegios = 0 Then NumNoGMs = NumNoGMs - 1
        Call CloseUser(UserIndex)
    End If
    frmMain.CantUsuarios.Caption = NumUsers + NumBots
    Call SendData(ToAll, 0, 0, "NON" & NumUsers + NumBots)
    'Dim NpcIndex As Integer
    'If Npclist(NpcIndex).flags.NPCActive Then
    '    If NumUsers > 0 Then NumBots = NumBots - 1
    'End If

    'ANTI TIRADA DE LOGIN
    Dim Hay_Socket As Long
    Dim i As Long

    For i = 1 To MAX_CONEX
        If Anti_DDOS(i).ip = UserList(UserIndex).ip Then
            Anti_DDOS(i).Desconectadas = Anti_DDOS(i).Desconectadas + 1
            If Anti_DDOS(i).Desconectadas >= 12 Then    ' 12 conex maximas
                Exit Sub
            End If
        End If
    Next i
    'ANTI TIRADA DE LOGIN

    Call ControlarPortalLum(UserIndex)    'matute
    UserList(UserIndex).flags.TiroPortalL = 0
    UserList(UserIndex).Counters.TimeTeleport = 0
    UserList(UserIndex).Counters.CreoTeleport = False
    If UserList(UserIndex).ConnID <> -1 Then Call ApiCloseSocket(UserList(UserIndex).ConnID)

    UserList(UserIndex) = UserOffline


    Exit Sub

errhandler:
    UserList(UserIndex) = UserOffline
    Call LogError("Error en CloseSocket " & Err.Description)

End Sub

Sub SendData(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, sndData As String)
    Dim LoopC As Integer
    Dim AUX$
    Dim dec$
    Dim nfile As Integer
    Dim Ret As Long

    'sndData = Mod_DesEncript.Encriptar(sndData)
    sndData = sndData & ENDC

    Select Case sndRoute

        Case ToIndex
            If UserList(sndIndex).ConnID > -1 Then
                Call WsApiEnviar(sndIndex, sndData)
                Exit Sub
            End If
            Exit Sub

        Case ToMap
            If sndMap <> 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    Call WsApiEnviar(MapInfo(sndMap).UserIndex(LoopC), sndData)
                Next
            End If
            Exit Sub

        Case ToPCArea


            For LoopC = 1 To MapInfo(UserList(sndIndex).pos.Map).NumUsers
                If EnPantalla(UserList(sndIndex).pos, UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).pos, 1) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC), sndData)
            Next
            Exit Sub

        Case ToNone
            Exit Sub

        Case ToConse    'matute
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID > -1 And (UserList(LoopC).flags.EsConseCaos Or UserList(LoopC).flags.EsConseReal) Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub


        Case ToAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.Privilegios Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub

        Case ToMoreAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.Privilegios >= UserList(sndIndex).flags.Privilegios Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub

        Case ToParty
            Dim MiembroIndex As Integer
            If UserList(sndIndex).PartyIndex = 0 Then Exit Sub
            For LoopC = 1 To MAXPARTYUSERS
                MiembroIndex = Party(UserList(sndIndex).PartyIndex).MiembrosIndex(LoopC)
                If MiembroIndex > 0 Then
                    If UserList(MiembroIndex).ConnID > -1 And UserList(MiembroIndex).flags.UserLogged And UserList(MiembroIndex).flags.Party > 0 Then Call WsApiEnviar(MiembroIndex, sndData)
                End If
            Next

            Exit Sub

        Case ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.UserLogged Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub

        Case ToActGlobal
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.UserLogged And UserList(LoopC).flags.ActGlobal = True Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub

        Case ToAllButIndex
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID > -1) And (LoopC <> sndIndex) And UserList(LoopC).flags.UserLogged Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub

        Case ToMapButIndex
            For LoopC = 1 To MapInfo(UserList(sndIndex).pos.Map).NumUsers
                If MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC), sndData)
            Next
            Exit Sub

        Case ToGuildMembers
            If Len(UserList(sndIndex).GuildInfo.GuildName) = 0 Then Exit Sub
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID > -1) And UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub

        Case ToGMArea
            For LoopC = 1 To MapInfo(UserList(sndIndex).pos.Map).NumUsers
                If EnPantalla(UserList(sndIndex).pos, UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).pos, 1) And UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).flags.Privilegios Then Call WsApiEnviar(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC), sndData)
            Next
            Exit Sub

        Case ToPCAreaVivos
            For LoopC = 1 To MapInfo(UserList(sndIndex).pos.Map).NumUsers
                If EnPantalla(UserList(sndIndex).pos, UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).pos, 1) Then
                    If Not UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).flags.Muerto Or UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).Clase = CLERIGO Then Call WsApiEnviar(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC), sndData)
                End If
            Next
            Exit Sub

        Case ToMuertos
            For LoopC = 1 To MapInfo(UserList(sndIndex).pos.Map).NumUsers
                If EnPantalla(UserList(sndIndex).pos, UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).pos, 1) Then
                    If UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).Clase = CLERIGO Or UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).flags.Muerto Or UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).flags.Privilegios Then Call WsApiEnviar(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC), sndData)
                End If
            Next
            Exit Sub

        Case ToPCAreaButIndex
            For LoopC = 1 To MapInfo(UserList(sndIndex).pos.Map).NumUsers
                If EnPantalla(UserList(sndIndex).pos, UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).pos, 1) And MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC), sndData)
            Next
            Exit Sub

        Case ToPCAreaButIndexG
            For LoopC = 1 To MapInfo(UserList(sndIndex).pos.Map).NumUsers
                If EnPantalla(UserList(sndIndex).pos, UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).pos, 3) And MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC), sndData)
            Next
            Exit Sub

        Case ToNPCArea
            For LoopC = 1 To MapInfo(Npclist(sndIndex).pos.Map).NumUsers
                If EnPantalla(Npclist(sndIndex).pos, UserList(MapInfo(Npclist(sndIndex).pos.Map).UserIndex(LoopC)).pos, 1) Then Call WsApiEnviar(MapInfo(Npclist(sndIndex).pos.Map).UserIndex(LoopC), sndData)
            Next
            Exit Sub

        Case ToNPCAreaG
            For LoopC = 1 To MapInfo(Npclist(sndIndex).pos.Map).NumUsers
                If EnPantalla(Npclist(sndIndex).pos, UserList(MapInfo(Npclist(sndIndex).pos.Map).UserIndex(LoopC)).pos, 3) Then Call WsApiEnviar(MapInfo(Npclist(sndIndex).pos.Map).UserIndex(LoopC), sndData)
            Next
            Exit Sub

        Case ToPCAreaG
            For LoopC = 1 To MapInfo(UserList(sndIndex).pos.Map).NumUsers
                If EnPantalla(UserList(sndIndex).pos, UserList(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC)).pos, 3) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).pos.Map).UserIndex(LoopC), sndData)
            Next
            Exit Sub

        Case ToAlianza
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID > -1 And UserList(LoopC).Faccion.Bando = Real Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub

        Case ToCaos
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID > -1 And UserList(LoopC).Faccion.Bando = Caos Then Call WsApiEnviar(LoopC, sndData)
            Next
            Exit Sub

    End Select

    Exit Sub
Error:
    Call LogError("Error en SendData: " & sndData & "-" & Err.Description & "-Ruta: " & sndRoute & "-Index:" & sndIndex & "-Mapa" & sndMap)

End Sub
Function HayPCarea(pos As WorldPos) As Boolean
    Dim i As Integer

    For i = 1 To MapInfo(pos.Map).NumUsers
        If EnPantalla(pos, UserList(MapInfo(pos.Map).UserIndex(i)).pos, 1) Then
            HayPCarea = True
            Exit Function
        End If
    Next

End Function
Function HayOBJarea(pos As WorldPos, OBJIndex As Integer) As Boolean
    Dim X As Integer, Y As Integer

    For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For X = pos.X - MinXBorder + 1 To pos.X + MinXBorder - 1
            If MapData(pos.Map, X, Y).OBJInfo.OBJIndex = OBJIndex Then
                HayOBJarea = True
                Exit Function
            End If
        Next
    Next

End Function

Sub CorregirSkills(UserIndex As Integer)
    Dim k As Integer

    For k = 1 To NUMSKILLS
        If UserList(UserIndex).Stats.UserSkills(k) > MAXSKILLPOINTS Then UserList(UserIndex).Stats.UserSkills(k) = MAXSKILLPOINTS
    Next

    For k = 1 To NUMATRIBUTOS
        If UserList(UserIndex).Stats.UserAtributos(k) > MAXATRIBUTOS Then
            Call SendData(ToIndex, UserIndex, 0, "ERREl personaje tiene atributos invalidos.")
            Exit Sub
        End If
    Next

End Sub
Function ValidateChr(UserIndex As Integer) As Boolean

    ValidateChr = (UserList(UserIndex).Char.Head <> 0 Or UserList(UserIndex).flags.Navegando = 1) And _
                  UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

End Function
Sub ConnectUser(UserIndex As Integer, name As String, Disco As String, Password As String, Cuenta As String)
    On Error GoTo Error
    Dim Privilegios As Byte
    Dim N As Integer
    Dim LoopC As Integer
    Dim o As Integer
    Dim NpcIndex As Integer

    UserList(UserIndex).Counters.Protegido = 4
    UserList(UserIndex).flags.Protegido = 2
    UserList(UserIndex).flags.ActGlobal = True
    UserList(UserIndex).HD = Disco
    UserList(UserIndex).flags.Guerra = False
    UserList(UserIndex).flags.Stopped = False
    UserList(UserIndex).Char.Account = UCase$(Cuenta)

    Dim numeromail As Integer

    If NumUsers > MaxUsers2 Then
        If Not (EsDios(name) Or EsSemiDios(name)) Then
            Call SendData(ToIndex, UserIndex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
            Exit Sub
        End If
    End If

    If NumUsers >= MaxUsers Then
        Call SendData(ToIndex, UserIndex, 0, "ERRLímite de usuarios alcanzado.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    If AllowMultiLogins = 0 Then
        If CheckForSameIP(UserIndex, UserList(UserIndex).ip) Then
            Call SendData(ToIndex, UserIndex, 0, "ERRNo es posible usar más de un personaje al mismo tiempo.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If


    If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, UserIndex, 0, "ERREl personaje no existe.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(name) & ".chr", "INIT", "Password")) Then
        Call SendData(ToIndex, UserIndex, 0, "ERRPassword incorrecto.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If


    If BANCheck(name) Then
        For LoopC = 1 To Baneos.Count
            If Baneos(LoopC).name = UCase$(name) Then
                Call SendData(ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a Lhirius AO.")
                Exit Sub
            End If
        Next
        Call SendData(ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a Lhirius AO debido a tu mal comportamiento.")
        Exit Sub
    End If


    If CheckForSameName(UserIndex, name) Then
        If NameIndex(name) = UserIndex Then Call CloseSocket(NameIndex(name))
        Call SendData(ToIndex, UserIndex, 0, "ERRPerdón, un usuario con el mismo nombre se encontraba logueado. INTENTE NUEVAMENTE")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If


    '¿Loguió mas de un personaje por cuenta? - Dylan.-
    Dim j As Long

    For j = 1 To LastUser

        If LastUser <> 0 Then
            If UCase$(UserList(UserIndex).Char.Account) = UCase$(UserList(j).Char.Account) And UserList(j).flags.UserLogged = True Then
                Call SendData(ToIndex, UserIndex, 0, "ERRNo se puede logear mas de un usuario por cuenta.")
                CloseSocket UserIndex
                CloseSocket j
                Exit Sub
            End If
        End If

    Next j

    If EsDios(name) Then
        Privilegios = 3
        Call LogGM(name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
    ElseIf EsSemiDios(name) Then
        Privilegios = 2
        Call LogGM(name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
    ElseIf EsConsejero(name) Then
        Privilegios = 1
        Call LogGM(name, "Se conecto con ip:" & UserList(UserIndex).ip, True)
    End If

    If Restringido And Privilegios = 0 Then
        If Not PuedeDenunciar(name) Then
            Call SendData(ToIndex, UserIndex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
            Exit Sub
        End If
    End If
    Dim Quest As Boolean
    Quest = PJQuest(name)

    Call LoadUser(UserIndex, CharPath & UCase$(name) & ".chr")

    UserList(UserIndex).Counters.IdleCount = Timer
    If UserList(UserIndex).Counters.TiempoPena Then UserList(UserIndex).Counters.Pena = Timer
    If UserList(UserIndex).Counters.TiempoSilenc Then UserList(UserIndex).Counters.PenaSilenc = Timer    'matute
    If UserList(UserIndex).flags.Envenenado Then UserList(UserIndex).Counters.Veneno = Timer
    UserList(UserIndex).Counters.AGUACounter = Timer
    UserList(UserIndex).Counters.COMCounter = Timer

    If Not ValidateChr(UserIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "ERRError en el personaje.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    For o = 1 To BanHDs.Count
        If BanHDs.Item(o) = UserList(UserIndex).HD Then
            Call SendData(ToIndex, UserIndex, 0, "ERRHas sido Baneado por el Disco Duro debido a tu Pesimo comportamiento, no podra loguear.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    Next

    For o = 1 To BanIps.Count
        If BanIps.Item(o) = UserList(UserIndex).ip Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    Next

    If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
    If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma

    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserHechizos(True, UserIndex, 0)

    If UserList(UserIndex).flags.Navegando = 1 Then
        UserList(UserIndex).Char.Alas = NingunAlas
        If UserList(UserIndex).flags.Muerto = 1 Then
            UserList(UserIndex).Char.Body = iFragataFantasmal
            UserList(UserIndex).Char.Head = 0
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.CascoAnim = NingunCasco

        Else
            UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
            UserList(UserIndex).Char.Head = 0
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.CascoAnim = NingunCasco
        End If
    End If

    UserList(UserIndex).flags.Privilegios = Privilegios
    UserList(UserIndex).flags.PuedeDenunciar = PuedeDenunciar(name)
    UserList(UserIndex).flags.Quest = Quest

    If UserList(UserIndex).flags.Privilegios > 1 Then
        If UCase$(name) = "DYLAN" Then
            UserList(UserIndex).flags.AdminInvisible = 1
            UserList(UserIndex).flags.Invisible = 1
        Else
            UserList(UserIndex).pos.Map = 9
            UserList(UserIndex).pos.X = 50
            UserList(UserIndex).pos.Y = 50
        End If
    End If

    If UserList(UserIndex).flags.Paralizado Then Call SendData(ToIndex, UserIndex, 0, "P9")

    If UserList(UserIndex).pos.Map = 0 Or UserList(UserIndex).pos.Map > NumMaps Then
        Select Case UserList(UserIndex).Hogar
            Case HOGAR_Hildegard
                UserList(UserIndex).pos = Hildegard
            Case HOGAR_Lonelerd
                UserList(UserIndex).pos = Lonelerd
            Case HOGAR_LINDOS
                UserList(UserIndex).pos = LINDOS
            Case HOGAR_ADELAIDE
                UserList(UserIndex).pos = ADELAIDE
            Case Else
                UserList(UserIndex).pos = Althalos
        End Select
        If UserList(UserIndex).pos.Map > NumMaps Then UserList(UserIndex).pos = Althalos
    End If

    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).UserIndex Then
        Dim TIndex As Integer
        Dim NposMap As Integer
        Dim NposX As Integer
        Dim NposY As Integer
        NposMap = UserList(UserIndex).pos.Map
        NposX = UserList(UserIndex).pos.X
        NposY = UserList(UserIndex).pos.Y + 1
        TIndex = MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).UserIndex
        Call SendData(ToIndex, TIndex, 0, "!!Un personaje se ha conectado en tu misma posición, reconectate.")
        Call WarpUserChar(UserIndex, NposMap, NposX, NposY, True)
    End If

    '    Dim nPos As WorldPos
    '    Call ClosestLegalPos(UserList(UserIndex).POS, nPos)
    '    UserList(UserIndex).POS = nPos
    'End If

    UserList(UserIndex).name = name

    If UserList(UserIndex).flags.Privilegios > 0 Then Call SendData(ToMoreAdmins, UserIndex, 0, "||" & UserList(UserIndex).name & " se conectó." & FONTTYPE_FENIX)

    Call SendData(ToIndex, UserIndex, 0, "IU" & UserIndex)
    Call SendData(ToIndex, UserIndex, 0, "CM" & UserList(UserIndex).pos.Map & "," & MapInfo(UserList(UserIndex).pos.Map).name & "," & MapInfo(UserList(UserIndex).pos.Map).TopPunto & "," & MapInfo(UserList(UserIndex).pos.Map).LeftPunto)
    Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).pos.Map).Music)

    Call SendUserStatsBox(UserIndex)
    Call EnviarHambreYsed(UserIndex)

    Call SendMOTD(UserIndex)

    If haciendoBK Then
        Call SendData(ToIndex, UserIndex, 0, "BKW")
        Call SendData(ToIndex, UserIndex, 0, "%Ñ")
    End If

    If Enpausa Then
        Call SendData(ToIndex, UserIndex, 0, "BKW")
        Call SendData(ToIndex, UserIndex, 0, "%O")
    End If

    UserList(UserIndex).flags.UserLogged = True
    Call AgregarAUsersPorMapa(UserIndex)

    If NumUsers + NumBots > recordusuarios Then
        Call SendData(ToAll, 0, 0, "2L" & NumUsers + NumBots)
        recordusuarios = NumUsers + NumBots
        Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    End If

    If UserList(UserIndex).flags.Privilegios > 0 Then UserList(UserIndex).flags.Ignorar = 1

    If UserIndex > LastUser Then LastUser = UserIndex

    NumUsers = NumUsers + 1

    If UserList(UserIndex).flags.Privilegios = 0 Then NumNoGMs = NumNoGMs + 1
    frmMain.CantUsuarios.Caption = NumUsers + NumBots
    Call SendData(ToAll, 0, 0, "NON" & NumUsers + NumBots)

    Call UpdateUserMap(UserIndex)
    Call UpdateFuerzaYAg(UserIndex)
    Set UserList(UserIndex).GuildRef = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

    UserList(UserIndex).flags.Seguro = True

    Call MakeUserChar(ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
    Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)
    If UserList(UserIndex).flags.Navegando = 1 Then Call SendData(ToIndex, UserIndex, 0, "NAVEG")

    If UserList(UserIndex).flags.AdminInvisible = 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFM" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 1)

    Call SendData(ToIndex, UserIndex, 0, "LOGEANDO" & Privilegios)
    UserList(UserIndex).Counters.Sincroniza = Timer

    If PuedeFaccion(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SUFA1")
    If PuedeSubirClase(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SUCL1")
    If PuedeRecompensa(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "SURE1")

    If UserList(UserIndex).Stats.SkillPts Then
        Call EnviarSkills(UserIndex)
        Call EnviarSubirNivel(UserIndex, UserList(UserIndex).Stats.SkillPts)
    End If

    Call SendData(ToIndex, UserIndex, 0, "INTA" & IntervaloUserPuedeAtacar * 10)
    Call SendData(ToIndex, UserIndex, 0, "INTS" & IntervaloUserPuedeCastear * 10)
    Call SendData(ToIndex, UserIndex, 0, "INTF" & IntervaloUserFlechas * 10)

    If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 And UserList(UserIndex).flags.Privilegios = 0 Then Call SendData(ToGuildMembers, UserIndex, 0, "4B" & UserList(UserIndex).name)
    If PuedeDestrabarse(UserIndex) Then Call SendData(ToIndex, UserIndex, 0, "||Estás encerrado, para destrabarte presiona la tecla Z." & FONTTYPE_INFO)


    Call SendData(ToIndex, UserIndex, 0, "||Bienvenido a la versión BETA de Lhirius AO necesitamos de tu colaboración para mejorar la calidad del servidor. Si ves algún error dentro del mismo escribe /BETA y su mensaje. Muchas gracias." & FONTTYPE_VENENO)
    Call SendData(ToIndex, UserIndex, 0, "||Nuevo comando /EDITAME para subir un nivel y 1.000.000 de oro." & FONTTYPE_FENIX)
    If UserList(UserIndex).Stats.PuntosDonador > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡Sos usuario PREMIUM! Escribe /CIUDADPREMIUM para obtener todos los beneficios de ser Premium. ¡¡Muchas Gracias!! ~255~255~0~1~0")
    End If

    If UserList(UserIndex).flags.SeguroCVC = True Then
        Call SendData(ToIndex, UserIndex, 0, "||SEGURO DE CVC ACTIVADO. ~0~255~0~1~0")
    Else
        Call SendData(ToIndex, UserIndex, 0, "||SEGURO DE CVC DESACTIVADO. ~255~0~0~1~0")
    End If


    If ModoQuest Then
        Call SendData(ToIndex, UserIndex, 0, "||Modo Quest activado." & FONTTYPE_FENIX)
        Call SendData(ToIndex, UserIndex, 0, "||Los neutrales pueden poner /MERCENARIO ALIANZA o /MERCENARIO HORDA para enlistarse en alguna facción temporalmente durante la quest." & FONTTYPE_FENIX)
        'Call SendData(ToIndex, UserIndex, 0, "||Al morir puedes poner /REGRESAR y serás teletransportado a Althalos." & FONTTYPE_FENIX)
    End If


    'Soporte Dylan.-
    Dim TieneSoporte As String
    TieneSoporte = GetVar(CharPath & UCase$(UserList(UserIndex).name) & ".chr", "STATS", "Respuesta")
    If Len(TieneSoporte) Then
        If Right$(TieneSoporte, 3) <> "0k1" Then
            Call SendData(ToIndex, UserIndex, 0, "TENSO")
        End If
    End If
    'soporte Dylan.-

    If Lloviendo Then Call SendData(ToIndex, UserIndex, 0, "LLU")

    N = FreeFile
    Open App.path & "\logs\numusers.log" For Output As N
    Print #N, NumUsers + NumBots
    Close #N

    UserList(UserIndex).flags.ClienteValido = 0    ' VERIFICACION CLIENTES
    Call SendData(ToIndex, UserIndex, 0, "ANCL")


    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Aura <> 0 Then
        UserList(UserIndex).Char.Aura = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Aura
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CRA" & UserList(UserIndex).Char.CharIndex & "," & UserList(UserIndex).Char.Aura)
    Else
        UserList(UserIndex).Char.Aura = 0
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CRA" & UserList(UserIndex).Char.CharIndex & "," & 0)
    End If
    Exit Sub
Error:
    Call LogError("Error en ConnectUser: " & name & " " & Err.Description)

End Sub

Sub SendMOTD(UserIndex As Integer)
    Dim j As Integer
    For j = 1 To MaxLines
        Call SendData(ToIndex, UserIndex, 0, "||" & MOTD(j).Texto)
    Next
    Call SendData(ToIndex, UserIndex, 0, "||Castillo Norte conquistado por: " & CastilloNorte & " Fecha: " & DateNorte & " Hora: " & HoraNorte & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "||Castillo Sur conquistado por: " & CastilloSur & " Fecha: " & DateSur & " Hora: " & HoraSur & FONTTYPE_INFO)

End Sub
Sub CloseUser(ByVal UserIndex As Integer)
    On Error GoTo errhandler
    Dim i As Integer, aN As Integer
    Dim name As String
    name = UCase$(UserList(UserIndex).name)

    aN = UserList(UserIndex).flags.AtacadoPorNpc

    If aN Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = 0
    End If

    If UserList(UserIndex).Tienda.NpcTienda Then
        Call DevolverItemsVenta(UserIndex)
        Npclist(UserList(UserIndex).Tienda.NpcTienda).flags.TiendaUser = 0
    End If

    If UserList(UserIndex).flags.Privilegios > 0 Then Call SendData(ToMoreAdmins, UserIndex, 0, "||" & UserList(UserIndex).name & " se desconectó." & FONTTYPE_FENIX)

    If UserList(UserIndex).flags.Party Then
        Call SendData(ToParty, UserIndex, 0, "||" & UserList(UserIndex).name & " se desconectó." & FONTTYPE_PARTY)
        If Party(UserList(UserIndex).PartyIndex).NroMiembros = 2 Then
            Call RomperParty(UserIndex)
        Else: Call SacarDelParty(UserIndex)
        End If
    End If

    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "CFX" & "0," & UserList(UserIndex).Char.CharIndex & ",0,0,0")


    If UserList(UserIndex).Caballos.Num And UserList(UserIndex).flags.Montado = 1 Then Call Desmontar(UserIndex)

    If UserList(UserIndex).flags.AdminInvisible Then Call DoAdminInvisible(UserIndex)
    If UserList(UserIndex).flags.Transformado Then Call DoTransformar(UserIndex, False)

    Call SaveUser(UserIndex, CharPath & name & ".chr")

    If MapInfo(UserList(UserIndex).pos.Map).NumUsers Then Call SendData(ToMapButIndex, UserIndex, UserList(UserIndex).pos.Map, "QDL" & UserList(UserIndex).Char.CharIndex)
    If UserList(UserIndex).Char.CharIndex Then Call EraseUserChar(ToMapButIndex, UserIndex, UserList(UserIndex).pos.Map, UserIndex)
    If UserList(UserIndex).Caballos.Num Then Call QuitarCaballos(UserIndex)

    For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
        If UserList(UserIndex).MascotasIndex(i) Then
            If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
               Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
        End If
    Next

    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If

    If Len(UserList(UserIndex).GuildInfo.GuildName) > 0 And UserList(UserIndex).flags.Privilegios = 0 Then Call SendData(ToGuildMembers, UserIndex, 0, "5B" & UserList(UserIndex).name)
    Dim NpcIndex As Integer
    Call QuitarDeUsersPorMapa(UserIndex)

    If MapInfo(UserList(UserIndex).pos.Map).NumUsers < 0 Then MapInfo(UserList(UserIndex).pos.Map).NumUsers = 0
    Exit Sub


errhandler:
    Call LogError("Error en CloseUser " & Err.Description)

End Sub
Function EsVigilado(Espiado As Integer) As Boolean
    Dim i As Integer

    For i = 1 To 10
        If UserList(Espiado).flags.Espiado(i) > 0 Then
            EsVigilado = True
            Exit Function
        End If
    Next

End Function
Sub ActivarTrampa(UserIndex As Integer)
    Dim i As Integer, TU As Integer

    For i = 1 To MapInfo(UserList(UserIndex).pos.Map).NumUsers
        TU = MapInfo(UserList(UserIndex).pos.Map).UserIndex(i)
        If UserList(TU).flags.Paralizado = 0 And Abs(UserList(UserIndex).pos.X - UserList(TU).pos.X) <= 3 And Abs(UserList(UserIndex).pos.Y - UserList(TU).pos.Y) <= 3 And TU <> UserIndex And PuedeAtacar(UserIndex, TU) Then
            UserList(TU).flags.QuienParalizo = UserIndex
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = Timer - 15 * Buleano(UserList(TU).Clase = GUERRERO And UserList(TU).Recompensas(3) = 2)
            Call SendData(ToIndex, TU, 0, "PU" & DesteEncripTE(UserList(TU).pos.X & "," & UserList(TU).pos.Y))
            Call SendData(ToIndex, TU, 0, ("P9"))
            Call SendData(ToPCArea, TU, UserList(TU).pos.Map, "CFX" & "0," & UserList(TU).Char.CharIndex & ",12" & "," & 0 & "1")
        End If
    Next

    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).pos.Map, "TW112")

End Sub
Sub DesactivarMercenarios()
    Dim UserIndex As Integer

    For UserIndex = 1 To LastUser
        If UserList(UserIndex).Faccion.Bando <> Neutral And UserList(UserIndex).Faccion.Bando <> UserList(UserIndex).Faccion.BandoOriginal Then
            Call SendData(ToIndex, UserIndex, 0, "||La quest ha terminado, has dejado de ser un mercenario." & FONTTYPE_FENIX)
            UserList(UserIndex).Faccion.Bando = Neutral
            Call UpdateUserChar(UserIndex)
        End If
    Next

End Sub
Function YaVigila(Espiado As Integer, Espiador As Integer) As Boolean
    Dim i As Integer

    For i = 1 To 10
        If UserList(Espiado).flags.Espiado(i) = Espiador Then
            UserList(Espiado).flags.Espiado(i) = 0
            YaVigila = True
            Exit Function
        End If
    Next

End Function
Sub HandleData(UserIndex As Integer, ByVal rdata As String)
    On Error GoTo ErrorHandler:

    Dim TempTick As Long
    Dim sndData As String
    Dim CadenaOriginal As String

    Dim LoopC As Integer
    Dim nPos As WorldPos
    Dim tStr As String
    Dim tInt As Integer
    Dim tLong As Long
    Dim numeromail As Integer
    Dim TIndex As Integer
    Dim tName As String
    Dim Clase As Byte
    Dim NumNPC As Integer
    Dim tMessage As String
    Dim i As Integer
    Dim auxind As Integer
    Dim Arg1 As String
    Dim Arg2 As String
    Dim arg3 As String
    Dim Arg4 As String
    Dim Arg5 As Integer
    Dim Arg6 As String
    Dim DummyInt As Integer
    Dim Antes As Boolean
    Dim Ver As String
    Dim encpass As String
    Dim Pass As String
    Dim mapa As Integer
    Dim usercon As String
    Dim nameuser As String
    Dim name As String
    Dim ind
    Dim GMDia As String
    Dim GMMapa As String
    Dim GMPJ As String
    Dim GMMail As String
    Dim GMGM As String
    Dim GMTitulo As String
    Dim GMMensaje As String
    Dim N As Integer
    Dim wpaux As WorldPos
    Dim mifile As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim cliMD5 As String
    Dim UserFile As String
    Dim UserName As String
    UserName = UserList(UserIndex).name
    UserFile = CharPath & UCase$(UserName) & ".chr"
    Dim ClientCRC As String
    Dim ServerSideCRC As Long
    Dim NombreIniChat As String
    Dim cantidadenmapa As Integer
    Dim Prueba1 As Integer
    Dim NpcIndex As Integer

    'rdata = Mod_DesEncript.DesEncriptar(rdata)
    CadenaOriginal = rdata

    If UserIndex <= 0 Then
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    If Recargando Then
        Call SendData(ToIndex, UserIndex, 0, "!!Recargando información, espere unos momentos.")
        Call CloseSocket(UserIndex)
    End If


    If Left$(rdata, 13) = "gIvEmEvAlcOde" Then
        UserList(UserIndex).flags.ValCoDe = CInt(RandomNumber(20000, 32000))
        UserList(UserIndex).RandKey = CLng(RandomNumber(145, 99999))
        UserList(UserIndex).PrevCRC = UserList(UserIndex).RandKey
        UserList(UserIndex).PacketNumber = 100

        UserList(UserIndex).PersonalPass = RandomNumber(100, 999)
        Dim S As String
        S = Encripta(CStr(UserList(UserIndex).PersonalPass), True)
        S = Encripta(S, True)
        Call SendData(ToIndex, UserIndex, 0, "VAL" & UserList(UserIndex).RandKey & "," & UserList(UserIndex).flags.ValCoDe & "," & S)
        UserList(UserIndex).PrevCRC = 0
        Exit Sub
    ElseIf Not UserList(UserIndex).flags.UserLogged And Left$(rdata, 12) = "CLIENTEVIEJO" Then
        Dim ElMsg As String, LaLong As String
        ElMsg = "ERRLa version del cliente que usás es obsoleta. Si deseas conectarte a este servidor entrá a www.fenixao.com.ar y allí podrás enterarte como hacer."
        If Len(ElMsg) > 255 Then ElMsg = Left$(ElMsg, 255)
        LaLong = Chr$(0) & Chr$(Len(ElMsg))
        Call SendData(ToIndex, UserIndex, 0, LaLong & ElMsg)
        Call CloseSocket(UserIndex)
        Exit Sub
    Else
        ClientCRC = Right$(rdata, Len(rdata) - InStrRev(rdata, Chr$(126)))
        tStr = Left$(rdata, Len(rdata) - Len(ClientCRC) - 1)

        rdata = tStr
        tStr = ""

    End If

    UserList(UserIndex).Counters.IdleCount = Timer

    If Not UserList(UserIndex).flags.UserLogged Then

        Select Case Left$(rdata, 6)


            Case "BORRAR"
                rdata = Right$(rdata, Len(rdata) - 6)
                Dim Password As String
                name = ReadField(1, rdata, 44)
                Password = MD5String(ReadField(2, rdata, 44))

                If CheckForSameName(UserIndex, name) Then
                    If NameIndex(name) = UserIndex Then Call CloseSocket(NameIndex(name))
                    Call SendData(ToIndex, UserIndex, 0, "ERRPerdón, un usuario con el mismo nombre se ha logeado.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                If Not AsciiValidos(name) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERREl nombre especificado es inválido.")
                    Exit Sub
                End If

                If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = False Then
                    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje no existe")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(name) & ".chr", "INIT", "Password")) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRLa contraseña no coinciden.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                If BANCheck(name) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje se encuentra baneado y por lo tanto no se podrá borrar. Haga su descargo en el foro o contáctese con la administración del juego.")
                    Exit Sub
                End If

                If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) Then
                    Kill CharPath & UCase$(name) & ".chr"
                    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje fué borrado correctamente!")
                    CloseSocket UserIndex
                    Exit Sub
                End If
                Exit Sub
            Case "OLOGIO"

                rdata = Right$(rdata, Len(rdata) - 6)
                tName = ReadField(1, rdata, 44)
                tName = RTrim(tName)


                If Not AsciiValidos(tName) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
                    Exit Sub
                End If

                If UserList(UserIndex).PersonalPass <> val(ReadField(3, rdata, 44)) Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                Call ConnectUser(UserIndex, tName, ReadField(4, rdata, 44), UserList(UserIndex).AccountedPass, ReadField(5, rdata, 44))



                Exit Sub

            Case "ALOGIN"
                rdata = Right$(rdata, Len(rdata) - 6)

                If Not AsciiValidos(ReadField(1, rdata, 44)) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
                    Call CloseSocket(UserIndex, True)
                    Exit Sub
                End If

                If Not CuentaExiste(ReadField(1, rdata, 44)) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRLa cuenta no existe.")
                    Call CloseSocket(UserIndex, True)
                    Exit Sub
                End If
                If UserList(UserIndex).PersonalPass <> val(ReadField(4, rdata, 44)) Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                Call ConnectAccount(UserIndex, ReadField(1, rdata, 44), ReadField(2, rdata, 44))
                Exit Sub

            Case "NACCNT"

                rdata = Right$(rdata, Len(rdata) - 6)

                Dim NCuenta As String
                Dim Passw As String
                Dim Mail As String

                'cuentas
                NCuenta = ReadField(1, rdata, Asc(","))
                Passw = ReadField(2, rdata, Asc(","))
                Mail = ReadField(3, rdata, Asc(","))

                If UserList(UserIndex).PersonalPass <> val(ReadField(5, rdata, 44)) Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                Call CreateAccount(NCuenta, Passw, Mail, UserIndex)


                Exit Sub
            Case "TIRDAD"
                If Restringido Then
                    Call SendData(ToIndex, UserIndex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
                    Exit Sub
                End If

                UserList(UserIndex).Stats.UserAtributosBackUP(1) = 18
                UserList(UserIndex).Stats.UserAtributosBackUP(2) = 18
                UserList(UserIndex).Stats.UserAtributosBackUP(3) = 18
                UserList(UserIndex).Stats.UserAtributosBackUP(4) = 18
                UserList(UserIndex).Stats.UserAtributosBackUP(5) = 18

                Call SendData(ToIndex, UserIndex, 0, ("DADOS" & UserList(UserIndex).Stats.UserAtributosBackUP(1) & "," & UserList(UserIndex).Stats.UserAtributosBackUP(2) & "," & UserList(UserIndex).Stats.UserAtributosBackUP(3) & "," & UserList(UserIndex).Stats.UserAtributosBackUP(4) & "," & UserList(UserIndex).Stats.UserAtributosBackUP(5)))

                Exit Sub

            Case "GMMAOP"

                If PuedeCrearPersonajes = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRNo se pueden crear más personajes en este servidor.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERRHas creado demasiados personajes.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                rdata = Right$(rdata, Len(rdata) - 6)
                cliMD5 = Right$(rdata, 8)



                If UserList(UserIndex).PersonalPass <> val(ReadField(35, rdata, 44)) Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                Call ConnectNewUser(UserIndex, ReadField(1, rdata, 44), val(ReadField(2, rdata, 44)), ReadField(3, rdata, 44), ReadField(5, rdata, 44), ReadField(6, rdata, 44), _
                                    ReadField(7, rdata, 44), ReadField(8, rdata, 44), ReadField(9, rdata, 44), ReadField(10, rdata, 44), ReadField(11, rdata, 44), ReadField(12, rdata, 44), _
                                    ReadField(13, rdata, 44), ReadField(14, rdata, 44), ReadField(15, rdata, 44), ReadField(16, rdata, 44), ReadField(17, rdata, 44), ReadField(18, rdata, 44), _
                                    ReadField(19, rdata, 44), ReadField(20, rdata, 44), ReadField(21, rdata, 44), ReadField(22, rdata, 44), ReadField(23, rdata, 44), ReadField(24, rdata, 44), _
                                    ReadField(25, rdata, 44), ReadField(26, rdata, 44), ReadField(27, rdata, 44), ReadField(28, rdata, 44), ReadField(29, rdata, 44), ReadField(30, rdata, 44), _
                                    ReadField(31, rdata, 44), ReadField(32, rdata, 44), ReadField(33, rdata, 44), ReadField(34, rdata, 44), ReadField(36, rdata, 44), ReadField(37, rdata, 44))

                Exit Sub
        End Select
    End If

    If Not UserList(UserIndex).flags.UserLogged Then
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    Dim Procesado As Boolean

    If UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = False
        UserList(UserIndex).Counters.Salir = 0
        Call SendData(ToIndex, UserIndex, 0, "{A")
    End If

    If Left$(rdata, 1) <> "#" Then
        Call HandleData1(UserIndex, rdata, Procesado)
        If Procesado Then Exit Sub
    Else
        Call HandleData2(UserIndex, rdata, Procesado)
        If Procesado Then Exit Sub
    End If

    If Left$(rdata, 7) = "BANEAME" Then
        Dim h As Integer
        rdata = Right$(rdata, Len(rdata) - 7)
        Arg1 = ReadField(1, rdata, Asc(","))
        Arg2 = ReadField(2, rdata, Asc(","))
        h = FreeFile
        Open App.path & "\Logs\CHITEROS.log" For Append Shared As h
        Print #h, "########################################################################"
        Print #h, "USUARIO: " & UserList(UserIndex).name
        Print #h, "FECHA: " & Date
        Print #h, "HORA: " & Time
        Print #h, "CHEAT: " & Arg1
        Print #h, "CLASS: " & Arg2
        Print #h, "########################################################################"
        Print #h, " "
        Close #h
        UserList(UserIndex).flags.Ban = 1
        Call SendData(ToAdmins, 0, 0, "||LhiriusSeC> " & UserList(UserIndex).name & " ha sido echado por uso de " & Arg1 & FONTTYPE_FIGHT)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    'Dylan.- Sistema de soporte.
    If UCase$(rdata) = "/MISOPORTE" Then
        Dim MiRespuesta As String
        MiRespuesta = GetVar(CharPath & UCase$(UserList(UserIndex).name) & ".CHR", "STATS", "Respuesta")
        If Len(MiRespuesta) Then
            If Right$(MiRespuesta, 3) = "0k1" Then
                Call SendData(ToIndex, UserIndex, 0, "VERSO" & Left$(MiRespuesta, Len(MiRespuesta) - 3))
            Else
                Call SendData(ToIndex, UserIndex, 0, "VERSO" & MiRespuesta)
                MiRespuesta = MiRespuesta & "0k1"
                Call WriteVar(CharPath & UCase$(UserList(UserIndex).name) & ".CHR", "STATS", "Respuesta", MiRespuesta)
            End If
        Else
            MiRespuesta = GetVar(CharPath & UCase$(UserList(UserIndex).name) & ".CHR", "STATS", "Soporte")

            If Len(MiRespuesta) Then
                Call SendData(ToIndex, UserIndex, 0, "||No respondida aún" & FONTTYPE_FENIX)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No has mandado ningun soporte!" & FONTTYPE_FENIX)
            End If

        End If

        Exit Sub
    End If

    If UCase$(Left$(rdata, 9)) = "/ZOPORTE " Then
        If SoporteDesactivado Then
            Call SendData(ToIndex, UserIndex, 0, "||El soporte se encuentra deshabilitado." & FONTTYPE_FENIX)
            Exit Sub
        End If
        If Len(rdata) > 310 Then Exit Sub
        If InStr(rdata, "°") Then Exit Sub
        If InStr(rdata, "~") Then Exit Sub
        'If UserList(userindex).flags.Silenciado > 0 Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 9)
        'Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " >" & "" & "SOPORTE:" & rdata & FONTTYPE_FIGHT)
        'Call SendData(ToIndex, userindex, 0, "||El soporte fue enviado. Rogamos que tengas paciencia y aguardes a ser atendido por un GM. No escribas más de un mensaje sobre el mismo tema." & FONTTYPE_fenix)

        Dim SoporteA As String

        SoporteA = GetVar(CharPath & UCase$(UserList(UserIndex).name) & ".CHR", "STATS", "Respuesta")

        'SI HAY RESPUESTA Y NO ESTA LEIDA LE AVISA.
        If Len(SoporteA) > 0 And Right$(SoporteA, 3) <> "0k1" Then
            Call SendData(ToIndex, UserIndex, 0, "||Primero debes leer la respuesta de tu anterior soporte." & FONTTYPE_FENIX)
            Exit Sub
        End If
        '/

        SoporteA = GetVar(CharPath & UCase$(UserList(UserIndex).name) & ".CHR", "STATS", "Soporte")

        'SI MANDO SOPORTE ANTES Y TODAVIA NO LE RESPONDIERON TIENE QE ESPERAR
        If Len(SoporteA) > 0 And Right$(SoporteA, 3) <> "0k1" Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya has mandado una consulta. Débes esperar la respuesta para enviar otro." & FONTTYPE_FENIX)
            Exit Sub
        End If
        '0K

        SoporteA = "Dia:" & Day(Now) & " Hora:" & Time & " - Soporte: " & Replace(Replace(rdata, ";", ":"), Chr$(13) & Chr$(10), Chr(32))




        Call WriteVar(CharPath & UCase$(UserList(UserIndex).name) & ".CHR", "STATS", "Soporte", SoporteA)
        Call WriteVar(CharPath & UCase$(UserList(UserIndex).name) & ".CHR", "STATS", "Respuesta", "")
        SoporteS.Add (UserList(UserIndex).name)
        Call SendData(ToIndex, UserIndex, 0, "||La consulta ha sido enviada con éxito. Aguarde hasta que un administrador le responda su consulta." & FONTTYPE_FENIX)
        Call SendData(ToAdmins, 0, 0, "||Nuevo SOS!" & FONTTYPE_FENIX)
        Exit Sub
    End If


    If UCase$(Left$(rdata, 8)) = "/SOPORTE" Then
        Call SendData(ToIndex, UserIndex, 0, "SHWSUP")
        Exit Sub
    End If

    'Dylan.- SISTEMA DE SOPORTE

    If UCase$(rdata) = "/PING" Then
        Call SendData(ToIndex, UserIndex, 0, "BUENO")
        Exit Sub
    End If

    If UCase$(rdata) = "/TIEMPO" Then
        Call SendData(ToIndex, UserIndex, 0, "||Faltan " & (60 - MinCasti) & " minutos para la siguiente entrega de Premios a Clanes." & FONTTYPE_GUILD)
        Call SendData(ToIndex, UserIndex, 0, "||Faltan " & (60 - MinutosTorneo) & " minutos para el siguiente Torneo Automático." & FONTTYPE_GUILD)
        Call SendData(ToIndex, UserIndex, 0, "||Faltan " & (30 - MinutosDeath) & " minutos para el siguiente Deathmatch Automático." & FONTTYPE_GUILD)
        Call SendData(ToIndex, UserIndex, 0, "||Faltan " & (TiempoEntreGuerra - TiempoGuerra) & " minutos para la siguiente Guerra Faccionaria." & FONTTYPE_GUILD)
        Exit Sub
    End If
    If UCase$(rdata) = "/ROSTRO" Then
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Primero debes hacer click en el NPC" & FONTTYPE_TALK)
            Exit Sub
        End If

        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).pos, UserList(UserIndex).pos) > 10 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estás lejos." & FONTTYPE_TALK)
            Exit Sub
        End If

        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_CIRUJANO _
           Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If Distancia(UserList(UserIndex).pos, Npclist(UserList(UserIndex).flags.TargetNpc).pos) > 10 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estás lejos." & FONTTYPE_TALK)
            Exit Sub
        End If

        Dim UserHead As Integer
        Dim QGENERO As Byte
        QGENERO = UserList(UserIndex).Genero
        Select Case QGENERO
            Case HOMBRE
                Select Case UserList(UserIndex).Raza
                    Case HUMANO
                        UserHead = CInt(RandomNumber(1, 24))
                        If UserHead > 24 Then UserHead = 24
                    Case ELFO
                        UserHead = CInt(RandomNumber(1, 7)) + 100
                        If UserHead > 107 Then UserHead = 107
                    Case ELFO_OSCURO
                        UserHead = CInt(RandomNumber(1, 4)) + 200
                        If UserHead > 204 Then UserHead = 204
                    Case ENANO
                        UserHead = RandomNumber(1, 4) + 300
                        If UserHead > 304 Then UserHead = 304
                    Case GNOMO
                        UserHead = RandomNumber(1, 3) + 400
                        If UserHead > 403 Then UserHead = 403
                    Case Else
                        UserHead = 1

                End Select
            Case MUJER
                Select Case UserList(UserIndex).Raza
                    Case HUMANO
                        UserHead = CInt(RandomNumber(1, 4)) + 69
                        If UserHead > 73 Then UserHead = 73
                    Case ELFO
                        UserHead = CInt(RandomNumber(1, 5)) + 169
                        If UserHead > 174 Then UserHead = 174
                    Case ELFO_OSCURO
                        UserHead = CInt(RandomNumber(1, 5)) + 269
                        If UserHead > 274 Then UserHead = 274
                    Case GNOMO
                        UserHead = RandomNumber(1, 4) + 469
                        If UserHead > 473 Then UserHead = 473
                    Case ENANO
                        UserHead = RandomNumber(1, 3) + 369
                        If UserHead > 372 Then UserHead = 372
                    Case Else
                        UserHead = 70
                End Select
        End Select

        If UserList(UserIndex).Char.Head = UserHead Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbRed & "°" & "Ahora no puedo operar , vuelva más tarde." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_TALK)
            Exit Sub
        End If

        UserList(UserIndex).Char.Head = UserHead
        UserList(UserIndex).OrigChar.Head = UserHead
        Call SendData(ToIndex, UserIndex, 0, "||" & vbRed & "°" & "Tu rostro ha sido operado." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_TALK)
        Call ChangeUserChar(ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).Char.Body, val(UserHead), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Alas)
        Exit Sub
    End If


    If UCase$(Left$(rdata, 6)) = "/IRCVC" Then
        If UserList(UserIndex).flags.SeguroCVC = True Then
            UserList(UserIndex).flags.SeguroCVC = False
            Call SendData(ToIndex, UserIndex, 0, "||SEGURO DE CVC DESACTIVADO. ~255~0~0~1~0")
            Call SendData(ToIndex, UserIndex, 0, "||No serás llevado a ninguna guerra de clanes que realice tu clan." & FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(UserIndex).flags.SeguroCVC = False Then
            UserList(UserIndex).flags.SeguroCVC = True
            Call SendData(ToIndex, UserIndex, 0, "||SEGURO DE CVC ACTIVADO. ~0~255~0~1~0")
            Call SendData(ToIndex, UserIndex, 0, "||Serás llevado a todas guerras de clanes que realice tu clan." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If


    If UCase$(Left$(rdata, 5)) = "/CVC " Then
        Dim ClanName As String

        ClanName = Right$(rdata, Len(rdata) - 5)
        CVC.NombreClan1 = ""
        CVC.NombreClan2 = ""

        If UserList(UserIndex).GuildInfo.GuildName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes clan." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Solo los lideres pueden mandar guerras de clanes." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UCase$(ClanName) = UCase$(UserList(UserIndex).GuildInfo.GuildName) Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes mandarte guerra de clan a tu mismo clan." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).pos.Map = 22 Or UserList(UserIndex).pos.Map = 205 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 19 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 7 Then Exit Sub


        If UserList(UserIndex).flags.EnCVC = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya estás en CVC." & FONTTYPE_INFO)
            Exit Sub
        End If

        If CVC.OcupadoCVC = True Then
            Call SendData(ToIndex, UserIndex, 0, "||La zona de CVCs está ocupada, intenta en otro momento." & FONTTYPE_INFO)
            Exit Sub
        End If

        Dim IndiceOtroClan As Integer
        For IndiceOtroClan = 1 To LastUser
            If UCase$(ClanName) = UCase$(UserList(IndiceOtroClan).GuildInfo.GuildName) And UserList(IndiceOtroClan).GuildInfo.EsGuildLeader = 1 Then
                Call SendData(ToIndex, IndiceOtroClan, 0, "||El clan " & UserList(UserIndex).GuildInfo.GuildName & " desafia a tu clan a una Guerra de Clanes, para aceptar el desafio Escribe /SICVC." & FONTTYPE_FIGHT)
                Call SendData(ToIndex, UserIndex, 0, "||El desafio al clan " & UserList(IndiceOtroClan).GuildInfo.GuildName & " ha sido enviado." & FONTTYPE_INFO)
                CVC.NombreClan1 = UserList(UserIndex).GuildInfo.GuildName
                CVC.NombreClan2 = UserList(IndiceOtroClan).GuildInfo.GuildName
                UserList(IndiceOtroClan).GuildInfo.LoDesafiaron = True
                Exit Sub    'si ya lo envio no es necesario volver a correr el for...
            End If
        Next IndiceOtroClan

        If CVC.NombreClan2 = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||Es posible que el clan no exista o el lider del clan oponente no se encuentre online." & FONTTYPE_INFO)
            Exit Sub
        End If

    End If

    If UCase$(Left$(rdata, 7)) = "/SICVC" Then
        If UserList(UserIndex).GuildInfo.GuildName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes clan." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Solo los lideres pueden aceptar guerras de clanes." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).flags.EnCVC = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya estás en CVC." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).pos.Map = 22 Or UserList(UserIndex).pos.Map = 205 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 19 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 7 Then Exit Sub


        If CVC.OcupadoCVC = True Then
            Call SendData(ToIndex, UserIndex, 0, "||La zona de CVCs está ocupada, intenta en otro momento." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(CVC.NombreClan1) Then Exit Sub

        If UserList(UserIndex).GuildInfo.LoDesafiaron = True Then

            CVC.CantidadDeParticipantes1 = 0
            CVC.CantidadDeParticipantes2 = 0
            For IndiceOtroClan = 1 To LastUser    'seteamos la cantidad de participantes de cada clan...
                If UCase$(UserList(IndiceOtroClan).GuildInfo.GuildName) = UCase$(CVC.NombreClan1) Then
                    If UserList(IndiceOtroClan).flags.SeguroCVC = True Then
                        CVC.CantidadDeParticipantes1 = CVC.CantidadDeParticipantes1 + 1
                    End If
                End If
                If UCase$(UserList(IndiceOtroClan).GuildInfo.GuildName) = UCase$(CVC.NombreClan2) Then
                    If UserList(IndiceOtroClan).flags.Seguro = True Then
                        CVC.CantidadDeParticipantes2 = CVC.CantidadDeParticipantes2 + 1
                    End If
                End If
            Next IndiceOtroClan

            If CVC.CantidadDeParticipantes1 <= CVC.CantidadDeParticipantes2 Then
                CVC.CantidadQueParticipa = CVC.CantidadDeParticipantes1
            ElseIf CVC.CantidadDeParticipantes1 > CVC.CantidadDeParticipantes2 Then
                CVC.CantidadQueParticipa = CVC.CantidadDeParticipantes2
            End If

            If CVC.CantidadQueParticipa = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||La guerra contra el clan " & CVC.NombreClan2 & " ha sido cancelada por falta de participantes de alguno de los clanes." & FONTTYPE_INFO)

                UserList(UserIndex).GuildInfo.LoDesafiaron = False
                Exit Sub
            End If


            'empieza el cvc
            CVC.OcupadoCVC = True
            CVC.CantidadDeParticipantes1 = 0
            CVC.CantidadDeParticipantes2 = 0
            Call SendData(ToAll, 0, 0, "||Los clanes " & CVC.NombreClan1 & " y " & CVC.NombreClan2 & " van a combatir en una Guerra de Clanes." & FONTTYPE_GUILD)
            'empezamos a sumonear
            For IndiceOtroClan = 1 To LastUser
                'clan 1
                If UCase$(UserList(IndiceOtroClan).GuildInfo.GuildName) = UCase$(CVC.NombreClan1) Then
                    If UserList(IndiceOtroClan).flags.SeguroCVC = True Then
                        If CVC.CantidadDeParticipantes1 <> CVC.CantidadQueParticipa Then
                            UserList(IndiceOtroClan).flags.ViejaPos.Map = UserList(IndiceOtroClan).pos.Map
                            UserList(IndiceOtroClan).flags.ViejaPos.X = UserList(IndiceOtroClan).pos.X
                            UserList(IndiceOtroClan).flags.ViejaPos.Y = UserList(IndiceOtroClan).pos.Y
                            WarpUserChar IndiceOtroClan, 70, RandomNumber(47, 49), RandomNumber(81, 82), True
                            CVC.CantidadDeParticipantes1 = CVC.CantidadDeParticipantes1 + 1
                            UserList(IndiceOtroClan).flags.EnCVC = True
                        End If
                    End If
                End If
                If UCase$(UserList(IndiceOtroClan).GuildInfo.GuildName) = UCase$(CVC.NombreClan2) Then
                    If UserList(IndiceOtroClan).flags.SeguroCVC = True Then
                        If CVC.CantidadDeParticipantes2 <> CVC.CantidadQueParticipa Then
                            UserList(IndiceOtroClan).flags.ViejaPos.Map = UserList(IndiceOtroClan).pos.Map
                            UserList(IndiceOtroClan).flags.ViejaPos.X = UserList(IndiceOtroClan).pos.X
                            UserList(IndiceOtroClan).flags.ViejaPos.Y = UserList(IndiceOtroClan).pos.Y
                            WarpUserChar IndiceOtroClan, 70, RandomNumber(47, 49), RandomNumber(31, 32), True
                            CVC.CantidadDeParticipantes2 = CVC.CantidadDeParticipantes2 + 1
                            UserList(IndiceOtroClan).flags.EnCVC = True
                        End If
                    End If
                End If
            Next IndiceOtroClan

            UserList(UserIndex).GuildInfo.LoDesafiaron = False
            Exit Sub
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Nadie te ha desafiado una guerra de clanes." & FONTTYPE_FIGHT)
            Exit Sub
        End If
    End If


    'DUELO POR ITEMS
    If UCase$(Left$(rdata, 12)) = "/DUELOITEMS " Then
        ClanName = Right$(rdata, Len(rdata) - 5)

        If NameIndex(ClanName) = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        Else
            dIndex = NameIndex(ClanName)
        End If

        If dIndex = UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes dueliar contra vos mismo." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).pos.Map = 22 Or UserList(UserIndex).pos.Map = 205 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 19 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 7 Then Exit Sub

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas muerto!!." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).pos.Map = 66 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas en carcel!!. " & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(dIndex).pos.Map = 66 Then
            Call SendData(ToIndex, UserIndex, 0, "||Esta en carcel tu oponente!!. " & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(dIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario està muerto." & FONTTYPE_INFO)
            Exit Sub
        End If

        If DueloXI.Ocupado = True Then
            Call SendData(ToIndex, UserIndex, 0, "||La zona de duelos por items está ocupado." & FONTTYPE_INFO)
            Exit Sub
        End If


        UserList(dIndex).flags.LeMandaronDueloXI = True
        UserList(dIndex).flags.UltimoEnMandarDueloXI = UserList(UserIndex).name
        Call SendData(ToIndex, (dIndex), 0, "||" & UserList(UserIndex).name & " [" & ListaClases(UserList(UserIndex).Clase) & " - " & UserList(UserIndex).Stats.ELV & "] - te està desafiando en un duelo por items, para aceptar escribi /ACEPTARDUELO." & "~124~124~124~1~0")
        Exit Sub
    End If



    'DUELO POR ORO.
    If UCase$(Left$(rdata, 7)) = "/DUELO " Then

        dMap = 5
        rdata = Right$(rdata, Len(rdata) - 7)
        dUser = ReadField(1, rdata, Asc("@"))

        If NameIndex(dUser) = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        Else
            dIndex = NameIndex(dUser)
        End If

        dMoney = ReadField(2, rdata, Asc("@"))
        If dIndex = UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes dueliar contra vos mismo." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(UserIndex).Stats.GLD < val(dMoney) Then
            Call SendData(ToIndex, UserIndex, 0, "||No tenes esa cantidad de oro." & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(dIndex).Stats.GLD < val(dMoney) Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no tiene esa cantidad de oro." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).pos.Map = 22 Or UserList(UserIndex).pos.Map = 205 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 19 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 7 Then Exit Sub

        If UserList(UserIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas muerto!!." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).pos.Map = 66 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas en carcel!!. " & FONTTYPE_INFO)
            Exit Sub
        End If

        If UserList(dIndex).pos.Map = 66 Then
            Call SendData(ToIndex, UserIndex, 0, "||Esta en carcel tu oponente!!. " & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(dIndex).flags.Muerto Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario està muerto." & FONTTYPE_INFO)
            Exit Sub
        End If

        If val(dMoney) < 100000 Then
            Call SendData(ToIndex, UserIndex, 0, "||El minimo de oro para duelear es de 100.000 monedas de oro." & FONTTYPE_INFO)
            Exit Sub
        End If

        If MapInfo(dMap).NumUsers = 2 Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya hay un duelo en curso, esperà que termine." & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(dIndex).flags.LeMandaronDuelo = True
        UserList(dIndex).flags.UltimoEnMandarDuelo = UserList(UserIndex).name
        Call SendData(ToIndex, (dIndex), 0, "||" & UserList(UserIndex).name & " [" & ListaClases(UserList(UserIndex).Clase) & " - " & UserList(UserIndex).Stats.ELV & "] - te està desafiando en un duelo por " & PonerPuntos(val(dMoney)) & " monedas de oro, para aceptar escribi /SIDUELO." & "~124~124~124~1~0")

    End If

    If UCase$(Left$(rdata, 8)) = "/RANKING" Then
        Call LeerRanking(UserIndex)
        Exit Sub
    End If
    If UCase$(rdata) = "/EDITAME" Then
        'If Not UserList(UserIndex).flags.Privilegios > 1 Then Exit Sub
        If UserList(UserIndex).Stats.ELV > 49 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡Ya tienes suficientes niveles! Para seguir subiendo niveles necesitas " & EFrags(UserList(UserIndex).Stats.ELV) & " frags." & FONTTYPE_INFO)
            
            Exit Sub
        Else
            'Nivel
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.ELU
            Call CheckUserLevel(UserIndex)
            'Nivel
            'oro
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + 1000000
            Dim Skills As Integer
            For Skills = 1 To NUMSKILLS
                UserList(UserIndex).Stats.UserSkills(Skills) = 100
            Next Skills
        End If
        Exit Sub
    End If


    If UCase$(Left$(rdata, 8)) = "/NOQUEST" Then
        If UserList(UserIndex).Quest.Index = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No estas haciendo ninguna quest." & FONTTYPE_FIGHT)
            Exit Sub
        End If
        UserList(UserIndex).Quest.Index = 0
        Call SendData(ToIndex, UserIndex, 0, "|| Has abandonado la Quest" & FONTTYPE_FENIX)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 8)) = "/SIDUELO" Then

        If UserList(UserIndex).pos.Map = 22 Or UserList(UserIndex).pos.Map = 205 Or UserList(UserIndex).pos.Map = 14 Or UserList(UserIndex).pos.Map = 19 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 7 Then Exit Sub
        If UserList(UserIndex).flags.LeMandaronDuelo = False Then
            Call SendData(ToIndex, UserIndex, 0, "||Nadie te ofreciò duelo." & FONTTYPE_INFO)
            Exit Sub
        Else

            If UserList(UserIndex).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas muerto!!." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).pos.Map = 66 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas en carcel!!. " & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(UserIndex).Stats.GLD < val(dMoney) Then
                Call SendData(ToIndex, UserIndex, 0, "||No tenes " & PonerPuntos(val(dMoney)) & " monedas de oro para aceptar el duelo." & FONTTYPE_INFO)
                Exit Sub
            End If

            If MapInfo(val(dMap)).NumUsers = 2 Then
                Call SendData(ToIndex, UserIndex, 0, "||Ya hay un duelo en curso, esperà que termine." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)).flags.Muerto Then
                Call SendData(ToIndex, UserIndex, 0, "||El usuario està muerto." & FONTTYPE_INFO)
                Exit Sub
            End If

            If UserList(NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)).pos.Map = 66 Then
                Call SendData(ToIndex, UserIndex, 0, "||Esta en carcel el otro usuario!!. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)).Stats.GLD < val(dMoney) Then
                Call SendData(ToIndex, UserIndex, 0, "||El usuario no tiene el oro suficiente para hacer el duelo." & FONTTYPE_INFO)
                Exit Sub
            End If

            If NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo) = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||El usuario que te mandò duelo, està offline." & FONTTYPE_INFO)
                Exit Sub
            End If

        End If

        Dim el As Integer
        el = NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)

        UserList(el).flags.LeMandaronDuelo = False
        UserList(el).flags.Endueloo = True
        UserList(UserIndex).flags.LeMandaronDuelo = False
        UserList(UserIndex).flags.Endueloo = True
        UserList(el).flags.DueliandoContra = UserList(UserIndex).name
        UserList(UserIndex).flags.DueliandoContra = UserList(el).name
        SendData ToAll, UserIndex, 0, "||" & UserList(UserIndex).name & " y " & UserList(NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)).name & " van a combatir en un duelo por " & PonerPuntos(val(dMoney)) & " monedas de oro." & FONTTYPE_TALK
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(dMoney)
        UserList(el).Stats.GLD = UserList(el).Stats.GLD - val(dMoney)
        UserList(el).flags.ViejaPos.Map = UserList(el).pos.Map
        UserList(el).flags.ViejaPos.X = UserList(el).pos.X
        UserList(el).flags.ViejaPos.Y = UserList(el).pos.Y
        Call WarpUserChar(el, 5, 64, 38, True)

        UserList(UserIndex).flags.ViejaPos.Map = UserList(UserIndex).pos.Map
        UserList(UserIndex).flags.ViejaPos.X = UserList(UserIndex).pos.X
        UserList(UserIndex).flags.ViejaPos.Y = UserList(UserIndex).pos.Y
        Call WarpUserChar(UserIndex, 5, 40, 56, True)
        Call SendUserStatsBox(UserIndex)
        Call SendUserStatsBox(el)
        Exit Sub
    End If

    If UCase$(rdata) = "/NOBLEZA" Then
        Call SendData(ToIndex, UserIndex, 0, "NBL")
        Exit Sub
    End If

    If UCase$(rdata) = "/SALIRPAREJA" Then
        If MapInfo(7).NumUsers = 2 And UserList(UserIndex).flags.EnPareja = True Then    'mapa de duelos 2vs2
            Call WarpUserChar(Pareja.Jugador1, 1, 50, 50)
            Call WarpUserChar(Pareja.Jugador2, 1, 50, 62)
            UserList(Pareja.Jugador1).Stats.GLD = UserList(Pareja.Jugador1).Stats.GLD - 350000
            UserList(Pareja.Jugador2).Stats.GLD = UserList(Pareja.Jugador2).Stats.GLD - 350000
            Call SendData(ToAll, 0, 0, "||Pareja > " & UserList(Pareja.Jugador1).name & " y " & UserList(Pareja.Jugador2).name & " abandonaron el duelo 2vs2." & FONTTYPE_GUILD)
            UserList(Pareja.Jugador1).flags.EnPareja = False
            UserList(Pareja.Jugador1).flags.EsperaPareja = False
            UserList(Pareja.Jugador1).flags.SuPareja = 0
            UserList(Pareja.Jugador2).flags.EnPareja = False
            UserList(Pareja.Jugador2).flags.EsperaPareja = False
            UserList(Pareja.Jugador2).flags.SuPareja = 0
            HayPareja = False
            Exit Sub
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No puedes utilizar este comando" & FONTTYPE_INFO)
            Exit Sub
        End If
    End If

    If UCase$(rdata) = "/PENAS" Then
        If UserList(UserIndex).Stats.Advertencias = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No tenés ninguna advertencia." & FONTTYPE_GUILD)
            Exit Sub
        End If
        If UserList(UserIndex).Stats.Advertencias <> 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usted tiene un total de " & UserList(UserIndex).Stats.Advertencias & " Advertencias. Te quedan " & (3 - UserList(UserIndex).Stats.Advertencias) & " más y serás baneado si te comportas mal.!" & FONTTYPE_GUILD)
            Exit Sub
        End If
    End If

    If UCase$(rdata) = "/CIUDADPREMIUM" Then
        If UserList(UserIndex).Stats.PuntosDonador = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No sos usuario PREMIUM. Si deseas ser Premium clikea el botón (PREMIUM) en la pantalla principal." & FONTTYPE_FIGHT)
            Exit Sub
        Else
            Call WarpUserChar(UserIndex, 218, 50, 50, True)
            Call SendData(ToIndex, UserIndex, 0, "||Has ingresado a la ciudad exclusiva PREMIUM." & FONTTYPE_INFO)
            Exit Sub
        End If
        Exit Sub
    End If

    If UCase$(rdata) = "/FORTALEZA" Then
        If Not UserList(UserIndex).GuildInfo.GuildName = CastilloNorte And Not UserList(UserIndex).GuildInfo.GuildName = CastilloSur Then Exit Sub
        Call WarpUserChar(UserIndex, 206, 51, 55, True)
        Call SendData(ToIndex, UserIndex, 0, "||Has ingresado a la fortaleza de tu clan." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(rdata) = "/CASTILLOS" Then
        Call SendData(ToIndex, UserIndex, 0, "||El castillo Norte está conquistado por el clan " & CastilloNorte & "." & FONTTYPE_SERVER)
        Call SendData(ToIndex, UserIndex, 0, "||El castillo Sur está conquistado por el clan " & CastilloSur & "." & FONTTYPE_SERVER)
        If CastilloSur = CastilloNorte Then
            Call SendData(ToIndex, UserIndex, 0, "||La fortaleza le pertenece al clan " & CastilloSur & "." & FONTTYPE_SERVER)
        End If
        Exit Sub
    End If

    If UCase$(rdata) = "/CASTILLONORTE" Then
        If Not UserList(UserIndex).GuildInfo.GuildName = CastilloNorte Then Exit Sub
        Call WarpUserChar(UserIndex, 213, 50, RandomNumber(22, 26), True)
        Call SendData(ToIndex, UserIndex, 0, "||Has sido transportado." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(rdata) = "/CASTILLOSUR" Then
        If Not UserList(UserIndex).GuildInfo.GuildName = CastilloSur Then Exit Sub
        Call WarpUserChar(UserIndex, 215, 50, RandomNumber(22, 26), True)
        Call SendData(ToIndex, UserIndex, 0, "||Has sido transportado." & FONTTYPE_INFO)
        Exit Sub
    End If

    'Información de los objetos
    If UCase$(Left$(rdata, 3)) = "IPX" Then
        rdata = Right$(rdata, Len(rdata) - 3)

        If val(rdata) > 0 And val(rdata) < UBound(PremiosList) + 1 Then
            Call SendData(ToIndex, UserIndex, 0, "INF" _
                                                 & PremiosList(val(rdata)).ObjRequiere & "," _
                                                 & PremiosList(rdata).ObjDescripcion & "," _
                                                 & UserList(UserIndex).Faccion.Quests & "," _
                                                 & ObjData(PremiosList(rdata).ObjIndexP).GrhIndex & "," _
                                                 & ObjData(PremiosList(rdata).ObjIndexP).MinDef & "," _
                                                 & ObjData(PremiosList(rdata).ObjIndexP).MaxDef & "," _
                                                 & ObjData(PremiosList(rdata).ObjIndexP).MinHit & "," _
                                                 & ObjData(PremiosList(rdata).ObjIndexP).MaxHit)
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 2)) = "AQ" Then
        rdata = Right$(rdata, Len(rdata) - 2)
        rdata = ReadField(1, rdata, 44)

        If rdata < 0 Or rdata > MaxQuest Then Exit Sub

        If UserList(UserIndex).Quest.Index > 0 Then Exit Sub

        If UserList(UserIndex).flags.Muerto > 0 Then Exit Sub

        If Quest(rdata).Premium <> 0 Then
            If UserList(UserIndex).Stats.PuntosDonador = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Quest exclusiva para usuarios Premium." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'acepta la quest

        UserList(UserIndex).Quest.Index = rdata
        UserList(UserIndex).Quest.NPCs = Quest(rdata).NPCs
        UserList(UserIndex).Quest.Users = Quest(rdata).Users
        UserList(UserIndex).Quest.IndexNPC = Quest(rdata).iNPCs

        SendData ToIndex, UserIndex, 0, "||Aceptaste la quest!!" & FONTTYPE_GUILD
        SendData ToIndex, UserIndex, 0, "||Para saber información sobre la misión escribe /INFOQUEST." & FONTTYPE_FENIX
        Exit Sub
    End If
    If UCase$(Left$(rdata, 10)) = "/INFOQUEST" Then
        Dim tmpName As String        'leemos el nombre del npc ;)
        If UserList(UserIndex).Quest.Index = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No estás haciendo ninguna quest." & FONTTYPE_FIGHT)
            Exit Sub
        End If

        If UserList(UserIndex).Quest.NPCs > 0 Then
            tmpName = GetVar(App.path & "\Dat\NPCs-HOSTILES.dat", "NPC" & UserList(UserIndex).Quest.IndexNPC, "Name")
            Call SendData(ToIndex, UserIndex, 0, "||Restan matar " & UserList(UserIndex).Quest.NPCs & " " & tmpName & " de un total de " & Quest(UserList(UserIndex).Quest.Index).NPCs & " " & tmpName & " y " & UserList(UserIndex).Quest.Users & " usuarios de un total de " & Quest(UserList(UserIndex).Quest.Index).Users & " usuarios para completar la misión." & FONTTYPE_FENIX)
            Exit Sub
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Restan matar " & UserList(UserIndex).Quest.Users & " de un total de " & Quest(UserList(UserIndex).Quest.Index).Users & " usuarios para completar la misión." & FONTTYPE_FENIX)
            Exit Sub
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/BETA " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        Call LogBalance(rdata, UserList(UserIndex).name, UserList(UserIndex).ip)
        Call SendData(ToIndex, UserIndex, 0, "||Mensaje enviado." & FONTTYPE_INFO)
        Call SendData(ToAdmins, 0, 0, "||Nuevo mensaje de BETA, Chequear URGENTE!! - " & rdata & " (" & UserList(UserIndex).name & ")." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 2)) = "IQ" Then    'Info quest
        rdata = Right$(rdata, Len(rdata) - 2)
        If rdata < 0 Or rdata > MaxQuest Then Exit Sub

        ' If UserList(UserIndex).Quest.Index = 0 Then Exit Sub


        If Quest(rdata).iNPCs > 0 Then
            tmpName = GetVar(App.path & "\Dat\NPCs-HOSTILES.dat", "NPC" & Quest(rdata).iNPCs, "Name") & ","
        Else
            tmpName = ","
        End If

        'send info quest !

        SendData ToIndex, UserIndex, 0, "RIQ" & rdata & "," & Quest(rdata).NPCs & "," & tmpName & Quest(rdata).Users & "," & Quest(rdata).Recompense
        Exit Sub
    End If

    'Requerimientos de los objetos
    If UCase$(Left$(rdata, 3)) = "SPX" Then
        rdata = Right$(rdata, Len(rdata) - 3)
        Dim Premio As Obj

        If val(rdata) > 0 And val(rdata) < UBound(PremiosList) + 1 Then

            Premio.Amount = 1
            Premio.OBJIndex = PremiosList(val(rdata)).ObjIndexP

        End If

        If PremiosList(val(rdata)).ObjPremium <> 0 Then
            If UserList(UserIndex).Stats.PuntosDonador = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Este objeto es exclusivo para usuarios Premium." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si no tiene los puntos necesarios
        If UserList(UserIndex).Faccion.Quests < PremiosList(val(rdata)).ObjRequiere Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes suficientes puntos para este objeto." & FONTTYPE_INFO)
            Exit Sub
        End If

        'Si no tenemoss lugar lo tiramos al piso
        'If Not MeterItemEnInventario(UserIndex, Premio) Then
        '   Call SendData(ToIndex, UserIndex, 0, "||No puedo cargar mas objetos." & FONTTYPE_INFO)
        'Exit Sub
        'End If

        'Metemos en inventario
        Call MeterItemEnInventario(UserIndex, Premio)
        Call UpdateUserInv(True, UserIndex, 0)

        'Avisamos por consola
        Call SendData(ToIndex, UserIndex, 0, "||Has obtenido: " & ObjData(Premio.OBJIndex).name & " (Cantidad: " & Premio.Amount & ")" & FONTTYPE_GUILD)

        'Restamos & actualizams
        UserList(UserIndex).Faccion.Quests = UserList(UserIndex).Faccion.Quests - PremiosList(val(rdata)).ObjRequiere
        Call SendUserStatsBox(UserIndex)
        Exit Sub
    End If
    'Dylan.- Sistema de Premios


    If UCase$(rdata) = "/DONACIONES" Then
        'abrimos el form.
        Dim Premios As Integer, SX As String

        SX = "PRM" & UBound(PremiosList) & ","

        For Premios = 1 To UBound(PremiosList)
            SX = SX & ObjData(PremiosList(Premios).ObjIndexP).name & ","
        Next Premios

        Call SendData(ToIndex, UserIndex, 0, SX & UserList(UserIndex).Faccion.Quests)
        Exit Sub
    End If

    If UCase$(rdata) = "/GUERRA" Then
        EntrarGuerra UserIndex
        Exit Sub
    End If

    If UCase$(rdata) = "/REGRESAR" Then
        If MapInfo(UserList(UserIndex).pos.Map).Pk = False Or UserList(UserIndex).pos.Map = 66 Or UserList(UserIndex).pos.Map = 60 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 22 Or UserList(UserIndex).pos.Map = 14 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes regresar a tu ciudad desde esta ubicación." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Muerto = 0 Then Call UserDie(UserIndex)
        WarpUserChar UserIndex, 1, 50, 50, True
        Call SendData(ToIndex, UserIndex, 0, "||Has regresado a Althalos." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UCase$(rdata) = "/SALA1" Then
        If UserList(UserIndex).pos.Map <> 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes estar en Althalos para Jugar" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.DueLeanDo = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya estas en la sala de Duelos!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Encarcelado = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas en la Carcel !!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If EsNewbie(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes ser mayor de nivel 13 para ingresar!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If MapInfo(3).NumUsers >= 2 Then
            Call SendData(ToIndex, UserIndex, 0, "||Sala Llena." & FONTTYPE_FENIX)
            Exit Sub
        End If

        If MapInfo(3).NumUsers = 0 Then
            Call SendData(ToAll, 0, 0, "||Sala 1>" & UserList(UserIndex).name & " espera contrincante en la sala de Duelos!" & FONTTYPE_FENIX)
            Call WarpUserChar(UserIndex, 3, 59, 43, True)
            Call SendData(ToIndex, UserIndex, 0, "||Has sido llevado a la sala de Duelos! cuando quieras salir, teclea: /SALIRDUELO" & FONTTYPE_INFO)
            UserList(UserIndex).flags.DueLeanDo = True
            Exit Sub
        Else
            If MapInfo(3).NumUsers = 1 Then
                Call WarpUserChar(UserIndex, 3, 45, 51, True)
                Call SendData(ToIndex, UserIndex, 0, "||Has sido llevado a la sala de Duelos! cuando quieras salir, teclea: /SALIRDUELO" & FONTTYPE_INFO)
                UserList(UserIndex).flags.DueLeanDo = True
                Call SendData(ToAll, 0, 0, "||Sala 1> Comenzara la Batalla!" & FONTTYPE_FENIX)
                Exit Sub
            End If
        End If
    End If

    If UCase$(rdata) = "/SALA2" Then
        If UserList(UserIndex).pos.Map <> 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes estar en Althalos para Jugar" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.DueLeanDo = True Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya estas en la sala de Duelos!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Encarcelado = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas en la Carcel !!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas muerto!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If EsNewbie(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes ser mayor de nivel 13 para ingresar!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If MapInfo(4).NumUsers >= 2 Then
            Call SendData(ToIndex, UserIndex, 0, "||Sala Llena." & FONTTYPE_FENIX)
            Exit Sub
        End If

        If MapInfo(4).NumUsers = 0 Then
            Call SendData(ToAll, 0, 0, "||Sala 2>" & UserList(UserIndex).name & " espera contrincante en la sala de Duelos!" & FONTTYPE_FENIX)
            Call WarpUserChar(UserIndex, 4, 59, 43, True)
            Call SendData(ToIndex, UserIndex, 0, "||Has sido llevado a la sala de Duelos! cuando quieras salir, teclea: /SALIRDUELO" & FONTTYPE_INFO)
            UserList(UserIndex).flags.DueLeanDo = True
            Exit Sub
        Else
            If MapInfo(4).NumUsers = 1 Then
                Call WarpUserChar(UserIndex, 4, 45, 51, True)
                Call SendData(ToIndex, UserIndex, 0, "||Has sido llevado a la sala de Duelos! cuando quieras salir, teclea: /SALIRDUELO" & FONTTYPE_INFO)
                UserList(UserIndex).flags.DueLeanDo = True
                Call SendData(ToAll, 0, 0, "||Sala 2> Comenzara la Batalla!" & FONTTYPE_FENIX)
                Exit Sub
            End If
        End If
    End If


    'sala de duelos
    'Informacion de Salas
    If UCase$(rdata) = "/INFSALAS" Then
        If MapInfo(3).NumUsers = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Sala 1. Disponible" & FONTTYPE_INFO)
        End If
        If MapInfo(3).NumUsers = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Sala 1. Esperando Oponente" & FONTTYPE_INFO)
        End If
        If MapInfo(3).NumUsers = 2 Then
            Call SendData(ToIndex, UserIndex, 0, "||Sala 1. Llena" & FONTTYPE_INFO)
        End If
        If MapInfo(4).NumUsers = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Sala 2. Disponible" & FONTTYPE_INFO)
        End If
        If MapInfo(4).NumUsers = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Sala 2. Esperando Oponente" & FONTTYPE_INFO)
        End If
        If MapInfo(4).NumUsers = 2 Then
            Call SendData(ToIndex, UserIndex, 0, "||Sala 2. Llena" & FONTTYPE_INFO)
        End If
        Exit Sub
    End If

    If UCase$(rdata) = "/CANCELARDESAFIO" Then
        If UserList(UserIndex).flags.Esperando = True Then
            UserList(UserIndex).flags.Esperando = False
            Call WarpUserChar(DesaFiante(1), 1, 50, 50, True)
            DesaFiante(1) = 0
            DeFenZas = 0
            Call SendData(ToAll, UserIndex, 0, "||Desafio Cancelado" & FONTTYPE_INFO)
            If DesaFiante(2) <> 0 Then
                If UserList(DesaFiante(2)).flags.Desafiando = True Then
                    Call WarpUserChar(DesaFiante(2), 1, 51, 51, True)
                    DesaFiante(2) = 0
                End If
            End If
            Exit Sub
        End If

        If UserList(UserIndex).flags.Desafiando = True Then
            UserList(UserIndex).flags.Desafiando = False
            Call WarpUserChar(DesaFiante(2), 1, 51, 51, True)
            Call SendData(ToAll, 0, 0, "||" & UserList(DesaFiante(2)).name & " ha abandonado el desafio, " & UserList(DesaFiante(1)).name & " espera su oponente. Escribe /DESAFIAR." & FONTTYPE_DESAFIO)
            DesaFiante(2) = 0
            Exit Sub
        End If
        Exit Sub
    End If

    If UCase$(rdata) = "/SALIRDUELO" Then
        If UserList(UserIndex).flags.DueLeanDo = True Then
            UserList(UserIndex).flags.DueLeanDo = False
            Call SendData(ToIndex, UserIndex, 0, "||Has Abandonado la Sala y Enviado a Althalos" & FONTTYPE_FENIX)
            Call WarpUserChar(UserIndex, 1, 50, 50, True)
        End If
        Exit Sub
    End If

    If UCase$(rdata) = "/PING" Then
        Call SendData(ToIndex, UserIndex, 0, "BUENO")
        Exit Sub
    End If

    If UCase$(Left$(rdata, 8)) = "/CONMSJ " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        If UserList(UserIndex).flags.EsConseCaos = 1 Then
            Call SendData(ToAll, UserIndex, 0, "||Concilio de Lhirius AO del mal > " & rdata & FONTTYPE_CONSEJOCAOS)
            Exit Sub
        End If
        If UserList(UserIndex).flags.EsConseReal = 1 Then
            Call SendData(ToAll, UserIndex, 0, "||Consejo de Lhirius AO del bien > " & rdata & FONTTYPE_CONSEJO)
            Exit Sub
        End If
        Exit Sub
    End If

    If UCase(rdata) = "/ACCEDER" Then
        ' aca le ponen las condiciones a su gusto, puede ser que sea mayor a tal lvl para entrar, que no puedan entrar invis, ni muertos, etc. a su gusto.
        If MapInfo(UserList(UserIndex).pos.Map).Pk = True Or UserList(UserIndex).pos.Map = 66 Or UserList(UserIndex).pos.Map = 60 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 22 Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes entrar al torneo automatico si estas en este mapa." & FONTTYPE_INFO)
            Exit Sub
        End If
        If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then Exit Sub
        If UserList(UserIndex).Stats.ELV <= 40 Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes entrar al torneo automatico si sos menor al nivel 41." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas muerto!." & FONTTYPE_INFO)
            Exit Sub
        End If
        Call Torneos_Entra(UserIndex)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/DEATH" Then
        If MapInfo(UserList(UserIndex).pos.Map).Pk = True Or UserList(UserIndex).pos.Map = 66 Or UserList(UserIndex).pos.Map = 60 Or UserList(UserIndex).pos.Map = 5 Or UserList(UserIndex).pos.Map = 22 Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes entrar al torneo automatico si estas en este mapa." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).Stats.ELV <= 40 Then
            Call SendData(ToIndex, UserIndex, 0, "||No podes entrar al torneo automatico si sos menor al nivel 41." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Estas muerto!." & FONTTYPE_INFO)
            Exit Sub
        End If
        Call death_entra(UserIndex)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 5)) = "/GANE" Then
        If UserList(UserIndex).flags.Death = True Then
            If terminodeat = True Then
                Call WarpUserChar(UserIndex, 1, 50, 50, True)
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + 1000000
                Call SendUserStatsBox(UserIndex)
                Call SendData(ToAll, UserIndex, 0, "||GANADOR DEATHMATCH: " & UserList(UserIndex).name & FONTTYPE_GUILD)
                Call SendData(ToAll, UserIndex, 0, "||PREMIO: 1.000.000 monedas de oro al ganador del DeathMatch." & FONTTYPE_GUILD)
                UserList(UserIndex).flags.Death = False
                terminodeat = False
                deathesp = False
                deathac = False
                Cantidad = 0
            End If
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 12)) = "/MERCENARIO " Then
        rdata = Right$(rdata, Len(rdata) - 12)
        If Not ModoQuest Then Exit Sub
        If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub
        Select Case UCase$(rdata)
            Case "ALIANZA"
                tInt = 1
            Case "HORDA"
                tInt = 2
            Case Else
                Call SendData(ToIndex, UserIndex, 0, "||La estructura del comando es /MERCENARIO ALIANZA o /MERCENARIO HORDA." & FONTTYPE_FENIX)
                Exit Sub
        End Select
        Select Case UserList(UserIndex).Faccion.BandoOriginal
            Case Neutral
                If UserList(UserIndex).Faccion.Bando <> Neutral Then
                    Call SendData(ToIndex, UserIndex, 0, "||Ya eres mercenario para " & ListaBandos(UserList(UserIndex).Faccion.Bando) & "." & FONTTYPE_FENIX)
                    Exit Sub
                End If
            Case Else
                Select Case UserList(UserIndex).Faccion.Bando
                    Case Neutral
                        If tInt = UserList(UserIndex).Faccion.BandoOriginal Then
                            Call SendData(ToIndex, UserIndex, 0, "||" & ListaBandos(tInt) & " no acepta desertores entre sus filas." & FONTTYPE_FENIX)
                            Exit Sub
                        End If

                    Case UserList(UserIndex).Faccion.BandoOriginal
                        Call SendData(ToIndex, UserIndex, 0, "||Ya perteneces a " & ListaBandos(UserList(UserIndex).Faccion.Bando) & ", no puedes ofrecerte como mercenario." & FONTTYPE_FENIX)
                        Exit Sub

                    Case Else
                        Call SendData(ToIndex, UserIndex, 0, "||Ya eres mercenario para " & ListaBandos(UserList(UserIndex).Faccion.Bando) & "." & FONTTYPE_FENIX)
                        Exit Sub
                End Select
        End Select
        Call SendData(ToIndex, UserIndex, 0, "||¡" & ListaBandos(tInt) & " te ha aceptado como un mercenario entre sus filas!" & FONTTYPE_FENIX)
        UserList(UserIndex).Faccion.Bando = tInt
        Call UpdateUserChar(UserIndex)
        Exit Sub
    End If
    If UserList(UserIndex).flags.Quest Then
        If UCase$(Left$(rdata, 3)) = "/M " Then
            rdata = Right$(rdata, Len(rdata) - 3)
            If Len(rdata) = 0 Then Exit Sub
            Select Case UserList(UserIndex).Faccion.Bando
                Case Real
                    tStr = FONTTYPE_ARMADA
                Case Caos
                    tStr = FONTTYPE_CAOS
            End Select
            Call SendData(ToAll, 0, 0, "||" & rdata & tStr)
            Exit Sub
        ElseIf UCase$(rdata) = "/TELEPLOC" Then
            Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
            Exit Sub
        ElseIf UCase$(rdata) = "/TRAMPA" Then
            Call ActivarTrampa(UserIndex)
            Exit Sub
        End If
    End If

    Call HandleDataTWO(UserIndex, rdata)    'comandos gms
    Exit Sub
ErrorHandler:
    If Err.Number = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Comando invalido..." & FONTTYPE_INFO)
    Else
        Call LogErrorUrgente("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).name & " UI:" & UserIndex & " N: " & Err.Number & " D: " & Err.Description)
    End If
End Sub
Sub HandleDataTWO(UserIndex As Integer, ByVal rdata As String)

    On Error GoTo ErrorHandler:

    Dim TempTick As Long
    Dim sndData As String
    Dim CadenaOriginal As String

    Dim LoopC As Integer
    Dim nPos As WorldPos
    Dim tStr As String
    Dim tInt As Integer
    Dim tLong As Long
    Dim numeromail As Integer
    Dim TIndex As Integer
    Dim tName As String
    Dim Clase As Byte
    Dim NumNPC As Integer
    Dim tMessage As String
    Dim i As Integer
    Dim auxind As Integer
    Dim Arg1 As String
    Dim Arg2 As String
    Dim arg3 As String
    Dim Arg4 As String
    Dim Arg5 As Integer
    Dim Arg6 As String
    Dim DummyInt As Integer
    Dim Antes As Boolean
    Dim Ver As String
    Dim encpass As String
    Dim Pass As String
    Dim mapa As Integer
    Dim usercon As String
    Dim nameuser As String
    Dim name As String
    Dim ind
    Dim GMDia As String
    Dim GMMapa As String
    Dim GMPJ As String
    Dim GMMail As String
    Dim GMGM As String
    Dim GMTitulo As String
    Dim GMMensaje As String
    Dim N As Integer
    Dim wpaux As WorldPos
    Dim mifile As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim cliMD5 As String
    Dim UserFile As String
    Dim UserName As String
    UserName = UserList(UserIndex).name
    UserFile = CharPath & UCase$(UserName) & ".chr"
    Dim ClientCRC As String
    Dim ServerSideCRC As Long
    Dim NombreIniChat As String
    Dim cantidadenmapa As Integer
    Dim Prueba1 As Integer
    Dim NpcIndex As Integer
    CadenaOriginal = rdata

    If UserList(UserIndex).flags.Privilegios = 0 Then Exit Sub

    'sistema de soporte dylan.-
    If UCase$(rdata) = "/DAMESOS" Then
        Dim LstU As String

        If SoporteS.Count = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No hay soportes para ver." & FONTTYPE_INFO)
            Exit Sub
        End If

        For i = 1 To SoporteS.Count
            LstU = LstU & "@" & SoporteS.Item(i)
            Debug.Print SoporteS.Item(i)
            DoEvents
        Next i

        LstU = SoporteS.Count & LstU

        LstU = "SHWSOP@" & LstU
        Call SendData(ToIndex, UserIndex, 0, LstU)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 7)) = "/BORSO " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        Call WriteVar(CharPath & UCase$(rdata) & ".chr", "STATS ", "Soporte", "")
        Call WriteVar(CharPath & UCase$(rdata) & ".chr", "STATS ", "Respuesta", "")
        For i = 1 To SoporteS.Count
            If UCase$(SoporteS.Item(i)) = UCase$(rdata) Then
                SoporteS.Remove (i)
                Exit For
            End If
            DoEvents
        Next i
        Call SendData(ToIndex, UserIndex, 0, "||Soporte y respuesta borrados con éxito" & FONTTYPE_INFO)
        Exit Sub
    End If


    If UCase$(Left$(rdata, 7)) = "/SOSDE " Then
        rdata = Right$(rdata, Len(rdata) - 7)

        Dim SosDe As String
        SosDe = GetVar(CharPath & UCase$(rdata) & ".chr", "STATS", "Soporte")


        If Len(SosDe) > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "SOPODE" & SosDe)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Error. Soporte no encontrado" & FONTTYPE_INFO)
        End If

        Exit Sub
    End If

    If UCase$(Left$(rdata, 7)) = "/RESOS " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        Dim Persona, Respuesta As String
        Persona = ReadField$(1, rdata, Asc(";"))    'GetVar(CharPath & UCase$(rdata) & ".chr", "STATS", "Soporte")
        Respuesta = Replace(ReadField$(2, rdata, Asc(";")), Chr$(13) & Chr$(10), Chr(32))
        If Len(Persona) = 0 Or Len(Respuesta) = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Error en la respuesta" & FONTTYPE_INFO)
            Exit Sub
        End If

        Call WriteVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Respuesta", Respuesta)
        Call WriteVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Soporte", GetVar(CharPath & UCase$(Persona) & ".chr", "STATS", "Soporte") & "0k1")


        TIndex = NameIndex(Persona)
        If TIndex > 0 Then
            Call SendData(ToIndex, TIndex, 0, "||Tu soporte ha sido respondido." & FONTTYPE_FENIX)
            Call SendData(ToIndex, TIndex, 0, "TENSO")
        End If

        Call SendData(ToIndex, UserIndex, 0, "||Soporte respondido con éxito" & FONTTYPE_INFO)
        For i = 1 To SoporteS.Count
            Debug.Print SoporteS.Item(1)

            If UCase$(SoporteS.Item(i)) = UCase$(Persona) Then
                SoporteS.Remove (i)
                Exit For
            End If
            DoEvents
        Next i

        Exit Sub
    End If

    'sistema de soporte dylan.-


    If UCase$(Left$(rdata, 4)) = "/GO " Then
        rdata = Right$(rdata, Len(rdata) - 4)
        mapa = val(ReadField(1, rdata, 32))
        If Not MapaValido(mapa) Then Exit Sub
        If UserList(UserIndex).flags.Privilegios = 1 And MapInfo(mapa).Pk Then Exit Sub
        Call WarpUserChar(UserIndex, mapa, 50, 50, True)
        Call SendData(ToIndex, UserIndex, 0, "2B" & UserList(UserIndex).name)
        Call LogGM(UserList(UserIndex).name, "Transporto a " & UserList(UserIndex).name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If

    If UCase$(Left$(rdata, 12)) = "/GUERRAAUTO " Then
        rdata = UCase$(Right$(rdata, Len(rdata) - 12))
        If UCase$(rdata) = "ON" Then
            GuerrasAuto UserIndex, 1
        ElseIf UCase$(rdata) = "OFF" Then
            GuerrasAuto UserIndex, 0
        End If
        Exit Sub
    End If

    If UCase$(rdata) = "/INICIARGUERRA" Then
        IniciarGuerra UserIndex
        Exit Sub
    End If

    If UCase$(rdata) = "/TERMINARGUERRA" Then
        TerminaGuerra "NONE"
        Exit Sub
    End If

    If UCase$(rdata) = "/TELEPLOC" Then
        Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
        Call LogGM(UserList(UserIndex).name, "/TELEPLOC a x:" & UserList(UserIndex).flags.TargetX & " Y:" & UserList(UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).pos.Map, (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If

    If UCase$(Left$(rdata, 13)) = "/DESADVERTIR " Then
        rdata = Right$(rdata, Len(rdata) - 13)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(TIndex).Stats.Advertencias = UserList(TIndex).Stats.Advertencias - 1
        Call SendData(ToAll, UserIndex, 0, "||Advertencias> " & UserList(TIndex).name & " ha sido desadvertido por " & UserList(UserIndex).name & ", con esta ya lleva: " & UserList(TIndex).Stats.Advertencias & " advertencias." & FONTTYPE_FIGHT)
        Call SendData(ToIndex, TIndex, 0, "||Adevertencias> " & UserList(UserIndex).name & " Te ha removido una advertencia." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 10)) = "/ADVERTIR " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(TIndex).Stats.Advertencias = UserList(TIndex).Stats.Advertencias + 1
        Call SendData(ToAll, UserIndex, 0, "||Advertencias> " & UserList(TIndex).name & " ha sido advertido por " & UserList(UserIndex).name & ", con esta ya lleva: " & UserList(TIndex).Stats.Advertencias & " advertencias." & FONTTYPE_FIGHT)
        Call SendData(ToIndex, TIndex, 0, "||Adevertencias> Recuerda que a las 3 advertencias acumuladas, serás baneado. Para ver tus Advertencias escribe /PENAS" & FONTTYPE_FIGHT)
        Call Encarcelar(TIndex, 5)
        If UserList(TIndex).Stats.Advertencias = 3 Then
            UserList(TIndex).flags.Ban = 1
            Call CloseSocket(TIndex)
            UserList(TIndex).Stats.Advertencias = 0
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 5)) = "/SUM " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(UserIndex).flags.Privilegios < UserList(TIndex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(UserIndex).flags.Privilegios = 1 And UserList(TIndex).pos.Map <> UserList(UserIndex).pos.Map Then Exit Sub
        Call SendData(ToIndex, UserIndex, 0, "%Z" & UserList(TIndex).name)
        Call WarpUserChar(TIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y + 1, True)
        Call LogGM(UserList(UserIndex).name, "/SUM " & UserList(TIndex).name & " Map:" & UserList(UserIndex).pos.Map & " X:" & UserList(UserIndex).pos.X & " Y:" & UserList(UserIndex).pos.Y, False)
        Exit Sub
    End If

    If UCase(rdata) = "/CANCELARDEATH" Then
        Call Death_Cancela
        Exit Sub
    End If

    If UCase$(Left$(rdata, 10)) = "/DEATHACT " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        If (CInt(rdata) > 2 And CInt(rdata) < 33) Then Call death_comienza(rdata)
    End If

    If UCase$(Left$(rdata, 5)) = "/IRA " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If ((UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1)) Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(TIndex).flags.AdminInvisible And Not UserList(UserIndex).flags.AdminInvisible Then Call DoAdminInvisible(UserIndex)
        Call WarpUserChar(UserIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X + 1, UserList(TIndex).pos.Y + 1, True)
        Call LogGM(UserList(UserIndex).name, "/IRA " & UserList(TIndex).name & " Mapa:" & UserList(TIndex).pos.Map & " X:" & UserList(TIndex).pos.X & " Y:" & UserList(TIndex).pos.Y, (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If
    If UCase$(Left$(rdata, 11)) = "/SILENCIAR " Then
        rdata = Right$(rdata, Len(rdata) - 11)
        name = ReadField(1, rdata, 32)
        i = val(ReadField(1, rdata, 32))
        name = Right$(rdata, Len(rdata) - (Len(name) + 1))
        TIndex = NameIndex(name)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If i > 15 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puede silenciar al usuario por más de 15min." & FONTTYPE_INFO)
            Exit Sub
        End If
        Call Silenciar(TIndex, i)
        Call SendData(ToIndex, TIndex, 0, "!!ESTIMADO USUARIO: Usted ha sido silenciado momentaneamente, no podra hablar, ni mandar soporte. Gracias. Lhirius AO Staff.")

        Exit Sub
    End If
    If UCase$(rdata) = "/TRABAJANDO" Then
        For LoopC = 1 To LastUser
            If Len(UserList(LoopC).name) > 0 And UserList(LoopC).flags.Trabajando Then
                DummyInt = DummyInt + 1
                tStr = tStr & UserList(LoopC).name & ", "
            End If
        Next
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||Usuarios trabajando: " & tStr & FONTTYPE_INFO)
            Call SendData(ToIndex, UserIndex, 0, "||Número de usuarios trabajando: " & DummyInt & "." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "%)")
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 8)) = "/CARCEL " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        name = ReadField(1, rdata, 32)
        i = val(ReadField(1, rdata, 32))
        name = Right$(rdata, Len(rdata) - (Len(name) + 1))
        TIndex = NameIndex(name)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
            Call SendData(ToIndex, UserIndex, 0, "1B")
            Exit Sub
        End If
        If i > 120 Then
            Call SendData(ToIndex, UserIndex, 0, "1C")
            Exit Sub
        End If
        Call Encarcelar(TIndex, i, UserList(UserIndex).name)
        Exit Sub
    End If
    If UserList(UserIndex).flags.Privilegios < 2 Then Exit Sub
    If UCase$(Left$(rdata, 4)) = "/REM" Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Call LogGM(UserList(UserIndex).name, "Comentario: " & rdata, (UserList(UserIndex).flags.Privilegios = 1))
        Call SendData(ToIndex, UserIndex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 7)) = "/STAFF " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        Call LogGM(UserList(UserIndex).name, "Mensaje a Gms:" & rdata, (UserList(UserIndex).flags.Privilegios = 1))
        If Len(rdata) > 0 Then
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).name & "> " & rdata & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 10)) = "/DARPUNTO " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        TIndex = UserList(UserIndex).flags.TargetUser
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar al Jugador para Darle sus Puntos!" & FONTTYPE_INFO)
            Exit Sub
        End If
        Call SendData(ToAll, 0, 0, "||" & UserList(UserList(UserIndex).flags.TargetUser).name & " gano " & rdata & " puntos de Canje" & FONTTYPE_FENIX)
        UserList(UserList(UserIndex).flags.TargetUser).Faccion.Quests = UserList(UserList(UserIndex).flags.TargetUser).Faccion.Quests + rdata
        Call LogGM(UserList(UserIndex).name, "Puntos de Canje: " & rdata & UserList(TIndex).name & " Map:" & UserList(UserIndex).pos.Map & " X:" & UserList(UserIndex).pos.X & " Y:" & UserList(UserIndex).pos.Y, False)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 5)) = "/HORA" Then
        Call LogGM(UserList(UserIndex).name, "Hora.", (UserList(UserIndex).flags.Privilegios = 1))
        rdata = Right$(rdata, Len(rdata) - 5)
        Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(rdata) = "/ONLINEGM" Then
        For LoopC = 1 To LastUser
            If Len(UserList(LoopC).name) > 0 Then
                If UserList(LoopC).flags.Privilegios > 0 And (UserList(LoopC).flags.Privilegios <= UserList(UserIndex).flags.Privilegios Or UserList(LoopC).flags.AdminInvisible = 0) Then
                    tStr = tStr & UserList(LoopC).name & ", "
                End If
            End If

        Next
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "%P")
        End If
        Exit Sub
    End If
    If UCase$(rdata) = "/GLOBALACT" Then
        GlobalAct = 1
        Call SendData(ToAll, 0, 0, "||El Global fue activado. Para hablar por Global presiona el Numero 8." & FONTTYPE_TALK)
        Exit Sub
    End If
    If UCase$(rdata) = "/GLOBALDES" Then
        GlobalAct = 0
        Call SendData(ToAll, 0, 0, "||El Global fue desactivado." & FONTTYPE_TALK)
        Exit Sub

    End If
    If UCase$(Left$(rdata, 7)) = "/DONDE " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Ubicacion de " & UserList(TIndex).name & ": " & UserList(TIndex).pos.Map & ", " & UserList(TIndex).pos.X & ", " & UserList(TIndex).pos.Y & "." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).name, "/Donde", (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If
    If UCase$(Left$(rdata, 6)) = "/NENE " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        If MapaValido(val(rdata)) Then
            Call SendData(ToIndex, UserIndex, 0, "NENE" & NPCHostiles(val(rdata)))
            Call LogGM(UserList(UserIndex).name, "Numero enemigos en mapa " & rdata, (UserList(UserIndex).flags.Privilegios = 1))
        End If
        Exit Sub
    End If
    If UCase$(rdata) = "/DESCONGELAR" Then
        Call Congela(True)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 9)) = "/VIGILAR " Then
        rdata = Right$(rdata, Len(rdata) - 9)
        TIndex = NameIndex(rdata)
        If TIndex > 0 Then
            If TIndex = UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes vigilarte a ti mismo." & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(TIndex).flags.Privilegios >= UserList(UserIndex).flags.Privilegios Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes vigilar a alguien con igual o mayor jerarquia que tú." & FONTTYPE_INFO)
                Exit Sub
            End If
            If YaVigila(TIndex, UserIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "||Dejaste de vigilar a " & UserList(TIndex).name & "." & FONTTYPE_INFO)
                If Not EsVigilado(TIndex) Then Call SendData(ToIndex, TIndex, 0, "VIG")
                Exit Sub
            End If
            If Not EsVigilado(TIndex) Then Call SendData(ToIndex, TIndex, 0, "VIG")
            Call SendData(ToIndex, UserIndex, 0, "||Estás vigilando a " & UserList(TIndex).name & "." & FONTTYPE_INFO)
            For i = 1 To 10
                If UserList(TIndex).flags.Espiado(i) = 0 Then
                    UserList(TIndex).flags.Espiado(i) = UserIndex
                    Exit For
                End If
            Next
            If i = 11 Then
                Call SendData(ToIndex, UserIndex, 0, "||Demasiados GM's están vigilando a este usuario." & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "1A")
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 7)) = "/VERPC " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios And UserList(UserIndex).flags.AdminInvisible = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(TIndex).flags.Privilegios >= UserList(UserIndex).flags.Privilegios Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes ver la PC de un GM con mayor jerarquia." & FONTTYPE_FIGHT)
            Exit Sub
        End If
        UserList(TIndex).flags.EsperandoLista = UserIndex
        Call SendData(ToIndex, TIndex, 0, "VPRC")
    End If
    If UCase$(Left$(rdata, 7)) = "/TELEP " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        mapa = val(ReadField(2, rdata, 32))
        If Not MapaValido(mapa) Then Exit Sub
        name = ReadField(1, rdata, 32)
        If Len(name) = 0 Then Exit Sub
        If UCase$(name) <> "YO" Then
            If UserList(UserIndex).flags.Privilegios = 1 Then
                Exit Sub
            End If
            TIndex = NameIndex(name)
        Else
            TIndex = UserIndex
        End If
        X = val(ReadField(3, rdata, 32))
        Y = val(ReadField(4, rdata, 32))
        If Not InMapBounds(X, Y) Then Exit Sub
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios And UserList(UserIndex).flags.AdminInvisible = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        Call WarpUserChar(TIndex, mapa, X, Y, True)
        Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).name & " te ha transportado." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).name, "Transporto a " & UserList(TIndex).name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If
    If UCase$(Left$(rdata, 4)) = "/GO " Then
        rdata = Right$(rdata, Len(rdata) - 4)
        mapa = val(ReadField(1, rdata, 32))
        If Not MapaValido(mapa) Then Exit Sub
        Call WarpUserChar(UserIndex, mapa, 50, 50, True)
        Call SendData(ToIndex, UserIndex, 0, "2B" & UserList(UserIndex).name)
        Call LogGM(UserList(UserIndex).name, "Transporto a " & UserList(UserIndex).name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If
    If UCase$(rdata) = "/OMAP" Then
        For LoopC = 1 To MapInfo(UserList(UserIndex).pos.Map).NumUsers
            If UserList(MapInfo(UserList(UserIndex).pos.Map).UserIndex(LoopC)).flags.Privilegios <= UserList(UserIndex).flags.Privilegios Then
                tStr = tStr & UserList(MapInfo(UserList(UserIndex).pos.Map).UserIndex(LoopC)).name & ","
            End If
        Next
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 1)
            Call SendData(ToIndex, UserIndex, 0, "||Usuarios en este mapa: " & tStr & "." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "%R")
        End If
        Exit Sub
    End If

    ' lo limite a torneode 5 rondas de 32 participantes, pero si quierren de mas participantes, cambien el < 6 por un numero mayor.
    If UCase$(Left$(rdata, 12)) = "/AUTOTORNEO " Then
        rdata = Right$(rdata, Len(rdata) - 12)
        Dim Torneos As Integer
        Torneos = CInt(rdata)
        If (Torneos > 0 And Torneos < 6) Then Call Torneos_Inicia(UserIndex, Torneos)
    End If
    If UCase(rdata) = "/CANCELARTORNEO" Then
        Call Rondas_Cancela
        Exit Sub
    End If
    If UCase$(rdata) = "/PANELGM" Then
        Call SendData(ToIndex, UserIndex, 0, "PGM" & UserList(UserIndex).flags.Privilegios)
        Exit Sub
    End If
    If UCase$(rdata) = "/CMAP" Then
        If MapInfo(UserList(UserIndex).pos.Map).NumUsers Then
            Call SendData(ToIndex, UserIndex, 0, "||Hay " & MapInfo(UserList(UserIndex).pos.Map).NumUsers & " usuarios en este mapa." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "%R")
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 8)) = "/TORNEO " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        PTorneo = val(ReadField(1, rdata, 32))
        If entorneo = 0 Then
            entorneo = 1
            If FileExist(App.path & "/logs/torneo.log", vbNormal) Then Kill (App.path & "/logs/torneo.log")
            Call SendData(ToIndex, UserIndex, 0, "||Has activado el torneo" & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "||Torneo para " & PTorneo & " jugadores! /PARTICIPAR para entrar al Torneo." & FONTTYPE_FENIX)
        Else
            entorneo = 0
            Call SendData(ToIndex, UserIndex, 0, "||Has desactivado el torneo" & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 10)) = "/VERTORNEO" Then
        Dim jugadores As Integer
        Dim jugador As Integer
        Dim stri As String
        stri = ""
        jugadores = val(GetVar(App.path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD"))
        For jugador = 1 To jugadores
            stri = stri & GetVar(App.path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugador) & ","
        Next
        Call SendData(ToIndex, UserIndex, 0, "||Quieren participar: " & stri & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(rdata) = "/NTORNEO" Then
        If AutoTorneo = False Then
            Call SendData(ToAll, 0, 0, "||Torneo Activado." & FONTTYPE_FENIX)
            Call SendData(ToAll, 0, 0, "||Para Ingresar, tipea /ENTRAR" & FONTTYPE_FENIX)
            AutoTorneo = True
            Exit Sub
        Else
            Call SendData(ToAll, 0, 0, "||Torneo Desactivado." & FONTTYPE_WARNING)
            AutoTorneo = False
            PosicionTorneo = 0
            Exit Sub
        End If
    End If
    If UCase$(rdata) = "/MTORNEO" Then
        If MegaTorneo = False Then
            Call SendData(ToAll, 0, 0, "||Torneo 2 vs 2 Activado." & FONTTYPE_FENIX)
            Call SendData(ToAll, 0, 0, "||Para Ingresar, tipea /INSCRIBIR TUPAREJA" & FONTTYPE_FENIX)
            Call SendData(ToAll, 0, 0, "||En TUPAREJA va el Nombre de tu Pareja. Acordate de colocar bien el Nombre. Si tiene Espacios, agregales un Signo +" & FONTTYPE_FENIX)
            MegaTorneo = True
            Exit Sub
        Else
            Call SendData(ToAll, 0, 0, "||Torneo Desactivado." & FONTTYPE_WARNING)
            MegaTorneo = False
            iPosicionTorneo = 0
            Exit Sub
        End If
    End If
    If UCase$(rdata) = "/PROMEDIO" Then
        Dim Promedio
        Promedio = Round(UserList(UserIndex).Stats.MaxHP / UserList(UserIndex).Stats.ELV, 2)
        Call SendData(ToIndex, UserIndex, 0, "||El Promedio de vida de tu Personaje es de " & Promedio & FONTTYPE_TALK)
        Exit Sub
    End If

    If UCase$(rdata) = "/INVISIBLE" Then
        Call DoAdminInvisible(UserIndex)
        Call LogGM(UserList(UserIndex).name, "/INVISIBLE", (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If
    If UCase$(Left$(rdata, 6)) = "/INFO " Then
        Call LogGM(UserList(UserIndex).name, rdata, False)
        rdata = Right$(rdata, Len(rdata) - 6)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        SendUserSTAtsTxt UserIndex, TIndex
        Call SendData(ToIndex, UserIndex, 0, "||Mail: " & UserList(TIndex).Email & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Ip: " & UserList(TIndex).ip & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        tStr = ""
        For LoopC = 1 To LastUser
            If UserList(LoopC).ip = rdata And Len(UserList(LoopC).name) > 0 And UserList(LoopC).flags.UserLogged Then
                If (UserList(UserIndex).flags.Privilegios > 0 And UserList(LoopC).flags.Privilegios = 0) Or (UserList(UserIndex).flags.Privilegios = 3) Then
                    tStr = tStr & UserList(LoopC).name & ", "
                End If
            End If
        Next
        Call SendData(ToIndex, UserIndex, 0, "||Los personajes con ip " & rdata & " son: " & tStr & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 10)) = "/MAILNICK " Then
        rdata = Right$(rdata, Len(rdata) - 10)
        tStr = ""
        For LoopC = 1 To LastUser
            If UCase$(UserList(LoopC).ip) = UCase$(rdata) And Len(UserList(LoopC).name) > 0 And UserList(LoopC).flags.UserLogged Then
                If (UserList(UserIndex).flags.Privilegios > 0 And UserList(LoopC).flags.Privilegios = 0) Or (UserList(UserIndex).flags.Privilegios = 3) Then
                    tStr = tStr & UserList(LoopC).name & ", "
                End If
            End If
        Next
        Call SendData(ToIndex, UserIndex, 0, "||Los personajes con mail " & rdata & " son: " & tStr & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 5)) = "/INV " Then
        Call LogGM(UserList(UserIndex).name, rdata, False)
        rdata = Right$(rdata, Len(rdata) - 5)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        SendUserInvTxt UserIndex, TIndex
        Exit Sub
    End If
    If UCase$(Left$(rdata, 8)) = "/SKILLS " Then
        Call LogGM(UserList(UserIndex).name, rdata, False)
        rdata = Right$(rdata, Len(rdata) - 8)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        SendUserSkillsTxt UserIndex, TIndex
        Exit Sub
    End If
    If UCase$(Left$(rdata, 5)) = "/ATR " Then
        Call LogGM(UserList(UserIndex).name, rdata, False)
        rdata = Right$(rdata, Len(rdata) - 5)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Atributos de " & UserList(TIndex).name & FONTTYPE_INFO)
        For i = 1 To NUMATRIBUTOS
            Call SendData(ToIndex, UserIndex, 0, "|| " & AtributosNames(i) & " = " & UserList(TIndex).Stats.UserAtributosBackUP(1) & FONTTYPE_INFO)
        Next
        Exit Sub
    End If
    If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
        rdata = Right$(rdata, Len(rdata) - 9)
        name = rdata
        If UCase$(name) <> "YO" Then
            TIndex = NameIndex(name)
        Else
            TIndex = UserIndex
        End If
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        Call RevivirUsuarioNPC(TIndex)
        Call SendData(ToIndex, TIndex, 0, "%T" & UserList(UserIndex).name)
        Call LogGM(UserList(UserIndex).name, "Resucito a " & UserList(TIndex).name, False)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/ITEM " Then
        Dim ET As Obj
        rdata = Right$(rdata, Len(rdata) - 6)
        ET.OBJIndex = val(ReadField(1, rdata, Asc(" ")))
        ET.Amount = val(ReadField(2, rdata, Asc(" ")))
        If ET.Amount <= 0 Then ET.Amount = 1
        If ET.OBJIndex < 1 Or ET.OBJIndex > NumObjDatas Then Exit Sub
        If ET.Amount > MAX_INVENTORY_OBJS Then Exit Sub
        If Not MeterItemEnInventario(UserIndex, ET) Then Call TirarItemAlPiso(UserList(UserIndex).pos, ET)
        Call LogGM(UserList(UserIndex).name, "Creo objeto:" & ObjData(ET.OBJIndex).name & " (" & ET.Amount & ")", False)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 9)) = "/DEBUGCVC" Then
        If CVC.OcupadoCVC = True Then
            CVC.CantidadDeParticipantes2 = 0
            CVC.CantidadQueParticipa = 0
            CVC.NombreClan1 = ""
            CVC.NombreClan2 = ""
            CVC.OcupadoCVC = False
            Call SendData(ToIndex, UserIndex, 0, "||CVC DEBUGUEADO." & FONTTYPE_FENIX)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||CVC NO ESTÁ BUGUEADO." & FONTTYPE_FIGHT)
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 13)) = "/DEBUGPAREJA" Then
        If HayPareja = True Then
            HayPareja = False
            Call SendData(ToIndex, UserIndex, 0, "||PAREJA DEBUGUEADA." & FONTTYPE_FENIX)
            Exit Sub
        Else
            Call SendData(ToIndex, UserIndex, 0, "||PAREJA NO ESTÁ BUGUEADA." & FONTTYPE_FIGHT)
            Exit Sub
        End If
    End If

    If UCase$(Left$(rdata, 9)) = "/PREMIUM " Then    'DONACION NICK@CANTIDAD
        rdata = Right$(rdata, Len(rdata) - 9)
        name = ReadField(1, rdata, Asc("@"))
        Arg1 = ReadField(2, rdata, Asc("@"))
        Dim CanjesParciales As Long

        If Arg1 > 4 Or Arg1 < 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||El rango de entrega de premium es 1 a 4, considerando: 1: (15 Dias + 30 puntos de canjes) - 2: (30 Dias + 70 puntos de canjes) - 3: (60 Dias + 100 puntos de canjes) - 4: (90 Dias + 150 puntos de canjes)." & FONTTYPE_FENIX)
            Exit Sub
        End If

        If NameIndex(name) = 0 Then
            If FileExist(App.path & "\CHARFILE\" & UCase$(name) & ".chr", vbArchive) Then
                If Arg1 = 1 Then
                    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "STATS", "PuntosDonador", 1)    'guardo los dias premium
                    CanjesParciales = val(GetVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "FACCIONES", "Quests"))
                    CanjesParciales = CanjesParciales + 30
                    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "FACCIONES", "Quests", str$(CanjesParciales))    'guardo los canjes
                    Call SendData(ToIndex, UserIndex, 0, "||Donación entregada con éxito." & FONTTYPE_FENIX)
                    GuardarPremium False, UCase$(name), 15
                    Exit Sub
                ElseIf Arg1 = 2 Then
                    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "STATS", "PuntosDonador", 1)    'guardo los dias premium
                    CanjesParciales = val(GetVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "FACCIONES", "Quests"))
                    CanjesParciales = CanjesParciales + 70
                    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "FACCIONES", "Quests", str$(CanjesParciales))    'guardo los canjes
                    Call SendData(ToIndex, UserIndex, 0, "||Donación entregada con éxito." & FONTTYPE_FENIX)
                    GuardarPremium False, UCase$(name), 30
                    Exit Sub
                ElseIf Arg1 = 3 Then
                    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "STATS", "PuntosDonador", 1)    'guardo los dias premium
                    CanjesParciales = val(GetVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "FACCIONES", "Quests"))
                    CanjesParciales = CanjesParciales + 100
                    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "FACCIONES", "Quests", str$(CanjesParciales))    'guardo los canjes
                    Call SendData(ToIndex, UserIndex, 0, "||Donación entregada con éxito." & FONTTYPE_FENIX)
                    GuardarPremium False, UCase$(name), 60
                    Exit Sub
                ElseIf Arg1 = 3 Then
                    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "STATS", "PuntosDonador", 1)    'guardo los dias premium
                    CanjesParciales = val(GetVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "FACCIONES", "Quests"))
                    CanjesParciales = CanjesParciales + 150
                    Call WriteVar(App.path & "\CHARFILE\" & UCase$(name) & ".chr", "FACCIONES", "Quests", str$(CanjesParciales))    'guardo los canjes
                    Call SendData(ToIndex, UserIndex, 0, "||Donación entregada con éxito." & FONTTYPE_FENIX)
                    GuardarPremium False, UCase$(name), 90
                    Exit Sub
                End If
                'guardar el registro de nick premium
            Else
                Call SendData(ToIndex, UserIndex, 0, "||El usuario NO existe, no es posible entregar la donación." & FONTTYPE_FENIX)
                Exit Sub
            End If
        Else
            If Arg1 = 1 Then
                UserList(NameIndex(name)).Stats.PuntosDonador = 1
                UserList(NameIndex(name)).Faccion.Quests = UserList(NameIndex(name)).Faccion.Quests + 30
                Call SendData(ToIndex, NameIndex(name), 0, "||¡Gracias por colaborar con el servidor! A partir de hoy sos usuario PREMIUM y vence según el plan que has abonado. Para disfrutar de los beneficios PREMIUM escribe /CIUDADPREMIUM. ~255~255~0~1~0")
                Call SendData(ToIndex, UserIndex, 0, "||Donación entregada con éxito." & FONTTYPE_FENIX)
                GuardarPremium False, UCase$(name), 15
                Exit Sub
            ElseIf Arg1 = 2 Then
                UserList(NameIndex(name)).Stats.PuntosDonador = 1
                UserList(NameIndex(name)).Faccion.Quests = UserList(NameIndex(name)).Faccion.Quests + 70
                Call SendData(ToIndex, NameIndex(name), 0, "||¡Gracias por colaborar con el servidor! A partir de hoy sos usuario PREMIUM y vence según el plan que has abonado. Para disfrutar de los beneficios PREMIUM escribe /CIUDADPREMIUM. ~255~255~0~1~0")
                Call SendData(ToIndex, UserIndex, 0, "||Donación entregada con éxito." & FONTTYPE_FENIX)
                GuardarPremium False, UCase$(name), 30
                Exit Sub
            ElseIf Arg1 = 3 Then
                UserList(NameIndex(name)).Stats.PuntosDonador = 1
                UserList(NameIndex(name)).Faccion.Quests = UserList(NameIndex(name)).Faccion.Quests + 100
                Call SendData(ToIndex, NameIndex(name), 0, "||¡Gracias por colaborar con el servidor! A partir de hoy sos usuario PREMIUM y vence según el plan que has abonado. Para disfrutar de los beneficios PREMIUM escribe /CIUDADPREMIUM. ~255~255~0~1~0")
                Call SendData(ToIndex, UserIndex, 0, "||Donación entregada con éxito." & FONTTYPE_FENIX)
                GuardarPremium False, UCase$(name), 60
                Exit Sub
            ElseIf Arg1 = 4 Then
                UserList(NameIndex(name)).Stats.PuntosDonador = 1
                UserList(NameIndex(name)).Faccion.Quests = UserList(NameIndex(name)).Faccion.Quests + 150
                Call SendData(ToIndex, NameIndex(name), 0, "||¡Gracias por colaborar con el servidor! A partir de hoy sos usuario PREMIUM y vence según el plan que has abonado. Para disfrutar de los beneficios PREMIUM escribe /CIUDADPREMIUM. ~255~255~0~1~0")
                Call SendData(ToIndex, UserIndex, 0, "||Donación entregada con éxito." & FONTTYPE_FENIX)
                GuardarPremium False, UCase$(name), 90
                Exit Sub
            End If
            'guardar el registro de nick premium
        End If
    End If

    If UCase$(Left$(rdata, 6)) = "/BANT " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        Arg1 = ReadField(1, rdata, 64)
        name = ReadField(2, rdata, 64)
        i = val(ReadField(3, rdata, 64))
        If Len(Arg1) = 0 Or Len(name) = 0 Or i = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||La estructura del comando es /BANT CAUSA@NICK@DIAS." & FONTTYPE_FENIX)
            Exit Sub
        End If
        TIndex = NameIndex(name)
        If TIndex > 0 Then
            If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
                Call SendData(ToIndex, UserIndex, 0, "1B")
                Exit Sub
            End If
            Call BanTemporal(name, i, Arg1, UserList(UserIndex).name)
            Call SendData(ToAdmins, 0, 0, "%X" & UserList(UserIndex).name & "," & UserList(TIndex).name)
            UserList(TIndex).flags.Ban = 1
            Call WarpUserChar(TIndex, Althalos.Map, Althalos.X, Althalos.Y)

            Call CloseSocket(TIndex)
        Else
            If Not FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = False Then
                Call SendData(ToIndex, UserIndex, 0, "||Offline, baneando" & FONTTYPE_INFO)

                If GetVar(CharPath & name & ".chr", "FLAGS", "Ban") <> "0" Then
                    Call SendData(ToIndex, UserIndex, 0, "||El personaje ya se encuentra baneado." & FONTTYPE_INFO)
                    Exit Sub
                End If

                Call BanTemporal(name, i, Arg1, UserList(UserIndex).name)

                Call ChangeBan(name, 1)
                Call ChangePos(name)

                Call SendData(ToAdmins, 0, 0, "%X" & UserList(UserIndex).name & "," & name)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe." & FONTTYPE_INFO)
            End If
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 7)) = "/ECHAR " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1E")
            Exit Sub
        End If
        If TIndex = UserIndex Then Exit Sub
        If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
            Call SendData(ToIndex, UserIndex, 0, "1F")
            Exit Sub
        End If
        Call SendData(ToAdmins, 0, 0, "%U" & UserList(UserIndex).name & "," & UserList(TIndex).name)
        Call LogGM(UserList(UserIndex).name, "Echo a " & UserList(TIndex).name, False)
        Call CloseSocket(TIndex)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 5)) = "/GLD " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        CantidadOro = rdata
        Call SendData(ToAdmins, 0, 0, "||EL ORO FUE AUMENTADO X" & rdata & "." & FONTTYPE_SERVER)
        Call WriteVar(IniPath & "Server.ini", "INIT", "ORO", rdata)
        Exit Sub
    End If

    If UCase$(Left$(rdata, 5)) = "/EXP " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        CantidadEXP = rdata
        Call SendData(ToAdmins, 0, 0, "||LA EXPERIENCIA FUE AUMENTADA X" & rdata & "." & FONTTYPE_SERVER)
        Call WriteVar(IniPath & "Server.ini", "INIT", "EXP", rdata)
        Exit Sub
    End If


    If UCase$(Left$(rdata, 5)) = "/BAN " Then
        Dim Razon As String
        rdata = Right$(rdata, Len(rdata) - 5)
        Razon = ReadField(1, rdata, Asc("@"))
        name = ReadField(2, rdata, Asc("@"))
        TIndex = NameIndex(name)
        '/ban motivo@nombre
        If TIndex Then
            If TIndex = UserIndex Then Exit Sub
            name = UserList(TIndex).name
            If UserList(TIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
                Call SendData(ToIndex, UserIndex, 0, "%V")
                Exit Sub
            End If
            Call LogBan(TIndex, UserIndex, Razon)
            UserList(TIndex).flags.Ban = 1
            If UserList(TIndex).flags.Privilegios Then
                UserList(UserIndex).flags.Ban = 1
                Call SendData(ToAdmins, 0, 0, "%W" & UserList(UserIndex).name)
                Call LogBan(UserIndex, UserIndex, "Baneado por banear a otro GM.")
                Call CloseSocket(UserIndex)
            End If

            Call SendData(ToAdmins, 0, 0, "%X" & UserList(UserIndex).name & "," & UserList(TIndex).name)
            Call SendData(ToAdmins, 0, 0, "||IP: " & UserList(TIndex).ip & " Mail: " & UserList(TIndex).Email & "." & FONTTYPE_FIGHT)

            Call CloseSocket(TIndex)
        Else
            If Not FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = False Then
                Call ChangeBan(name, 1)
                Call LogBanOffline(UCase$(name), UserIndex, Razon)
                Call SendData(ToAdmins, 0, 0, "%X" & UserList(UserIndex).name & "," & name)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe." & FONTTYPE_INFO)
            End If
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        If Not FileExist(CharPath & UCase$(rdata) & ".chr", vbNormal) = False Then
            Call ChangeBan(rdata, 0)
            Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " unbanned." & FONTTYPE_INFO)
            For i = 1 To Baneos.Count
                If Baneos(i).name = UCase$(rdata) Then
                    Call Baneos.Remove(i)
                    Exit Sub
                End If
            Next
        Else
            Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe" & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    If UCase$(rdata) = "/SEGUIR" Then
        If UserList(UserIndex).flags.TargetNpc Then
            Call DoFollow(UserList(UserIndex).flags.TargetNpc, UserIndex)
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 3)) = "/CC" Then
        Call EnviarSpawnList(UserIndex)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 3)) = "SPA" Then
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) > 0 And val(rdata) < UBound(SpawnList) + 1 Then _
           Call SpawnNpc(SpawnList(val(rdata)).NpcIndex, UserList(UserIndex).pos, True, False)
        Call LogGM(UserList(UserIndex).name, "Sumoneo " & SpawnList(val(rdata)).NpcName, False)
        Exit Sub
    End If
    If UCase$(rdata) = "/RESETINV" Then
        rdata = Right$(rdata, Len(rdata) - 9)
        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        Call ResetNpcInv(UserList(UserIndex).flags.TargetNpc)
        Call LogGM(UserList(UserIndex).name, "/RESETINV " & Npclist(UserList(UserIndex).flags.TargetNpc).name, False)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 6)) = "/RMSG " Then    '[LmB/Mercury 2008]
        rdata = Right$(rdata, Len(rdata) - 6)
        If UserList(UserIndex).flags.Privilegios < 1 Then Exit Sub
        Call LogGM(UserList(UserIndex).name, "Mensaje Broadcast:" & rdata, False)
        If rdata <> "" Then
            Call SendData(ToAll, 0, 0, "|$" & UserList(UserIndex).name & ">> " & rdata)
        End If
        Exit Sub
    End If    '[LmB/Mercury 2008]

    If UCase$(Left$(rdata, 7)) = "/RMSGT " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        If UCase$(rdata) = "NO" Then
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).name & " ha anulado la repetición del mensaje: " & MensajeRepeticion & "." & FONTTYPE_FENIX)
            IntervaloRepeticion = 0
            TiempoRepeticion = 0
            MensajeRepeticion = ""
            Exit Sub
        End If
        tName = ReadField(1, rdata, 64)
        tInt = ReadField(2, rdata, 64)
        Prueba1 = ReadField(3, rdata, 64)
        If Len(tName) = 0 Or val(Prueba1) = 0 Or (Prueba1 >= tInt And tInt <> 0) Then
            Call SendData(ToIndex, UserIndex, 0, "||La estructura del comando es: /RMSGT MENSAJE@TIEMPO TOTAL@INTERVALO DE REPETICION." & FONTTYPE_INFO)
            Exit Sub
        End If
        If val(tInt) > 10000 Or val(Prueba1) > 10000 Then
            Call SendData(ToIndex, UserIndex, 0, "||La cantidad de tiempo establecida es demasiado grande." & FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(UserList(UserIndex).name, "Mensaje Broadcast repetitivo:" & rdata, False)
        MensajeRepeticion = tName
        TiempoRepeticion = tInt
        IntervaloRepeticion = Prueba1
        If TiempoRepeticion = 0 Then
            Call SendData(ToAdmins, 0, 0, "||El mensaje " & MensajeRepeticion & " será repetido cada " & IntervaloRepeticion & " minutos durante tiempo indeterminado." & FONTTYPE_FENIX)
            TiempoRepeticion = -IntervaloRepeticion
        Else
            Call SendData(ToAdmins, 0, 0, "||El mensaje " & MensajeRepeticion & " será repetido cada " & IntervaloRepeticion & " minutos durante un total de " & TiempoRepeticion & " minutos." & FONTTYPE_FENIX)
            TiempoRepeticion = TiempoRepeticion - TiempoRepeticion Mod IntervaloRepeticion
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 8)) = "/BUSCAR " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        For i = 1 To UBound(ObjData)
            If InStr(1, Tilde(ObjData(i).name), Tilde(rdata)) Then
                Call SendData(ToIndex, UserIndex, 0, "||" & i & " " & ObjData(i).name & "." & FONTTYPE_INFO)
                N = N + 1
            End If
        Next
        If N = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No hubo resultados de la búsqueda: " & rdata & "." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Hubo " & N & " resultados de la busqueda: " & rdata & "." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 8)) = "/CUENTA " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        CuentaRegresiva = val(ReadField(1, rdata, 32)) + 1
        GMCuenta = UserList(UserIndex).pos.Map
        Exit Sub
    End If

    If UCase$(rdata) = "/MATA" Then
        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
        Call LogGM(UserList(UserIndex).name, "/MATA " & Npclist(UserList(UserIndex).flags.TargetNpc).name, False)
        Exit Sub
    End If
    If UCase$(rdata) = "/MATARBOT" Then
        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        Call QuitarNPCBOT(UserList(UserIndex).flags.TargetNpc)
        Call LogGM(UserList(UserIndex).name, "/MATA " & Npclist(UserList(UserIndex).flags.TargetNpc).name, False)
        frmMain.CantUsuarios.Caption = NumUsers + NumBots
        Call SendData(ToAll, 0, 0, "NON" & NumUsers + NumBots)
        Exit Sub
    End If
    If UCase$(rdata) = "/IGNORAR" Then
        If UserList(UserIndex).flags.Ignorar = 1 Then
            UserList(UserIndex).flags.Ignorar = 0
            Call SendData(ToIndex, UserIndex, 0, "||Ahora las criaturas te persiguen." & FONTTYPE_INFO)
        Else
            UserList(UserIndex).flags.Ignorar = 1
            Call SendData(ToIndex, UserIndex, 0, "||Ahora las criaturas te ignoran." & FONTTYPE_INFO)
        End If
    End If
    If UCase$(rdata) = "/PROTEGER" Then
        TIndex = UserList(UserIndex).flags.TargetUser
        If TIndex > 0 Then
            If UserList(TIndex).flags.Privilegios > 1 Then Exit Sub
            If UserList(TIndex).flags.Protegido = 1 Then
                UserList(TIndex).flags.Protegido = 0
                Call SendData(ToIndex, UserIndex, 0, "||Desprotegiste a " & UserList(TIndex).name & "." & FONTTYPE_INFO)
                Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).name & " te desprotegió." & FONTTYPE_FIGHT)
            Else
                UserList(TIndex).flags.Protegido = 1
                Call SendData(ToIndex, UserIndex, 0, "||Protegiste a " & UserList(TIndex).name & "." & FONTTYPE_INFO)
                Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).name & " te protegió. No puedes atacar ni ser atacado." & FONTTYPE_FIGHT)
            End If
        End If
    End If
    If Left$(UCase$(rdata), 5) = "/PRO " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        TIndex = NameIndex(rdata)
        If TIndex > 0 Then
            If UserList(TIndex).flags.Privilegios > 1 Then Exit Sub
            If UserList(TIndex).flags.Protegido = 1 Then
                UserList(TIndex).flags.Protegido = 0
                Call SendData(ToIndex, UserIndex, 0, "||Desprotegiste a " & UserList(TIndex).name & "." & FONTTYPE_INFO)
                Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).name & " te desprotegió." & FONTTYPE_FIGHT)
            Else
                UserList(TIndex).flags.Protegido = 1
                Call SendData(ToIndex, UserIndex, 0, "||Protegiste a " & UserList(TIndex).name & "." & FONTTYPE_INFO)
                Call SendData(ToIndex, TIndex, 0, "||" & UserList(UserIndex).name & " te protegió. No puedes atacar ni ser atacado." & FONTTYPE_FIGHT)
            End If
        End If
    End If
    If UCase$(Left$(rdata, 6)) = "/STOP " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        TIndex = NameIndex(rdata)

        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(TIndex).flags.Stopped = True Then
            UserList(TIndex).flags.Stopped = False
            Call SendData(ToIndex, UserIndex, 0, "||Has sacado el Stop a " & UserList(TIndex).name & "." & FONTTYPE_FENIX)
            Call SendData(ToIndex, TIndex, 0, "NT")
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Has Stoppeado a " & UserList(TIndex).name & "." & FONTTYPE_FENIX)
            UserList(TIndex).flags.Stopped = True
            Call SendData(ToIndex, TIndex, 0, "ST")
        End If
    End If
    If UCase$(Left$(rdata, 5)) = "/DEST" Then
        Call LogGM(UserList(UserIndex).name, "/DEST", False)
        rdata = Right$(rdata, Len(rdata) - 5)
        Call EraseObj(ToMap, UserIndex, UserList(UserIndex).pos.Map, 10000, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y)
        Exit Sub
    End If
    If UCase$(rdata) = "/MASSDEST" Then
        For Y = UserList(UserIndex).pos.Y - MinYBorder + 1 To UserList(UserIndex).pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).pos.X - MinXBorder + 1 To UserList(UserIndex).pos.X + MinXBorder - 1
                If InMapBounds(X, Y) Then _
                   If MapData(UserList(UserIndex).pos.Map, X, Y).OBJInfo.OBJIndex > 0 And Not ItemEsDeMapa(UserList(UserIndex).pos.Map, X, Y) Then Call EraseObj(ToMap, UserIndex, UserList(UserIndex).pos.Map, 10000, UserList(UserIndex).pos.Map, X, Y)
            Next
        Next
        Call LogGM(UserList(UserIndex).name, "/MASSDEST", (UserList(UserIndex).flags.Privilegios = 1))
        Exit Sub
    End If
    If UCase$(Left$(rdata, 6)) = "/KILL " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        TIndex = NameIndex(rdata)
        If TIndex Then
            If UserList(TIndex).flags.Privilegios < UserList(UserIndex).flags.Privilegios Then Call UserDie(TIndex)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 11)) = "/GANOTORNEO" Then
        Dim spll As Obj
        rdata = Right$(rdata, Len(rdata) - 5)
        TIndex = UserList(UserIndex).flags.TargetUser
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.EsNoble = 0 Then
            Call SendData(ToAll, 0, 0, "||" & UserList(UserList(UserIndex).flags.TargetUser).name & " ganó un Torneo organizado por un GM." & FONTTYPE_INFO)
            spll.Amount = 1  'Cantidad de copas
            spll.OBJIndex = 824  'Numero de item
            Call MeterItemEnInventario(TIndex, spll)    'Metemos Item En Inventario.
            Call SendData(ToIndex, TIndex, 0, "||Has recibido una Copa de Oro, felicitaciones!." & FONTTYPE_INFO)
            UserList(UserList(UserIndex).flags.TargetUser).Faccion.Torneos = UserList(UserList(UserIndex).flags.TargetUser).Faccion.Torneos + 1
            UserList(UserIndex).Ranking.TorneosGanados = UserList(UserIndex).Ranking.TorneosGanados + 1
            Call GuardarRanking("Torneo", UserIndex)
            Call LogGM(UserList(UserIndex).name, "Gano torneo: " & UserList(TIndex).name & " Map:" & UserList(UserIndex).pos.Map & " X:" & UserList(UserIndex).pos.X & " Y:" & UserList(UserIndex).pos.Y, False)
            Exit Sub
        End If
        If UserList(UserIndex).flags.EsNoble = 1 Then
            Call SendData(ToAll, 0, 0, "||" & UserList(UserList(UserIndex).flags.TargetUser).name & " ganó un Torneo organizado por un GM." & FONTTYPE_INFO)
            spll.Amount = 1  'Cantidad de copas
            spll.OBJIndex = 824  'Numero de item
            Call MeterItemEnInventario(TIndex, spll)    'Metemos Item En Inventario.
            Call SendData(ToIndex, TIndex, 0, "||Has recibido dos Copas de Oro, felicitaciones!." & FONTTYPE_INFO)
            UserList(UserList(UserIndex).flags.TargetUser).Faccion.Torneos = UserList(UserList(UserIndex).flags.TargetUser).Faccion.Torneos + 2
            UserList(UserIndex).Ranking.TorneosGanados = UserList(UserIndex).Ranking.TorneosGanados + 2
            Call GuardarRanking("Torneo", UserIndex)
            Call LogGM(UserList(UserIndex).name, "Gano torneo: " & UserList(TIndex).name & " Map:" & UserList(UserIndex).pos.Map & " X:" & UserList(UserIndex).pos.X & " Y:" & UserList(UserIndex).pos.Y, False)
            Exit Sub
        End If
    End If
    'Canjeo
    If UCase$(Left$(rdata, 12)) = "/SACARPUNTO " Then
        rdata = Right$(rdata, Len(rdata) - 12)
        TIndex = UserList(UserIndex).flags.TargetUser
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Debes seleccionar al Jugador para Sacarle sus Puntos!" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(TIndex).Faccion.Quests < rdata Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes Sacar esa Cantidad de Puntos, Genera Variable Muerta!" & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserList(UserIndex).flags.TargetUser).Faccion.Quests = UserList(UserList(UserIndex).flags.TargetUser).Faccion.Quests - rdata
        Call LogGM(UserList(UserIndex).name, "Restó puntos de canje: " & rdata & UserList(TIndex).name & " Map:" & UserList(UserIndex).pos.Map & " X:" & UserList(UserIndex).pos.X & " Y:" & UserList(UserIndex).pos.Y, False)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 13)) = "/VERPROCESOS " Then
        rdata = Right$(rdata, Len(rdata) - 13)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, TIndex, 0, "PCGR" & UserIndex)
        End If
        Exit Sub
    End If
    If UserList(UserIndex).flags.Privilegios < 3 Then Exit Sub
    If UCase$(Left$(rdata, 8)) = "/VERFPS " Then
        rdata = Right$(rdata, Len(rdata) - 8)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, TIndex, 0, "MFPS" & UserIndex)
        End If
        Exit Sub
    End If
    If UCase$(rdata) = "/RESTRINGIR" Then
        If Restringido Then
            Call SendData(ToAll, 0, 0, "||La restricción de GameMasters fue desactivada." & FONTTYPE_FENIX)
            Call LogGM(UserList(UserIndex).name, "Desrestringió el servidor.", False)
        Else
            Call SendData(ToAll, 0, 0, "||La restricción de GameMasters fue activada." & FONTTYPE_FENIX)
            For i = 1 To LastUser
                DoEvents
                If UserList(i).flags.UserLogged And UserList(i).flags.Privilegios = 0 And Not UserList(i).flags.PuedeDenunciar Then Call CloseSocket(i)
            Next
            Call LogGM(UserList(UserIndex).name, "Restringió el servidor.", False)
        End If
        Restringido = Not Restringido
        Exit Sub
    End If
    If UCase$(Left$(rdata, 7)) = "/VERHD " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        TIndex = NameIndex(rdata)
        If TIndex = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||El disco es: " & UserList(TIndex).HD & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 7)) = "/BANHD " Then
        rdata = Right$(rdata, Len(rdata) - 7)
        TIndex = NameIndex(rdata)
        If TIndex = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario Offline." & FONTTYPE_INFO)
            Exit Sub
        Else
            For LoopC = 1 To BanHDs.Count
                If BanHDs.Item(LoopC) = UserList(TIndex).HD Then
                    Call SendData(ToIndex, UserIndex, 0, "||Disco duro ya baneado" & FONTTYPE_INFO)
                    Exit Sub
                End If
            Next
            BanHDs.Add UserList(TIndex).HD
            Call SendData(ToIndex, UserIndex, 0, "||Has baneado al disco duro: " & UserList(TIndex).HD & " del usuario " & UserList(TIndex).name & "." & FONTTYPE_INFO)
            Dim numHD As Integer
            numHD = val(GetVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad"))
            If FileExist(App.path & "\Logs\BanHDs.dat", vbNormal) Then
                Call WriteVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad", numHD + 1)
                Call WriteVar(App.path & "\Logs\BanHDs.dat", "BANS", "Disco" & numHD + 1, UserList(TIndex).HD)
                Call LogGM(UserList(UserIndex).name, "/BanHD " & UserList(TIndex).name & " " & UserList(TIndex).HD, False)
            Else
                Call WriteVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad", 1)
                Call WriteVar(App.path & "\Logs\BanHDs.dat", "BANS", "Disco1", UserList(TIndex).HD)
                Call LogGM(UserList(UserIndex).name, "/BanHD " & UserList(TIndex).name & " " & UserList(TIndex).HD, False)
            End If
            Call CloseSocket(TIndex)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 9)) = "/UNBANHD " Then
        rdata = Right$(rdata, Len(rdata) - 9)
        Dim numHD2 As Integer
        numHD2 = val(GetVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad"))
        For LoopC = 1 To BanHDs.Count
            If BanHDs.Item(LoopC) = rdata Then
                BanHDs.Remove LoopC
                Call SendData(ToIndex, UserIndex, 0, "||Has desbaneado el disco de " & rdata & FONTTYPE_INFO)
                Call WriteVar(App.path & "\Logs\BanHDs.dat", "INIT", "Cantidad", numHD2 - 1)
                Call WriteVar(App.path & "\Logs\BanHDs.dat", "BANS", "Disco" & numHD2 - 1, "")
                Call LogGM(UserList(UserIndex).name, "/UNBanHD " & UserList(TIndex).name & " " & UserList(TIndex).HD, False)
            End If
        Next
        Exit Sub
    End If

    If UCase$(Left$(rdata, 7)) = "/BANPC " Then
        rdata = Right$(rdata, Len(rdata) - 7)    'nick
        Call BanPC(UserIndex, NameIndex(rdata))
        Exit Sub
    End If

    If UCase$(Left$(rdata, 7)) = "/BANIP " Then
        Dim BanIP As String, XNick As Boolean
        rdata = Right$(rdata, Len(rdata) - 7)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            XNick = False
            Call LogGM(UserList(UserIndex).name, "/BanIP " & rdata, False)
            BanIP = rdata
        Else
            XNick = True
            Call LogGM(UserList(UserIndex).name, "/BanIP " & UserList(TIndex).name & " - " & UserList(TIndex).ip, False)
            BanIP = UserList(TIndex).ip
        End If
        For LoopC = 1 To BanIps.Count
            If BanIps.Item(LoopC) = BanIP Then
                Call SendData(ToIndex, UserIndex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
                Exit Sub
            End If
        Next
        BanIps.Add BanIP
        Call SendData(ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
        If XNick Then
            Call LogBan(TIndex, UserIndex, "Ban por IP desde Nick")
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).name & " echo a " & UserList(TIndex).name & "." & FONTTYPE_FIGHT)
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).name & " Banned a " & UserList(TIndex).name & "." & FONTTYPE_FIGHT)
            UserList(TIndex).flags.Ban = 1
            Call LogGM(UserList(UserIndex).name, "Echo a " & UserList(TIndex).name, False)
            Call LogGM(UserList(UserIndex).name, "BAN a " & UserList(TIndex).name, False)
            Call CloseSocket(TIndex)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 9)) = "/UNBANIP " Then
        rdata = Right$(rdata, Len(rdata) - 9)
        Call LogGM(UserList(UserIndex).name, "/UNBANIP " & rdata, False)
        For LoopC = 1 To BanIps.Count
            If BanIps.Item(LoopC) = rdata Then
                BanIps.Remove LoopC
                Call SendData(ToIndex, UserIndex, 0, "||La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
                Exit Sub
            End If
        Next
        Call SendData(ToIndex, UserIndex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 9)) = "/BANMAIL " Then
        Dim BanMail As String, XXNick As Boolean
        rdata = Right$(rdata, Len(rdata) - 9)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            XXNick = False
            Call LogGM(UserList(UserIndex).name, "/BanMail " & rdata, False)
            BanMail = rdata
        Else
            XXNick = True
            Call LogGM(UserList(UserIndex).name, "/BanMail " & UserList(TIndex).name & " - " & UserList(TIndex).Email, False)
            BanMail = UserList(TIndex).Email
        End If
        numeromail = GetVar(App.path & "\logs\" & "BanMail.dat", "INIT", "Mails")
        For LoopC = 1 To numeromail
            If GetVar(App.path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail") = BanMail Then
                Call SendData(ToIndex, UserIndex, 0, "||El mail " & BanMail & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
                Exit Sub
            End If
        Next
        Call WriteVar(App.path & "\logs\" & "BanMail.dat", "Mail" & numeromail + 1, "Mail", BanMail)
        If XXNick Then Call WriteVar(App.path & "\logs\" & "BanMail.dat", "Mail" & numeromail + 1, "User", UserList(TIndex).name)
        Call WriteVar(App.path & "\logs\" & "BanMail.dat", "INIT", "Mails", numeromail + 1)
        Call SendData(ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).name & " Baneo el mail " & BanMail & FONTTYPE_FIGHT)
        If XXNick Then
            Call LogBan(TIndex, UserIndex, "Ban por mail desde Nick")
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).name & " echo a " & UserList(TIndex).name & "." & FONTTYPE_FIGHT)
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).name & " Banned a " & UserList(TIndex).name & "." & FONTTYPE_FIGHT)
            UserList(TIndex).flags.Ban = 1
            Call LogGM(UserList(UserIndex).name, "Echo a " & UserList(TIndex).name, False)
            Call LogGM(UserList(UserIndex).name, "BAN a " & UserList(TIndex).name, False)
            Call CloseSocket(TIndex)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 11)) = "/UNBANMAIL " Then
        numeromail = GetVar(App.path & "\logs\" & "BanMail.dat", "INIT", "Mails")
        rdata = Right$(rdata, Len(rdata) - 11)
        Call LogGM(UserList(UserIndex).name, "/UNBanMail " & rdata, False)
        For LoopC = 1 To numeromail
            If GetVar(App.path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail") = rdata Then
                Call WriteVar(App.path & "\logs\" & "BanMail.dat", "Mail" & numeromail, "Mail", "Desbaneado por " & UserList(UserIndex).name)
                Call SendData(ToIndex, UserIndex, 0, "||El mail " & rdata & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
                Exit Sub
            End If
        Next
        Call SendData(ToIndex, UserIndex, 0, "||El mail " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 3)) = "/CT" Then
        rdata = Right$(rdata, Len(rdata) - 4)
        Call LogGM(UserList(UserIndex).name, "/CT: " & rdata, False)
        mapa = ReadField(1, rdata, 32)
        X = ReadField(2, rdata, 32)
        Y = ReadField(3, rdata, 32)
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).OBJInfo.OBJIndex Then
            Exit Sub
        End If
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map Then
            Exit Sub
        End If
        If Not MapaValido(mapa) Or Not InMapBounds(X, Y) Then Exit Sub
        ET.Amount = 1
        ET.OBJIndex = Teleport
        Call MakeObj(ToMap, 0, UserList(UserIndex).pos.Map, ET, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1)
        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Map = mapa
        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.X = X
        MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y - 1).TileExit.Y = Y
        Exit Sub
    End If
    If UCase$(Left$(rdata, 3)) = "/DT" Then
        Call LogGM(UserList(UserIndex).name, "/DT", False)
        mapa = UserList(UserIndex).flags.TargetMap
        X = UserList(UserIndex).flags.TargetX
        Y = UserList(UserIndex).flags.TargetY
        If ObjData(MapData(mapa, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT And _
           MapData(mapa, X, Y).TileExit.Map Then
            Call EraseObj(ToMap, 0, mapa, MapData(mapa, X, Y).OBJInfo.Amount, mapa, X, Y)
            MapData(mapa, X, Y).TileExit.Map = 0
            MapData(mapa, X, Y).TileExit.X = 0
            MapData(mapa, X, Y).TileExit.Y = 0
        End If

        Exit Sub
    End If
    If UCase$(Left$(rdata, 13)) = "/KILLPROCESS " Then
        Dim NombreProceso As String
        rdata = Right$(rdata, Len(rdata) - 13)
        TIndex = NameIndex(ReadField(1, rdata, Asc(" ")))
        NombreProceso = Right$(rdata, Len(rdata) - (Len(ReadField(1, rdata, Asc(" "))) + 1))
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Else
            Call SendData(ToAdmins, UserIndex, 0, "||Estas borrando el Proceso: " & NombreProceso & " de " & UserList(TIndex).name & FONTTYPE_WARNING)
            Call SendData(ToIndex, TIndex, 0, "VPDM" & " " & NombreProceso)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
        Call LogGM(UserList(UserIndex).name, "/BLOQ", False)
        rdata = Right$(rdata, Len(rdata) - 5)
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 0 Then
            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 1
            Call Bloquear(ToMap, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y, 1)
        Else
            MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y).Blocked = 0
            Call Bloquear(ToMap, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.X, UserList(UserIndex).pos.Y, 0)
        End If
        Exit Sub
    End If
    If UCase$(rdata) = "/MASSKILL" Then
        For Y = UserList(UserIndex).pos.Y - MinYBorder + 1 To UserList(UserIndex).pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).pos.X - MinXBorder + 1 To UserList(UserIndex).pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                   If MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex Then Call QuitarNPC(MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex)
            Next
        Next
        Call LogGM(UserList(UserIndex).name, "/MASSKILL", False)
        Exit Sub
    End If
    If UCase$(rdata) = "/MASSBOT" Then
        For Y = UserList(UserIndex).pos.Y - MinYBorder + 1 To UserList(UserIndex).pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).pos.X - MinXBorder + 1 To UserList(UserIndex).pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                   If MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex Then Call QuitarNPCBOT(MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex)
            Next
        Next
        NumBots = NumBots - (MapData(UserList(UserIndex).pos.Map, X, Y).NpcIndex)
        Call LogGM(UserList(UserIndex).name, "/MASSKILL", False)
        frmMain.CantUsuarios.Caption = NumUsers + NumBots
        Call SendData(ToAll, 0, 0, "NON" & NumUsers + NumBots)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 6)) = "/SMSG " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        Call LogGM(UserList(UserIndex).name, "Mensaje de sistema:" & rdata, False)
        Call SendData(ToAll, 0, 0, "!!" & rdata & ENDC)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 5)) = "/ACC " Then
        rdata = val(Right$(rdata, Len(rdata) - 5))
        NumNPC = val(GetVar(App.path & "\Dat\NPCs-HOSTILES.dat", "INIT", "NumNPCs")) + 500
        If rdata < 0 Or rdata > NumNPC Then
            Call SendData(ToIndex, UserIndex, 0, "||La criatura no existe." & FONTTYPE_INFO)
        Else
            Call SpawnNpc(val(rdata), UserList(UserIndex).pos, True, False)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 6)) = "/RACC " Then
        rdata = val(Right$(rdata, Len(rdata) - 6))
        NumNPC = val(GetVar(App.path & "\Dat\NPCs-HOSTILES.dat", "INIT", "NumNPCs")) + 500
        If rdata < 0 Or rdata > NumNPC Then
            Call SendData(ToIndex, UserIndex, 0, "||La criatura no existe." & FONTTYPE_INFO)
        Else
            Call SpawnNpc(val(rdata), UserList(UserIndex).pos, True, True)
        End If
        Exit Sub
    End If
    If UCase$(rdata) = "/NAVE" Then
        If UserList(UserIndex).flags.Navegando Then
            UserList(UserIndex).flags.Navegando = 0
        Else
            UserList(UserIndex).flags.Navegando = 1
        End If
        Exit Sub
    End If

    If UCase$(rdata) = "/APAGAR" Then
        Call LogMain(" Server apagado por " & UserList(UserIndex).name & ".")
        Call SendData(ToAll, 0, 0, "||APAGANDO SISTEMA..." & FONTTYPE_FIGHT)
        DoBackUp (True)
        Call ApagarSistema
        End
    End If


    If UCase$(rdata) = "/INTERVALOS" Then
        Call SendData(ToIndex, UserIndex, 0, "||Golpe-Golpe: " & IntervaloUserPuedeAtacar & " segundos." & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Golpe-Hechizo: " & IntervaloUserPuedeGolpeHechi & " segundos." & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Hechizo-Hechizo: " & IntervaloUserPuedeCastear & " segundos." & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Hechizo-Golpe: " & IntervaloUserPuedeHechiGolpe & " segundos." & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Arco-Arco: " & IntervaloUserFlechas & " segundos." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 6)) = "/MODS " Then
        Dim PreInt As Single
        rdata = Right$(rdata, Len(rdata) - 6)
        TIndex = ClaseIndex(ReadField(1, rdata, 64))
        If TIndex = 0 Then Exit Sub
        tInt = ReadField(2, rdata, 64)
        If tInt < 1 Or tInt > 6 Then Exit Sub
        Arg5 = ReadField(3, rdata, 64)
        If Arg5 < 40 Or Arg5 > 125 Then Exit Sub
        PreInt = Mods(tInt, TIndex)
        Mods(tInt, TIndex) = Arg5 / 100
        Call SendData(ToAdmins, 0, 0, "||El modificador n° " & tInt & " de la clase " & ListaClases(TIndex) & " fue cambiado de " & PreInt & " a " & Mods(tInt, TIndex) & "." & FONTTYPE_FIGHT)
        Call SaveMod(tInt, TIndex)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 4)) = "/INT" Then
        rdata = Right$(rdata, Len(rdata) - 4)
        Select Case UCase$(Left$(rdata, 2))
            Case "GG"
                rdata = Right$(rdata, Len(rdata) - 3)
                PreInt = IntervaloUserPuedeAtacar
                IntervaloUserPuedeAtacar = val(rdata) / 10
                Call SendData(ToAdmins, 0, 0, "||El intervalo Golpe-Golpe fue cambiado de " & PreInt & " a " & IntervaloUserPuedeAtacar & " segundos." & FONTTYPE_INFO)
                Call SendData(ToAll, 0, 0, "INTA" & IntervaloUserPuedeAtacar * 10)
                Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", IntervaloUserPuedeAtacar * 10)
            Case "GH"
                rdata = Right$(rdata, Len(rdata) - 3)
                PreInt = IntervaloUserPuedeGolpeHechi
                IntervaloUserPuedeGolpeHechi = val(rdata) / 10
                Call SendData(ToAdmins, 0, 0, "||El intervalo Golpe-Hechizo fue cambiado de " & PreInt & " a " & IntervaloUserPuedeGolpeHechi & " segundos." & FONTTYPE_INFO)
                Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeGolpeHechi", IntervaloUserPuedeGolpeHechi * 10)
            Case "HH"
                rdata = Right$(rdata, Len(rdata) - 3)
                PreInt = IntervaloUserPuedeCastear
                IntervaloUserPuedeCastear = val(rdata) / 10
                Call SendData(ToAdmins, 0, 0, "||El intervalo Hechizo-Hechizo fue cambiado de " & PreInt & " a " & IntervaloUserPuedeCastear & " segundos." & FONTTYPE_INFO)
                Call SendData(ToAll, 0, 0, "INTS" & IntervaloUserPuedeCastear * 10)
                Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", IntervaloUserPuedeCastear * 10)
            Case "HG"
                rdata = Right$(rdata, Len(rdata) - 3)
                PreInt = IntervaloUserPuedeHechiGolpe
                IntervaloUserPuedeHechiGolpe = val(rdata) / 10
                Call SendData(ToAdmins, 0, 0, "||El intervalo Hechizo-Golpe fue cambiado de " & PreInt & " a " & IntervaloUserPuedeHechiGolpe & " segundos." & FONTTYPE_INFO)
                Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeHechiGolpe", IntervaloUserPuedeHechiGolpe * 10)
            Case "AA"
                rdata = Right$(rdata, Len(rdata) - 2)
                PreInt = IntervaloUserFlechas
                IntervaloUserFlechas = val(rdata) / 10
                Call SendData(ToAdmins, 0, 0, "||El intervalo de flechas fue cambiado de " & PreInt & " a " & IntervaloUserFlechas & " segundos." & FONTTYPE_INFO)
                Call SendData(ToIndex, UserIndex, 0, "INTF" & IntervaloUserFlechas * 10)

                Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserFlechas", IntervaloUserFlechas * 10)
            Case "SH"
                rdata = Right$(rdata, Len(rdata) - 2)
                PreInt = IntervaloUserSH
                IntervaloUserSH = val(rdata)
                Call SendData(ToAdmins, 0, 0, "||Intervalo de SH cambiado de " & PreInt & " a " & IntervaloUserSH & " segundos de tardanza." & FONTTYPE_INFO)
                Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserSH", str(IntervaloUserSH))
        End Select
    End If
    If UCase$(rdata) = "/DATS" Then
        Call CargarHechizos
        Call LoadOBJData
        Call DescargaNpcsDat
        Call CargaNpcsDat
        Exit Sub
    End If
    If UCase$(Left$(rdata, 13)) = "/ENVENCUESTA " Then
        If encuestas.activa = 1 Then Call SendData(ToIndex, UserIndex, 0, "||Ya hay una encuesta, espera a que termine.." & FONTTYPE_INFO)
        rdata = Right$(rdata, Len(rdata) - 13)
        encuestas.votosNP = 0
        encuestas.votosSI = 0
        encuestas.Tiempo = 0
        encuestas.activa = 1
        Call SendData(ToAll, 0, 0, "||ENCUESTA> " & rdata & FONTTYPE_FENIX)
        Call SendData(ToAll, 0, 0, "||OPCIONES: /SI - /NO | La encuesta durará 30 segundos." & FONTTYPE_FENIX)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 12)) = "/ACEPTCONSE " Then
        If UserList(UserIndex).flags.Privilegios > 1 Then
            rdata = Right$(rdata, Len(rdata) - 12)
            TIndex = NameIndex(rdata)
            If TIndex <= 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
            Else
                Call SendData(ToAll, 0, 0, "||" & rdata & " fue aceptado en el honorable Consejo del Rey." & FONTTYPE_CONSEJO)
                UserList(TIndex).flags.EsConseReal = 1
                Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y, False)
            End If
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 16)) = "/ACEPTCONSECAOS " Then
        If UserList(UserIndex).flags.Privilegios > 1 Then
            rdata = Right$(rdata, Len(rdata) - 16)
            TIndex = NameIndex(rdata)
            If TIndex <= 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
            Else
                Call SendData(ToAll, 0, 0, "||" & rdata & " fue aceptado en el Concilio de ADELAIDE." & FONTTYPE_CONSEJOCAOS)
                UserList(TIndex).flags.EsConseCaos = 1
                Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y, False)
            End If
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 12)) = "/ECHARCONSE " Then
        If UserList(UserIndex).flags.Privilegios > 1 Then
            'If UserList(UserIndex).Flags.EsConseReal/Caos > 0 Or UserList(UserIndex).flags.Privilegios > 1 Then
            rdata = Right$(rdata, Len(rdata) - 12)
            TIndex = NameIndex(rdata)
            If TIndex <= 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Usuario offline" & FONTTYPE_INFO)
            Else
                If UserList(TIndex).flags.EsConseCaos = 1 Then
                    Call SendData(ToAll, 0, 0, "||" & rdata & " fue echado del Concilio del Mal." & FONTTYPE_CONSEJOCAOS)
                    UserList(TIndex).flags.EsConseCaos = 0
                    Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y, False)
                    Exit Sub
                End If
                If UserList(TIndex).flags.EsConseReal = 1 Then
                    Call SendData(ToAll, 0, 0, "||" & rdata & " fue echado del honorable Consejo del Rey." & FONTTYPE_CONSEJO)
                    UserList(TIndex).flags.EsConseReal = 0
                    Call WarpUserChar(TIndex, UserList(TIndex).pos.Map, UserList(TIndex).pos.X, UserList(TIndex).pos.Y, False)
                    Exit Sub
                End If
                If UserList(TIndex).flags.EsConseReal = 0 And UserList(TIndex).flags.EsConseCaos = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " no es consejero de la Alianza ni pertenece al consejo de la Horda." & FONTTYPE_FENIX)
                End If
            End If
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 7)) = "/NOMANA" Then
        rdata = Right$(rdata, Len(rdata) - 7)
        UserList(UserIndex).Stats.MinMAN = 0
        Call SendUserMANA(UserIndex)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 7)) = "/SECAEN" Then
        If MapInfo(UserList(UserIndex).pos.Map).SeCaenItems = 1 Then
            MapInfo(UserList(UserIndex).pos.Map).SeCaenItems = 0
            Call SendData(ToIndex, UserIndex, 0, "||Los items de los usuarios  se caerán en este mapa." & FONTTYPE_INFO)
            Exit Sub
        Else
            MapInfo(UserList(UserIndex).pos.Map).SeCaenItems = 1
            Call SendData(ToIndex, UserIndex, 0, "||Los items de los usuarios NO se caerán en este mapa." & FONTTYPE_INFO)
            Exit Sub
        End If
        Exit Sub
    End If

    If UCase$(rdata) = "/SOPORTEACTIVADO" Then
        SoporteDesactivado = Not SoporteDesactivado
        Call SendData(ToIndex, UserIndex, 0, "||El soporte está desactivado : " & SoporteDesactivado & FONTTYPE_FENIX)
        Exit Sub
    End If


    If UCase$(rdata) = "/MODOQUEST" Then
        ModoQuest = Not ModoQuest
        If ModoQuest Then
            Call SendData(ToAll, 0, 0, "||Modo Quest activado." & FONTTYPE_FENIX)
            Call SendData(ToAll, 0, 0, "||Los neutrales pueden poner /MERCENARIO ALIANZA o /MERCENARIO HORDA para enlistarse en alguna facción temporalmente durante la quest." & FONTTYPE_FENIX)
            Call SendData(ToAll, 0, 0, "||Al morir puedes poner /HOGAR y serás teletransportado a Althalos." & FONTTYPE_FENIX)
        Else
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).name & " desactivó el modo quest." & FONTTYPE_FENIX)
            Call DesactivarMercenarios
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 7)) = "/SEGURO" Then
        If MapInfo(UserList(UserIndex).pos.Map).Pk = True Then
            MapInfo(UserList(UserIndex).pos.Map).Pk = False
            Call SendData(ToIndex, UserIndex, 0, "||Ahora es zona segura." & FONTTYPE_INFO)
            Exit Sub
        Else
            MapInfo(UserList(UserIndex).pos.Map).Pk = True
            Call SendData(ToIndex, UserIndex, 0, "||Ahora es zona insegura." & FONTTYPE_INFO)
            Exit Sub
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 10)) = "/VERCLIENT" Then
        Call SendData(ToIndex, UserIndex, 0, "||VC > Verificando los clientes de los usuarios.." & FONTTYPE_INFO)
        Dim Chiteros As Integer
        Chiteros = 0
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.UserLogged And UserList(LoopC).flags.ClienteValido = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & UserList(LoopC).name & " con Cliente Invalido." & FONTTYPE_INFO)
                Chiteros = Chiteros + 1
            End If
        Next LoopC
        If Chiteros > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Finalizado. Nro de Clientes Editados: " & Chiteros & FONTTYPE_VENENO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||VERIFICANDO CLIENTES EDITADOS... Finalizado. Nro de Clientes Invalidos: 0" & FONTTYPE_VENENO)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rdata, 9)) = "/DOBACKUP" Then
        Call DoBackUp(True)
        Exit Sub
    End If
    If UCase$(Left$(rdata, 9)) = "/PAUSA" Then
        If haciendoBK Then Exit Sub
        Enpausa = Not Enpausa
        If Enpausa Then
            Call SendData(ToAll, 0, 0, "TL" & 197)
            Call SendData(ToAll, 0, 0, "||Servidor> El mundo ha sido detenido." & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "BKW")
            Call SendData(ToAll, 0, 0, "TM" & "0")
        Else
            Call SendData(ToAll, 0, 0, "TL")
            Call SendData(ToAll, 0, 0, "||Servidor> Juego reanudado." & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "BKW")
            Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).pos.Map).Music)
        End If
        Exit Sub
    End If
    If UCase$(rdata) = "/LIMPIARMUNDO" Then
        If UserList(UserIndex).flags.Privilegios = 3 Then
            Call SendData(ToAll, 0, 0, "||Se realizará una limpieza del Mundo en 1 minuto. Por favor recojan sus pertenencias." & FONTTYPE_VENENO)
            frmMain.Tlimpiar.Enabled = True
            Call LogGM(UserList(UserIndex).name, "Ejecutó una limpieza del Mundo.", True)
        End If
        Exit Sub
    End If
    If UCase$(rdata) = "/LLUVIA" Then
        Lloviendo = Not Lloviendo
        Call SendData(ToAll, 0, 0, "LLU")
        Exit Sub
    End If


    ' Actulizar archivos DAT's
    If UCase$(Left$(rdata, 5)) = "/MOD " Then
        Call LogGM(UserList(UserIndex).name, rdata, False)
        rdata = Right$(rdata, Len(rdata) - 5)
        TIndex = NameIndex(ReadField(1, rdata, 32))
        Arg1 = ReadField(2, rdata, 32)
        Arg2 = ReadField(3, rdata, 32)
        arg3 = ReadField(4, rdata, 32)
        Arg4 = ReadField(5, rdata, 32)
        If TIndex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "1A")
            Exit Sub
        End If
        If UserList(TIndex).flags.Privilegios > 2 And UserIndex <> TIndex Then Exit Sub

        Select Case UCase$(Arg1)
            Case "RAZA"
                If val(Arg2) < 6 Then
                    UserList(TIndex).Raza = val(Arg2)
                    Call DarCuerpoDesnudo(TIndex)
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Alas)
                End If
            Case "JER"
                UserList(UserIndex).Faccion.Jerarquia = 0
            Case "BANDO"
                If val(Arg2) < 3 Then
                    If val(Arg2) > 0 Then Call SendData(ToIndex, TIndex, 0, Mensajes(val(Arg2), 10))
                    UserList(TIndex).Faccion.Bando = val(Arg2)
                    UserList(TIndex).Faccion.BandoOriginal = val(Arg2)
                    If Not PuedeFaccion(TIndex) Then Call SendData(ToIndex, TIndex, 0, "SUFA0")
                    Call UpdateUserChar(TIndex)
                    If val(Arg2) = 0 Then UserList(TIndex).Faccion.Jerarquia = 0
                End If
            Case "SKI"
                If val(Arg2) >= 0 And val(Arg2) <= 100 Then
                    For i = 1 To NUMSKILLS
                        UserList(TIndex).Stats.UserSkills(i) = val(Arg2)
                    Next
                End If
            Case "CLASE"
                i = ClaseIndex(Arg2)
                If i = 0 Then Exit Sub
                UserList(TIndex).Clase = i
                UserList(TIndex).Recompensas(1) = 0
                UserList(TIndex).Recompensas(2) = 0
                UserList(TIndex).Recompensas(3) = 0
                Call SendData(ToIndex, TIndex, 0, "||Ahora eres " & ListaClases(i) & "." & FONTTYPE_INFO)
                If PuedeRecompensa(UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "SURE1")
                Else: Call SendData(ToIndex, UserIndex, 0, "SURE0")
                End If
                If PuedeSubirClase(UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "SUCL1")
                Else: Call SendData(ToIndex, UserIndex, 0, "SUCL0")
                End If

            Case "ORO"
                If val(Arg2) > 10000000 Then Arg2 = 10000000
                UserList(TIndex).Stats.GLD = val(Arg2)
                Call SendUserORO(TIndex)
            Case "EXP"
                If val(Arg2) > 10000000 Then Arg2 = 10000000
                UserList(TIndex).Stats.Exp = val(Arg2)
                Call CheckUserLevel(TIndex)
                Call SendUserEXP(TIndex)
            Case "MEX"
                If val(Arg2) > 10000000 Then Arg2 = 10000000
                UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + val(Arg2)
                Call CheckUserLevel(TIndex)
                Call SendUserEXP(TIndex)
            Case "BODY"
                Call ChangeUserBody(ToMap, 0, UserList(TIndex).pos.Map, TIndex, val(Arg2))
            Case "HEAD"
                Call ChangeUserHead(ToMap, 0, UserList(TIndex).pos.Map, TIndex, val(Arg2))
                UserList(TIndex).OrigChar.Head = val(Arg2)
            Case "PHEAD"
                UserList(TIndex).OrigChar.Head = val(Arg2)
                Call ChangeUserHead(ToMap, 0, UserList(TIndex).pos.Map, TIndex, val(Arg2))
            Case "TOR"
                UserList(TIndex).Faccion.Torneos = val(Arg2)
            Case "QUE"
                UserList(TIndex).Faccion.Quests = val(Arg2)
            Case "NEU"
                UserList(TIndex).Faccion.Matados(Neutral) = val(Arg2)
            Case "CRI"
                UserList(TIndex).Faccion.Matados(Caos) = val(Arg2)
            Case "CIU"
                UserList(TIndex).Faccion.Matados(Real) = val(Arg2)
            Case "HP"
                If val(Arg2) > 999 Then Exit Sub
                UserList(TIndex).Stats.MaxHP = val(Arg2)
                Call SendUserMAXHP(UserIndex)
            Case "MAN"
                If val(Arg2) > 2200 + 800 * Buleano(UserList(TIndex).Clase = MAGO And UserList(TIndex).Recompensas(2) = 2) Then Exit Sub
                UserList(TIndex).Stats.MaxMAN = val(Arg2)
                Call SendUserMAXMANA(UserIndex)
            Case "STA"
                If val(Arg2) > 999 Then Exit Sub
                UserList(TIndex).Stats.MaxSta = val(Arg2)
            Case "HAM"
                UserList(TIndex).Stats.MinHam = val(Arg2)
            Case "SED"
                UserList(TIndex).Stats.MinAGU = val(Arg2)
            Case "ATF"
                If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
                UserList(TIndex).Stats.UserAtributos(fuerza) = val(Arg2)
                UserList(TIndex).Stats.UserAtributosBackUP(fuerza) = val(Arg2)
                Call UpdateFuerzaYAg(TIndex)
            Case "ATI"
                If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
                UserList(TIndex).Stats.UserAtributos(Inteligencia) = val(Arg2)
                UserList(TIndex).Stats.UserAtributosBackUP(Inteligencia) = val(Arg2)
            Case "ATA"
                If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
                UserList(TIndex).Stats.UserAtributos(Agilidad) = val(Arg2)
                UserList(TIndex).Stats.UserAtributosBackUP(Agilidad) = val(Arg2)
                Call UpdateFuerzaYAg(TIndex)
            Case "ATC"
                If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
                UserList(TIndex).Stats.UserAtributos(Carisma) = val(Arg2)
                UserList(TIndex).Stats.UserAtributosBackUP(Carisma) = val(Arg2)
            Case "ATV"
                If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
                UserList(TIndex).Stats.UserAtributos(Constitucion) = val(Arg2)
                UserList(TIndex).Stats.UserAtributosBackUP(Constitucion) = val(Arg2)
            Case "LEVEL"
                If val(Arg2) < 1 Or val(Arg2) > STAT_MAXELV Then Exit Sub
                UserList(TIndex).Stats.ELV = val(Arg2)
                If val(Arg2) < 49 Then
                    UserList(TIndex).Stats.ELU = ELUs(UserList(TIndex).Stats.ELV)
                    Call SendData(ToIndex, TIndex, 0, "5O" & UserList(TIndex).Stats.ELV & "," & UserList(TIndex).Stats.ELU)
                Else
                    UserList(TIndex).Stats.FragsLVL = EFrags(UserList(TIndex).Stats.ELV)
                    Call SendData(ToIndex, TIndex, 0, "5O" & UserList(TIndex).Stats.ELV & "," & UserList(TIndex).Stats.FragsLVL)
                End If

                If PuedeRecompensa(UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "SURE1")
                Else: Call SendData(ToIndex, UserIndex, 0, "SURE0")
                End If
                If PuedeSubirClase(UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "SUCL1")
                Else: Call SendData(ToIndex, UserIndex, 0, "SUCL0")
                End If

            Case Else
                Call SendData(ToIndex, UserIndex, 0, "||Comando inexistente." & FONTTYPE_INFO)
        End Select
        Exit Sub
    End If

    If UCase$(rdata) = "/DATSFULL" Then
        Call SendData(ToAll, 0, 0, "||" & "SERVIDOR:" & UserList(UserIndex).name & " está actulizado todos los archivos DATS" & "~0~255~255~0~0")
        Call LoadSoportes
        Call LoadSini
        Call LoadQuest
        Call CargarPremiosList
        Call CargaNpcsDat
        Call CargarHechizos
        Call LoadArmasHerreria
        Call LoadArmadurasHerreria
        Call LoadEscudosHerreria
        Call LoadCascosHerreria
        Call LoadObjCarpintero
        Call LoadObjSastre
        Call LoadVentas
        Call LoadCasino
        Call LoadOBJData
        Call SendData(ToAll, 0, 0, "||" & "SERVIDOR: Archivos DATS actualizados. Gracias por la espera." & "~0~255~255~0~0")

        Exit Sub
    End If

ErrorHandler:
    If Err.Number = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Comando invalido..." & FONTTYPE_INFO)
    Else
        Call LogErrorUrgente("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).name & " UI:" & UserIndex & " N: " & Err.Number & " D: " & Err.Description)
    End If
End Sub
