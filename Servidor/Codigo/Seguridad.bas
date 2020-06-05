Attribute VB_Name = "Seguridad"
'ANTI EDITORES DE PACKET Y HUEVADAS
Sub DataFalsaNo(UserIndex As Integer, EsDataFalsa As Integer)

    If Not EsDataFalsa = 1 Then
        UserList(UserIndex).flags.DataSTRINGGENM = UserList(UserIndex).flags.DataSTRINGGENM + 1

        'reseteo las diferentes entonces .-. (ni ganas de explicar) lo mismo para todas las otras datas.
        UserList(UserIndex).flags.DataJffsdfgdrt = 0
        UserList(UserIndex).flags.DataWEJDJz = 0
        UserList(UserIndex).flags.DatagfsdewS = 0
        UserList(UserIndex).flags.DataUEUSIDx = 0

        If UserList(UserIndex).flags.DataSTRINGGENM > 45 Then    ' Si Envia la data igual 45 veces KB
            Call BaneoAuto(UserIndex)
        End If

    ElseIf Not EsDataFalsa = 2 Then
        UserList(UserIndex).flags.DataJffsdfgdrt = UserList(UserIndex).flags.DataJffsdfgdrt + 1

        UserList(UserIndex).flags.DataSTRINGGENM = 0
        UserList(UserIndex).flags.DataWEJDJz = 0
        UserList(UserIndex).flags.DatagfsdewS = 0
        UserList(UserIndex).flags.DataUEUSIDx = 0

        If UserList(UserIndex).flags.DataJffsdfgdrt > 45 Then    ' Si Envia la data igual 45 veces KB
            Call BaneoAuto(UserIndex)
        End If

    ElseIf Not EsDataFalsa = 3 Then
        UserList(UserIndex).flags.DataWEJDJz = UserList(UserIndex).flags.DataWEJDJz + 1

        UserList(UserIndex).flags.DataSTRINGGENM = 0
        UserList(UserIndex).flags.DataJffsdfgdrt = 0
        UserList(UserIndex).flags.DatagfsdewS = 0
        UserList(UserIndex).flags.DataUEUSIDx = 0

        If UserList(UserIndex).flags.DataWEJDJz > 45 Then    ' Si Envia la data igual 45 veces KB
            Call BaneoAuto(UserIndex)
        End If

    ElseIf Not EsDataFalsa = 4 Then
        UserList(UserIndex).flags.DatagfsdewS = UserList(UserIndex).flags.DatagfsdewS + 1

        UserList(UserIndex).flags.DataSTRINGGENM = 0
        UserList(UserIndex).flags.DataJffsdfgdrt = 0
        UserList(UserIndex).flags.DataWEJDJz = 0
        UserList(UserIndex).flags.DataUEUSIDx = 0

        If UserList(UserIndex).flags.DatagfsdewS > 45 Then    ' Si Envia la data igual 45 veces KB
            Call BaneoAuto(UserIndex)
        End If

    ElseIf Not EsDataFalsa = 5 Then
        UserList(UserIndex).flags.DataUEUSIDx = UserList(UserIndex).flags.DataUEUSIDx + 1

        UserList(UserIndex).flags.DataSTRINGGENM = 0
        UserList(UserIndex).flags.DataJffsdfgdrt = 0
        UserList(UserIndex).flags.DataWEJDJz = 0
        UserList(UserIndex).flags.DatagfsdewS = 0

        If UserList(UserIndex).flags.DataUEUSIDx > 45 Then    ' Si Envia la data igual 45 veces KB
            Call BaneoAuto(UserIndex)
        End If

    End If

End Sub

Sub BaneoAuto(UserIndex As Integer)
    Call SendData(ToAdmins, 0, 0, "||BAN AUTOMÁTICO> " & UserList(UserIndex).name & " por uso de Inyector o cliente editado." & FONTTYPE_FIGHT)
    Call LogBan(UserIndex, UserIndex, "Uso de cheats/programas externos.")
    UserList(UserIndex).flags.Ban = 1
    Call CloseSocket(UserIndex)
End Sub
'ANTI EDITORES DE PACKET Y HUEVADAS
'EnCriptacion
' Text1.Text = THeDEnCripTe("DATO STRING", "asdasd")
Function THeDEnCripTe(ByVal S As String, ByVal P As String) As String
    Dim i As Integer, R As String
    Dim C1 As Integer, C2 As Integer
    R = ""
    If Len(P) > 0 Then
        For i = 1 To Len(S)
            C1 = Asc(mid(S, i, 1))
            If i > Len(P) Then
                C2 = Asc(mid(P, i Mod Len(P) + 1, 1))
            Else
                C2 = Asc(mid(P, i, 1))
            End If
            C1 = C1 + C2 + 64
            If C1 > 255 Then C1 = C1 - 256
            R = R + Chr(C1)
        Next i
    Else
        R = S
    End If
    THeDEnCripTe = R
End Function
'EnCriptacion


'ENCRIPTT

Private Function MamasiTEEX(X As Integer) As String
    If X > 9 Then
        MamasiTEEX = Chr(X + 55)
    Else
        MamasiTEEX = CStr(X)
    End If
End Function
Function DesteEncripTE(DataValue As Variant) As Variant

    Dim X As Long
    Dim temp As String
    Dim TempNum As Integer
    Dim TempChar As String
    Dim TempChar2 As String

    For X = 1 To Len(DataValue)
        TempChar2 = mid(DataValue, X, 1)
        TempNum = Int(Asc(TempChar2) / 16)

        If ((TempNum * 16) < Asc(TempChar2)) Then

            TempChar = MamasiTEEX(Asc(TempChar2) - (TempNum * 16))
            temp = temp & MamasiTEEX(TempNum) & TempChar
        Else
            temp = temp & MamasiTEEX(TempNum) & "0"

        End If
    Next X


    DesteEncripTE = temp
End Function
Private Function MoveEltoto(X As String) As Integer

    Dim X1 As String
    Dim X2 As String
    Dim temp As Integer

    X1 = mid(X, 1, 1)
    X2 = mid(X, 2, 1)

    If IsNumeric(X1) Then
        temp = 16 * Int(X1)
    Else
        temp = (Asc(X1) - 55) * 16
    End If

    If IsNumeric(X2) Then
        temp = temp + Int(X2)
    Else
        temp = temp + (Asc(X2) - 55)
    End If

    ' retorno
    MoveEltoto = temp

End Function

'ENCRIPTT
