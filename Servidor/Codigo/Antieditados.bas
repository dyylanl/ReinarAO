Attribute VB_Name = "Module1"
Function Encripta(Text As String, Encriptar As Boolean) As String

    On Error GoTo a:

    Dim a() As Integer
    Dim b() As Integer
    Dim Contraseñas(9) As String
    Dim i As Integer
    Dim ii As Integer
    Dim R As String
    Dim CI As Byte
    Dim ss As Integer


    Contraseñas(0) = "pepetuvieja"
    Contraseñas(1) = "soreteinmundo"
    Contraseñas(2) = "alentragaleche"
    Contraseñas(3) = "hermanadealen"
    Contraseñas(4) = "catamonetta"    'amor platonico :3
    Contraseñas(5) = "teamohermosa"    'awww
    Contraseñas(6) = "dylancapo"
    Contraseñas(7) = "catamonettaa"
    Contraseñas(8) = "nachotpaopete"
    Contraseñas(9) = "tpaoserverchoto"

    '********* que contraseña hay q usar? *********
    If Not Encriptar Then
        CI = val(Asc(Left(Text, 1))) - 10
        Text = Right(Text, Len(Text) - 1)
    End If
    '**********************************************

    'para no llamar a cada rato a la function
    ss = Len(Text)

    'Por las dudas
    If ss <= 0 Then Exit Function

    ReDim a(1 To ss) As Integer

    For i = 1 To ss
        a(i) = Asc(mid(Text, i, 1))
    Next i


    If Encriptar Then

        '****** Separamos la Contraseña ******
        CI = RandomNumber(0, 9)
        ReDim b(1 To Len(Contraseñas(CI))) As Integer

        For i = 1 To Len(Contraseñas(CI))
            b(i) = Asc(mid(Contraseñas(CI), i, 1))
        Next i
        '*************************************

        For i = 1 To ss
            If ii >= UBound(b) Then ii = 0
            ii = ii + 1
            a(i) = a(i) + b(ii)
            If a(i) > 255 Then a(i) = a(i) - 255
            R = R + Chr(a(i))
        Next i

        Encripta = Chr(CI + 10) & R

    Else

        '****** Separamos la Contraseña ******
        ReDim b(1 To Len(Contraseñas(CI))) As Integer

        For i = 1 To Len(Contraseñas(CI))
            b(i) = Asc(mid(Contraseñas(CI), i, 1))
        Next i
        '*************************************

        For i = 1 To ss
            If ii >= UBound(b) Then ii = 0
            ii = ii + 1
            a(i) = a(i) - b(ii)
            If a(i) < 0 Then
                a(i) = a(i) + 255
            End If
            R = R + Chr(a(i))
        Next i

        Encripta = R

    End If

a:

End Function


