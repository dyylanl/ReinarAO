Attribute VB_Name = "Module1"
Function Encripta(Text As String, Encriptar As Boolean) As String

    On Error GoTo a:

    Dim a() As Integer
    Dim b() As Integer
    Dim Contraseņas(9) As String
    Dim i As Integer
    Dim ii As Integer
    Dim R As String
    Dim CI As Byte
    Dim ss As Integer


    Contraseņas(0) = "pepetuvieja"
    Contraseņas(1) = "soreteinmundo"
    Contraseņas(2) = "alentragaleche"
    Contraseņas(3) = "hermanadealen"
    Contraseņas(4) = "catamonetta"    'amor platonico :3
    Contraseņas(5) = "teamohermosa"    'awww
    Contraseņas(6) = "dylancapo"
    Contraseņas(7) = "catamonettaa"
    Contraseņas(8) = "nachotpaopete"
    Contraseņas(9) = "tpaoserverchoto"

    '********* que contraseņa hay q usar? *********
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

        '****** Separamos la Contraseņa ******
        CI = RandomNumber(0, 9)
        ReDim b(1 To Len(Contraseņas(CI))) As Integer

        For i = 1 To Len(Contraseņas(CI))
            b(i) = Asc(mid(Contraseņas(CI), i, 1))
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

        '****** Separamos la Contraseņa ******
        ReDim b(1 To Len(Contraseņas(CI))) As Integer

        For i = 1 To Len(Contraseņas(CI))
            b(i) = Asc(mid(Contraseņas(CI), i, 1))
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


