Attribute VB_Name = "Module4"
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

    Contraseñas(0) = Chr$(112) & Chr$(101) & Chr$(112) & Chr$(101) & Chr$(116) & Chr$(117) & Chr$(118) & Chr$(105) _
                     & Chr$(101) & Chr$(106) & Chr$(97)
    Contraseñas(1) = Chr$(115) & Chr$(111) & Chr$(114) & Chr$(101) & Chr$(116) & Chr$(101) & Chr$(105) & Chr$(110) _
                     & Chr$(109) & Chr$(117) & Chr$(110) & Chr$(100) & Chr$(111)
    Contraseñas(2) = Chr$(97) & Chr$(108) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(114) & Chr$(97) & Chr$(103) _
                     & Chr$(97) & Chr$(108) & Chr$(101) & Chr$(99) & Chr$(104) & Chr$(101)
    Contraseñas(3) = Chr$(104) & Chr$(101) & Chr$(114) & Chr$(109) & Chr$(97) & Chr$(110) & Chr$(97) & Chr$(100) _
                     & Chr$(101) & Chr$(97) & Chr$(108) & Chr$(101) & Chr$(110)
    Contraseñas(4) = Chr$(99) & Chr$(97) & Chr$(116) & Chr$(97) & Chr$(109) & Chr$(111) & Chr$(110) & Chr$(101) & Chr$(116) _
                     & Chr$(116) & Chr$(97)
    Contraseñas(5) = Chr$(116) & Chr$(101) & Chr$(97) & Chr$(109) & Chr$(111) & Chr$(104) & Chr$(101) & Chr$(114) _
                     & Chr$(109) & Chr$(111) & Chr$(115) & Chr$(97)
    Contraseñas(6) = Chr$(100) & Chr$(121) & Chr$(108) & Chr$(97) & Chr$(110) & Chr$(99) & Chr$(97) & Chr$(112) & Chr$(111) _

Contraseñas(7) = Chr$(99) & Chr$(97) & Chr$(116) & Chr$(97) & Chr$(109) & Chr$(111) & Chr$(110) & Chr$(101) & Chr$(116) _
                     & Chr$(116) & Chr$(97) & Chr$(97)
    Contraseñas(8) = Chr$(110) & Chr$(97) & Chr$(99) & Chr$(104) & Chr$(111) & Chr$(116) & Chr$(112) & Chr$(97) _
                     & Chr$(111) & Chr$(112) & Chr$(101) & Chr$(116) & Chr$(101)
    Contraseñas(9) = Chr$(116) & Chr$(112) & Chr$(97) & Chr$(111) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(118) _
                     & Chr$(101) & Chr$(114) & Chr$(99) & Chr$(104) & Chr$(111) & Chr$(116) & Chr$(111)



    '********* que contraseña hay q usar? *********
    If Not Encriptar Then
        CI = Val(Asc(Left(Text, 1))) - 10
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

