Attribute VB_Name = "ModHash"
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
'elpresi@Dragoonao.com.ar
'www.Dragoonao.com.ar

Option Explicit
Public Function GenHash(filename As String) As String

    Dim cStream As New cBinaryFileStream
    Dim cCRC32 As New cCRC32
    Dim lCRC32 As Long

    cStream.File = filename
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    GenHash = Hex$(lCRC32)

End Function

'ENCRIPTACION
' Text1.Text = THeEnCripTe(Text1.Text, "asdasd")
Function THeEnCripTe(ByVal s As String, ByVal P As String) As String
    Dim i As Integer, R As String
    Dim C1 As Integer, C2 As Integer
    R = ""
    If Len(P) > 0 Then
        For i = 1 To Len(s)
            C1 = Asc(mid(s, i, 1))
            If i > Len(P) Then
                C2 = Asc(mid(P, i Mod Len(P) + 1, 1))
            Else
                C2 = Asc(mid(P, i, 1))
            End If
            C1 = C1 - C2 - 64
            If Sgn(C1) = -1 Then C1 = 256 + C1
            R = R + Chr(C1)
        Next i
    Else
        R = s
    End If
    THeEnCripTe = R
End Function
'ENCRIPTACION

'ENCRIPTT

Private Function MamasiTEEX(X As Integer) As String
    If X > 9 Then
        MamasiTEEX = Chr(X + 55)
    Else
        MamasiTEEX = CStr(X)
    End If
End Function
Private Function MoveEltoto(X As String) As Integer

    Dim X1 As String
    Dim X2 As String
    Dim Temp As Integer

    X1 = mid(X, 1, 1)
    X2 = mid(X, 2, 1)

    If IsNumeric(X1) Then
        Temp = 16 * Int(X1)
    Else
        Temp = (Asc(X1) - 55) * 16
    End If

    If IsNumeric(X2) Then
        Temp = Temp + Int(X2)
    Else
        Temp = Temp + (Asc(X2) - 55)
    End If

    ' retorno
    MoveEltoto = Temp

End Function

Function TeEncripTE(DataValue As Variant) As Variant

    Dim X As Long
    Dim Temp As String
    Dim HexByte As String

    For X = 1 To Len(DataValue) Step 2

        HexByte = mid(DataValue, X, 2)
        Temp = Temp & Chr(MoveEltoto(HexByte))

    Next X
    ' retorno
    TeEncripTE = Temp

End Function
'ENCRIPTT
