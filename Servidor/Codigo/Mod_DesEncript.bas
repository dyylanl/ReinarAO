Attribute VB_Name = "Mod_DesEncript"
    Option Explicit
     
    Dim RandomNum As Integer
     
    Public Function Encriptar(ByVal Cadena As String) As String
    Dim i As Long
     
    RandomNum = 99 * Rnd
    If RandomNum < 10 Then RandomNum = 10
     
    For i = 1 To Len(Cadena)
    Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) + RandomNum)
    Next i
     
    Encriptar = Cadena & Chr$(Asc(Left$(RandomNum, 1)) + 10) & Chr$(Asc(Right$(RandomNum, 1)) + 10)
     
    'DoEvents
     
    End Function
     
    Public Function DesEncriptar(ByVal Cadena As String) As String
    Dim i As Long, NumDesencriptar As String
     
    NumDesencriptar = Chr$(Asc(Left$((Right(Cadena, 2)), 1)) - 10) & Chr$(Asc(Right$((Right(Cadena, 2)), 1)) - 10)
     
    Cadena = (Left$(Cadena, Len(Cadena) - 2))
     
    For i = 1 To Len(Cadena)
     
    Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) - NumDesencriptar)
    Next i
    DesEncriptar = Cadena
    'DoEvents
     
    End Function
