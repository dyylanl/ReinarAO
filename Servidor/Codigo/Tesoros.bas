Attribute VB_Name = "Module2"
'******************************************************************************
'Zefron AO v0.1.0
'Mod_Tesoros.bas
'You can contact the programmer of Zefron AO at gabi.13@live.com.ar
'******************************************************************************
 
Public MapaTesoro As Integer
Public RecompenzaTesoro As Integer
Public MapaTesoroMap As Integer
Public MapaTesoroX As Integer
Public MapaTesoroY As Integer
Public TiempoTesoro As Integer
Public TESOROCONTANDO As Boolean
Public SepuedeDesenterrar As Boolean
Public Const LlaveTesoro As Integer = 14 'num de la llave en el obj.dat
Public ObjetoT As Obj
Public objetoCofreAbierto As Obj
 
Public Sub Tesoros()
       
    ObjetoT.Amount = 1
    ObjetoT.OBJIndex = 11 'Cofre Cerrado
   
    objetoCofreAbierto.Amount = 1
    objetoCofreAbierto.OBJIndex = 10 'Cofre abierto
MapaTesoro = RandomNumber(1, 3)
 
If MapaTesoro = 1 Then
    MapaTesoroMap = 36 ' mapa . Les dejo este ejemplo para que se guien
    MapaTesoroX = RandomNumber(20, 80) '  rango de posicion de X. Les dejo este ejemplo para que se guien
    MapaTesoroY = RandomNumber(20, 80) '  rango de posicion de Y. Les dejo este ejemplo para que se guien
ElseIf MapaTesoro = 2 Then
    MapaTesoroMap = 37 'cambien como el anterior por el que quieran
    MapaTesoroX = RandomNumber(20, 80) 'cambien como el anterior por lo que quieran
    MapaTesoroY = RandomNumber(20, 80) 'cambien como el anterior por lo que quieran
ElseIf MapaTesoro = 3 Then
    MapaTesoroMap = 21
    MapaTesoroX = RandomNumber(20, 80)
    MapaTesoroY = RandomNumber(20, 80)
'ElseIf MapaTesoro = 4 Then
 '   MapaTesoroMap = M
  '  MapaTesoroX = RandomNumber(x, x)
   ' MapaTesoroY = RandomNumber(y, y)
End If
    SepuedeDesenterrar = True
    TESOROCONTANDO = True
    TiempoTesoro = 45
    Call SendData(ToAll, 0, 0, "||Apareció un tesoro enterrado en el mapa " & MapaTesoroMap & " en las coordenadas " & MapaTesoroX & ", " & MapaTesoroY & "  El que lo pueda desenterrar ganará premios importantes." & FONTTYPE_INFO)

End Sub
 
 
Public Sub DondeTesoros()
    Call SendData(ToAll, 0, 0, "||El tesoro se encuentra en el mapa " & MapaTesoro & " en las coordenadas " & MapaTesoroX & ", " & MapaTesoroY & " El que lo pueda desenterrar ganará premios importantes." & FONTTYPE_INFO)
End Sub
 
Public Sub CofreAbierto()
Call EraseObj(ToMap, UserIndex, MapaTesoroMap, 10000, MapaTesoroMap, MapaTesoroX, MapaTesoroY)
Call MakeObj(ToMap, 0, MapaTesoroMap, objetoCofreAbierto, MapaTesoroMap, MapaTesoroX, MapaTesoroY)
End Sub
 
