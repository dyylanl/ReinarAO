Attribute VB_Name = "modGSec_Servidor"
Option Explicit

'*********************************************
'*********************************************
'********** GSec v1.42 - Anti-cheat **********
'************** GS-Zone (c) 2012 *************
'********** http://www.gs-zone.org ***********
'*********************************************
'*********************************************

' Procedimientos
Public Declare Sub gsCredits Lib "GSec.dll" () ' Abre la ventana de Creditos
Public Declare Sub gsStart Lib "GSec.dll" () ' Inicia la protección
Public Declare Sub gsStop Lib "GSec.dll" () ' Detiene la protección

' Funciones
Public Declare Function gsStatus Lib "GSec.dll" () As Byte ' Devuelve el estado del anticheat
    ' RECOMENDADO: Se recomienda realizar esta función cada 1 seguno o 5 segundos... en un Timer talvez.
    ' ACLARACIÓN: Esta funcion no hace nada especial, solo se fija que esta haciendo el anticheat,
    ' por lo tanto, si se ejecuta una vez cada minuto, no afecta en nada al funcionamiento del anticheat.
    ' Estado:
    ' 0 - Desactivado
    ' 1 - Activado
    ' 2 - Cheat detectado
Public Declare Function gsCheatName Lib "GSec.dll" () As String ' Devuelve el Nombre del cheat asociado a la detección (solo si el estado fue igual 2)
Public Declare Function gsCheatPath Lib "GSec.dll" () As String ' Devuelve el Path del cheat detectado (solo si el estado fue igual 2)
Public Declare Function gsGetGSEC_ID Lib "GSec.dll" () As String  ' Devuelve el ID de identificación unica del usuario

' INSTALACIÓN

' GUÍA BASADA EN FÉNIX

' - PASO 1 -
' En el módulo Declaraciones, buscar:
'   Type UserFlags
' Agregar justo debajo.
'   GSEC_ID As String

'IMPORTANTE: EN LA VERSIÓN 1.42 BETA 2 NO HAY QUE HACER ESTE PASO!
' - PASO 2 -
' En el mismo módulo TCP, buscar:
'   If Left$(rData, 13) = "gIvEmEvAlcOde" Then
' Agregar justo arriba.
'   If Len(rdata) > 3 Then
'   Select Case Left$(rdata, 3)
'   Case "GID"
'   rData = Right$(rData, Len(rData) - 3)
'   ClientChecksum = Right$(rData, Len(rData) - InStrRev(rData, Chr$(126)))
'   rData = Left$(rData, Len(ClientChecksum))
'   If LenB(UserList(UserIndex).flags.GSec_ID) = 0 Then
'   UserList(UserIndex).flags.GSec_ID = rData
'   Else
'   Call CloseSocket(UserIndex, True)
'   End If
'   Exit Sub
'   Case "GAC"
'   rData = Right$(rData, Len(rData) - 3)
'   ClientChecksum = Right$(rData, Len(rData) - InStrRev(rData, Chr$(126)))
'   rData = Left$(rData, Len(ClientChecksum))
'   UserList(UserIndex).flags.Ban = 1
'   Call LogBanFromName("GSec-Anticheat", UserIndex, "ANTICHEAT detectó " & rData)
'   Call SendData(SendTarget.ToAdmins, 0, 0, "||GSec> ANTICHEAT ha baneado a " & UserList(UserIndex).Name & "." & FONTTYPE_SERVER)
'   Call CloseSocket(UserIndex)
'   Exit Sub
'   End Select
'   End If
'   If LenB(UserList(UserIndex).flags.GSec_ID) = 0 Then
'   Call CloseSocket(UserIndex)
'   Exit Sub
'   End If
