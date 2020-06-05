Attribute VB_Name = "modGSec"
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
' En el módulo General, buscar:
'   Unload frmCargando
' Agregar justo debajo.
'   Call gsStart

' - PASO 2 -
' En el mismo módulo General, buscar:
'   If DirectX.TickCount - lFrameTimer > 1000 Then
' Agregar justo debajo.
'   loopc = gsStatus
'   If loopc <> 0 Then
'   If loopc = 2 Then
'   If Connected = True Then Call SendData(gsInformar)
'   Sleep 5
'   prgRun = False
'   End If
'   Else
'   prgRun = False
'   End If

' - PASO 3 -
' En el mismo módulo General, buscar:
'   Call UnloadAllForms
' Agregar justo arriba.
'    Call gsStop

' - PASO 4 -
' En el formulario frmMain, buscar:
' If EstadoLogin = CrearNuevoPj Then
' Agregar justo arriba.
' Call SendData(gsEnviarID)

'IMPORTANTE: EN LA VERSIÓN 1.42 BETA 2 NO HAY QUE HACER ESTE PASO!
' - PASO 5 -
' En el módulo Mod_TCP, buscar:
'   If Left$(Rdata, 1) = "Ç" Then Rdata = (Right$(Rdata, Len(Rdata) - 1))
' Agregar justo arriba.
'    Call gsProcesar(RData)

Public Function gsInformar() As String
    ' Informa del cheat detectado al servidor!
    gsInformar = "GAC" & gsCheatName() & "~" & gsCheatPath()
End Function

Public Function gsEnviarID() As String
    ' Le envia el GSEC_ID al servidor...
    'Dim GSEC_ID As String * 32
    'GSEC_ID = gsGetGSEC_ID()
    'gsEnviarID = "GID" & GSEC_ID
End Function
