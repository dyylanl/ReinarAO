Attribute VB_Name = "Mod_Resolucion"
Option Explicit

Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const CDS_TEST As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1

Private Type typDevMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private oldResHeight As Long
Private oldResWidth As Long
Private oldDepth As Integer
Private oldFrequency As Long


Private Declare Function EnumDisplaySettings Lib "USER32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "USER32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Public Sub SetResolucion()
'***************************************************
'Autor: Unknown
'Last Modification: 03/29/08
'Changes the display resolution if needed.
'Last Modified By: Juan Mart�n Sotuyo Dodero (Maraxus)
' 03/29/2008: Maraxus - Retrieves current settings storing display depth and frequency for proper restoration.
'***************************************************
    Dim lRes As Long
    Dim MidevM As typDevMODE
    Dim CambiarResolucion As Boolean
   
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
   
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
   
    'If Opciones.NoRes Then
        CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
    'Else
    '    CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
    'End If
   
    'If CambiarResolucion Then
       
        With MidevM
            oldDepth = .dmBitsPerPel
            oldFrequency = .dmDisplayFrequency
           
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 32
        End With
       
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
    'Else
    '    bNoResChange = True
    'End If
End Sub