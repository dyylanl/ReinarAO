VERSION 5.00
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Carteles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   6735
      Begin VB.PictureBox PictureOcultarse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   335
         Left            =   120
         MouseIcon       =   "frmOpciones.frx":324A
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   36
         Top             =   240
         Width           =   335
      End
      Begin VB.PictureBox PictureNoHayNada 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   335
         Left            =   120
         MouseIcon       =   "frmOpciones.frx":3554
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   35
         Top             =   1200
         Width           =   335
      End
      Begin VB.PictureBox PictureMenosCansado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   335
         Left            =   3840
         MouseIcon       =   "frmOpciones.frx":385E
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   34
         Top             =   1200
         Width           =   335
      End
      Begin VB.PictureBox PictureVestirse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   335
         Left            =   3840
         MouseIcon       =   "frmOpciones.frx":3B68
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   33
         Top             =   240
         Width           =   335
      End
      Begin VB.PictureBox PictureRecuMana 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   335
         Left            =   120
         MouseIcon       =   "frmOpciones.frx":3E72
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   32
         Top             =   720
         Width           =   335
      End
      Begin VB.PictureBox PictureSanado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   335
         Left            =   3840
         MouseIcon       =   "frmOpciones.frx":417C
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   31
         Top             =   720
         Width           =   335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   30
         Text            =   "Meditación"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   29
         Text            =   "Ocultarse"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4320
         TabIndex        =   28
         Text            =   "Abrigarse"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Text            =   "Menos Cansado"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4320
         TabIndex        =   26
         Text            =   "Has Sanado"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   600
         TabIndex        =   25
         Text            =   "No Hay Nada Aquí"
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jugabilidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   3015
      Begin VB.OptionButton FPS1 
         Caption         =   "18 FPS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton FPS2 
         Caption         =   "32 FPS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton FPS3 
         Caption         =   "64 FPS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton FPS4 
         Caption         =   "FPS Libres"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Partículas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   3840
      TabIndex        =   10
      Top             =   2880
      Width           =   3015
      Begin VB.CheckBox CheckHechiz 
         Caption         =   "Hechizos con Particulas"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   225
         Index           =   0
         Left            =   2445
         TabIndex        =   13
         Top             =   720
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option2"
         Height          =   255
         Index           =   1
         Left            =   1365
         TabIndex        =   12
         Top             =   720
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option3"
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   11
         Top             =   720
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bajo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   18
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Medio"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   17
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Densidad de Particulas en Hechizos"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Alto"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   15
         Top             =   960
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Audio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6735
      Begin VB.CheckBox chkop 
         Caption         =   "Sonidos Ambientales"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonidos"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Música"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Manual"
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   6600
      Width           =   1575
   End
   Begin VB.OptionButton DesactivarGlobal 
      Caption         =   "Desactivar Global"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.OptionButton ActivarGlobal 
      Caption         =   "Activar Global"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdKeys 
      Caption         =   "Configurar teclas"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Desactivar Consola General"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Left            =   5520
      Top             =   -2040
   End
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   -2640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir y Guardar"
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   6600
      Width           =   1575
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private ret As Long
Private cont2 As Integer

Private Sub CheckHechiz_Click()
    Opciones.Particulas = Val(frmOpciones.CheckHechiz.value)
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG VIDEO", "Particulas", Val(frmOpciones.CheckHechiz.value))
End Sub

Private Sub cmdKeys_Click()
    Unload Me
    Call frmCustomKeys.Show(vbModeless, frmMain)
End Sub

Private Sub Command10_Click()
    Novedades.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Me.Visible = False
End Sub

Private Sub Command8_Click()
    If Opciones.ConsolaActivada = True Then
        AddtoRichTextBox frmMain.rectxt, "Consola General Desactivada.", 255, 0, 0, True, False
        Opciones.ConsolaActivada = False
        'Aca donde dice "Command8" ponganle el nombre de su commandbutton
        Command8.Caption = "Activar Consola General"
        Exit Sub
    Else

        Command8.Caption = "Desactivar Consola General"
        Opciones.ConsolaActivada = True
        AddtoRichTextBox frmMain.rectxt, "Consola General Activada.", 0, 255, 0, True, False
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    cont2 = 255
    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    Timer1.Interval = 1
    Timer2.Interval = 1
    Timer2.Enabled = False
    Timer1.Enabled = True

    'Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\OpcionesDelJuego.gif")
    
    If Opciones.Particulas = 1 Then
        CheckHechiz.value = vbChecked
    Else
        CheckHechiz.value = vbUnchecked
    End If
    If Opciones.Audio = 1 Then
        chkop(1).value = vbChecked
    Else
        chkop(1).value = vbUnchecked
    End If
    If Opciones.sMusica = CONST_DESHABILITADA Then
        chkop(0).value = vbUnchecked
    Else
        chkop(0).value = vbChecked
    End If
    If Opciones.Ambient = 0 Then
        chkop(2).value = vbUnchecked
    Else
        chkop(2).value = vbChecked
    End If
    If Opciones.CartelOcultarse = 1 Then
        PictureOcultarse.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        PictureOcultarse.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If

    If Opciones.CartelMenosCansado = 1 Then
        PictureMenosCansado.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        PictureMenosCansado.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If

    If Opciones.CartelVestirse = 1 Then
        PictureVestirse.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        PictureVestirse.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If

    If Opciones.CartelNoHayNada = 1 Then
        PictureNoHayNada.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        PictureNoHayNada.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If

    If Opciones.CartelRecuMana = 1 Then
        PictureRecuMana.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        PictureRecuMana.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If

    If Opciones.CartelSanado = 1 Then
        PictureSanado.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        PictureSanado.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If

    If Opciones.FPSConfig = 1 Then
        frmOpciones.FPS1.value = True
    ElseIf Opciones.FPSConfig = 2 Then
        frmOpciones.FPS2.value = True
    ElseIf Opciones.FPSConfig = 3 Then
        frmOpciones.FPS3.value = True
    ElseIf Opciones.FPSConfig = 4 Then
        frmOpciones.FPS4.value = True
    End If
End Sub
Private Sub chkop_Click(Index As Integer)
Call Sound.Sound_Play(SND_CLICK)
 
    Select Case Index
        Case 0
             
            If chkop(Index).value = vbUnchecked Then
                Sound.Music_Stop
                Opciones.sMusica = CONST_DESHABILITADA
                'scrMidi.Enabled = False
            Else
                Opciones.sMusica = CONST_MP3
                'scrMidi.Enabled = True
                Sound.Music_Play
            End If
 
        Case 1
 
            If chkop(Index).value = vbUnchecked Then
                chkop(2).Enabled = False
                'scrAmbient.Enabled = False
                'scrVolume.Enabled = False
                Opciones.Audio = 0
            Else
                Opciones.Audio = 1
                chkop(2).Enabled = True
                'scrVolume.Enabled = True
            End If
     
        Case 2
     
            If chkop(Index).value = vbUnchecked Then
                Opciones.Ambient = 0
                Call Sound.Sound_Stop_All
            Else
                Opciones.Ambient = 1
                'scrAmbient.Enabled = True
                Call Sound.Ambient_Load(Sound.AmbienteActual, Opciones.AmbientVol)
                Call Sound.Ambient_Play
            End If
    End Select
End Sub
 
Private Sub scrMidi_Change()
 
    If Opciones.sMusica <> CONST_DESHABILITADA Then
        'Sound.Music_Volume_Set scrMidi.value
        'Sound.VolumenActualMusicMax = scrMidi.value
        Opciones.MusicVolume = Sound.VolumenActualMusicMax
    End If
 
End Sub
 
Private Sub scrAmbient_Change()
    If Opciones.Ambient = 1 Then
        'Sound.VolumenActualAmbient_set scrAmbient.value
        Opciones.AmbientVol = Sound.VolumenActualAmbient
    End If
End Sub
 
Private Sub scrVolume_Change()
 
If Opciones.Audio = 1 Then
    'Sound.VolumenActual = scrVolume.value
    Opciones.FXVolume = Sound.VolumenActual
End If
 
End Sub

Private Sub Option1_Click(Index As Integer)
    Opciones.bGraphics = Index
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG VIDEO", "Densidad", Val(Index))
End Sub
Private Sub PictureMenosCansado_Click()

    If Opciones.CartelMenosCansado = 0 Then
        Opciones.CartelMenosCansado = 1
        PictureMenosCansado.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        Opciones.CartelMenosCansado = 0
        PictureMenosCansado.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If

    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "MenosCansado", Str(Opciones.CartelMenosCansado))

End Sub

Private Sub PictureNoHayNada_Click()
    If Opciones.CartelNoHayNada = 0 Then
        Opciones.CartelNoHayNada = 1
        PictureNoHayNada.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        Opciones.CartelNoHayNada = 0
        PictureNoHayNada.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "NoHayNada", Str(Opciones.CartelNoHayNada))

End Sub

Private Sub PictureOcultarse_Click()

    If Opciones.CartelOcultarse = 0 Then
        Opciones.CartelOcultarse = 1
        PictureOcultarse.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        Opciones.CartelOcultarse = 0
        PictureOcultarse.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "Ocultarse", Str(Opciones.CartelOcultarse))
End Sub

Private Sub PictureRecuMana_Click()
    If Opciones.CartelRecuMana = 0 Then
        Opciones.CartelRecuMana = 1
        PictureRecuMana.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        Opciones.CartelRecuMana = 0
        PictureRecuMana.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "RecuMana", Str(Opciones.CartelRecuMana))

End Sub

Private Sub PictureSanado_Click()
    If Opciones.CartelSanado = 0 Then
        Opciones.CartelSanado = 1
        PictureSanado.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        Opciones.CartelSanado = 0
        PictureSanado.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "Sanado", Str(Opciones.CartelSanado))

End Sub

Private Sub PictureVestirse_Click()
    If Opciones.CartelVestirse = 0 Then
        Opciones.CartelVestirse = 1
        PictureVestirse.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick1.gif")
    Else
        Opciones.CartelVestirse = 0
        PictureVestirse.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\tick2.gif")
    End If
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CARTELES", "Vestirse", Str(Opciones.CartelVestirse))

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving = False And Button = vbLeftButton Then

        Dx3 = X

        dy = Y

        bmoving = True

    End If



End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving And ((X <> Dx3) Or (Y <> dy)) Then

        Move Left + (X - Dx3), Top + (Y - dy)

    End If



End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then

        bmoving = False

    End If



End Sub
Private Sub Timer1_Timer()
    Static Cont As Integer
    Cont = Cont + 5
    If Cont > 255 Then
        Cont = 0
        Timer1.Enabled = False
    Else
        SetLayeredWindowAttributes Me.hwnd, 0, Cont, LWA_ALPHA
    End If
End Sub

Private Sub Timer2_Timer()
    cont2 = cont2 - 5
    If cont2 < 0 Then
        Timer2.Enabled = False
        End
    Else
        SetLayeredWindowAttributes Me.hwnd, 0, cont2, LWA_ALPHA
    End If
End Sub
Private Sub ActivarGlobal_Click()
    Call SendData("GLB")
End Sub

Private Sub DesactivarGlobal_Click()
    Call SendData("GLC")
End Sub

'FPS
Private Sub FPS1_Click()
    Opciones.FPSConfig = 1
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG VIDEO", "FPS", "1")

End Sub

Private Sub FPS2_Click()
    Opciones.FPSConfig = 2
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG VIDEO", "FPS", "2")

End Sub

Private Sub FPS3_Click()
    Opciones.FPSConfig = 3
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG VIDEO", "FPS", "3")

End Sub

Private Sub FPS4_Click()
    Opciones.FPSConfig = 4
    Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG VIDEO", "FPS", "4")

End Sub
'FPS
