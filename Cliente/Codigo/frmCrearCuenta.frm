VERSION 5.00
Begin VB.Form frmCrearAccount 
   BorderStyle     =   0  'None
   Caption         =   "Crear Cuenta"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Mail 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   2865
      Width           =   2620
   End
   Begin VB.TextBox RePass 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2295
      Width           =   2620
   End
   Begin VB.TextBox Pass 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   2620
   End
   Begin VB.TextBox Nombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3120
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1320
      Width           =   2620
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3480
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   960
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Crear_Cuenta.bmp")
    MsgBox "Deberas ingresar tu MAIL correcto, de lo contrario, no tendras respuesta de los Game Master en tus Soportes"
End Sub

Private Sub Image1_Click()
   Call Sound.Sound_Play(SND_CLICK)
    If Opciones.sMusica <> CONST_DESHABILITADA Then
        If Opciones.sMusica <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_VolverInicio
            Sound.Fading = 350
        End If
    End If
    frmMain.Socket1.Disconnect
    frmConnect.MousePointer = 1
    frmConnect.Show
    Unload Me
End Sub

Private Sub Image2_Click()
    If Pass <> RePass Then
        MsgBox "Lass passwords que tipeo no coinciden", , "MD5 Changed Info Tip"
        Exit Sub
    End If

    If Not CheckMailString(Mail) Then
        MsgBox "Direccion de mail invalida."
        Exit Sub
    End If

    If Nombre = "" Or Pass = "" Or RePass = "" Or Mail = "" Then
        MsgBox "Completa todo!"
        Exit Sub
    End If

    Pass = MD5String(Pass.Text)
    EstadoLogin = CrearAccount
    Call Login
    frmConnect.Show
    Unload Me
End Sub

