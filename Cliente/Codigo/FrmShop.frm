VERSION 5.00
Begin VB.Form frmShop 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Canjeo"
   ClientHeight    =   6000
   ClientLeft      =   420
   ClientTop       =   315
   ClientWidth     =   5580
   Icon            =   "FrmShop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2640
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox lDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.ListBox ListaPremios 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   4155
      Left            =   195
      TabIndex        =   0
      Top             =   625
      Width           =   2295
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lPuntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   980
      Width           =   645
   End
   Begin VB.Label lAtaque 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   1410
      Width           =   735
   End
   Begin VB.Label lDef 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   1850
      Width           =   735
   End
   Begin VB.Label lAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lDM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   2700
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Requiere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   4890
      Width           =   885
   End
   Begin VB.Image bSalir 
      Height          =   435
      Left            =   2880
      Top             =   5280
      Width           =   2310
   End
   Begin VB.Image bAceptar 
      Height          =   435
      Left            =   240
      Top             =   5280
      Width           =   2310
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("SPX" & ListaPremios.ListIndex + 1)
End Sub

Private Sub ListaPremios_Click()
    If ListaPremios.ListIndex + 1 <> 0 Then
        Call SendData("IPX" & ListaPremios.ListIndex + 1)
    End If
End Sub

Private Sub bSalir_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me
End Sub

Private Sub Form_Load()

    Call SendData("IPX" & ListaPremios.ListIndex + 1)
    bAceptar.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_BcanjearN.jpg")
    bSalir.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_BsalirN.jpg")
    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_main.jpg")
End Sub

Private Sub baceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bAceptar.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_BcanjearA.jpg")
End Sub

Private Sub baceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bAceptar.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_BcanjearI.jpg")
End Sub

Private Sub bsalir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bSalir.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_BsalirA.jpg")
End Sub

Private Sub bsalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bSalir.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_BsalirI.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bSalir.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_BsalirN.jpg")
    bAceptar.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Canjear_BcanjearN.jpg")
End Sub

