VERSION 5.00
Begin VB.Form frmParty 
   BorderStyle     =   0  'None
   Caption         =   "Party"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListaIntegrantes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      ItemData        =   "Party.frx":0000
      Left            =   720
      List            =   "Party.frx":0007
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No te encuentras en ningún party. Para formar uno, clickea en el botón ""Invitar al party""."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3600
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Salir 
      Height          =   360
      Left            =   810
      MouseIcon       =   "Party.frx":001D
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   2370
   End
   Begin VB.Image Echar 
      Height          =   360
      Left            =   840
      MouseIcon       =   "Party.frx":0327
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   2370
   End
   Begin VB.Image Invitar 
      Height          =   360
      Left            =   840
      MouseIcon       =   "Party.frx":0631
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   2355
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()


    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\sistemaparty.gif")
    Invitar.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\invitar.gif")
    Echar.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\echarparty.gif")
    Salir.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\salirparty.gif")

End Sub

Private Sub Salir_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("PARSAL")
    Unload frmParty
End Sub
Private Sub Echar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("PARECH" & frmParty.ListaIntegrantes.Text)
    Unload frmParty
End Sub
Private Sub Invitar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("UK" & Invita)
    Unload frmParty
End Sub

Private Sub Image1_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me
End Sub
