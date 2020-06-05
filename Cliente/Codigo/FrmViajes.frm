VERSION 5.00
Begin VB.Form FrmViajes 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Viajes"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Ciudad Perdida"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hildegard"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Adelaide"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lonelerd"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Althalos"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Ciudades"
      ForeColor       =   &H0000C0C0&
      Height          =   3255
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Los viajes cuestan 10.000 Monedas de oro"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   3060
   End
End
Attribute VB_Name = "FrmViajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("#Z")
    Unload Me
End Sub


Private Sub Command2_Click()
    Call SendData("#¥")
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me
End Sub

Private Sub Command3_Click()
    Call SendData("#Ø")
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me
End Sub

Private Sub Command4_Click()
    Call SendData("#X")
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me
End Sub


Private Sub Command5_Click()
    Call SendData("#®")
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me
End Sub

