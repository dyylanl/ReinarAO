VERSION 5.00
Begin VB.Form frmNobleza 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nobleza"
   ClientHeight    =   4320
   ClientLeft      =   105
   ClientTop       =   315
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   4440
      Picture         =   "frmNobleza.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   600
      Width           =   540
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3120
      Picture         =   "frmNobleza.frx":1044
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   600
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1800
      Picture         =   "frmNobleza.frx":1C86
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   600
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   480
      Picture         =   "frmNobleza.frx":28C8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   600
      Width           =   540
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hacerme Noble"
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Construir"
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Construir"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Contruir"
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Construir"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNobleza.frx":350A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   4935
   End
End
Attribute VB_Name = "frmNobleza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("/ESPADA")
End Sub

Private Sub Command2_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("/ARMADURA")
End Sub

Private Sub Command3_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("/ESCUDO")
End Sub

Private Sub Command4_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("/ANILLO")
End Sub

Private Sub Command5_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("/NOBLE")
End Sub

