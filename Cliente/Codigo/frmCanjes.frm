VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Sistema de Canje"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bajos"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6000
      TabIndex        =   21
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "25 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Image Image39 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":0000
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4200
      TabIndex        =   19
      Top             =   5520
      Width           =   1680
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "40 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   960
      TabIndex        =   18
      Top             =   6120
      Width           =   1680
   End
   Begin VB.Image Image38 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":0C42
      Top             =   5400
      Width           =   480
   End
   Begin VB.Image Image37 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":11FE
      Top             =   6000
      Width           =   480
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "30 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "30 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "30 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "20 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "10 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "30 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "30 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "30 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "20 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "20 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "4 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "4 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "45 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "45 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "45 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "15 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "15 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "15 Puntos de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image27 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":1A40
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image Image26 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":2684
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Image21 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":32C8
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image20 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":3F0C
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Image19 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":4B4E
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image18 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":538E
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image17 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":5968
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image16 
      Height          =   480
      Left            =   3480
      Picture         =   "frmCanjes.frx":61AB
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image13 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":69EE
      Top             =   5400
      Width           =   480
   End
   Begin VB.Image Image12 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":7632
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image11 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":8276
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":8AB8
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":96FA
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":9F3C
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":AB80
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":B3C2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":C004
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   240
      Picture         =   "frmCanjes.frx":C846
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image10_Click()
Call SendData("/CANJEO T7")
End Sub

Private Sub Image11_Click()
Call SendData("/CANJEO T8")
End Sub

Private Sub Image12_Click()
Call SendData("/CANJEO T9")
End Sub

Private Sub Image13_Click()
Call SendData("/CANJEO T10")
End Sub

Private Sub Image14_Click()
Call SendData("/CANJEO T17")
End Sub

Private Sub Image15_Click()
Call SendData("/CANJEO T18")
End Sub

Private Sub Image16_Click()
Call SendData("/CANJEO T11")
End Sub

Private Sub Image17_Click()
Call SendData("/CANJEO T12")
End Sub

Private Sub Image18_Click()
Call SendData("/CANJEO T13")
End Sub

Private Sub Image19_Click()
Call SendData("/CANJEO T14")
End Sub

Private Sub Image20_Click()
Call SendData("/CANJEO T15")
End Sub

Private Sub Image21_Click()
Call SendData("/CANJEO T16")
End Sub

Private Sub Image26_Click()
Call SendData("/CANJEO T19")
End Sub

Private Sub Image27_Click()
Call SendData("/CANJEO T20")
End Sub

Private Sub Image28_Click()
Call SendData("/CANJEO T24")
End Sub

Private Sub Image29_Click()
Call SendData("/CANJEO T25")
End Sub

Private Sub Image30_Click()
Call SendData("/CANJEO T26")
End Sub

Private Sub Image31_Click()
Call SendData("/CANJEO T29")
End Sub

Private Sub Image32_Click()
Call SendData("/CANJEO T27")
End Sub

Private Sub Image33_Click()
Call SendData("/CANJEO T28")
End Sub

Private Sub Image34_Click()
Call SendData("/CANJEO T30")
End Sub

Private Sub Image35_Click()
Call SendData("/CANJEO T31")
End Sub

Private Sub Image36_Click()
Call SendData("/CANJEO T32")
End Sub

Private Sub Image37_Click()
Call SendData("/CANJEO T21")
End Sub

Private Sub Image38_Click()
Call SendData("/CANJEO T22")
End Sub

Private Sub Image39_Click()
Call SendData("/CANJEO T23")
End Sub

Private Sub Image4_Click()
Call SendData("/CANJEO T1")
End Sub

Private Sub Image5_Click()
Call SendData("/CANJEO T2")
End Sub

Private Sub Image6_Click()
Call SendData("/CANJEO T3")
End Sub

Private Sub Image7_Click()
Call SendData("/CANJEO T4")
End Sub

Private Sub Image8_Click()
Call SendData("/CANJEO T5")
End Sub

Private Sub Image9_Click()
Call SendData("/CANJEO T6")
End Sub
