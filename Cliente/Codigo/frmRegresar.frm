VERSION 5.00
Begin VB.Form frmRegresar 
   Caption         =   "Regreso"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Has Muerto!"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command5 
         Caption         =   "Continuar como fantasma"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Arghal"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Banderbill"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Nix"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ciudad Principal (Ullathorpe)"
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Has muerto!, si lo deseas, Elije a la ciudad que quieres ir,"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmRegresar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call SendData("/AULLA")
Unload Me
End Sub

Private Sub Command2_Click()
Call SendData("/ANIXX")
Unload Me
End Sub

Private Sub Command3_Click()
Call SendData("/ABNDR")
Unload Me
End Sub

Private Sub Command4_Click()
Call SendData("/AARGL")
Unload Me
End Sub

Private Sub Command5_Click()
Unload Me

End Sub


