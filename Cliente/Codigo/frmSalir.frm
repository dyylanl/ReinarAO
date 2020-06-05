VERSION 5.00
Begin VB.Form frmSalir 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   Icon            =   "frmSalir.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton No 
      Caption         =   "No"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Si 
      Caption         =   "Si"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "¿Estas seguro que quieres salir de Lhirius AO?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "frmSalir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Aceptar_Click()

Call SendData("#B")
Unload Me
Unload frmMain

End Sub

Private Sub Cancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Redondear_Formulario(Me, 80)
End Sub

Private Sub No_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub

Private Sub Si_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("#B")
Me.Hide

End Sub
Public Sub Redondear_Formulario(El_Form As Form, Radio As Long)
 
 Dim Region As Long
 Dim ret As Long
 Dim Ancho As Long
 Dim Alto As Long
 Dim old_Scale As Integer
 
     old_Scale = El_Form.ScaleMode
     El_Form.ScaleMode = vbPixels
     Ancho = El_Form.ScaleWidth
     Alto = El_Form.ScaleHeight
     Region = CreateRoundRectRgn(0, 0, Ancho, Alto, Radio, Radio)
     ret = SetWindowRgn(El_Form.hWnd, Region, True)
     El_Form.ScaleMode = old_Scale
 
 End Sub
