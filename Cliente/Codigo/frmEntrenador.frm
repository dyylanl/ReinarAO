VERSION 5.00
Begin VB.Form frmEntrenador 
   BorderStyle     =   0  'None
   Caption         =   "Entrenar"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4470
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      ItemData        =   "frmEntrenador.frx":0000
      Left            =   360
      List            =   "frmEntrenador.frx":0002
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Image command2 
      Height          =   375
      Left            =   4200
      MouseIcon       =   "frmEntrenador.frx":0004
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   495
      Left            =   1080
      MouseIcon       =   "frmEntrenador.frx":030E
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   2295
   End
End
Attribute VB_Name = "frmEntrenador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call SendData("ENTR" & lstCriaturas.ListIndex + 1)
    Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving = False And Button = vbLeftButton Then
        Dx3 = X
        dy = Y
        bmoving = True
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving And ((X <> Dx3) Or (Y <> dy)) Then Move Left + (X - Dx3), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\RECURSOS\Interfaces\entrenar.gif")

End Sub

